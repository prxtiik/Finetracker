// ===========================================================
// FineTracker — app.js
// All app logic: room auth, transactions, dashboards, cycles
// ===========================================================

// ---------- Reference data (mirrors the original Excel sheet) ----------
const EXPENSE_CATS = {
  "Grocery": ["Blinkit/Zepto", "Offline"],
  "Food": ["Zomato/Swiggy", "Offline"],
  "Bills": ["Electricity", "Mobile"],
  "Rent": ["Apartment"],
  "Shopping": ["Clothes", "Jewellry", "Accessories", "Cosmetics", "House/Décor", "Art", "Miscellaneous"],
  "Drinks": ["Alcohol", "Tea/juice"],
  "Vehicle": ["Fuel", "Repair"],
  "Family": ["Help", "Gift"],
  "Debt": ["Debt"]
};
const INCOME_TYPES = ["Salary", "Gift", "Miscellaneous"];
const SAVINGS_TYPES = {
  "Liquid": ["Cash", "Savings", "FD"],
  "Non - Liquid": ["Chit fund", "Investment", "Gold"]
};
const LOAN_CATEGORIES = ["EMI", "Personal Loan", "Home Loan", "Vehicle Loan", "Education Loan", "Credit Card", "Borrowing (Family/Friends)", "Other"];

const TX_ICON = { income: "↑", expense: "↓", savings: "◆", borrow: "▲", repay: "▼" };

// ---------- Local state ----------
let fb = null; // set once firebase-ready fires
let state = {
  screen: "boot",       // boot | auth | app
  authError: "",
  room: null,            // { code, members: [{name}], cycles: {...} }
  me: null,              // my name within the room
  currentCycleId: null,
  view: "combined",      // 'me' | 'partner' | 'combined'
  nav: "home",           // home | history | analytics | room
  sheetOpen: null,       // 'add' | 'newloan' | null
  addType: "expense",    // expense | income | savings | loan
  loanAction: "repay",   // 'borrow' | 'repay' — which action within the loan sheet
  selectedLoanId: null,  // which loan the current borrow/repay entry targets
  newLoanReturnTo: null, // where to go after creating a loan: 'add' or null (Room tab)
  toast: null,
  unsub: null
};

const $ = (sel) => document.querySelector(sel);
const app = () => document.getElementById("app");

function showToast(msg) {
  state.toast = msg;
  render();
  setTimeout(() => { state.toast = null; }, 2000);
}

function fmtMoney(n) {
  n = Number(n) || 0;
  return n.toLocaleString("en-IN", { maximumFractionDigits: 0 });
}

function genCycleId() {
  return "cycle_" + Date.now();
}

// ---------- Persistence (localStorage for "who am I" only — never for room data) ----------
function saveSession(roomCode, name) {
  localStorage.setItem("ft_room", roomCode);
  localStorage.setItem("ft_name", name);
}
function clearSession() {
  localStorage.removeItem("ft_room");
  localStorage.removeItem("ft_name");
}
function getSession() {
  return {
    room: localStorage.getItem("ft_room"),
    name: localStorage.getItem("ft_name")
  };
}

// ===========================================================
// FIREBASE-BACKED ROOM OPERATIONS
// ===========================================================

async function createRoom(code, name) {
  code = code.trim().toUpperCase();
  name = name.trim();
  if (!code || !name) return { ok: false, err: "Enter a room code and your name." };

  const ref = fb.doc(fb.db, "rooms", code);
  const snap = await fb.getDoc(ref);
  if (snap.exists()) {
    return { ok: false, err: "That room code is already taken. Try another one." };
  }

  const cycleId = genCycleId();
  const roomData = {
    code,
    members: [{ name }],
    createdAt: Date.now(),
    currentCycleId: cycleId,
    loans: {},
    cycles: {
      [cycleId]: {
        startedAt: Date.now(),
        startedBy: name,
        transactions: []
      }
    }
  };
  await fb.setDoc(ref, roomData);
  saveSession(code, name);
  return { ok: true, room: roomData, name };
}

async function joinRoom(code, name) {
  code = code.trim().toUpperCase();
  name = name.trim();
  if (!code || !name) return { ok: false, err: "Enter a room code and your name." };

  const ref = fb.doc(fb.db, "rooms", code);
  const snap = await fb.getDoc(ref);
  if (!snap.exists()) {
    return { ok: false, err: "No room found with that code. Check it and try again." };
  }
  const data = snap.data();
  const existing = data.members.find(m => m.name.toLowerCase() === name.toLowerCase());

  if (existing) {
    saveSession(code, existing.name);
    return { ok: true, room: data, name: existing.name };
  }

  if (data.members.length >= 2) {
    return { ok: false, err: "This room already has two people in it." };
  }

  const updatedMembers = [...data.members, { name }];
  await fb.updateDoc(ref, { members: updatedMembers });
  saveSession(code, name);
  return { ok: true, room: { ...data, members: updatedMembers }, name };
}

async function relogin(code, name) {
  code = code.trim().toUpperCase();
  name = name.trim();
  const ref = fb.doc(fb.db, "rooms", code);
  const snap = await fb.getDoc(ref);
  if (!snap.exists()) return { ok: false, err: "Room not found." };
  const data = snap.data();
  const existing = data.members.find(m => m.name.toLowerCase() === name.toLowerCase());
  if (!existing) {
    return { ok: false, err: "That name doesn't match anyone in this room." };
  }
  saveSession(code, existing.name);
  return { ok: true, room: data, name: existing.name };
}

function listenToRoom(code) {
  if (state.unsub) state.unsub();
  const ref = fb.doc(fb.db, "rooms", code);
  state.unsub = fb.onSnapshot(ref, (snap) => {
    if (snap.exists()) {
      state.room = snap.data();
      if (!state.currentCycleId) state.currentCycleId = state.room.currentCycleId;
      // Don't yank the add-entry sheet out from under someone mid-type when
      // the partner's update arrives in the background — just let it apply
      // silently and re-render once they close the sheet.
      if (state.sheetOpen !== 'add') render();
    }
  });
}

async function addTransaction(tx) {
  const ref = fb.doc(fb.db, "rooms", state.room.code);
  const cycleId = state.room.currentCycleId;
  const path = `cycles.${cycleId}.transactions`;
  const txWithMeta = { ...tx, id: "tx_" + Date.now() + "_" + Math.random().toString(36).slice(2, 7), by: state.me, ts: Date.now() };
  await fb.updateDoc(ref, { [path]: fb.arrayUnion(txWithMeta) });
  showToast("Entry added");
}

async function deleteTransaction(txId, cycleId) {
  // Firestore arrayUnion has no "remove by predicate" — only arrayRemove with an
  // exact-match object. Safer + simpler: read the doc, filter the array in JS,
  // and write the whole array back. Rooms are tiny (two people, one cycle's
  // worth of entries) so this is cheap.
  const ref = fb.doc(fb.db, "rooms", state.room.code);
  const snap = await fb.getDoc(ref);
  if (!snap.exists()) return;
  const data = snap.data();
  const cycle = data.cycles[cycleId];
  if (!cycle) return;
  const updatedTxs = (cycle.transactions || []).filter(t => t.id !== txId);
  await fb.updateDoc(ref, { [`cycles.${cycleId}.transactions`]: updatedTxs });
  showToast("Entry deleted");
}

async function startNewCycle() {
  const ref = fb.doc(fb.db, "rooms", state.room.code);
  const newId = genCycleId();
  await fb.updateDoc(ref, {
    currentCycleId: newId,
    [`cycles.${newId}`]: {
      startedAt: Date.now(),
      startedBy: state.me,
      transactions: []
    }
  });
  state.currentCycleId = newId;
  showToast("New cycle started — fresh balance!");
}

// ===========================================================
// LOANS — named, persistent debt trackers
// Borrowing more increases outstanding only (no effect on balance).
// Repaying decreases outstanding AND counts as an expense (real money out).
// Outstanding lives across cycles — a loan doesn't reset when a new
// monthly cycle starts, since the debt itself doesn't disappear.
// ===========================================================

function genLoanId() {
  return "loan_" + Date.now() + "_" + Math.random().toString(36).slice(2, 7);
}

async function createLoan(name, category, openingOutstanding) {
  name = (name || "").trim();
  if (!name) return { ok: false, err: "Give the loan a name." };
  const ref = fb.doc(fb.db, "rooms", state.room.code);
  const loanId = genLoanId();
  const opening = parseFloat(openingOutstanding) || 0;
  const loan = {
    id: loanId, name, category: category || "Other",
    createdBy: state.me, createdAt: Date.now(),
    openingOutstanding: opening // what was already owed before tracking started — not a transaction, just a starting point
  };
  await fb.updateDoc(ref, { [`loans.${loanId}`]: loan });
  showToast("Loan added");
  return { ok: true, loanId };
}

function listLoans() {
  if (!state.room || !state.room.loans) return [];
  return Object.values(state.room.loans).sort((a, b) => b.createdAt - a.createdAt);
}

// Outstanding for one loan = opening balance + total borrowed − total repaid,
// across ALL cycles. The opening balance is a starting fact (what you
// already owed before using the app), not a transaction — it never touches
// available balance, it just sets where this loan's outstanding starts from.
function computeLoanOutstanding(loanId) {
  const all = getAllTxAcrossCycles(null);
  const loan = state.room.loans && state.room.loans[loanId];
  const opening = (loan && loan.openingOutstanding) || 0;
  let borrowed = 0, repaid = 0;
  for (const t of all) {
    if (t.loanId !== loanId) continue;
    if (t.kind === 'borrow') borrowed += t.amount;
    else if (t.kind === 'repay') repaid += t.amount;
  }
  return { borrowed, repaid, outstanding: opening + borrowed - repaid };
}

// Combined outstanding across every loan in the room (including each loan's
// opening balance) — this is the number shown as "Debt Outstanding" on the Ledger.
//
// Opening balances belong to the loan itself, not to either person specifically
// (you didn't track who originally took it on day one) — so they only appear
// in the Combined view. The "me"/"partner" views show just that person's own
// borrow/repay activity, so the same opening amount never shows up fully under
// both tabs (which would make it look like double the actual debt).
function computeTotalOutstanding(filterBy) {
  if (!state.room) return 0;
  let total = 0;
  if (!filterBy) { // combined view only
    const loans = listLoans();
    for (const loan of loans) total += loan.openingOutstanding || 0;
  }
  const all = getAllTxAcrossCycles(filterBy);
  for (const t of all) {
    if (t.kind === 'borrow') total += t.amount;
    else if (t.kind === 'repay') total -= t.amount;
  }
  return total;
}

// ===========================================================
// EXPORT (CSV — opens directly in Excel / Sheets / Numbers)
// ===========================================================

function listAllCycles() {
  if (!state.room || !state.room.cycles) return [];
  return Object.entries(state.room.cycles)
    .map(([id, c]) => ({ id, ...c }))
    .sort((a, b) => b.startedAt - a.startedAt);
}

function csvEscape(val) {
  if (val === null || val === undefined) return "";
  const s = String(val);
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function buildCycleCsv(cycle) {
  const txs = (cycle.transactions || []).slice().sort((a, b) => a.ts - b.ts);
  const rows = [];

  // ---- Summary block first, same spirit as her original Dashboard sheet ----
  const members = state.room.members.map(m => m.name);
  rows.push(["FineTracker export"]);
  rows.push(["Room", state.room.code]);
  rows.push(["Cycle started", new Date(cycle.startedAt).toLocaleString("en-IN")]);
  rows.push(["Exported on", new Date().toLocaleString("en-IN")]);
  rows.push([]);

  rows.push(["Summary", "Income", "Expenses (incl. debt repayment)", "Savings", "Borrowed", "Repaid", "Available balance"]);
  const allSummary = summarizeTxs(txs, null);
  rows.push(["Combined", allSummary.income, allSummary.expense, allSummary.savings, allSummary.borrowed, allSummary.repaid, allSummary.available]);
  for (const name of members) {
    const s = summarizeTxs(txs, name);
    rows.push([name, s.income, s.expense, s.savings, s.borrowed, s.repaid, s.available]);
  }
  rows.push([]);

  // ---- Full transaction log ----
  rows.push(["Date", "Time", "By", "Type", "Category", "Subcategory", "Loan", "Amount", "Note"]);
  for (const t of txs) {
    const d = new Date(t.ts);
    const loanName = t.loanId && state.room.loans && state.room.loans[t.loanId] ? state.room.loans[t.loanId].name : "";
    rows.push([
      d.toLocaleDateString("en-IN"),
      d.toLocaleTimeString("en-IN"),
      t.by,
      cap(t.kind),
      t.kind === 'repay' ? 'Debt Repayment' : (t.category || ""),
      t.subcategory || "",
      loanName,
      t.amount,
      t.note || ""
    ]);
  }

  return rows.map(row => row.map(csvEscape).join(",")).join("\r\n");
}

function summarizeTxs(txs, byName) {
  const list = byName ? txs.filter(t => t.by === byName) : txs;
  let income = 0, expense = 0, savings = 0, borrowed = 0, repaid = 0;
  for (const t of list) {
    if (t.kind === "income") income += t.amount;
    else if (t.kind === "expense") expense += t.amount;
    else if (t.kind === "savings") savings += t.amount;
    else if (t.kind === "repay") { expense += t.amount; repaid += t.amount; }
    else if (t.kind === "borrow") borrowed += t.amount;
  }
  return { income, expense, savings, borrowed, repaid, available: income - expense - savings };
}

function downloadCycleCsv(cycle, cycleId) {
  const csv = buildCycleCsv(cycle);
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const dateLabel = new Date(cycle.startedAt).toLocaleDateString("en-IN", { month: "short", year: "numeric" }).replace(" ", "_");
  const a = document.createElement("a");
  a.href = url;
  a.download = `FineTracker_${state.room.code}_${dateLabel}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  showToast("Exported — check your Downloads");
}

// ===========================================================
// DERIVED DATA / CALCULATIONS
// ===========================================================

function getCurrentCycle() {
  if (!state.room) return null;
  const cid = state.room.currentCycleId;
  return state.room.cycles[cid] || { transactions: [] };
}

function getTx(filterBy) {
  const cid = state.room.currentCycleId;
  const cycle = getCurrentCycle();
  if (!cycle) return [];
  let txs = (cycle.transactions || []).map(t => ({ ...t, cycleId: cid }));
  if (filterBy === "me") txs = txs.filter(t => t.by === state.me);
  if (filterBy === "partner") txs = txs.filter(t => t.by !== state.me);
  return txs.sort((a, b) => b.ts - a.ts);
}

function getPartnerName() {
  if (!state.room) return null;
  const other = state.room.members.find(m => m.name !== state.me);
  return other ? other.name : null;
}

function computeSummary(filterBy) {
  const txs = getTx(filterBy);
  let income = 0, expense = 0, savings = 0, repaid = 0;
  for (const t of txs) {
    if (t.kind === "income") income += t.amount;
    else if (t.kind === "expense") expense += t.amount;
    else if (t.kind === "savings") savings += t.amount;
    else if (t.kind === "repay") { expense += t.amount; repaid += t.amount; }
    // 'borrow' deliberately does NOT touch income/expense/balance —
    // taking a loan isn't income you earned, it's money you'll owe back.
  }

  // Opening balances are starting facts declared once via "Set opening balance",
  // not transactions — they never get subtracted again. Opening cash-on-hand
  // adds straight to available; opening savings adds straight to the savings
  // total, without ever touching available balance (declaring savings you
  // already had shouldn't make it look like you just spent that money).
  const opening = getOpeningBalance(filterBy);

  const available = opening.available + income - expense - savings;
  const totalSavings = opening.savings + savings;
  const debt = computeTotalOutstanding(filterBy);
  return { income, expense, savings: totalSavings, debt, repaid, available };
}

// Sums opening balances for whoever filterBy resolves to. Combined view sums
// both members' opening declarations; "me"/"partner" show just that person's.
function getOpeningBalance(filterBy) {
  const obs = (state.room && state.room.openingBalances) || {};
  let names;
  if (filterBy === 'me') names = [state.me];
  else if (filterBy === 'partner') { const p = getPartnerName(); names = p ? [p] : []; }
  else names = state.room ? state.room.members.map(m => m.name) : [];

  let available = 0, savings = 0;
  for (const n of names) {
    const o = obs[n];
    if (o) { available += o.available || 0; savings += o.savings || 0; }
  }
  return { available, savings };
}

async function setOpeningBalance(available, savings) {
  const ref = fb.doc(fb.db, "rooms", state.room.code);
  await fb.updateDoc(ref, {
    [`openingBalances.${state.me}`]: {
      available: parseFloat(available) || 0,
      savings: parseFloat(savings) || 0,
      setAt: Date.now()
    }
  });
  showToast("Opening balance set");
}

function computeCategoryBreakdown(filterBy, kind) {
  let txs;
  if (kind === 'expense') {
    // Repayments are real money out, so they belong in the expense
    // breakdown too — labeled distinctly as "Debt Repayment".
    txs = getTx(filterBy).filter(t => t.kind === 'expense' || t.kind === 'repay');
  } else {
    txs = getTx(filterBy).filter(t => t.kind === kind);
  }
  const map = {};
  for (const t of txs) {
    const key = t.kind === 'repay' ? 'Debt Repayment' : (t.category || "Other");
    map[key] = (map[key] || 0) + t.amount;
  }
  return Object.entries(map).sort((a, b) => b[1] - a[1]);
}

// ---------- Analytics: flatten transactions across ALL cycles ----------
function getAllTxAcrossCycles(filterBy) {
  if (!state.room || !state.room.cycles) return [];
  let all = [];
  for (const [cid, cycle] of Object.entries(state.room.cycles)) {
    const txs = (cycle.transactions || []).map(t => ({ ...t, cycleId: cid }));
    all = all.concat(txs);
  }
  if (filterBy === "me") all = all.filter(t => t.by === state.me);
  if (filterBy === "partner") all = all.filter(t => t.by !== state.me);
  return all;
}

function monthKey(ts) {
  const d = new Date(ts);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}
function monthLabel(key) {
  const [y, m] = key.split('-').map(Number);
  return new Date(y, m - 1, 1).toLocaleDateString('en-IN', { month: 'short', year: '2-digit' });
}

// Returns last N months (chronological order) each with income/expense/savings/borrowed/repaid totals
function computeMonthlySeries(filterBy, months = 6) {
  const txs = getAllTxAcrossCycles(filterBy);
  const byMonth = {};
  for (const t of txs) {
    const key = monthKey(t.ts);
    if (!byMonth[key]) byMonth[key] = { income: 0, expense: 0, savings: 0, borrowed: 0, repaid: 0 };
    if (t.kind === 'income') byMonth[key].income += t.amount;
    else if (t.kind === 'expense') byMonth[key].expense += t.amount;
    else if (t.kind === 'savings') byMonth[key].savings += t.amount;
    else if (t.kind === 'repay') { byMonth[key].expense += t.amount; byMonth[key].repaid += t.amount; }
    else if (t.kind === 'borrow') byMonth[key].borrowed += t.amount;
  }

  // Build a continuous range of the last N months ending this month, even if
  // some months have zero entries — keeps the chart's x-axis honest.
  const now = new Date();
  const series = [];
  for (let i = months - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    const vals = byMonth[key] || { income: 0, expense: 0, savings: 0, borrowed: 0, repaid: 0 };
    series.push({ key, label: monthLabel(key), ...vals });
  }
  return series;
}

function computeCategoryBreakdownAllTime(filterBy, kind) {
  let txs;
  if (kind === 'expense') {
    txs = getAllTxAcrossCycles(filterBy).filter(t => t.kind === 'expense' || t.kind === 'repay');
  } else {
    txs = getAllTxAcrossCycles(filterBy).filter(t => t.kind === kind);
  }
  const map = {};
  for (const t of txs) {
    const key = t.kind === 'repay' ? 'Debt Repayment' : (t.category || "Other");
    map[key] = (map[key] || 0) + t.amount;
  }
  return Object.entries(map).sort((a, b) => b[1] - a[1]);
}

// Month-wise totals broken down BY CATEGORY — this is what powers the
// per-category mini-charts in Analytics, so you can see how "Grocery" or
// "Salary" trended over the last few months, not just its all-time total.
function computeCategoryMonthlySeries(filterBy, kind, months = 6) {
  let txs;
  if (kind === 'expense') {
    txs = getAllTxAcrossCycles(filterBy).filter(t => t.kind === 'expense' || t.kind === 'repay');
  } else {
    txs = getAllTxAcrossCycles(filterBy).filter(t => t.kind === kind);
  }

  // Build the continuous month range first (so every category series lines
  // up against the same x-axis, including zero months).
  const now = new Date();
  const monthKeys = [];
  for (let i = months - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    monthKeys.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  }

  // category -> { monthKey -> amount }
  const byCategory = {};
  for (const t of txs) {
    const cat = t.kind === 'repay' ? 'Debt Repayment' : (t.category || 'Other');
    const mKey = monthKey(t.ts);
    if (!byCategory[cat]) byCategory[cat] = {};
    byCategory[cat][mKey] = (byCategory[cat][mKey] || 0) + t.amount;
  }

  // Order categories by all-time total, descending — same order as the list view
  const ordered = Object.entries(byCategory)
    .map(([cat, monthMap]) => {
      const total = Object.values(monthMap).reduce((a, b) => a + b, 0);
      const series = monthKeys.map(k => ({ key: k, label: monthLabel(k), amount: monthMap[k] || 0 }));
      return { category: cat, total, series };
    })
    .sort((a, b) => b.total - a.total);

  return ordered;
}

// Loan-category-wise breakdown (EMI, Personal Loan, Credit Card, etc.) for
// the Analytics > Loans section — separate from the expense/income/savings
// metrics since borrowing/repaying isn't "spending" in the same sense.
function computeLoanCategoryBreakdown(filterBy, action) {
  // action: 'borrow' | 'repay'
  const txs = getAllTxAcrossCycles(filterBy).filter(t => t.kind === action);
  const map = {};
  for (const t of txs) {
    const loan = state.room.loans && state.room.loans[t.loanId];
    const key = loan ? loan.category : "Other";
    map[key] = (map[key] || 0) + t.amount;
  }
  return Object.entries(map).sort((a, b) => b[1] - a[1]);
}

// Same idea as computeCategoryMonthlySeries, but for loan borrow/repay
// activity grouped by loan category (EMI, Credit Card, etc.) over time.
function computeLoanCategoryMonthlySeries(filterBy, action, months = 6) {
  const txs = getAllTxAcrossCycles(filterBy).filter(t => t.kind === action);

  const now = new Date();
  const monthKeys = [];
  for (let i = months - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    monthKeys.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  }

  const byCategory = {};
  for (const t of txs) {
    const loan = state.room.loans && state.room.loans[t.loanId];
    const cat = loan ? loan.category : 'Other';
    const mKey = monthKey(t.ts);
    if (!byCategory[cat]) byCategory[cat] = {};
    byCategory[cat][mKey] = (byCategory[cat][mKey] || 0) + t.amount;
  }

  return Object.entries(byCategory)
    .map(([cat, monthMap]) => {
      const total = Object.values(monthMap).reduce((a, b) => a + b, 0);
      const series = monthKeys.map(k => ({ key: k, label: monthLabel(k), amount: monthMap[k] || 0 }));
      return { category: cat, total, series };
    })
    .sort((a, b) => b.total - a.total);
}

// ===========================================================
// RENDER FUNCTIONS
// ===========================================================

function render() {
  const root = app();
  if (state.screen === "boot") {
    root.innerHTML = `<div class="boot-screen"><div class="boot-mark">FT</div></div>`;
    return;
  }
  if (state.screen === "auth") {
    root.innerHTML = renderAuth();
    bindAuthEvents();
    return;
  }
  if (state.screen === "app") {
    root.innerHTML = renderApp();
    bindAppEvents();
    return;
  }
}

// ---------- AUTH SCREEN ----------
let authMode = "create"; // create | join

function renderAuth() {
  return `
  <div class="auth-wrap">
    <div class="brand-block">
      <div class="brand-mark">FineTracker</div>
      <div class="brand-tag">A shared ledger for two. Track separately, see it together.</div>
    </div>

    <div class="tab-row">
      <div class="tab-btn ${authMode === 'create' ? 'active' : ''}" data-authtab="create">Create room</div>
      <div class="tab-btn ${authMode === 'join' ? 'active' : ''}" data-authtab="join">Join / Log in</div>
    </div>

    ${state.authError ? `<div class="error-msg">${state.authError}</div>` : ''}

    ${authMode === 'create' ? `
      <div class="field">
        <label>Room code (you make this up)</label>
        <input id="inCode" type="text" placeholder="e.g. PRATIK26" maxlength="16" autocapitalize="characters" />
      </div>
      <div class="field">
        <label>Your name</label>
        <input id="inName" type="text" placeholder="e.g. Pratik" maxlength="20" />
      </div>
      <div class="hint-msg">You'll log in later using this exact code + name. There's no recovery if you forget it, so keep it somewhere safe.</div>
      <button class="btn btn-primary" id="btnCreate">Create room</button>
    ` : `
      <div class="field">
        <label>Room code</label>
        <input id="inCode2" type="text" placeholder="Enter the shared code" maxlength="16" autocapitalize="characters" />
      </div>
      <div class="field">
        <label>Your name</label>
        <input id="inName2" type="text" placeholder="Use the exact name on this room" maxlength="20" />
      </div>
      <div class="hint-msg">If you're joining for the first time, pick any name — that becomes your login name forever. If you've joined before, use the same name exactly.</div>
      <button class="btn btn-primary" id="btnJoin">Join / Log in</button>
    `}
  </div>`;
}

function bindAuthEvents() {
  document.querySelectorAll('[data-authtab]').forEach(el => {
    el.onclick = () => { authMode = el.dataset.authtab; state.authError = ""; render(); };
  });

  const btnCreate = $("#btnCreate");
  if (btnCreate) {
    btnCreate.onclick = async () => {
      const code = $("#inCode").value;
      const name = $("#inName").value;
      btnCreate.disabled = true; btnCreate.textContent = "Creating...";
      const res = await createRoom(code, name);
      if (!res.ok) {
        state.authError = res.err; render(); return;
      }
      enterApp(res.room, res.name);
    };
  }

  const btnJoin = $("#btnJoin");
  if (btnJoin) {
    btnJoin.onclick = async () => {
      const code = $("#inCode2").value;
      const name = $("#inName2").value;
      btnJoin.disabled = true; btnJoin.textContent = "Checking...";
      const res = await joinRoom(code, name);
      if (!res.ok) {
        state.authError = res.err; render(); return;
      }
      enterApp(res.room, res.name);
    };
  }
}

function enterApp(room, name) {
  state.room = room;
  state.me = name;
  state.currentCycleId = room.currentCycleId;
  state.screen = "app";
  state.authError = "";
  listenToRoom(room.code);
  render();
}

// ---------- MAIN APP SCREEN ----------
function renderApp() {
  if (!state.room) return `<div class="center-msg"><div class="big-emoji">⏳</div><p>Loading your room…</p></div>`;

  const partnerName = getPartnerName();
  const waitingForPartner = state.room.members.length < 2;

  let body = "";
  if (state.nav === "home") body = renderHome(partnerName, waitingForPartner);
  else if (state.nav === "history") body = renderHistory();
  else if (state.nav === "analytics") body = renderAnalytics(partnerName, waitingForPartner);
  else if (state.nav === "room") body = renderRoomTab(partnerName, waitingForPartner);

  return `
    <div class="screen">
      ${body}
    </div>
    <button class="fab" id="fabAdd">+</button>
    ${renderBottomNav()}
    ${state.sheetOpen === 'add' ? renderAddSheet() : ''}
    ${state.sheetOpen === 'newloan' ? renderNewLoanSheet() : ''}
    ${state.sheetOpen === 'opening' ? renderOpeningBalanceSheet() : ''}
    ${state.toast ? `<div class="toast">${state.toast}</div>` : ''}
  `;
}

function renderOpeningBalanceSheet() {
  const existing = (state.room.openingBalances && state.room.openingBalances[state.me]) || { available: '', savings: '' };
  return `
  <div class="sheet-overlay" id="openingOverlay">
    <div class="sheet">
      <div class="sheet-handle"></div>
      <h2 class="sheet-title">My opening balance</h2>
      <p class="page-sub" style="margin-bottom:20px;">This is a one-time starting point — money you already had before you started using FineTracker. It won't get subtracted from your balance the way a regular Savings entry would.</p>
      <div class="field">
        <label>Cash on hand right now</label>
        <div class="amount-input-wrap"><span class="rupee">₹</span><input id="fOpenAvailable" type="number" inputmode="decimal" placeholder="0" value="${existing.available || ''}" /></div>
        <div class="hint-msg" style="margin-top:8px;margin-bottom:0;">Spendable money you currently have — gets added straight to your available balance.</div>
      </div>
      <div class="field">
        <label>Existing savings</label>
        <div class="amount-input-wrap"><span class="rupee">₹</span><input id="fOpenSavings" type="number" inputmode="decimal" placeholder="0" value="${existing.savings || ''}" /></div>
        <div class="hint-msg" style="margin-top:8px;margin-bottom:0;">Money already in FD, gold, cash savings etc. — counts in your savings total, but never reduces your available balance.</div>
      </div>
      <div class="btn-row">
        <button class="btn btn-ghost" id="btnCancelOpening">Cancel</button>
        <button class="btn btn-primary" id="btnSaveOpening">Save</button>
      </div>
    </div>
  </div>`;
}

function renderNewLoanSheet() {
  return `
  <div class="sheet-overlay" id="newLoanOverlay">
    <div class="sheet">
      <div class="sheet-handle"></div>
      <h2 class="sheet-title">New loan</h2>
      <div class="field">
        <label>Name</label>
        <input id="fLoanName" type="text" placeholder="e.g. HDFC Credit Card" maxlength="30" />
      </div>
      <div class="field">
        <label>Category</label>
        <div class="chip-grid" id="loanCatChips">
          ${LOAN_CATEGORIES.map(c => `<div class="chip" data-loancat="${c}">${c}</div>`).join('')}
        </div>
      </div>
      <div class="field">
        <label>Already outstanding? (optional)</label>
        <div class="amount-input-wrap"><span class="rupee">₹</span><input id="fLoanOpening" type="number" inputmode="decimal" placeholder="0" /></div>
        <div class="hint-msg" style="margin-top:8px;margin-bottom:0;">If you're mid-way through this loan already — e.g. an EMI you're partway done with — put what's currently owed here. Leave at 0 if this is a brand new loan.</div>
      </div>
      <div class="btn-row">
        <button class="btn btn-ghost" id="btnCancelNewLoan">Cancel</button>
        <button class="btn btn-primary" id="btnSaveNewLoan">Create loan</button>
      </div>
    </div>
  </div>`;
}

function renderBottomNav() {
  const tabs = [
    { id: "home", icon: "𝓛", label: "Ledger" },
    { id: "history", icon: "≡", label: "History" },
    { id: "analytics", icon: "▲", label: "Analytics" },
    { id: "room", icon: "⚭", label: "Room" }
  ];
  return `<div class="bottom-nav">
    ${tabs.map(t => `
      <button class="nav-btn ${state.nav === t.id ? 'active' : ''}" data-nav="${t.id}">
        <span class="nav-icon">${t.icon}</span>${t.label}
      </button>
    `).join('')}
  </div>`;
}

function renderHome(partnerName, waiting) {
  const filterBy = state.view === 'combined' ? null : state.view === 'me' ? 'me' : 'partner';
  const summary = computeSummary(filterBy);
  const recentTx = getTx(filterBy).slice(0, 6);

  return `
    <div class="eyebrow">Room ${state.room.code}</div>
    <h1 class="page-title">Hey, ${state.me}</h1>
    <p class="page-sub">${waiting ? "Waiting for your partner to join with the room code." : `Sharing this ledger with ${partnerName}.`}</p>

    <div class="view-switch">
      <button data-view="me" class="${state.view === 'me' ? 'active' : ''}">${state.me}</button>
      <button data-view="partner" class="${state.view === 'partner' ? 'active' : ''}" ${waiting ? 'disabled style="opacity:.4"' : ''}>${partnerName || 'Partner'}</button>
      <button data-view="combined" class="${state.view === 'combined' ? 'active' : ''}">Combined</button>
    </div>

    <div class="ledger-band">
      <div class="ledger-room">Current cycle · <b>${state.view === 'combined' ? 'Both' : state.view === 'me' ? state.me : partnerName}</b></div>
      <div class="ledger-balance-label">Available balance</div>
      <div class="ledger-balance"><span class="rupee">₹</span>${fmtMoney(summary.available)}</div>
      <div class="ledger-meta-row">
        <div class="ledger-meta"><div class="lbl">Income</div><div class="val pos">₹${fmtMoney(summary.income)}</div></div>
        <div class="ledger-meta"><div class="lbl">Expenses</div><div class="val neg">₹${fmtMoney(summary.expense)}</div></div>
        <div class="ledger-meta"><div class="lbl">Savings</div><div class="val">₹${fmtMoney(summary.savings)}</div></div>
      </div>
    </div>

    <div class="mini-grid">
      <div class="mini-card"><div class="lbl">Total Debt Outstanding</div><div class="val">₹${fmtMoney(summary.debt)}</div></div>
      <div class="mini-card"><div class="lbl">Entries this cycle</div><div class="val">${getTx(filterBy).length}</div></div>
    </div>

    ${renderCategoryBreakdown()}

    <div class="section-head">
      <h3>Recent entries</h3>
      <span class="link-sm" data-nav="history">See all</span>
    </div>
    ${recentTx.length ? recentTx.map(renderTxItem).join('') : `<div class="tx-empty">No entries yet in this cycle. Tap + to add one.</div>`}
  `;
}

function renderCategoryBreakdown() {
  const filterBy = state.view === 'combined' ? null : state.view === 'me' ? 'me' : 'partner';
  const cats = computeCategoryBreakdown(filterBy, 'expense');
  if (!cats.length) return '';
  const max = cats[0][1];
  return `
    <div class="section-head"><h3>Expense breakdown</h3></div>
    ${cats.slice(0, 6).map(([name, amt]) => `
      <div class="cat-row">
        <div class="cat-row-top"><span class="name">${name}</span><span class="amt">₹${fmtMoney(amt)}</span></div>
        <div class="bar-track"><div class="bar-fill" style="width:${(amt / max) * 100}%"></div></div>
      </div>
    `).join('')}
  `;
}

function renderTxItem(t) {
  const sign = t.kind === 'income' ? '+' : (t.kind === 'savings' || t.kind === 'borrow') ? '' : '−';
  const cls = t.kind === 'income' ? 'pos' : (t.kind === 'expense' || t.kind === 'repay') ? 'neg' : '';
  const dateStr = new Date(t.ts).toLocaleDateString('en-IN', { day: 'numeric', month: 'short' });

  let title, sub;
  if (t.kind === 'borrow' || t.kind === 'repay') {
    const loan = state.room.loans && state.room.loans[t.loanId];
    title = (loan ? loan.name : 'Loan') + (t.kind === 'borrow' ? ' · Borrowed' : ' · Repaid');
    sub = `${t.by} · ${dateStr}${t.note ? ' · ' + t.note : ''}`;
  } else {
    title = `${t.category || cap(t.kind)}${t.subcategory ? ' · ' + t.subcategory : ''}`;
    sub = `${t.by} · ${dateStr}${t.note ? ' · ' + t.note : ''}`;
  }

  return `
    <div class="tx-item" data-txid="${t.id}" data-cycleid="${t.cycleId}">
      <div class="tx-icon ${t.kind}">${TX_ICON[t.kind]}</div>
      <div class="tx-body">
        <div class="tx-title">${title}</div>
        <div class="tx-sub">${sub}</div>
      </div>
      <div class="tx-amt ${cls}">${sign}₹${fmtMoney(t.amount)}</div>
      <button class="tx-delete" data-deltx="${t.id}" data-delcycle="${t.cycleId}" aria-label="Delete entry">✕</button>
    </div>
  `;
}

function cap(s) { return s ? s[0].toUpperCase() + s.slice(1) : s; }

function renderHistory() {
  const filterBy = state.view === 'combined' ? null : state.view === 'me' ? 'me' : 'partner';
  const txs = getTx(filterBy);
  return `
    <div class="eyebrow">All entries</div>
    <h1 class="page-title">History</h1>
    <p class="page-sub">Every entry logged in the current cycle.</p>
    <div class="view-switch">
      <button data-view="me" class="${state.view === 'me' ? 'active' : ''}">${state.me}</button>
      <button data-view="partner" class="${state.view === 'partner' ? 'active' : ''}">${getPartnerName() || 'Partner'}</button>
      <button data-view="combined" class="${state.view === 'combined' ? 'active' : ''}">Combined</button>
    </div>
    ${txs.length ? txs.map(renderTxItem).join('') : `<div class="tx-empty">No entries yet.</div>`}
  `;
}

// ---------- ANALYTICS TAB ----------
let analyticsMetric = "expense"; // expense | income | savings | loans

function renderAnalytics(partnerName, waiting) {
  const filterBy = state.view === 'combined' ? null : state.view === 'me' ? 'me' : 'partner';
  const metricColor = { income: 'var(--sage)', expense: 'var(--terracotta)', savings: 'var(--marigold)' };
  const metricLabel = { income: 'Income', expense: 'Expenses', savings: 'Savings', loans: 'Loans' };

  const viewSwitch = `
    <div class="view-switch">
      <button data-view="me" class="${state.view === 'me' ? 'active' : ''}">${state.me}</button>
      <button data-view="partner" class="${state.view === 'partner' ? 'active' : ''}" ${waiting ? 'disabled style="opacity:.4"' : ''}>${partnerName || 'Partner'}</button>
      <button data-view="combined" class="${state.view === 'combined' ? 'active' : ''}">Combined</button>
    </div>`;

  const metricTabs = `
    <div class="type-tab-row">
      ${['expense', 'income', 'savings', 'loans'].map(m => `
        <div class="type-tab ${analyticsMetric === m ? 'active' : ''}" data-metric="${m}">${metricLabel[m]}</div>
      `).join('')}
    </div>`;

  if (analyticsMetric === 'loans') {
    return `
      <div class="eyebrow">Analytics</div>
      <h1 class="page-title">Loans over time</h1>
      <p class="page-sub">Borrowed vs repaid, separate from your income/expense balance.</p>
      ${viewSwitch}
      ${metricTabs}
      ${renderLoanAnalytics(filterBy)}
    `;
  }

  const series = computeMonthlySeries(filterBy, 6);
  const max = Math.max(1, ...series.map(s => s[analyticsMetric]));
  const totalForMetric = series.reduce((sum, s) => sum + s[analyticsMetric], 0);
  const avg = series.length ? totalForMetric / series.length : 0;

  return `
    <div class="eyebrow">Analytics</div>
    <h1 class="page-title">Spending over time</h1>
    <p class="page-sub">Month-wise trend across every cycle, not just the current one.</p>
    ${viewSwitch}
    ${metricTabs}

    <div class="mini-grid">
      <div class="mini-card"><div class="lbl">Last 6 months total</div><div class="val">₹${fmtMoney(totalForMetric)}</div></div>
      <div class="mini-card"><div class="lbl">Monthly average</div><div class="val">₹${fmtMoney(avg)}</div></div>
    </div>

    <div class="chart-card">
      <div class="bar-chart">
        ${series.map(s => {
          const h = Math.round((s[analyticsMetric] / max) * 100);
          return `
            <div class="bar-col">
              <div class="bar-col-amt">${s[analyticsMetric] > 0 ? '₹' + fmtMoneyShort(s[analyticsMetric]) : ''}</div>
              <div class="bar-col-track">
                <div class="bar-col-fill" style="height:${h}%; background:${metricColor[analyticsMetric]};"></div>
              </div>
              <div class="bar-col-label">${s.label}</div>
            </div>
          `;
        }).join('')}
      </div>
    </div>

    <div class="section-head"><h3>${metricLabel[analyticsMetric]}${analyticsMetric === 'expense' ? ' (incl. debt repayment)' : ''} by category — last 6 months</h3></div>
    ${renderCategoryCharts(filterBy, analyticsMetric, metricColor[analyticsMetric])}
  `;
}

function renderCategoryCharts(filterBy, kind, color) {
  const catSeries = computeCategoryMonthlySeries(filterBy, kind, 6);
  if (!catSeries.length) {
    const label = kind === 'expense' ? 'expense' : kind;
    return `<div class="tx-empty">No ${label} entries yet.</div>`;
  }
  return catSeries.map(({ category, total, series }) => {
    const max = Math.max(1, ...series.map(s => s.amount));
    return `
      <div class="mini-chart-card">
        <div class="mini-chart-head">
          <span class="mini-chart-name">${category}</span>
          <span class="mini-chart-total">₹${fmtMoney(total)}</span>
        </div>
        <div class="mini-bar-row">
          ${series.map(s => `
            <div class="mini-bar-col" title="${s.label}: ₹${fmtMoney(s.amount)}">
              <div class="mini-bar-track">
                <div class="mini-bar-fill" style="height:${Math.round((s.amount / max) * 100)}%; background:${color};"></div>
              </div>
              <div class="mini-bar-label">${s.label}</div>
            </div>
          `).join('')}
        </div>
      </div>
    `;
  }).join('');
}

function renderLoanAnalytics(filterBy) {
  const series = computeMonthlySeries(filterBy, 6);
  const max = Math.max(1, ...series.map(s => Math.max(s.borrowed, s.repaid)));
  const totalBorrowed = series.reduce((sum, s) => sum + s.borrowed, 0);
  const totalRepaid = series.reduce((sum, s) => sum + s.repaid, 0);
  const outstanding = computeTotalOutstanding(filterBy);

  return `
    <div class="mini-grid">
      <div class="mini-card"><div class="lbl">Outstanding now</div><div class="val">₹${fmtMoney(outstanding)}</div></div>
      <div class="mini-card"><div class="lbl">Repaid (6 months)</div><div class="val">₹${fmtMoney(totalRepaid)}</div></div>
    </div>

    <div class="chart-card">
      <div class="bar-chart">
        ${series.map(s => {
          const hBorrow = Math.round((s.borrowed / max) * 100);
          const hRepay = Math.round((s.repaid / max) * 100);
          return `
            <div class="bar-col">
              <div class="bar-col-amt">${s.borrowed > 0 ? '₹' + fmtMoneyShort(s.borrowed) : ''}</div>
              <div class="bar-col-track dual">
                <div class="bar-col-fill" style="height:${hBorrow}%; background:#9098AE;"></div>
                <div class="bar-col-fill" style="height:${hRepay}%; background:var(--sage);"></div>
              </div>
              <div class="bar-col-label">${s.label}</div>
            </div>
          `;
        }).join('')}
      </div>
      <div class="legend-row">
        <span class="legend-dot" style="background:#9098AE;"></span> Borrowed
        <span class="legend-dot" style="background:var(--sage); margin-left:16px;"></span> Repaid
      </div>
    </div>

    <div class="section-head"><h3>Borrowed by category — last 6 months</h3></div>
    ${renderLoanCategoryCharts(filterBy, 'borrow', '#9098AE')}

    <div class="section-head"><h3>Repaid by category — last 6 months</h3></div>
    ${renderLoanCategoryCharts(filterBy, 'repay', 'var(--sage)')}
  `;
}

function renderLoanCategoryCharts(filterBy, action, color) {
  const catSeries = computeLoanCategoryMonthlySeries(filterBy, action, 6);
  if (!catSeries.length) {
    return `<div class="tx-empty">No ${action === 'borrow' ? 'borrowing' : 'repayments'} logged yet.</div>`;
  }
  return catSeries.map(({ category, total, series }) => {
    const max = Math.max(1, ...series.map(s => s.amount));
    return `
      <div class="mini-chart-card">
        <div class="mini-chart-head">
          <span class="mini-chart-name">${category}</span>
          <span class="mini-chart-total">₹${fmtMoney(total)}</span>
        </div>
        <div class="mini-bar-row">
          ${series.map(s => `
            <div class="mini-bar-col" title="${s.label}: ₹${fmtMoney(s.amount)}">
              <div class="mini-bar-track">
                <div class="mini-bar-fill" style="height:${Math.round((s.amount / max) * 100)}%; background:${color};"></div>
              </div>
              <div class="mini-bar-label">${s.label}</div>
            </div>
          `).join('')}
        </div>
      </div>
    `;
  }).join('');
}

function fmtMoneyShort(n) {
  if (n >= 100000) return (n / 100000).toFixed(1).replace(/\.0$/, '') + 'L';
  if (n >= 1000) return (n / 1000).toFixed(1).replace(/\.0$/, '') + 'k';
  return fmtMoney(n);
}

function renderRoomTab(partnerName, waiting) {
  return `
    <div class="eyebrow">Room settings</div>
    <h1 class="page-title">${state.room.code}</h1>
    <p class="page-sub">Share this code with your partner so they can join.</p>

    <div class="code-box">
      <span class="code">${state.room.code}</span>
      <button id="btnCopyCode">Copy</button>
    </div>

    <div class="section-head"><h3>People in this room</h3></div>
    ${state.room.members.map(m => `
      <div class="profile-row">
        <div class="avatar ${m.name === state.me ? '' : 'dim'}">${m.name[0].toUpperCase()}</div>
        <div>
          <div class="tx-title">${m.name}${m.name === state.me ? ' (you)' : ''}</div>
          <div class="tx-sub">${m.name === state.me ? 'This device' : 'Joined the room'}</div>
        </div>
      </div>
    `).join('')}
    ${waiting ? `
      <div class="profile-row">
        <div class="avatar dim">?</div>
        <div>
          <div class="tx-title">Waiting...</div>
          <div class="tx-sub"><span class="waiting-pulse"></span>Not joined yet</div>
        </div>
      </div>
    ` : ''}

    <div class="divider-label">Opening balance</div>
    <p class="page-sub" style="margin-bottom:14px;">For money you already had before using this app — cash on hand and existing savings. Set once, it won't get subtracted again like a regular entry.</p>
    ${state.room.members.map(m => {
      const ob = (state.room.openingBalances && state.room.openingBalances[m.name]) || null;
      return `
        <div class="profile-row">
          <div class="avatar ${m.name === state.me ? '' : 'dim'}">${m.name[0].toUpperCase()}</div>
          <div style="flex:1;">
            <div class="tx-title">${m.name}${m.name === state.me ? ' (you)' : ''}</div>
            <div class="tx-sub">${ob ? `₹${fmtMoney(ob.available)} cash · ₹${fmtMoney(ob.savings)} savings` : 'Not set yet'}</div>
          </div>
        </div>
      `;
    }).join('')}
    <button class="btn btn-ghost" id="btnSetOpening" style="margin-top:8px;">${(state.room.openingBalances && state.room.openingBalances[state.me]) ? 'Update my opening balance' : 'Set my opening balance'}</button>

    <div class="divider-label">Loans</div>
    <p class="page-sub" style="margin-bottom:14px;">Each loan tracks its own outstanding balance. Borrowing more doesn't touch your available balance — repaying does, since that's real money out.</p>
    ${listLoans().map(loan => {
      const o = computeLoanOutstanding(loan.id);
      return `
        <div class="profile-row">
          <div class="avatar dim">${loan.name[0].toUpperCase()}</div>
          <div style="flex:1;">
            <div class="tx-title">${loan.name}</div>
            <div class="tx-sub">${loan.category}</div>
          </div>
          <div class="val" style="font-size:16px;">₹${fmtMoney(o.outstanding)}</div>
        </div>
      `;
    }).join('') || `<div class="tx-empty">No loans added yet.</div>`}
    <button class="btn btn-ghost" id="btnAddLoan" style="margin-top:8px;">+ Add a loan</button>

    <div class="divider-label">Cycle</div>
    <p class="page-sub" style="margin-bottom:14px;">Starting a new cycle resets the available balance to zero going forward, but keeps all past history saved.</p>
    <button class="btn btn-ghost" id="btnNewCycle">Start new cycle</button>

    <div class="divider-label">Export</div>
    <p class="page-sub" style="margin-bottom:14px;">Download a cycle as a spreadsheet (.csv) — opens directly in Excel, Google Sheets, or Numbers. Includes both your entries and your partner's, plus the combined totals.</p>
    ${listAllCycles().map((c, i) => {
      const label = new Date(c.startedAt).toLocaleDateString("en-IN", { month: "long", year: "numeric" });
      const count = (c.transactions || []).length;
      const isCurrent = c.id === state.room.currentCycleId;
      return `
        <div class="profile-row">
          <div class="avatar dim">${i + 1}</div>
          <div style="flex:1;">
            <div class="tx-title">${label}${isCurrent ? ' · current' : ''}</div>
            <div class="tx-sub">${count} ${count === 1 ? 'entry' : 'entries'}</div>
          </div>
          <button class="btn btn-sm btn-ghost" data-exportcycle="${c.id}" ${count === 0 ? 'disabled style="opacity:.4"' : ''}>Export</button>
        </div>
      `;
    }).join('')}

    <div class="divider-label">Account</div>
    <button class="btn btn-danger-ghost" id="btnLogout">Log out of this device</button>
  `;
}

// ---------- ADD ENTRY SHEET ----------
function renderAddSheet() {
  const type = state.addType;
  return `
  <div class="sheet-overlay" id="sheetOverlay">
    <div class="sheet" id="sheetBody">
      <div class="sheet-handle"></div>
      <h2 class="sheet-title">Add an entry</h2>
      <div class="type-tab-row">
        ${["expense", "income", "savings", "loan"].map(t => `
          <div class="type-tab ${type === t ? 'active' : ''}" data-addtype="${t}">${t === 'loan' ? 'Loan' : cap(t)}</div>
        `).join('')}
      </div>
      <div id="addFormHost">${renderAddForm(type)}</div>
      <div class="btn-row">
        <button class="btn btn-ghost" id="btnCancelAdd">Cancel</button>
        <button class="btn btn-primary" id="btnSaveAdd">Save entry</button>
      </div>
    </div>
  </div>`;
}

function renderAddForm(type) {
  if (type === "expense") {
    const cats = Object.keys(EXPENSE_CATS);
    return `
      <div class="field">
        <label>Amount</label>
        <div class="amount-input-wrap"><span class="rupee">₹</span><input id="fAmount" type="number" inputmode="decimal" placeholder="0" /></div>
      </div>
      <div class="field">
        <label>Category</label>
        <div class="chip-grid" id="catChips">
          ${cats.map(c => `<div class="chip" data-cat="${c}">${c}</div>`).join('')}
        </div>
      </div>
      <div class="field" id="subcatField" style="display:none;">
        <label>Subtype</label>
        <div class="chip-grid" id="subcatChips"></div>
      </div>
      <div class="field"><label>Note (optional)</label><input id="fNote" type="text" placeholder="e.g. auto fare, chai" /></div>
    `;
  }
  if (type === "income") {
    return `
      <div class="field">
        <label>Amount</label>
        <div class="amount-input-wrap"><span class="rupee">₹</span><input id="fAmount" type="number" inputmode="decimal" placeholder="0" /></div>
      </div>
      <div class="field">
        <label>Type</label>
        <div class="chip-grid" id="catChips">
          ${INCOME_TYPES.map(c => `<div class="chip" data-cat="${c}">${c}</div>`).join('')}
        </div>
      </div>
      <div class="field"><label>Note (optional)</label><input id="fNote" type="text" placeholder="e.g. June salary" /></div>
    `;
  }
  if (type === "savings") {
    const cats = Object.keys(SAVINGS_TYPES);
    return `
      <div class="field">
        <label>Amount</label>
        <div class="amount-input-wrap"><span class="rupee">₹</span><input id="fAmount" type="number" inputmode="decimal" placeholder="0" /></div>
      </div>
      <div class="field">
        <label>Type</label>
        <div class="chip-grid" id="catChips">
          ${cats.map(c => `<div class="chip" data-cat="${c}">${c}</div>`).join('')}
        </div>
      </div>
      <div class="field" id="subcatField" style="display:none;">
        <label>Subtype</label>
        <div class="chip-grid" id="subcatChips"></div>
      </div>
      <div class="field"><label>Note (optional)</label><input id="fNote" type="text" placeholder="e.g. RD this month" /></div>
    `;
  }
  if (type === "loan") {
    const loans = listLoans();
    return `
      <div class="field">
        <label>Action</label>
        <div class="chip-grid" id="loanActionChips">
          <div class="chip ${state.loanAction === 'borrow' ? 'active' : ''}" data-loanaction="borrow">Borrowed more</div>
          <div class="chip ${state.loanAction === 'repay' ? 'active' : ''}" data-loanaction="repay">Repaid</div>
        </div>
      </div>
      <div class="field">
        <label>Amount</label>
        <div class="amount-input-wrap"><span class="rupee">₹</span><input id="fAmount" type="number" inputmode="decimal" placeholder="0" /></div>
      </div>
      <div class="field">
        <label>Which loan?</label>
        <div class="chip-grid" id="loanPickChips">
          ${loans.map(l => `<div class="chip ${state.selectedLoanId === l.id ? 'active' : ''}" data-loanpick="${l.id}">${l.name}</div>`).join('')}
          <div class="chip" id="chipNewLoanInline">+ New loan</div>
        </div>
        ${!loans.length ? `<div class="hint-msg" style="margin-top:8px;">No loans yet — tap "+ New loan" to create one first.</div>` : ''}
      </div>
      <div class="field"><label>Note (optional)</label><input id="fNote" type="text" placeholder="e.g. June EMI" /></div>
    `;
  }
  return "";
}

let selectedCat = null;
let selectedSubcat = null;

function bindAddFormEvents(type) {
  selectedCat = null; selectedSubcat = null;
  document.querySelectorAll('#catChips .chip').forEach(chip => {
    chip.onclick = () => {
      document.querySelectorAll('#catChips .chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      selectedCat = chip.dataset.cat;
      selectedSubcat = null;

      const subField = $("#subcatField");
      if (subField) {
        let subs = [];
        if (type === 'expense') subs = EXPENSE_CATS[selectedCat] || [];
        if (type === 'savings') subs = SAVINGS_TYPES[selectedCat] || [];
        if (subs.length) {
          subField.style.display = 'block';
          $("#subcatChips").innerHTML = subs.map(s => `<div class="chip" data-subcat="${s}">${s}</div>`).join('');
          document.querySelectorAll('#subcatChips .chip').forEach(sc => {
            sc.onclick = () => {
              document.querySelectorAll('#subcatChips .chip').forEach(x => x.classList.remove('active'));
              sc.classList.add('active');
              selectedSubcat = sc.dataset.subcat;
            };
          });
        } else {
          subField.style.display = 'none';
        }
      }
    };
  });

  // ---- Loan-specific bindings ----
  document.querySelectorAll('#loanActionChips .chip').forEach(chip => {
    chip.onclick = () => {
      state.loanAction = chip.dataset.loanaction;
      document.querySelectorAll('#loanActionChips .chip').forEach(c => c.classList.toggle('active', c.dataset.loanaction === state.loanAction));
    };
  });
  document.querySelectorAll('#loanPickChips .chip[data-loanpick]').forEach(chip => {
    chip.onclick = () => {
      state.selectedLoanId = chip.dataset.loanpick;
      document.querySelectorAll('#loanPickChips .chip[data-loanpick]').forEach(c => c.classList.toggle('active', c.dataset.loanpick === state.selectedLoanId));
    };
  });
  const chipNewLoan = $("#chipNewLoanInline");
  if (chipNewLoan) chipNewLoan.onclick = () => {
    state.newLoanReturnTo = 'add';
    state.sheetOpen = 'newloan';
    render();
    setTimeout(bindNewLoanEvents, 0);
  };
}

// ===========================================================
// EVENT BINDING
// ===========================================================

let selectedLoanCat = null;

function bindNewLoanEvents() {
  selectedLoanCat = null;
  document.querySelectorAll('#loanCatChips .chip').forEach(chip => {
    chip.onclick = () => {
      document.querySelectorAll('#loanCatChips .chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      selectedLoanCat = chip.dataset.loancat;
    };
  });

  const overlay = $("#newLoanOverlay");
  if (overlay) overlay.onclick = (e) => { if (e.target.id === 'newLoanOverlay') closeNewLoanSheet(); };

  const btnCancel = $("#btnCancelNewLoan");
  if (btnCancel) btnCancel.onclick = closeNewLoanSheet;

  const btnSave = $("#btnSaveNewLoan");
  if (btnSave) {
    btnSave.onclick = async () => {
      const name = $("#fLoanName").value;
      if (!name.trim()) { showToast("Give the loan a name"); return; }
      const opening = $("#fLoanOpening") ? $("#fLoanOpening").value : 0;
      btnSave.disabled = true; btnSave.textContent = "Creating...";
      const res = await createLoan(name, selectedLoanCat, opening);
      if (res.ok) {
        state.selectedLoanId = res.loanId;
      }
      closeNewLoanSheet();
    };
  }
}

function closeNewLoanSheet() {
  selectedLoanCat = null;
  if (state.newLoanReturnTo === 'add') {
    state.sheetOpen = 'add'; // return to the add-entry sheet, now with the new loan selected
    render();
    setTimeout(() => bindAddFormEvents(state.addType), 0);
  } else {
    state.sheetOpen = null; // opened from the Room tab — just close back to it
    render();
  }
}

function bindOpeningBalanceEvents() {
  const overlay = $("#openingOverlay");
  if (overlay) overlay.onclick = (e) => { if (e.target.id === 'openingOverlay') closeOpeningSheet(); };

  const btnCancel = $("#btnCancelOpening");
  if (btnCancel) btnCancel.onclick = closeOpeningSheet;

  const btnSave = $("#btnSaveOpening");
  if (btnSave) {
    btnSave.onclick = async () => {
      const available = $("#fOpenAvailable").value;
      const savings = $("#fOpenSavings").value;
      btnSave.disabled = true; btnSave.textContent = "Saving...";
      await setOpeningBalance(available, savings);
      closeOpeningSheet();
    };
  }
}

function closeOpeningSheet() {
  state.sheetOpen = null;
  render();
}

function bindAppEvents() {
  const fab = $("#fabAdd");
  if (fab) fab.onclick = () => { state.sheetOpen = 'add'; render(); setTimeout(() => bindAddFormEvents(state.addType), 0); };

  document.querySelectorAll('[data-nav]').forEach(el => {
    el.onclick = () => { state.nav = el.dataset.nav; render(); };
  });

  document.querySelectorAll('[data-view]').forEach(el => {
    el.onclick = () => {
      if (el.disabled) return;
      state.view = el.dataset.view; render();
    };
  });

  const overlay = $("#sheetOverlay");
  if (overlay) {
    overlay.onclick = (e) => { if (e.target.id === 'sheetOverlay') closeSheet(); };
  }
  const btnCancel = $("#btnCancelAdd");
  if (btnCancel) btnCancel.onclick = closeSheet;

  document.querySelectorAll('[data-addtype]').forEach(el => {
    el.onclick = () => {
      state.addType = el.dataset.addtype;
      if (state.addType !== 'loan') state.selectedLoanId = null;
      $("#addFormHost").innerHTML = renderAddForm(state.addType);
      document.querySelectorAll('[data-addtype]').forEach(t => t.classList.toggle('active', t.dataset.addtype === state.addType));
      bindAddFormEvents(state.addType);
    };
  });
  bindAddFormEvents(state.addType);

  const btnSave = $("#btnSaveAdd");
  if (btnSave) {
    btnSave.onclick = async () => {
      const amount = parseFloat($("#fAmount").value);
      if (!amount || amount <= 0) { showToast("Enter a valid amount"); return; }
      const note = $("#fNote") ? $("#fNote").value.trim() : "";

      if (state.addType === 'loan') {
        if (!state.selectedLoanId) { showToast("Pick a loan, or create one"); return; }
        btnSave.disabled = true; btnSave.textContent = "Saving...";
        await addTransaction({
          kind: state.loanAction, // 'borrow' or 'repay'
          amount,
          loanId: state.selectedLoanId,
          note: note || null
        });
        closeSheet();
        return;
      }

      if (!selectedCat) { showToast("Pick a category"); return; }
      btnSave.disabled = true; btnSave.textContent = "Saving...";
      await addTransaction({
        kind: state.addType,
        amount,
        category: selectedCat,
        subcategory: selectedSubcat || null,
        note: note || null
      });
      closeSheet();
    };
  }

  document.querySelectorAll('[data-metric]').forEach(el => {
    el.onclick = () => { analyticsMetric = el.dataset.metric; render(); };
  });

  document.querySelectorAll('[data-deltx]').forEach(el => {
    el.onclick = async (e) => {
      e.stopPropagation();
      const txId = el.dataset.deltx;
      const cycleId = el.dataset.delcycle;
      if (confirm("Delete this entry? This can't be undone.")) {
        await deleteTransaction(txId, cycleId);
      }
    };
  });

  const btnCopy = $("#btnCopyCode");
  if (btnCopy) btnCopy.onclick = () => {
    navigator.clipboard.writeText(state.room.code).then(() => showToast("Code copied"));
  };

  document.querySelectorAll('[data-exportcycle]').forEach(el => {
    el.onclick = () => {
      if (el.disabled) return;
      const cycleId = el.dataset.exportcycle;
      const cycle = state.room.cycles[cycleId];
      if (cycle) downloadCycleCsv(cycle, cycleId);
    };
  });

  const btnNewCycle = $("#btnNewCycle");
  if (btnNewCycle) btnNewCycle.onclick = async () => {
    if (confirm("Start a new cycle? This resets the available balance, but all history stays saved.")) {
      await startNewCycle();
    }
  };

  const btnAddLoan = $("#btnAddLoan");
  if (btnAddLoan) btnAddLoan.onclick = () => {
    state.newLoanReturnTo = null;
    state.sheetOpen = 'newloan';
    render();
    setTimeout(bindNewLoanEvents, 0);
  };

  const btnSetOpening = $("#btnSetOpening");
  if (btnSetOpening) btnSetOpening.onclick = () => {
    state.sheetOpen = 'opening';
    render();
    setTimeout(bindOpeningBalanceEvents, 0);
  };

  const btnLogout = $("#btnLogout");
  if (btnLogout) btnLogout.onclick = () => {
    if (confirm("Log out of this device? You'll need your room code and exact name to log back in.")) {
      clearSession();
      if (state.unsub) state.unsub();
      state.room = null; state.me = null; state.screen = 'auth';
      render();
    }
  };
}

function closeSheet() {
  state.sheetOpen = null;
  selectedCat = null; selectedSubcat = null;
  state.selectedLoanId = null; state.loanAction = 'repay';
  render();
}

// ===========================================================
// BOOTSTRAP
// ===========================================================

function init() {
  fb = window.__fb;
  const session = getSession();
  if (session.room && session.name) {
    relogin(session.room, session.name).then(res => {
      if (res.ok) {
        enterApp(res.room, res.name);
      } else {
        clearSession();
        state.screen = "auth";
        render();
      }
    });
  } else {
    state.screen = "auth";
    render();
  }
}

if (window.__fb) {
  init();
} else {
  window.addEventListener("firebase-ready", init, { once: true });
}
