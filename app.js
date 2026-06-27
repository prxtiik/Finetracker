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
    txs = getAl
