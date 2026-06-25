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
const DEBT_TYPES = ["Home Loan", "Personal Loan", "Vehicle Loan", "Education Loan", "Credit Card", "Borrowed (Family/Friends)", "Other"];

const TX_ICON = { income: "↑", expense: "↓", savings: "◆", debt: "●" };

// ---------- Local state ----------
let fb = null; // set once firebase-ready fires
let state = {
  screen: "boot",       // boot | auth | app
  authError: "",
  room: null,            // { code, members: [{name}], cycles: {...} }
  me: null,              // my name within the room
  currentCycleId: null,
  view: "combined",      // 'me' | 'partner' | 'combined'
  nav: "home",           // home | history | room
  sheetOpen: null,       // 'add' | null
  addType: "expense",    // expense | income | savings | debt
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
// DERIVED DATA / CALCULATIONS
// ===========================================================

function getCurrentCycle() {
  if (!state.room) return null;
  const cid = state.room.currentCycleId;
  return state.room.cycles[cid] || { transactions: [] };
}

function getTx(filterBy) {
  const cycle = getCurrentCycle();
  if (!cycle) return [];
  let txs = cycle.transactions || [];
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
  let income = 0, expense = 0, savings = 0, debt = 0;
  for (const t of txs) {
    if (t.kind === "income") income += t.amount;
    else if (t.kind === "expense") expense += t.amount;
    else if (t.kind === "savings") savings += t.amount;
    else if (t.kind === "debt") debt += t.amount;
  }
  const available = income - expense - savings;
  return { income, expense, savings, debt, available };
}

function computeCategoryBreakdown(filterBy, kind) {
  const txs = getTx(filterBy).filter(t => t.kind === kind);
  const map = {};
  for (const t of txs) {
    const key = t.category || "Other";
    map[key] = (map[key] || 0) + t.amount;
  }
  return Object.entries(map).sort((a, b) => b[1] - a[1]);
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
  else if (state.nav === "room") body = renderRoomTab(partnerName, waitingForPartner);

  return `
    <div class="screen">
      ${body}
    </div>
    <button class="fab" id="fabAdd">+</button>
    ${renderBottomNav()}
    ${state.sheetOpen === 'add' ? renderAddSheet() : ''}
    ${state.toast ? `<div class="toast">${state.toast}</div>` : ''}
  `;
}

function renderBottomNav() {
  const tabs = [
    { id: "home", icon: "𝓛", label: "Ledger" },
    { id: "history", icon: "≡", label: "History" },
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
  const sign = t.kind === 'income' ? '+' : (t.kind === 'savings' || t.kind === 'debt') ? '' : '−';
  const cls = t.kind === 'income' ? 'pos' : t.kind === 'expense' ? 'neg' : '';
  const dateStr = new Date(t.ts).toLocaleDateString('en-IN', { day: 'numeric', month: 'short' });
  return `
    <div class="tx-item">
      <div class="tx-icon ${t.kind}">${TX_ICON[t.kind]}</div>
      <div class="tx-body">
        <div class="tx-title">${t.category || cap(t.kind)}${t.subcategory ? ' · ' + t.subcategory : ''}</div>
        <div class="tx-sub">${t.by} · ${dateStr}${t.note ? ' · ' + t.note : ''}</div>
      </div>
      <div class="tx-amt ${cls}">${sign}₹${fmtMoney(t.amount)}</div>
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

    <div class="divider-label">Cycle</div>
    <p class="page-sub" style="margin-bottom:14px;">Starting a new cycle resets the available balance to zero going forward, but keeps all past history saved.</p>
    <button class="btn btn-ghost" id="btnNewCycle">Start new cycle</button>

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
        ${["expense", "income", "savings", "debt"].map(t => `
          <div class="type-tab ${type === t ? 'active' : ''}" data-addtype="${t}">${cap(t)}</div>
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
  if (type === "debt") {
    return `
      <div class="field">
        <label>Outstanding amount</label>
        <div class="amount-input-wrap"><span class="rupee">₹</span><input id="fAmount" type="number" inputmode="decimal" placeholder="0" /></div>
      </div>
      <div class="field">
        <label>Type</label>
        <div class="chip-grid" id="catChips">
          ${DEBT_TYPES.map(c => `<div class="chip" data-cat="${c}">${c}</div>`).join('')}
        </div>
      </div>
      <div class="field"><label>Lender (optional)</label><input id="fNote" type="text" placeholder="e.g. Slice, HDFC" /></div>
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
}

// ===========================================================
// EVENT BINDING
// ===========================================================

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
      if (!selectedCat) { showToast("Pick a category"); return; }
      const note = $("#fNote") ? $("#fNote").value.trim() : "";
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

  const btnCopy = $("#btnCopyCode");
  if (btnCopy) btnCopy.onclick = () => {
    navigator.clipboard.writeText(state.room.code).then(() => showToast("Code copied"));
  };

  const btnNewCycle = $("#btnNewCycle");
  if (btnNewCycle) btnNewCycle.onclick = async () => {
    if (confirm("Start a new cycle? This resets the available balance, but all history stays saved.")) {
      await startNewCycle();
    }
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
