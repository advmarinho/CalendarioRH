/* ============================================================
   Calendário Operacional RH Sonova (Single Page)
   Ajustes desta versão:
   - Sticky columns sem "left fixo": offsets calculados via JS
   - Botão para sair do Modo Impressão (A3) (barra flutuante)
   - Topbar/Filtros sticky no scroll da página
   - Zebra na tabela (CSS) sem quebrar sticky backgrounds
   - Modelo financeiro:
       * provisioned (provisionado) e executed (executado) por mês
       * compatibilidade com bases antigas (value)
       * compatibilidade com HTML antigo (cellValue)
   ============================================================ */

(() => {
  "use strict";

  const APP_VERSION = "1.0.2";
  const STORAGE_KEY = "sonova_calendario_operacional_v1";
  const BASE_ONLINE_NAME_KEY = "sonova_calendario_base_online_name_v1";

  const MONTHS = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];

  const $ = (id) => document.getElementById(id);

  const state = {
    year: new Date().getFullYear(),
    month: new Date().getMonth(),
    activities: [],
    archives: [],
    filters: {
      category: "",
      owner: "",
      supplier: "",
      period: "",
      text: "",
      onlyPending: false,
      onlyApplicable: true,
    },
    ui: {
      printMode: false,
      editingActivityId: null,
      cellContext: null,
      followContext: null,
      prevImportBuffer: null,
    }
  };

  function nowISO() { return new Date().toISOString(); }
  function pad2(n) { return String(n).padStart(2, "0"); }

  function formatDateShort(iso) {
    if (!iso) return "";
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return "";
    return `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}`;
  }

  function formatDateLong(iso) {
    if (!iso) return "";
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return "";
    return `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}/${d.getFullYear()} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
  }

  function parseBRNumber(text) {
    if (text == null) return null;
    const raw = String(text).trim();
    if (!raw) return null;
    const normalized = raw.replace(/\s/g, "").replace(/\./g, "").replace(",", ".");
    const val = Number(normalized);
    if (!Number.isFinite(val)) return null;
    return val;
  }

  function formatBRL(value) {
    const v = Number(value || 0);
    return v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  }

// === Linhas de Provisão/Pagamento + Ordenação por Prazo ===
function normalizeEntryLines(entry) {
  if (!entry) return;

  if (!Array.isArray(entry.provisions)) entry.provisions = [];
  if (!Array.isArray(entry.payments)) entry.payments = [];

  // Backward compatibility: older fields -> seed arrays once
  const prov = entryProvisioned(entry);
  const exec = Number(entry.executed);

  if (entry.provisions.length === 0 && Number.isFinite(prov) && prov > 0) {
    entry.provisions.push({ amount: prov, note: "" });
  }

  if (entry.payments.length === 0 && Number.isFinite(exec) && exec > 0) {
    entry.payments.push({ amount: exec, note: "" });
  }
}

function ensureCellLinesWrap() {
  const modal = $("modalCell");
  if (!modal) return null;

  let wrap = $("cellLinesWrap");
  if (wrap) return wrap;

  // Place right after the executed field if it exists, otherwise at end of form-grid
  const grid = modal.querySelector(".form-grid");
  if (!grid) return null;

  wrap = document.createElement("div");
  wrap.id = "cellLinesWrap";
  wrap.className = "form-row span2";
  wrap.style.marginTop = "4px";

  grid.appendChild(wrap);
  return wrap;
}

function escapeHtml(s) {
  const str = safeText(s);
  return str.replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
}

function formatBRNumberInput(n) {
  if (!Number.isFinite(n)) return "";
  // keep 2 decimals when needed
  const s = (Math.round(n * 100) / 100).toString();
  return s.replace(".", ",");
}

function renderLinesSection(title, lines, kind) {
  const rows = Array.isArray(lines) ? lines : [];
  const idList = (kind === "prov") ? "cellProvLines" : "cellPayLines";
  const idTotal = (kind === "prov") ? "cellProvTotal" : "cellPayTotal";
  const idBtn = (kind === "prov") ? "btnAddProvLine" : "btnAddPayLine";

  const htmlRows = rows.map((ln, idx) => {
    const v = Number(ln?.amount);
    const note = safeText(ln?.note);
    return `
      <div class="line-row" data-kind="${kind}" data-idx="${idx}">
        <input class="line-amount" type="text" value="${formatBRNumberInput(v)}" placeholder="0,00" />
        <input class="line-note" type="text" value="${escapeHtml(note)}" placeholder="Descrição (opcional)" />
        <button class="btn btn-outline btn-sm line-del" type="button">Remover</button>
      </div>`;
  }).join("");

  return `
    <div class="lines-card" data-kind="${kind}">
      <div class="lines-head">
        <div class="lines-title">${escapeHtml(title)}</div>
        <div class="lines-actions">
          <button class="btn btn-outline btn-sm" id="${idBtn}" type="button">Adicionar linha</button>
        </div>
      </div>

      <div class="lines-table" id="${idList}">
        <div class="line-row line-row-head">
          <div class="line-h">Valor (R$)</div>
          <div class="line-h">Descrição</div>
          <div class="line-h"></div>
        </div>
        ${htmlRows || `<div class="muted" style="font-size:12px;padding:6px 2px;">Sem linhas lançadas.</div>`}
      </div>

      <div class="lines-foot">
        <div class="muted" style="font-size:12px;">Total</div>
        <div class="lines-total" id="${idTotal}">R$ 0</div>
      </div>
    </div>
  `;
}

function renderCellLinesUI(entry) {
  normalizeEntryLines(entry);

  // Hide the single inputs (keeps HTML compatibility)
  const elProv = $("cellProvisioned");
  const elExec = $("cellExecuted");
  if (elProv && elProv.closest(".form-row")) elProv.closest(".form-row").style.display = "none";
  if (elExec && elExec.closest(".form-row")) elExec.closest(".form-row").style.display = "none";

  const wrap = ensureCellLinesWrap();
  if (!wrap) return;

  // Minimal inline styles for layout, without touching styles.css
  wrap.innerHTML = `
    <div style="display:grid;gap:10px;">
      ${renderLinesSection("Provisões (planejado)", entry.provisions, "prov")}
      ${renderLinesSection("Pagamentos (executado)", entry.payments, "pay")}
      <div class="hint" style="margin-top:-4px;">
        Dica: use múltiplas linhas quando houver complementos, ajustes, parcelas ou pagamentos separados.
        O cartão do mês usa o total das provisões e dos pagamentos.
      </div>
    </div>
  `;

  // Wire buttons + delegation
  const btnProv = $("btnAddProvLine");
  const btnPay = $("btnAddPayLine");

  if (btnProv) {
    btnProv.onclick = () => {
      entry.provisions.push({ amount: null, note: "" });
      renderCellLinesUI(entry);
      updateCellLinesTotals(entry);
    };
  }
  if (btnPay) {
    btnPay.onclick = () => {
      entry.payments.push({ amount: null, note: "" });
      renderCellLinesUI(entry);
      updateCellLinesTotals(entry);
    };
  }

  wrap.onclick = (ev) => {
    const t = ev.target;
    if (!(t instanceof HTMLElement)) return;
    if (!t.classList.contains("line-del")) return;

    const row = t.closest(".line-row");
    if (!row) return;

    const kind = row.getAttribute("data-kind");
    const idx = Number(row.getAttribute("data-idx"));
    if (!Number.isFinite(idx) || idx < 0) return;

    if (kind === "prov") {
      entry.provisions.splice(idx, 1);
    } else if (kind === "pay") {
      entry.payments.splice(idx, 1);
    }
    renderCellLinesUI(entry);
    updateCellLinesTotals(entry);
  };

  wrap.oninput = () => {
    // live update totals (without saving)
    syncEntryLinesFromUI(entry);
    updateCellLinesTotals(entry);
  };

  updateCellLinesTotals(entry);
}

function syncEntryLinesFromUI(entry) {
  const wrap = $("cellLinesWrap");
  if (!wrap) return;

  const provRows = wrap.querySelectorAll('.line-row[data-kind="prov"]');
  const payRows = wrap.querySelectorAll('.line-row[data-kind="pay"]');

  const prov = [];
  provRows.forEach(r => {
    const amountEl = r.querySelector(".line-amount");
    const noteEl = r.querySelector(".line-note");
    const amount = parseBRNumber(amountEl ? amountEl.value : "");
    const note = safeText(noteEl ? noteEl.value : "").trim();
    if (amount != null && amount >= 0) prov.push({ amount, note });
    else if (note) prov.push({ amount: null, note });
  });

  const pay = [];
  payRows.forEach(r => {
    const amountEl = r.querySelector(".line-amount");
    const noteEl = r.querySelector(".line-note");
    const amount = parseBRNumber(amountEl ? amountEl.value : "");
    const note = safeText(noteEl ? noteEl.value : "").trim();
    if (amount != null && amount >= 0) pay.push({ amount, note });
    else if (note) pay.push({ amount: null, note });
  });

  entry.provisions = prov;
  entry.payments = pay;

  // keep totals in legacy fields
  entry.provisioned = sumLines(entry.provisions);
  entry.executed = sumLines(entry.payments);
  entry.value = entry.provisioned;
}

function sumLines(lines) {
  if (!Array.isArray(lines)) return 0;
  return lines.reduce((acc, it) => {
    const v = Number(it?.amount);
    return acc + (Number.isFinite(v) ? v : 0);
  }, 0);
}

function updateCellLinesTotals(entry) {
  const p = sumLines(entry.provisions);
  const e = sumLines(entry.payments);

  const elP = $("cellProvTotal");
  const elE = $("cellPayTotal");
  if (elP) elP.textContent = formatBRL(p);
  if (elE) elE.textContent = formatBRL(e);
}

function propagateNextMonths(activity, year, month, baseProvisioned, monthsAhead = 6) {
  if (!Number.isFinite(baseProvisioned) || baseProvisioned <= 0) return;

  for (let i = 1; i <= monthsAhead; i++) {
    const m = month + i;
    const y = year + Math.floor(m / 12);
    const mm = m % 12;

    const e = getEntry(activity, y, mm);

    const curProv = entryProvisioned(e);
    if (!Number.isFinite(curProv) || curProv === 0) {
      // also keep arrays for new model
      if (!Array.isArray(e.provisions)) e.provisions = [];
      if (e.provisions.length === 0) e.provisions.push({ amount: baseProvisioned, note: "" });
      e.provisioned = baseProvisioned;
      e.value = baseProvisioned;
    }
  }
}

function sortActivitiesByDueDate() {
  const year = state.year;
  const month = state.month;

  state.activities.sort((a, b) => {
    const da = getDueDate(a, year, month);
    const db = getDueDate(b, year, month);

    if (!da && !db) return 0;
    if (!da) return 1;
    if (!db) return -1;

    return da.getTime() - db.getTime();
  });
}



  function abbrBRL(value) {
    const v = Number(value);
    if (!Number.isFinite(v) || v === 0) return "—";
    const abs = Math.abs(v);
    if (abs >= 1000) {
      const k = (v / 1000);
      const kStr = k.toLocaleString("pt-BR", { maximumFractionDigits: 1, minimumFractionDigits: 0 });
      return `R$ ${kStr}k`;
    }
    const s = v.toLocaleString("pt-BR", { maximumFractionDigits: 0 });
    return `R$ ${s}`;
  }

  function uid() {
    return "A" + Math.random().toString(16).slice(2) + Date.now().toString(16);
  }

  function clampDay(day) {
    const d = Number(day);
    if (!Number.isFinite(d)) return null;
    if (d < 1) return 1;
    if (d > 31) return 31;
    return d;
  }

  function safeText(s) { return (s == null) ? "" : String(s); }

  function getBaseOnlineName() { return localStorage.getItem(BASE_ONLINE_NAME_KEY) || ""; }
  function setBaseOnlineName(name) { localStorage.setItem(BASE_ONLINE_NAME_KEY, String(name || "")); }

  function emptyEntry() {
    return {
      done: false,
      doneAt: "",
      // compatibilidade antiga
      value: null,
      // novos campos
      provisioned: null,
      executed: null,
      note: "",
      evidence: { pipefy: "", ssa: "" }
    };
  }

  // Preferência: provisioned/executed; compatibilidade: value = provisionado
  function entryProvisioned(entry) {
    if (!entry) return 0;
    const p = Number(entry.provisioned);
    if (Number.isFinite(p)) return p;

    const v = Number(entry.value);
    if (Number.isFinite(v)) return v;

    return 0;
  }

  function entryExecuted(entry) {
    if (!entry) return 0;

    const ex = Number(entry.executed);
    if (Number.isFinite(ex)) return ex;

    // Se concluído e não informou executado, assume provisionado
    if (entry.done) return entryProvisioned(entry);

    return 0;
  }

  function getEntry(activity, year, month) {
    if (!activity.entries) activity.entries = {};
    if (!activity.entries[year]) activity.entries[year] = {};
    if (!activity.entries[year][month]) activity.entries[year][month] = emptyEntry();

    // retrocompat: caso tenha entrada antiga sem novos campos
    const e = activity.entries[year][month];
    if (e && typeof e === "object") {
      if (!("provisioned" in e)) e.provisioned = null;
      if (!("executed" in e)) e.executed = null;
      if (!("evidence" in e) || !e.evidence) e.evidence = { pipefy: "", ssa: "" };
      if (!("note" in e)) e.note = "";
      if (!("done" in e)) e.done = false;
      if (!("doneAt" in e)) e.doneAt = "";
      if (!("value" in e)) e.value = null;
    }
    return activity.entries[year][month];
  }

  function isApplicable(activity, month) {
    if (!activity.active) return false;
    if (activity.periodicity === "Anual") {
      if (activity.dueType === "annualMonthDay") return Number(activity.dueMonth) === Number(month);
      return true;
    }
    return true;
  }

  function getDueDate(activity, year, month) {
    if (!activity.active) return null;

    if (activity.dueType === "monthlyDay") {
      const day = clampDay(activity.dueDay);
      if (!day) return null;
      return new Date(year, month, day, 23, 59, 59, 999);
    }

    if (activity.dueType === "annualMonthDay") {
      const m = Number(activity.dueMonth);
      const d = clampDay(activity.dueDay);
      if (!Number.isFinite(m) || !d) return null;
      if (m !== month) return null;
      return new Date(year, m, d, 23, 59, 59, 999);
    }

    return null;
  }

  function isOverdue(activity, year, month) {
    if (!isApplicable(activity, month)) return false;
    const due = getDueDate(activity, year, month);
    if (!due) return false;
    const entry = getEntry(activity, year, month);
    if (entry.done) return false;
    return new Date().getTime() > due.getTime();
  }

  function hasGapBefore(activity, year, month) {
    for (let m = 0; m < month; m++) {
      if (!isApplicable(activity, m)) continue;
      const e = getEntry(activity, year, m);
      if (!e.done) return true;
    }
    return false;
  }

  function evidenceState(activity, entry) {
    const reqP = !!activity.requirePipefy;
    const reqS = !!activity.requireSSA;
    const pFilled = !!safeText(entry?.evidence?.pipefy).trim();
    const sFilled = !!safeText(entry?.evidence?.ssa).trim();

    return {
      P: reqP ? (pFilled ? "ok" : "missing") : (pFilled ? "ok" : "hidden"),
      S: reqS ? (sFilled ? "ok" : "missing") : (sFilled ? "ok" : "hidden")
    };
  }

  function seed() {
    const y = new Date().getFullYear();
    state.year = y;
    state.month = new Date().getMonth();

    state.activities = [
      //mkActivity("Folha", "Fechamento da folha mensal", "DP", "", "Mensal", { type: "monthlyDay", day: 25 }, false, false),
      //mkActivity("Folha", "Conferência de encargos (INSS/FGTS/IRRF)", "DP", "", "Mensal", { type: "monthlyDay", day: 20 }, false, true),
      //mkActivity("Benefícios", "Conferência fatura Saúde", "RH/Financeiro", "SulAmérica", "Mensal", { type: "monthlyDay", day: 10 }, true, true),
      //mkActivity("Benefícios", "Compra VR/VA", "RH", "Alelo", "Mensal", { type: "monthlyDay", day: 5 }, true, false),
      //mkActivity("Fiscal", "Recebimento de notas fiscais de benefícios", "Financeiro", "", "Mensal", { type: "monthlyDay", day: 12 }, false, true),
      //mkActivity("RH", "Atualização indicadores (HC, Turnover)", "RH", "", "Mensal", { type: "monthlyDay", day: 8 }, false, false),
      //mkActivity("Anual", "Informe de Rendimentos", "DP", "", "Anual", { type: "annualMonthDay", month: 1, day: 28 }, false, false),
      //mkActivity("Anual", "13º Salário - 1ª Parcela", "DP/Financeiro", "", "Anual", { type: "annualMonthDay", month: 10, day: 30 }, false, false),
      //mkActivity("Anual", "13º Salário - 2ª Parcela", "DP/Financeiro", "", "Anual", { type: "annualMonthDay", month: 11, day: 20 }, false, false),

// =========================
// FORNECEDORES
// =========================

// SST - Essence
mkActivity("Fornecedores", "Conferência Fatura SST", "RH/Financeiro", "Essence", "Mensal", { type: "monthlyDay", day: 10 }, true, true),

// Cia Estágio
mkActivity("Fornecedores", "Conferência Fatura Cia Estágio", "RH/Financeiro", "Cia Estágio", "Mensal", { type: "monthlyDay", day: 10 }, true, true),

// ADP
mkActivity("Fornecedores", "Conferência Fatura ADP", "RH/Financeiro", "ADP", "Mensal", { type: "monthlyDay", day: 10 }, true, true),

// Soulan
mkActivity("Fornecedores", "Conferência Fatura Soulan", "RH/Financeiro", "Soulan", "Mensal", { type: "monthlyDay", day: 10 }, true, true),

// TV Corporativa
mkActivity("Fornecedores", "Pagamento TV Corporativa", "RH/Financeiro", "TV Corporativa", "Mensal", { type: "monthlyDay", day: 10 }, true, true),

// Frutas
mkActivity("Fornecedores", "Pagamento Fornecedor de Frutas", "RH/Financeiro", "Fornecedor Frutas", "Mensal", { type: "monthlyDay", day: 10 }, true, true),


// =========================
// BENEFÍCIOS
// =========================

// VR / VA - Alelo
mkActivity("Benefícios", "Conferência Fatura VR/VA", "RH/Financeiro", "Alelo", "Mensal", { type: "monthlyDay", day: 10 }, true, true),

// Saúde - SulAmérica
mkActivity("Benefícios", "Conferência Fatura Saúde", "RH/Financeiro", "SulAmérica", "Mensal", { type: "monthlyDay", day: 10 }, true, true),

// Odonto
mkActivity("Benefícios", "Conferência Fatura Odonto", "RH/Financeiro", "Operadora Odonto", "Mensal", { type: "monthlyDay", day: 10 }, true, true),

// VT
mkActivity("Benefícios", "Compra VT", "RH/Financeiro", "Operadora VT", "Mensal", { type: "monthlyDay", day: 3 }, true, true),

// Previdência
mkActivity("Benefícios", "Conferência Previdência", "RH/Financeiro", "BrasilPrev", "Mensal", { type: "monthlyDay", day: 10 }, true, true),

// Farmácia
mkActivity("Benefícios", "Conferência Fatura Farmácia", "RH/Financeiro", "Univers", "Mensal", { type: "monthlyDay", day: 10 }, true, true),

// Seguro de Vida
mkActivity("Benefícios", "Conferência Fatura Seguro de Vida", "RH/Financeiro", "MetLife", "Mensal", { type: "monthlyDay", day: 10 }, true, true),

// Petlove
mkActivity("Benefícios", "Conferência Fatura Petlove", "RH/Financeiro", "Petlove", "Mensal", { type: "monthlyDay", day: 10 }, true, true),

// Wellhub (Gympass)
mkActivity("Benefícios", "Conferência Fatura Wellhub", "RH/Financeiro", "Wellhub", "Mensal", { type: "monthlyDay", day: 10 }, true, true),

// Sonova Cuida
mkActivity("Benefícios", "Conferência Sonova Cuida", "RH/Financeiro", "Wellhub", "Mensal", { type: "monthlyDay", day: 10 }, true, true),



    ];

    state.archives = [];
  }

  function mkActivity(category, title, owner, supplier, periodicity, due, requirePipefy, requireSSA) {
    const id = uid();
    return {
      id,
      category: safeText(category),
      title: safeText(title),
      owner: safeText(owner),
      supplier: safeText(supplier),
      periodicity: periodicity || "Mensal",
      dueType: due?.type || "none",
      dueDay: due?.day || null,
      dueMonth: (due?.month != null ? due.month : 0),
      notes: "",
      requirePipefy: !!requirePipefy,
      requireSSA: !!requireSSA,
      active: true,
      entries: {},
      followUps: []
    };
  }

  function load() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      seed();
      save();
      return;
    }
    try {
      const obj = JSON.parse(raw);
      if (!obj || typeof obj !== "object") throw new Error("Formato inválido");
      state.year = Number(obj.year) || state.year;
      state.month = Number(obj.month);
      if (!Number.isFinite(state.month) || state.month < 0 || state.month > 11) state.month = new Date().getMonth();
      state.activities = Array.isArray(obj.activities) ? obj.activities : [];
      state.archives = Array.isArray(obj.archives) ? obj.archives : [];

      // normaliza entradas antigas (garante campos novos)
      state.activities.forEach(a => {
        if (!a || !a.entries) return;
        Object.keys(a.entries).forEach(y => {
          const byMonth = a.entries[y];
          if (!byMonth || typeof byMonth !== "object") return;
          Object.keys(byMonth).forEach(m => {
            const e = byMonth[m];
            if (!e || typeof e !== "object") return;
            if (!("provisioned" in e)) e.provisioned = null;
            if (!("executed" in e)) e.executed = null;
            if (!("evidence" in e) || !e.evidence) e.evidence = { pipefy: "", ssa: "" };
            if (!("note" in e)) e.note = "";
            if (!("done" in e)) e.done = false;
            if (!("doneAt" in e)) e.doneAt = "";
            if (!("value" in e)) e.value = null;
          });
        });
      });

    } catch {
      seed();
      save();
    }
  }

  function save() {
    const payload = {
      version: APP_VERSION,
      savedAt: nowISO(),
      year: state.year,
      month: state.month,
      activities: state.activities,
      archives: state.archives
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  }

  function ensureYearOptions() {
    const sel = $("yearSelect");
    sel.innerHTML = "";
    const current = new Date().getFullYear();
    const years = [];
    for (let y = current - 2; y <= current + 2; y++) years.push(y);
    if (!years.includes(state.year)) years.push(state.year);
    years.sort((a,b) => a-b);
    years.forEach(y => {
      const opt = document.createElement("option");
      opt.value = String(y);
      opt.textContent = String(y);
      sel.appendChild(opt);
    });
    sel.value = String(state.year);
  }

  function ensureMonthOptions() {
    const sel = $("monthSelect");
    sel.innerHTML = "";
    MONTHS.forEach((m, i) => {
      const opt = document.createElement("option");
      opt.value = String(i);
      opt.textContent = m;
      sel.appendChild(opt);
    });
    sel.value = String(state.month);
  }

  function ensureCategoryFilterOptions() {
    const sel = $("filterCategory");
    const map = new Map();
    state.activities.forEach(a => {
      const c = safeText(a?.category).trim();
      if (c) map.set(c.toLowerCase(), c);
    });
    const values = Array.from(map.values()).sort((a,b) => a.localeCompare(b, "pt-BR"));
    const currentValue = sel.value || "";
    sel.innerHTML = `<option value="">Todas</option>`;
    values.forEach(v => {
      const opt = document.createElement("option");
      opt.value = v.toLowerCase();
      opt.textContent = v;
      sel.appendChild(opt);
    });
    if (currentValue) sel.value = currentValue;
  }

  function filteredActivities() {
    const f = state.filters;
    return state.activities.filter(a => {
      if (!a || !a.active) return false;

      if (f.category) {
        if (String(a.category || "").toLowerCase() !== String(f.category).toLowerCase()) return false;
      }
      if (f.owner) {
        if (!String(a.owner || "").toLowerCase().includes(String(f.owner).toLowerCase())) return false;
      }
      if (f.supplier) {
        if (!String(a.supplier || "").toLowerCase().includes(String(f.supplier).toLowerCase())) return false;
      }
      if (f.period) {
        if (String(a.periodicity || "") !== f.period) return false;
      }
      if (f.text) {
        const needle = String(f.text).toLowerCase();
        const hay = `${a.title || ""} ${a.notes || ""}`.toLowerCase();
        if (!hay.includes(needle)) return false;
      }

      const app = isApplicable(a, state.month);
      if (f.onlyApplicable && !app) return false;

      if (f.onlyPending) {
        const e = getEntry(a, state.year, state.month);
        const over = isOverdue(a, state.year, state.month);
        const pending = !e.done && !over;
        if (!(pending || over)) return false;
      }

      return true;
    });
  }

  function mkTh(text, className) {
    const th = document.createElement("th");
    th.textContent = text;
    if (className) th.className = className;
    return th;
  }

  function mkTd(text, className) {
    const td = document.createElement("td");
    td.textContent = text;
    if (className) td.className = className;
    return td;
  }

  function renderDueText(a) {
    const type = a.dueType || "none";
    if (type === "monthlyDay") {
      const d = clampDay(a.dueDay);
      if (!d) return `<span class="badge muted">Sem prazo</span>`;
      return `Dia ${d}`;
    }
    if (type === "annualMonthDay") {
      const d = clampDay(a.dueDay);
      const m = Number(a.dueMonth);
      if (!d || !Number.isFinite(m)) return `<span class="badge muted">Sem prazo</span>`;
      return `${MONTHS[m]} dia ${d}`;
    }
    return `<span class="badge muted">Sem prazo</span>`;
  }

  function computeUptoMonth(a, year) {
    let last = -1;
    for (let m = 0; m < 12; m++) {
      if (!isApplicable(a, m)) continue;
      const e = getEntry(a, year, m);
      if (!e.done) break;
      last = m;
    }
    return last;
  }

  function renderActivityCell(a) {
    const wrap = document.createElement("div");
    wrap.style.display = "grid";
    wrap.style.gap = "4px";

    const title = document.createElement("div");
    title.style.fontWeight = "800";
    title.textContent = safeText(a.title);

    const sub = document.createElement("div");
    sub.className = "muted";
    sub.style.fontSize = "11px";
    const parts = [];
    if (a.notes) parts.push(a.notes);
    if (a.requirePipefy || a.requireSSA) {
      const evs = [];
      if (a.requirePipefy) evs.push("Pipefy");
      if (a.requireSSA) evs.push("SSA");
      parts.push("Evidências: " + evs.join(" / "));
    }
    sub.textContent = parts.join(" | ");

    const bar = document.createElement("div");
    bar.style.display = "flex";
    bar.style.gap = "8px";
    bar.style.flexWrap = "wrap";
    bar.style.alignItems = "center";

    const editBtn = document.createElement("button");
    editBtn.className = "link-btn";
    editBtn.textContent = "Editar";
    editBtn.addEventListener("click", () => openActivity(a.id));

    const upto = computeUptoMonth(a, state.year);
    const badge = document.createElement("span");
    badge.className = "badge muted";
    badge.textContent = (upto >= 0) ? `Em dia até: ${MONTHS[upto]}` : "Sem conclusão no ano";

    bar.append(editBtn, badge);
    wrap.append(title, sub, bar);
    return wrap;
  }

  function renderMonthCell(a, year, month, forceNA) {
    const box = document.createElement("div");
    box.className = "cell-box";

    const top = document.createElement("div");
    top.className = "cell-top";

    const check = document.createElement("div");
    check.className = "cell-check";

    const e = getEntry(a, year, month);

    if (!forceNA && e.done) {
      check.classList.add("done");
      check.textContent = "✓";
    } else {
      check.textContent = "";
    }

    const val = document.createElement("div");
    val.className = "cell-value";

    if (forceNA) {
      val.textContent = "—";
    } else {
      const p = entryProvisioned(e);
      const ex = entryExecuted(e);
      const show = e.done ? (ex > 0 ? ex : p) : p;
      val.textContent = (show > 0) ? abbrBRL(show) : "—";
    }

    top.append(check, val);

    const sub = document.createElement("div");
    sub.className = "cell-sub";

    const left = document.createElement("div");
    if (forceNA) left.textContent = "N/A";
    else if (e.done) left.textContent = formatDateShort(e.doneAt);
    else left.textContent = isOverdue(a, year, month) ? "Vencido" : "";

    const right = document.createElement("div");
    right.className = "evs";

    if (!forceNA) {
      const st = evidenceState(a, e);
      if (st.P !== "hidden") {
        const p = document.createElement("div");
        p.className = "ev";
        p.textContent = "P";
        if (st.P === "ok") p.classList.add("ok");
        if (st.P === "missing") p.classList.add("req-missing");
        right.appendChild(p);
      }
      if (st.S !== "hidden") {
        const s = document.createElement("div");
        s.className = "ev";
        s.textContent = "S";
        if (st.S === "ok") s.classList.add("ok");
        if (st.S === "missing") s.classList.add("req-missing");
        right.appendChild(s);
      }
    }

    sub.append(left, right);
    box.append(top, sub);
    return box;
  }

  function renderGrid() {
    sortActivitiesByDueDate();
    const head = $("gridHead");
    const body = $("gridBody");

    head.innerHTML = "";
    body.innerHTML = "";

    const trh = document.createElement("tr");

    trh.append(
      mkTh("Categoria", "col-cat sticky-left-1"),
      mkTh("Atividade", "col-act sticky-left-2"),
      mkTh("Responsável", "col-owner sticky-left-3"),
      mkTh("Fornecedor", "col-supplier"),
      mkTh("Periodicidade", "col-period"),
      mkTh("Prazo", "col-due"),
      mkTh("Ações", "col-actions")
    );

    MONTHS.forEach(m => trh.appendChild(mkTh(m, "cell-month")));
    head.appendChild(trh);

    const list = filteredActivities();

    list.forEach((a, idx) => {
      const row = document.createElement("tr");

      // Zebra visual: se seu CSS já faz, ok. Se não, ajuda em impressão/visual
      row.classList.toggle("row-zebra", (idx % 2) === 1);

      const catTd = mkTd(safeText(a.category), "sticky-left-1 col-cat");

      const actTd = document.createElement("td");
      actTd.className = "sticky-left-2 col-act";
      actTd.appendChild(renderActivityCell(a));

      const ownerTd = mkTd(safeText(a.owner), "sticky-left-3 col-owner");
      const supplierTd = mkTd(safeText(a.supplier), "col-supplier");

      const periodTd = document.createElement("td");
      periodTd.className = "col-period";
      periodTd.innerHTML = `<span class="badge">${safeText(a.periodicity || "Mensal")}</span>`;

      const dueTd = document.createElement("td");
      dueTd.className = "col-due";
      dueTd.innerHTML = renderDueText(a);

      const actionsTd = document.createElement("td");
      actionsTd.className = "col-actions";
      const btnFollow = document.createElement("button");
      btnFollow.className = "link-btn";
      btnFollow.textContent = "Follow-up";
      btnFollow.addEventListener("click", () => openFollow(a.id));
      actionsTd.appendChild(btnFollow);

      row.append(catTd, actTd, ownerTd, supplierTd, periodTd, dueTd, actionsTd);

      const entry = getEntry(a, state.year, state.month);
      const over = isOverdue(a, state.year, state.month);
      if (isApplicable(a, state.month)) {
        if (entry.done) row.classList.add("row-highlight-ok");
        else if (over) row.classList.add("row-highlight-over");
      }

      for (let m = 0; m < 12; m++) {
        const td = document.createElement("td");
        td.className = "cell-month";

        const app = isApplicable(a, m);
        if (!app) {
          td.classList.add("cell-na");
          td.appendChild(renderMonthCell(a, state.year, m, true));
          row.appendChild(td);
          continue;
        }

        const e = getEntry(a, state.year, m);
        const overM = isOverdue(a, state.year, m);
        const gap = (!e.done && hasGapBefore(a, state.year, m) && m <= state.month);

        if (e.done) td.classList.add("cell-ok");
        else if (overM) td.classList.add("cell-over");
        else if (gap) td.classList.add("cell-gap");

        td.appendChild(renderMonthCell(a, state.year, m, false));
        td.addEventListener("click", (ev) => {
          ev.preventDefault();
          openCell(a.id, state.year, m);
        });

        row.appendChild(td);
      }

      body.appendChild(row);
    });

    requestAnimationFrame(() => {
      updateStickyOffsets();
    });
  }

  function renderSummary() {
    const list = state.activities.filter(a => a && a.active);

    let totalApplicable = 0;
    let done = 0;
    let overdue = 0;
    let pending = 0;

    let provisioned = 0;
    let executed = 0;

    list.forEach(a => {
      if (!isApplicable(a, state.month)) return;
      totalApplicable++;

      const e = getEntry(a, state.year, state.month);
      const over = isOverdue(a, state.year, state.month);

      if (e.done) done++;
      else if (over) overdue++;
      else pending++;

      const p = entryProvisioned(e);
      const ex = entryExecuted(e);

      if (p > 0) provisioned += p;
      if (ex > 0) executed += ex;
    });

    const gap = (provisioned - executed);

    $("kpiTotal").textContent = String(totalApplicable);
    $("kpiDone").textContent = String(done);
    $("kpiPending").textContent = String(pending);
    $("kpiOverdue").textContent = String(overdue);

    $("moneyProvisioned").textContent = formatBRL(provisioned);
    $("moneyExecuted").textContent = formatBRL(executed);
    $("moneyGap").textContent = formatBRL(gap);

    drawChart(provisioned, executed, gap);
  }

  function drawChart(provisioned, executed, gap) {
    const canvas = $("summaryChart");
    const ctx = canvas.getContext("2d");
    const w = canvas.width;
    const h = canvas.height;

    ctx.clearRect(0,0,w,h);

    const pad = 18;
    const barH = 22;
    const gapY = 18;

    const max = Math.max(1, provisioned, executed);
    const scale = (w - pad*2) / max;

    ctx.fillStyle = "rgba(17,24,39,0.85)";
    ctx.font = "12px Calibri, Arial";
    ctx.fillText(`Mês: ${MONTHS[state.month]} / ${state.year}`, pad, 14);

    const y1 = 45;
    const y2 = y1 + barH + gapY;

    ctx.strokeStyle = "rgba(0,74,153,0.35)";
    ctx.lineWidth = 2;
    ctx.strokeRect(pad, y1, (provisioned * scale), barH);

    ctx.fillStyle = "rgba(0,74,153,0.20)";
    ctx.fillRect(pad, y2, (executed * scale), barH);

    ctx.strokeStyle = "rgba(0,0,0,0.10)";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(pad + (provisioned * scale), y2);
    ctx.lineTo(pad + (provisioned * scale), y2 + barH);
    ctx.stroke();

    ctx.fillStyle = "rgba(17,24,39,0.75)";
    ctx.font = "12px Calibri, Arial";
    ctx.fillText(`Provisionado: ${formatBRL(provisioned)}`, pad, y1 - 6);
    ctx.fillText(`Executado: ${formatBRL(executed)}`, pad, y2 - 6);
    ctx.fillText(`Gap: ${formatBRL(gap)}`, pad, y2 + barH + 16);

        // Área hachurada do GAP (diferença entre Provisionado e Executado)
    // gap > 0  => faltou executar (Provisionado maior)
    // gap < 0  => excedeu (Executado maior)
    if (gap !== 0) {
      const minVal = Math.min(provisioned, executed);
      const maxVal = Math.max(provisioned, executed);

      const start = pad + (minVal * scale);
      const end = pad + (maxVal * scale);

      ctx.strokeStyle = "rgba(17,24,39,0.20)";
      ctx.lineWidth = 1;

      // Direção da hachura muda para diferenciar "faltou" vs "excedeu" mesmo em PB
      const down = (gap > 0);

      for (let x = start; x < end; x += 6) {
        ctx.beginPath();
        if (down) {
          ctx.moveTo(x, y2);
          ctx.lineTo(x + 6, y2 + barH);
        } else {
          ctx.moveTo(x, y2 + barH);
          ctx.lineTo(x + 6, y2);
        }
        ctx.stroke();
      }
    }
  }

  // Sticky offsets calculados (corrige o “Responsável deslocado/sobreposto”)
  function updateStickyOffsets() {
    const table = $("gridTable");
    const headRow = $("gridHead")?.querySelector("tr");
    if (!table || !headRow) return;

    const ths = headRow.querySelectorAll("th");
    if (ths.length < 3) return;

    const th1 = ths[0]; // Categoria
    const th2 = ths[1]; // Atividade
    const th3 = ths[2]; // Responsável

    const w1 = Math.ceil(th1.getBoundingClientRect().width);
    const w2 = Math.ceil(th2.getBoundingClientRect().width);
    const w3 = Math.ceil(th3.getBoundingClientRect().width);

    table.style.setProperty("--sticky-left-2", `${w1}px`);
    table.style.setProperty("--sticky-left-3", `${w1 + w2}px`);

    updateTopbarHeight();
    void w3;
  }

  function updateTopbarHeight() {
    const topbar = $("topbar");
    if (!topbar) return;
    const h = Math.ceil(topbar.getBoundingClientRect().height);
    document.documentElement.style.setProperty("--topbar-h", `${h}px`);
  }

  // Modals
  function openModal(id) { $(id).classList.remove("hidden"); }
  function closeModal(id) { $(id).classList.add("hidden"); }

  function wireModalClose() {
    document.querySelectorAll("[data-close]").forEach(el => {
      el.addEventListener("click", () => closeModal(el.getAttribute("data-close")));
    });
  }

  // Activity CRUD
  function openActivity(activityId) {
    const isEdit = !!activityId;
    state.ui.editingActivityId = activityId || null;

    $("modalActivityTitle").textContent = isEdit ? "Editar Atividade" : "Nova Atividade";
    $("btnDeleteActivity").classList.toggle("hidden", !isEdit);
    $("actIdRow").classList.toggle("hidden", !isEdit);

    if (!isEdit) {
      $("actCategory").value = "";
      $("actTitle").value = "";
      $("actOwner").value = "";
      $("actSupplier").value = "";
      $("actPeriod").value = "Mensal";
      $("actDueType").value = "monthlyDay";
      $("actDueDay").value = "10";
      $("actDueMonth").value = String(state.month);
      $("actNotes").value = "";
      $("actReqPipefy").checked = false;
      $("actReqSSA").checked = false;
      $("actActive").checked = true;
      $("actId").value = "";
      syncDueFields();
      openModal("modalActivity");
      return;
    }

    const a = state.activities.find(x => x.id === activityId);
    if (!a) return;

    $("actCategory").value = safeText(a.category);
    $("actTitle").value = safeText(a.title);
    $("actOwner").value = safeText(a.owner);
    $("actSupplier").value = safeText(a.supplier);
    $("actPeriod").value = safeText(a.periodicity || "Mensal");
    $("actDueType").value = safeText(a.dueType || "none");
    $("actDueDay").value = (a.dueDay != null ? String(a.dueDay) : "");
    $("actDueMonth").value = String(a.dueMonth != null ? a.dueMonth : 0);
    $("actNotes").value = safeText(a.notes);
    $("actReqPipefy").checked = !!a.requirePipefy;
    $("actReqSSA").checked = !!a.requireSSA;
    $("actActive").checked = !!a.active;
    $("actId").value = safeText(a.id);

    syncDueFields();
    openModal("modalActivity");
  }

  function syncDueFields() {
    const type = $("actDueType").value;
    const isAnnual = (type === "annualMonthDay");
    $("actDueDay").disabled = (type === "none");
    $("actDueMonth").disabled = !isAnnual;
    if (type === "monthlyDay") $("actDueMonth").value = String(state.month);
  }

  function saveActivityFromForm() {
    const id = state.ui.editingActivityId;

    const category = safeText($("actCategory").value).trim() || "Outro";
    const title = safeText($("actTitle").value).trim();
    if (!title) { alert("Informe a Atividade."); return; }

    const owner = safeText($("actOwner").value).trim();
    const supplier = safeText($("actSupplier").value).trim();
    const periodicity = $("actPeriod").value;
    const dueType = $("actDueType").value;
    const dueDay = clampDay($("actDueDay").value);
    const dueMonth = Number($("actDueMonth").value);

    const notes = safeText($("actNotes").value).trim();
    const requirePipefy = $("actReqPipefy").checked;
    const requireSSA = $("actReqSSA").checked;
    const active = $("actActive").checked;

    if (!id) {
      const a = mkActivity(category, title, owner, supplier, periodicity, {}, requirePipefy, requireSSA);
      a.periodicity = periodicity;
      a.dueType = dueType;
      a.dueDay = (dueType === "none") ? null : dueDay;
      a.dueMonth = (Number.isFinite(dueMonth) ? dueMonth : 0);
      a.notes = notes;
      a.active = active;
      state.activities.unshift(a);
    } else {
      const a = state.activities.find(x => x.id === id);
      if (!a) return;
      a.category = category;
      a.title = title;
      a.owner = owner;
      a.supplier = supplier;
      a.periodicity = periodicity;
      a.dueType = dueType;
      a.dueDay = (dueType === "none") ? null : dueDay;
      a.dueMonth = (Number.isFinite(dueMonth) ? dueMonth : 0);
      a.notes = notes;
      a.requirePipefy = requirePipefy;
      a.requireSSA = requireSSA;
      a.active = active;
    }

    save();
    closeModal("modalActivity");
    render();
  }

  function deleteActivity() {
    const id = state.ui.editingActivityId;
    if (!id) return;
    const ok = confirm("Excluir esta atividade? Esta ação não pode ser desfeita.");
    if (!ok) return;
    state.activities = state.activities.filter(x => x.id !== id);
    save();
    closeModal("modalActivity");
    render();
  }

  // Cell edit
  function openCell(activityId, year, month) {
    const a = state.activities.find(x => x.id === activityId);
    if (!a) return;
    if (!isApplicable(a, month)) return;

    state.ui.cellContext = { activityId, year, month };
    const entry = getEntry(a, year, month);

    $("modalCellTitle").textContent = `Detalhe do Mês: ${MONTHS[month]} / ${year}`;
    $("cellMeta").textContent = `${a.category} | ${a.title} | Resp: ${a.owner || "-"} | Forn: ${a.supplier || "-"}`;

    $("cellDone").checked = !!entry.done;
    $("cellDoneHint").textContent = entry.doneAt ? `Concluído em: ${formatDateLong(entry.doneAt)}` : "";

    // Compatibilidade de HTML:
    // - novo: cellProvisioned / cellExecuted
    // - antigo: cellValue (considera como provisionado)
    const elProv = $("cellProvisioned");
    const elExec = $("cellExecuted");
    const elOld = $("cellValue");

    const prov = entryProvisioned(entry);
    const exec = Number(entry.executed);
    const hasExec = Number.isFinite(exec);

    if (elProv) elProv.value = prov > 0 ? String(prov).replace(".", ",") : "";
    if (elExec) elExec.value = hasExec ? String(exec).replace(".", ",") : "";

    if (!elProv && elOld) {
      elOld.value = prov > 0 ? String(prov).replace(".", ",") : "";
    }

    $("cellNote").value = safeText(entry.note);
    $("cellPipefy").value = safeText(entry?.evidence?.pipefy);
    $("cellSSA").value = safeText(entry?.evidence?.ssa);

    renderCellLinesUI(entry);

    openModal("modalCell");
  }

  function saveCellFromForm() {
    const ctx = state.ui.cellContext;
    if (!ctx) return;

    const a = state.activities.find(x => x.id === ctx.activityId);
    if (!a) return;

    const entry = getEntry(a, ctx.year, ctx.month);
    const done = $("cellDone").checked;
    const prevDone = !!entry.done;

    entry.done = done;
    if (done && !prevDone) entry.doneAt = nowISO();
    if (!done) entry.doneAt = "";


    const hasLinesUI = !!$("cellLinesWrap");

    if (hasLinesUI) {
      // Novo modelo: múltiplas linhas (provisões/pagamentos)
      syncEntryLinesFromUI(entry);
    } else {
      // Compatibilidade: um ou dois campos (provisionado/executado)
      const elProv = $("cellProvisioned");
      const elExec = $("cellExecuted");
      const elOld = $("cellValue");

      if (elProv || elExec) {
        const p = parseBRNumber(elProv ? elProv.value : "");
        const ex = parseBRNumber(elExec ? elExec.value : "");

        entry.provisioned = (p != null && p >= 0) ? p : null;
        entry.executed = (ex != null && ex >= 0) ? ex : null;

        // compatibilidade antiga: value = provisionado
        entry.value = entry.provisioned;

        // sementes para arrays (para transição suave)
        if (!Array.isArray(entry.provisions)) entry.provisions = [];
        if (!Array.isArray(entry.payments)) entry.payments = [];
        if (entry.provisions.length === 0 && Number.isFinite(Number(entry.provisioned)) && Number(entry.provisioned) > 0) {
          entry.provisions.push({ amount: Number(entry.provisioned), note: "" });
        }
        if (entry.payments.length === 0 && Number.isFinite(Number(entry.executed)) && Number(entry.executed) > 0) {
          entry.payments.push({ amount: Number(entry.executed), note: "" });
        }

      } else if (elOld) {
        // HTML antigo: um campo só (provisão)
        const val = parseBRNumber(elOld.value);
        entry.value = (val != null && val >= 0) ? val : null;
        entry.provisioned = entry.value;
        entry.executed = null;

        if (!Array.isArray(entry.provisions)) entry.provisions = [];
        if (entry.provisions.length === 0 && Number.isFinite(Number(entry.provisioned)) && Number(entry.provisioned) > 0) {
          entry.provisions.push({ amount: Number(entry.provisioned), note: "" });
        }
        if (!Array.isArray(entry.payments)) entry.payments = [];
      }
    }

    // Regra: se marcou concluído e não informou executado/pagamentos, assume executado = provisionado
    const pTotal = entryProvisioned(entry);
    const eTotal = Number(entry.executed);
    const hasETotal = Number.isFinite(eTotal) && eTotal > 0;

    if (done && (!hasETotal) && Number.isFinite(pTotal) && pTotal > 0) {
      entry.executed = pTotal;
      if (!Array.isArray(entry.payments)) entry.payments = [];
      if (entry.payments.length === 0) entry.payments.push({ amount: pTotal, note: "Auto (Concluído)" });
    }

    entry.note = safeText($("cellNote").value).trim();
    // Replica provisão para próximos meses (6) sem sobrescrever meses já preenchidos
    propagateNextMonths(a, ctx.year, ctx.month, entryProvisioned(entry), 6);

    entry.evidence = {
      pipefy: safeText($("cellPipefy").value).trim(),
      ssa: safeText($("cellSSA").value).trim()
    };

    save();
    closeModal("modalCell");
    render();
  }

  // Follow-up
  function openFollow(activityId) {
    const a = state.activities.find(x => x.id === activityId);
    if (!a) return;

    state.ui.followContext = { activityId };
    $("modalFollowTitle").textContent = `Follow-up: ${a.title}`;
    $("followMeta").textContent = `${a.category} | Resp: ${a.owner || "-"} | Forn: ${a.supplier || "-"} | Period: ${a.periodicity || "-"}`;

    $("followType").value = "Cobrança fornecedor";
    $("followNextAction").value = "";
    $("followNextDate").value = "";
    $("followText").value = "";
    $("copyHint").textContent = "";

    renderFollowHistory(a);
    openModal("modalFollow");
  }

  function renderFollowHistory(a) {
    const box = $("followHistory");
    box.innerHTML = "";

    const list = Array.isArray(a.followUps) ? a.followUps.slice() : [];
    list.sort((x,y) => (safeText(y.ts)).localeCompare(safeText(x.ts)));

    if (list.length === 0) {
      const empty = document.createElement("div");
      empty.className = "muted";
      empty.style.fontSize = "12px";
      empty.textContent = "Sem histórico registrado.";
      box.appendChild(empty);
      return;
    }

    list.forEach(item => {
      const card = document.createElement("div");
      card.className = "history-item";

      const head = document.createElement("div");
      head.className = "history-head";

      const p1 = document.createElement("span");
      p1.className = "pill blue";
      p1.textContent = safeText(item.type || "Follow-up");

      const p2 = document.createElement("span");
      p2.className = "pill";
      p2.textContent = safeText(formatDateLong(item.ts));

      head.append(p1, p2);

      const body = document.createElement("div");
      body.className = "history-body";
      const lines = [];
      if (item.text) lines.push(item.text);
      if (item.nextAction) lines.push(`Próxima ação: ${item.nextAction}`);
      if (item.nextDate) lines.push(`Data próxima ação: ${item.nextDate}`);
      body.textContent = lines.join("\n");

      card.append(head, body);
      box.appendChild(card);
    });
  }

  function saveFollowUp() {
    const ctx = state.ui.followContext;
    if (!ctx) return;

    const a = state.activities.find(x => x.id === ctx.activityId);
    if (!a) return;

    const text = safeText($("followText").value).trim();
    if (!text) { alert("Informe o registro do follow-up."); return; }

    const item = {
      ts: nowISO(),
      type: $("followType").value,
      text,
      nextAction: safeText($("followNextAction").value).trim(),
      nextDate: safeText($("followNextDate").value).trim(),
    };

    if (!Array.isArray(a.followUps)) a.followUps = [];
    a.followUps.push(item);

    save();
    renderFollowHistory(a);
    $("followText").value = "";
    $("followNextAction").value = "";
    $("followNextDate").value = "";
  }

  async function copyToClipboard(text) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      try {
        const ta = document.createElement("textarea");
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        const ok = document.execCommand("copy");
        document.body.removeChild(ta);
        return ok;
      } catch {
        return false;
      }
    }
  }

  function buildEmailText(a) {
    const month = state.month;
    const year = state.year;
    const entry = getEntry(a, year, month);
    const over = isOverdue(a, year, month);

    const status = entry.done ? "Concluído" : (over ? "Vencido" : "Pendente");

    const p = entryProvisioned(entry);
    const ex = entryExecuted(entry);
    const valProv = (p > 0) ? formatBRL(p) : "—";
    const valExec = (ex > 0) ? formatBRL(ex) : "—";

    const due = getDueDate(a, year, month);
    const dueStr = due ? `${pad2(due.getDate())}/${pad2(due.getMonth()+1)}/${due.getFullYear()}` : "Sem prazo";

    const lines = [];
    lines.push(`Assunto: Follow-up – ${a.title} – ${MONTHS[month]}/${year}`);
    lines.push("");
    lines.push(`Atividade: ${a.title}`);
    lines.push(`Categoria: ${a.category || "-"}`);
    lines.push(`Responsável: ${a.owner || "-"}`);
    lines.push(`Fornecedor: ${a.supplier || "-"}`);
    lines.push(`Mês/Ano: ${MONTHS[month]}/${year}`);
    lines.push(`Status: ${status}`);
    lines.push(`Prazo: ${dueStr}`);
    lines.push(`Valor provisionado: ${valProv}`);
    lines.push(`Valor executado: ${valExec}`);
    lines.push(`Pipefy: ${safeText(entry?.evidence?.pipefy) || "-"}`);
    lines.push(`SSA: ${safeText(entry?.evidence?.ssa) || "-"}`);

    const followText = safeText($("followText").value).trim();
    if (followText) {
      lines.push("");
      lines.push("Registro (follow-up):");
      lines.push(followText);
    }

    const nextAction = safeText($("followNextAction").value).trim();
    const nextDate = safeText($("followNextDate").value).trim();
    if (nextAction || nextDate) {
      lines.push("");
      lines.push(`Próxima ação: ${nextAction || "-"}`);
      lines.push(`Data próxima ação: ${nextDate || "-"}`);
    }

    return lines.join("\n");
  }

  async function copyEmail() {
    const ctx = state.ui.followContext;
    if (!ctx) return;
    const a = state.activities.find(x => x.id === ctx.activityId);
    if (!a) return;

    const text = buildEmailText(a);
    const ok = await copyToClipboard(text);
    $("copyHint").textContent = ok ? "Texto copiado. Cole no e-mail (Outlook/Gmail)." : "Não foi possível copiar automaticamente. Selecione e copie manualmente.";
  }

  // TXT Export/Import
  function downloadText(text, filename) {
    const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  function exportTxt() {
    const baseOnline = getBaseOnlineName().trim();
    const filename = baseOnline || `calendario_rh_sonova_${state.year}-${pad2(state.month+1)}.txt`;

    const payload = {
      type: "SONOVA_CALENDARIO_RH",
      version: APP_VERSION,
      exportedAt: nowISO(),
      year: state.year,
      month: state.month,
      baseOnlineName: baseOnline,
      activities: state.activities,
      archives: state.archives
    };

    downloadText(JSON.stringify(payload, null, 2), filename);
  }

  function importTxtFromFile(file, mode = "replace") {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const text = String(reader.result || "");
        const obj = JSON.parse(text);

        if (!obj || typeof obj !== "object") throw new Error("Arquivo inválido");
        if (!Array.isArray(obj.activities)) throw new Error("Base sem atividades");

        if (mode === "replace") {
          state.activities = obj.activities;
          state.archives = Array.isArray(obj.archives) ? obj.archives : [];
          if (Number.isFinite(Number(obj.year))) state.year = Number(obj.year);
          if (Number.isFinite(Number(obj.month))) state.month = Number(obj.month);
          save();
          render();
          return;
        }

        state.ui.prevImportBuffer = obj;
        applyPreviousImportMode(mode);
      } catch {
        alert("Não foi possível importar. Verifique o TXT exportado pelo sistema.");
      }
    };
    reader.readAsText(file, "utf-8");
  }

  function applyPreviousImportMode(mode) {
    const obj = state.ui.prevImportBuffer;
    if (!obj) return;

    if (mode === "structure") {
      state.activities = obj.activities.map(a => ({ ...a, entries: {}, followUps: [] }));
      save();
      render();
      return;
    }

    if (mode === "structurePlusValues") {
      const y = state.year;
      const m = state.month;

      state.activities = obj.activities.map(a => {
        const clone = { ...a, entries: {}, followUps: [] };

        const srcYear = Number.isFinite(Number(obj.year)) ? Number(obj.year) : y;
        const srcMonth = Number.isFinite(Number(obj.month)) ? Number(obj.month) : m;

        let srcEntry = null;
        try {
          if (a.entries && a.entries[srcYear] && a.entries[srcYear][srcMonth]) srcEntry = a.entries[srcYear][srcMonth];
          else if (a.entries && a.entries[y] && a.entries[y][m]) srcEntry = a.entries[y][m];
        } catch {}

        const entry = emptyEntry();

        const v = Number(srcEntry?.value);
        const p = Number(srcEntry?.provisioned);
        const ex = Number(srcEntry?.executed);

        if (Number.isFinite(p)) entry.provisioned = p;
        else if (Number.isFinite(v)) entry.provisioned = v;

        if (Number.isFinite(ex)) entry.executed = ex;

        entry.value = entry.provisioned; // compat
        entry.note = safeText(srcEntry?.note || "");
        entry.evidence = { pipefy: "", ssa: "" };
        entry.done = false;
        entry.doneAt = "";

        clone.entries[y] = {};
        clone.entries[y][m] = entry;

        return clone;
      });

      save();
      render();
      return;
    }

    if (mode === "archiveAll") {
      state.archives.push({
        ts: nowISO(),
        label: `Arquivo importado: ${safeText(obj.year)}-${pad2((Number(obj.month)||0)+1)}`,
        data: obj
      });
      save();
      alert("Base anterior arquivada. Histórico preservado (interno).");
      render();
      return;
    }
  }

  // Base Online modal
  function openBaseOnline() {
    $("baseOnlineName").value = getBaseOnlineName();
    $("baseOnlinePaste").value = "";
    openModal("modalBaseOnline");
  }

  async function copyBaseOnlineInstructions() {
    const name = getBaseOnlineName().trim() || "calendario_rh_sonova_oficial.txt";
    const lines = [];
    lines.push("Fluxo recomendado (Base Online):");
    lines.push(`1) Defina o nome do arquivo oficial: ${name}`);
    lines.push("2) Trabalhe normalmente (tudo salva no localStorage).");
    lines.push("3) Ao final do dia/mês, clique em 'Exportar Base (TXT)'.");
    lines.push("4) Substitua o arquivo oficial no projeto (GitHub/SharePoint/Drive).");
    lines.push("5) Em outra máquina, use 'Importar Base (TXT)' ou cole o TXT na Base Online e carregue.");
    const ok = await copyToClipboard(lines.join("\n"));
    alert(ok ? "Instruções copiadas." : "Não foi possível copiar automaticamente.");
  }

  function saveBaseOnlineNameFromModal() {
    setBaseOnlineName(safeText($("baseOnlineName").value).trim());
    closeModal("modalBaseOnline");
  }

  function loadPastedBase() {
    const text = safeText($("baseOnlinePaste").value).trim();
    if (!text) { alert("Cole o conteúdo do TXT primeiro."); return; }
    try {
      const obj = JSON.parse(text);
      if (!obj || typeof obj !== "object" || !Array.isArray(obj.activities)) {
        alert("Conteúdo inválido. Cole o TXT exportado pelo sistema.");
        return;
      }
      const ok = confirm("Carregar a base colada vai substituir sua base atual. Deseja continuar?");
      if (!ok) return;
      state.activities = obj.activities;
      state.archives = Array.isArray(obj.archives) ? obj.archives : [];
      if (Number.isFinite(Number(obj.year))) state.year = Number(obj.year);
      if (Number.isFinite(Number(obj.month))) state.month = Number(obj.month);
      save();
      closeModal("modalBaseOnline");
      render();
    } catch {
      alert("Conteúdo inválido. Cole o TXT exportado pelo sistema.");
    }
  }

  // Excel export (.xls via HTML)
  function escHtml(s) {
    return String(s || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function renderDueTextPlain(a) {
    const type = a.dueType || "none";
    if (type === "monthlyDay") {
      const d = clampDay(a.dueDay);
      return d ? `Dia ${d}` : "Sem prazo";
    }
    if (type === "annualMonthDay") {
      const d = clampDay(a.dueDay);
      const m = Number(a.dueMonth);
      return (d && Number.isFinite(m)) ? `${MONTHS[m]} dia ${d}` : "Sem prazo";
    }
    return "Sem prazo";
  }

  function computeTotalsForMonth(year, month) {
    let provisioned = 0;
    let executed = 0;
    state.activities.forEach(a => {
      if (!a || !a.active) return;
      if (!isApplicable(a, month)) return;
      const e = getEntry(a, year, month);
      const p = entryProvisioned(e);
      const ex = entryExecuted(e);
      if (p > 0) provisioned += p;
      if (ex > 0) executed += ex;
    });
    return { provisioned, executed, gap: (provisioned - executed) };
  }

  function buildExcelHtml() {
    const rows = state.activities
      .filter(a => a && a.active)
      .map(a => {
        const cols = [];
        cols.push(`<td>${escHtml(a.category)}</td>`);
        cols.push(`<td>${escHtml(a.title)}</td>`);
        cols.push(`<td>${escHtml(a.owner)}</td>`);
        cols.push(`<td>${escHtml(a.supplier)}</td>`);
        cols.push(`<td>${escHtml(a.periodicity)}</td>`);
        cols.push(`<td>${escHtml(renderDueTextPlain(a))}</td>`);

        for (let m = 0; m < 12; m++) {
          if (!isApplicable(a, m)) {
            cols.push(`<td style="background:#f3f4f6;">N/A</td>`);
            continue;
          }
          const e = getEntry(a, state.year, m);
          const status = e.done ? "OK" : (isOverdue(a, state.year, m) ? "VENCIDO" : "PENDENTE");
          const dt = e.doneAt ? formatDateShort(e.doneAt) : "";
          const p = entryProvisioned(e);
          const ex = entryExecuted(e);

          const valP = (p > 0) ? formatBRL(p) : "";
          const valE = (ex > 0) ? formatBRL(ex) : "";

          const ev = evidenceState(a, e);
          const pipe = (ev.P === "hidden") ? "" : (ev.P === "ok" ? "P:OK" : "P:FALTA");
          const ssa = (ev.S === "hidden") ? "" : (ev.S === "ok" ? "S:OK" : "S:FALTA");

          const cell = [
            status,
            dt,
            valP ? `Prov: ${valP}` : "",
            valE ? `Exec: ${valE}` : "",
            pipe,
            ssa
          ].filter(Boolean).join(" | ");

          cols.push(`<td>${escHtml(cell)}</td>`);
        }

        return `<tr>${cols.join("")}</tr>`;
      })
      .join("");

    const monthLabel = `${MONTHS[state.month]}/${state.year}`;
    const totals = computeTotalsForMonth(state.year, state.month);

    return `
      <html><head><meta charset="utf-8" /></head><body>
        <table border="1">
          <tr>
            <th colspan="18" style="background:#004A99;color:#fff;font-weight:bold;text-align:left;">
              Calendário Operacional RH Sonova - Exportação
            </th>
          </tr>
          <tr>
            <td colspan="18">
              Mês de referência: ${escHtml(monthLabel)} | Exportado em: ${escHtml(formatDateLong(nowISO()))}
            </td>
          </tr>
          <tr><td colspan="6">Total Provisionado (mês)</td><td colspan="12">${escHtml(formatBRL(totals.provisioned))}</td></tr>
          <tr><td colspan="6">Total Executado (mês)</td><td colspan="12">${escHtml(formatBRL(totals.executed))}</td></tr>
          <tr><td colspan="6">Gap (mês)</td><td colspan="12">${escHtml(formatBRL(totals.gap))}</td></tr>

          <tr style="background:#f3f4f6;font-weight:bold;">
            <th>Categoria</th>
            <th>Atividade</th>
            <th>Responsável</th>
            <th>Fornecedor</th>
            <th>Periodicidade</th>
            <th>Prazo</th>
            ${MONTHS.map(m => `<th>${m}</th>`).join("")}
          </tr>
          ${rows}
        </table>
      </body></html>
    `;
  }

  function exportExcel() {
    const html = buildExcelHtml();
    const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `calendario_operacional_rh_${state.year}.xls`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  // Print mode
  function setPrintMode(on) {
    state.ui.printMode = !!on;
    document.body.classList.toggle("print-mode", state.ui.printMode);
    $("printExitBar").classList.toggle("hidden", !state.ui.printMode);
    requestAnimationFrame(() => {
      updateStickyOffsets();
    });
  }

  function togglePrintMode() {
    setPrintMode(!state.ui.printMode);
  }

  // Eventos / Wiring
  function wireControls() {
    $("yearSelect").addEventListener("change", () => {
      state.year = Number($("yearSelect").value);
      save();
      render();
    });

    $("monthSelect").addEventListener("change", () => {
      state.month = Number($("monthSelect").value);
      save();
      render();
    });

    $("btnNewActivity").addEventListener("click", () => openActivity(null));
    $("btnExportTxt").addEventListener("click", exportTxt);
    $("btnImportTxt").addEventListener("click", () => $("fileImportTxt").click());

    $("fileImportTxt").addEventListener("change", (e) => {
      const file = e.target.files?.[0];
      e.target.value = "";
      if (!file) return;
      const ok = confirm("Importar vai substituir sua base atual. Deseja continuar?");
      if (!ok) return;
      importTxtFromFile(file, "replace");
    });

    $("btnImportPrevious").addEventListener("click", () => openModal("modalImportPrevious"));
    $("btnPickPreviousFile").addEventListener("click", () => $("fileImportPrevious").click());

    $("fileImportPrevious").addEventListener("change", (e) => {
      const file = e.target.files?.[0];
      e.target.value = "";
      if (!file) return;
      const mode = document.querySelector('input[name="prevMode"]:checked')?.value || "structure";
      importTxtFromFile(file, mode);
      closeModal("modalImportPrevious");
    });

    $("btnExportExcel").addEventListener("click", exportExcel);

    $("btnReset").addEventListener("click", () => {
      const ok = confirm("Resetar base apaga seus dados locais. Deseja continuar?");
      if (!ok) return;
      localStorage.removeItem(STORAGE_KEY);
      load();
      render();
    });

    $("btnPrintMode").addEventListener("click", togglePrintMode);
    $("btnExitPrint").addEventListener("click", () => setPrintMode(false));

    // Filtros
    $("filterCategory").addEventListener("change", () => {
      state.filters.category = $("filterCategory").value || "";
      renderGrid();
      renderSummary();
    });

    $("filterOwner").addEventListener("input", () => {
      state.filters.owner = $("filterOwner").value || "";
      renderGrid();
      renderSummary();
    });

    $("filterSupplier").addEventListener("input", () => {
      state.filters.supplier = $("filterSupplier").value || "";
      renderGrid();
      renderSummary();
    });

    $("filterPeriod").addEventListener("change", () => {
      state.filters.period = $("filterPeriod").value || "";
      renderGrid();
      renderSummary();
    });

    $("filterText").addEventListener("input", () => {
      state.filters.text = $("filterText").value || "";
      renderGrid();
      renderSummary();
    });

    $("toggleOnlyPending").addEventListener("change", () => {
      state.filters.onlyPending = $("toggleOnlyPending").checked;
      renderGrid();
      renderSummary();
    });

    $("toggleOnlyApplicable").addEventListener("change", () => {
      state.filters.onlyApplicable = $("toggleOnlyApplicable").checked;
      renderGrid();
      renderSummary();
    });

    // Activity modal
    $("actDueType").addEventListener("change", syncDueFields);
    $("btnSaveActivity").addEventListener("click", saveActivityFromForm);
    $("btnDeleteActivity").addEventListener("click", deleteActivity);

    // Cell modal
    $("btnSaveCell").addEventListener("click", saveCellFromForm);
    $("cellDone").addEventListener("change", () => {
      $("cellDoneHint").textContent = $("cellDone").checked ? "Ao salvar, registra a data/hora." : "";
    });

    // Follow-up modal
    $("btnSaveFollow").addEventListener("click", saveFollowUp);
    $("btnCopyEmail").addEventListener("click", copyEmail);

    // Base online
    $("btnBaseOnline").addEventListener("click", openBaseOnline);
    $("btnCopyBaseOnlineHint").addEventListener("click", copyBaseOnlineInstructions);
    $("btnSaveBaseOnlineName").addEventListener("click", saveBaseOnlineNameFromModal);
    $("btnLoadPastedBase").addEventListener("click", loadPastedBase);

    // Fechar modais por ESC
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        ["modalActivity","modalCell","modalFollow","modalBaseOnline","modalImportPrevious"].forEach(id => {
          const el = $(id);
          if (el && !el.classList.contains("hidden")) closeModal(id);
        });
        if (state.ui.printMode) setPrintMode(false);
      }
    });

    // Recalcular offsets sticky ao redimensionar
    window.addEventListener("resize", () => {
      updateStickyOffsets();
    });

    // Scroll horizontal
    const sc = $("tableScroll");
    if (sc) {
      sc.addEventListener("scroll", () => {
        updateTopbarHeight();
      });
    }

    // Observa mudanças de layout na tabela
    const table = $("gridTable");
    if (table && "ResizeObserver" in window) {
      const ro = new ResizeObserver(() => updateStickyOffsets());
      ro.observe(table);
    }
  }

  function render() {
    ensureYearOptions();
    ensureMonthOptions();
    ensureCategoryFilterOptions();
    renderGrid();
    renderSummary();
    updateStickyOffsets();
  }

  function init() {
    load();
    wireModalClose();
    wireControls();
    render();
    requestAnimationFrame(() => updateTopbarHeight());
  }

  init();

})();
