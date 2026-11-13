/* =====================================================================
   DASHBOARD  (runs after app.js)
   Reads every <work order>/checklist.json in the results folder, joins it with the
   work orders in the Status file and shows both checklists with filters.
   Checklist 1 = template checklist, Checklist 2 = maintenance checks (unit status items).
   ===================================================================== */
const C_OK = "#4a7c1f", C_NOK = "#c0504d", C_PROG = "#b7791f", C_NONE = "#8f9488", C_DONE = "#2f6fb5";
const STATES = ["Needs attention", "In progress", "Not started", "Complete"];
const STATE_C = { "Needs attention": C_NOK, "In progress": C_PROG, "Not started": C_NONE, "Complete": C_OK };
const STATE_ORD = { "Needs attention": 0, "In progress": 1, "Not started": 2, "Complete": 3 };
const FL0 = { q: "", wst: "open", customer: "", product: "", engine: "", state: "", sort: "state" };
const cmpWo = (a, b) => a.wo.localeCompare(b.wo, undefined, { numeric: true });

/* ---- counts for one work order ---- */
function summarize(d, items) {
  const rows = (d && d.rows) || {}, maint = (d && d.maint) || {};
  const s = { n: items.length, ok: 0, nok: 0, notChk: 0, notRet: 0, full: 0, photos: 0,
    mTotal: d && d.maintTotal != null ? d.maintTotal : null, mOk: 0, mNok: 0, mProg: 0, mNone: null, started: false };
  items.forEach((t, i) => {
    const a = rows[val(t, TCOLS.title) || String(i)] || {};
    if (a.chk === "OK") s.ok++; else if (a.chk === "NOK") s.nok++; else s.notChk++;
    if (a.ret === "NO") s.notRet++;
    if (a.del && a.ret && a.chk) s.full++;
    s.photos += (a.retP || []).length + (a.chkP || []).length;
    if (a.del || a.ret || a.chk) s.started = true;
  });
  Object.values(maint).forEach(o => {
    const lab = maintStateOf(o)[0];
    if (lab === "OK") s.mOk++; else if (lab === "NOK") s.mNok++; else if (lab === "In progress") s.mProg++;
    MAINT.forEach((q, i) => { s.photos += (o["q" + i + "P"] || []).length; });
    if (Object.keys(o).length) s.started = true;
  });
  if (s.mTotal != null) s.mNone = Math.max(0, s.mTotal - s.mOk - s.mNok - s.mProg);
  const attention = s.nok || s.notRet || s.mNok;
  s.state = !s.started ? "Not started" : attention ? "Needs attention"
    : (s.n && s.full === s.n && s.mTotal != null && s.mOk === s.mTotal) ? "Complete" : "In progress";
  return s;
}

/* ---- load every result file ---- */
async function loadDash() {
  const D = S.dash = S.dash || { byKey: {}, loaded: false };
  if (D.loading) return;
  D.loading = true; D.done = 0; D.total = 0; D.err = ""; render();
  try {
    const list = [];
    if (demoMode) Store.mem.forEach((d, wo) => list.push([key(wo), d]));
    else {
      const f = await folder(), H = { headers: { Authorization: "Bearer " + await token() } }, dirs = [];
      let url = `${G}/drives/${f.drive}/items/${f.id}/children?$top=999&$select=id,name,folder`;
      while (url) {
        const r = await fetch(url, H);
        if (!r.ok) throw new Error("Could not read the results folder (" + r.status + ").");
        const j = await r.json(); dirs.push(...j.value.filter(x => x.folder)); url = j["@odata.nextLink"];
      }
      D.total = dirs.length; let i = 0;
      const worker = async () => {
        while (i < dirs.length) {
          const x = dirs[i++];
          try { const r = await fetch(`${G}/drives/${f.drive}/items/${x.id}:/checklist.json:/content`, H); if (r.ok) list.push([x.name, await r.json()]); } catch (e) { /* folder without a checklist */ }
          D.done++; const el = $("#dprog"); if (el) el.textContent = `Reading results: ${D.done} of ${D.total}`;
        }
      };
      await Promise.all(Array.from({ length: 6 }, worker));
    }
    D.byKey = Object.fromEntries(list); D.loaded = true; D.at = new Date();
  } catch (e) { D.err = e.message; }
  D.loading = false;
  if (S.screen === "dash") render();
}
function openDash() { S.fl = S.fl || { ...FL0 }; go("dash"); if (!S.dash || !S.dash.loaded) loadDash(); }

/* ---- one row per work order of the Status file ---- */
function dashRows() {
  const D = S.dash, seen = new Set(), memo = new Map(), out = [];
  const lines = name => { if (!memo.has(name)) memo.set(name, name ? itemsFor(name) : []); return memo.get(name); };
  for (const r of S.status) {
    const wo = val(r, COLS.wo); if (!wo || seen.has(wo)) continue; seen.add(wo);
    const d = D.byKey[key(wo)] || null, cands = templateNames(r);
    const template = (d && d.template) || (cands.length === 1 ? cands[0] : "");
    out.push({ wo, sn: val(r, COLS.sn), product: val(r, COLS.product), engine: val(r, COLS.engine), customer: val(r, COLS.customer),
      open: !hasStatusCol() || isOpen(r), template, by: (d && d.inspector) || "", at: (d && d.updated) || "", ...summarize(d, lines(template)) });
  }
  return out;
}
function applyFilters(rows, skipState) {
  const f = S.fl, q = f.q.trim().toLowerCase();
  return rows.filter(r => (f.wst === "all" || r.open) && (!f.customer || r.customer === f.customer) && (!f.product || r.product === f.product)
    && (!f.engine || r.engine === f.engine) && (skipState || !f.state || r.state === f.state)
    && (!q || (r.wo + " " + r.sn + " " + r.product + " " + r.template).toLowerCase().includes(q)));
}
const SORTS = {
  state: (a, b) => STATE_ORD[a.state] - STATE_ORD[b.state] || cmpWo(a, b),
  wo: cmpWo,
  product: (a, b) => a.product.localeCompare(b.product) || cmpWo(a, b),
  nok: (a, b) => (b.nok + b.notRet + b.mNok) - (a.nok + a.notRet + a.mNok) || cmpWo(a, b),
  progress: (a, b) => (b.n ? b.full / b.n : 0) - (a.n ? a.full / a.n : 0) || cmpWo(a, b),
  saved: (a, b) => String(b.at).localeCompare(String(a.at)) || cmpWo(a, b)
};

/* ---- small pieces ---- */
const pc = (n, t) => t ? Math.round(100 * n / t) : 0;
function stack(segs, aria) {
  const tot = segs.reduce((a, x) => a + x.n, 0);
  if (!tot) return `<p class="note0">No data yet.</p>`;
  return `<div class="stk" role="img" aria-label="${esc(aria)}">${segs.filter(x => x.n > 0).map(x => `<span style="flex:${x.n};background:${x.c}" title="${esc(x.l)}: ${x.n}">${x.n}</span>`).join("")}</div>
  <ul class="lg">${segs.map(x => `<li><i style="background:${x.c}"></i>${esc(x.l)} <b>${x.n}</b> <small>${pc(x.n, tot)}%</small></li>`).join("")}</ul>`;
}
function mini(r) {
  if (!r.n) return `<span class="note0">no template</span>`;
  return `<div class="mini" title="${r.ok} OK, ${r.nok} NOK, ${r.notChk} not checked"><span style="flex:${r.ok || 0.0001};background:${C_OK}"></span><span style="flex:${r.nok || 0.0001};background:${C_NOK}"></span><span style="flex:${r.notChk || 0.0001};background:#cfd3c8"></span></div><small>${r.ok} / ${r.n} OK${r.nok ? `, <b class="bad">${r.nok} NOK</b>` : ""}</small>`;
}

/* ---- page ---- */
function dashHTML() {
  const D = S.dash || {}, err = S.err ? `<div class="err" role="alert">${esc(S.err)}</div>` : "";
  if (D.err) return `<div class="err" role="alert">${esc(D.err)}</div><button class="btn" onclick="loadDash()">Try again</button>`;
  if (!D.loaded) return `<p class="sub" id="dprog">Reading results${D.total ? `: ${D.done} of ${D.total}` : "…"}</p>`;
  D.rows = dashRows();
  const f = S.fl, base = D.rows.filter(r => f.wst === "all" || r.open), opt = k => [...new Set(base.map(r => r[k]).filter(Boolean))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  const sel = (id, k, label, list) => `<label>${label}<select id="${id}" onchange="setFl('${k}',this.value)"><option value="">All</option>${list.map(v => `<option ${v === f[k] ? "selected" : ""}>${esc(v)}</option>`).join("")}</select></label>`;
  return `${err}<div class="dfilters">
    <label>Search<input id="fq" type="search" placeholder="Work order, S/N, product" value="${esc(f.q)}" oninput="setFl('q',this.value)"></label>
    ${hasStatusCol() ? `<label>Work orders<select onchange="S.fl.wst=this.value;render()"><option value="open" ${f.wst === "open" ? "selected" : ""}>Open only</option><option value="all" ${f.wst === "all" ? "selected" : ""}>All, including closed</option></select></label>` : ""}
    ${sel("fcustomer", "customer", "Customer", opt("customer"))}${sel("fproduct", "product", "Product", opt("product"))}${sel("fengine", "engine", "Aircraft / engine", opt("engine"))}
    ${sel("fstate", "state", "Progress", STATES)}
    <div class="dact"><button class="btn alt" onclick="resetDash()">Reset filters</button><button class="btn alt" onclick="S.dash.loaded=false;loadDash()">Refresh</button></div>
  </div><p class="prog" style="margin:-6px 0 12px">Results read ${D.at ? D.at.toLocaleTimeString() : ""}. Opening a work order from the table brings you back here.</p>
  <div id="dbody">${dashBody()}</div>`;
}
function dashBody() {
  const rows = applyFilters(S.dash.rows), f = S.fl, facets = applyFilters(S.dash.rows, true);
  const cnt = st => facets.filter(r => r.state === st).length;
  const T = rows.reduce((a, r) => { ["n", "ok", "nok", "notChk", "notRet", "full", "photos"].forEach(k => a[k] += r[k]);
    if (r.mTotal != null) { a.mWo++; ["mTotal", "mOk", "mNok", "mProg", "mNone"].forEach(k => a[k] += r[k]); } else a.mUnknown++; return a; },
    { n: 0, ok: 0, nok: 0, notChk: 0, notRet: 0, full: 0, photos: 0, mWo: 0, mTotal: 0, mOk: 0, mNok: 0, mProg: 0, mNone: 0, mUnknown: 0 });
  const tile = (label, v, c, st) => `<button class="kpi" style="--c:${c}" aria-pressed="${f.state === st}" onclick="setFl('state','${st}')"><b>${v}</b><span>${label}</span></button>`;
  const kpis = `<div class="kpis"><button class="kpi" style="--c:${C_DONE}" aria-pressed="${!f.state}" onclick="setFl('state','')"><b>${facets.length}</b><span>Work orders</span></button>
    ${tile("Not started", cnt("Not started"), C_NONE, "Not started")}${tile("In progress", cnt("In progress"), C_PROG, "In progress")}
    ${tile("Needs attention", cnt("Needs attention"), C_NOK, "Needs attention")}${tile("Complete", cnt("Complete"), C_OK, "Complete")}</div>`;
  const c1 = `<div class="pn"><h3>Checklist 1: template checklist</h3><p class="sub2">${T.n} lines in ${rows.length} work order${rows.length === 1 ? "" : "s"}. ${T.full} lines fully answered, ${T.notRet} not returned, ${T.photos} photos.</p>
    ${stack([{ n: T.ok, l: "Checked OK", c: C_OK }, { n: T.nok, l: "NOK", c: C_NOK }, { n: T.notChk, l: "Not checked", c: C_NONE }], "Template checklist lines by result")}</div>`;
  const c2 = `<div class="pn"><h3>Checklist 2: maintenance checks</h3><p class="sub2">${T.mTotal} unit status items in ${T.mWo} work order${T.mWo === 1 ? "" : "s"}.${T.mUnknown ? ` ${T.mUnknown} work order${T.mUnknown === 1 ? " has" : "s have"} not had unit status opened yet, so ${T.mUnknown === 1 ? "its" : "their"} items are not counted.` : ""}</p>
    ${stack([{ n: T.mOk, l: "OK", c: C_OK }, { n: T.mNok, l: "NOK", c: C_NOK }, { n: T.mProg, l: "In progress", c: C_PROG }, { n: T.mNone, l: "Not started", c: C_NONE }], "Maintenance check items by result")}</div>`;
  // by product (click to filter)
  const by = {}; rows.forEach(r => { (by[r.product || "(no product)"] ||= []).push(r); });
  const prod = Object.entries(by).sort((a, b) => b[1].length - a[1].length).map(([p, rs]) => {
    const c = Object.fromEntries(STATES.map(s => [s, rs.filter(r => r.state === s).length]));
    return `<li><button class="plink" onclick="setFl('product',${esc(JSON.stringify(p === "(no product)" ? "" : p)).replace(/"/g, "&quot;")})"><span class="pname">${esc(p)}</span><span class="pcount">${rs.length}</span>
      <div class="mini big">${STATES.map(s => `<span style="flex:${c[s] || 0.0001};background:${STATE_C[s]}" title="${s}: ${c[s]}"></span>`).join("")}</div>
      <small>${STATES.filter(s => c[s]).map(s => `${c[s]} ${s.toLowerCase()}`).join(", ")}</small></button></li>`;
  }).join("");
  const th = (k, l) => `<th><button class="thb" onclick="setFl('sort','${k}')" aria-label="Sort by ${l}">${l}${f.sort === k ? " ▾" : ""}</button></th>`;
  const list = rows.slice().sort(SORTS[f.sort] || SORTS.state);
  const table = list.length ? `<div class="dtw"><table class="dt"><thead><tr>${th("wo", "Work order")}${th("product", "Product")}<th>Engine</th><th>S/N</th>
    ${th("progress", "Checklist 1")}<th>Not returned</th><th>Checklist 2 (maintenance)</th>${th("state", "Status")}${th("saved", "Last saved")}</tr></thead><tbody>
    ${list.map(r => `<tr onclick="openFromDash(${esc(JSON.stringify(r.wo)).replace(/"/g, "&quot;")})"><td><button class="link strong">${esc(r.wo)}</button></td><td>${esc(r.product)}<br><small>${esc(r.template || "no template rule")}</small></td>
      <td>${esc(r.engine)}</td><td>${esc(r.sn)}</td><td class="mc1">${mini(r)}</td><td class="num">${r.notRet ? `<b class="bad">${r.notRet}</b>` : "0"}</td>
      <td>${r.mTotal == null ? `<small class="note0">not opened</small>` : `${r.mOk} / ${r.mTotal} OK${r.mNok ? `, <b class="bad">${r.mNok} NOK</b>` : ""}`}</td>
      <td><span class="stt" style="--c:${STATE_C[r.state]}">${r.state}</span></td><td><small>${r.at ? esc(fmt(r.at)) : "-"}${r.by ? `<br>${esc(r.by)}` : ""}</small></td></tr>`).join("")}</tbody></table></div>`
    : `<p class="sub">No work orders match these filters.</p>`;
  return `${kpis}<div class="dgrid">${c1}${c2}</div>
    <div class="dgrid"><div class="pn"><h3>By product</h3><p class="sub2">Select a product to filter everything on this page.</p><ul class="plist">${prod || `<li class="note0">No data</li>`}</ul></div>
    <div class="pn"><h3>How to read this</h3><p class="sub2" style="margin:0">A work order is <b>Needs attention</b> when any line is NOK, any line is not returned, or any maintenance item is NOK.
    It is <b>Complete</b> when every checklist line is answered and every maintenance item is OK. <b>Not started</b> means nothing has been filled in yet.
    Click a tile above, a product, or a column heading to filter or sort.</p></div></div>
    <h3 class="tt">${list.length} work order${list.length === 1 ? "" : "s"}</h3>${table}`;
}
function setFl(k, v) {
  S.fl[k] = v;
  if (k === "q" || !$("#dbody")) { if ($("#dbody")) $("#dbody").innerHTML = dashBody(); return; }
  render();                       // other filters: redraw so the drop-downs and tiles stay in step
}
function resetDash() { S.fl = { ...FL0, wst: S.fl.wst }; render(); }
function afterDashRender() { /* filters keep their own state; nothing to restore */ }
function openFromDash(wo) {
  selectWo(wo);
  if (!S.choice) { S.err = `No template rule for "${S.product}"${S.engine ? " with engine " + S.engine : ""}, so this work order can't be opened.`; return render(); }
  S.from = "dash"; openWo();
}

/* ---- demo results so the dashboard has something to show ---- */
function seedDemo() {
  if (Store.mem.size) return;
  const now = new Date().toISOString(), A = (d, r, c, extra) => ({ del: d, ret: r, chk: c, ...(extra || {}) });
  const M = (bad) => Object.fromEntries([0, 1, 2, 3, 5].map(i => ["q" + i, bad === i ? "NOK" : "OK"]));
  const mk = (wo, sn, product, template, rows, maint, total) => Store.mem.set(wo, { wo, sn, product, template, rows, maint: maint || {}, maintTotal: total ?? null, updated: now, inspector: "Demo user" });
  mk("B24-804", "FHH00RC5", "Fan Stator Module", "Fan Stator Module",
    { "FSM-02": A("YES", "YES", "OK"), "FSM-03": A("YES", "NO", "NOK", { retC: "Specification missing" }), "FSM-04": A("YES", "YES", "OK") },
    { IWS01: { ...M(1), q1C: "Stamp missing" } }, 3);
  mk("A24-847", "ACL01003", "NACELLE- Thrust Reverser", "Thrust Reverser (A320neo)", { "TRA-01": A("YES", "YES", "OK") }, { IWS01: M(-1) }, 1);
  mk("A25-931", "TR02024", "NACELLE- Thrust Reverser", "Thrust Reverser (FEC|TR)", { "TR-01": A("YES", "YES", "OK"), "TR-02": A("YES", "", "") });
  mk("A25-016", "NOZ28547", "NACELLE- Exhaust Nozzle", "Exhaust Nozzle", { "NZ-01": A("N/A", "NO", "NOK") });
}
