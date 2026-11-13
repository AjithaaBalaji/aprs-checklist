/* =====================================================================
   APRS CHECKLIST
   1. Edit the column names below so they match your Excel files.
   2. Fill in Settings once (Azure client id + 3 SharePoint links).
   ===================================================================== */

// STATUS file (one row per work-order item). Header names, any capitalisation.
const COLS = {
  wo: "WO", sn: "S/N", product: "Product", engine: "Aircraft/Engine", customer: "Customer",
  // The 9 columns shown on the Unit Status screen (read from the per-WO unit status file). The FIRST one is the item id (IWS01...)
  status: ["Item","Item Status","Location","Finding","Corrective Action","JC Reference","JC Status","NDT","Shop Status"]
};
// TEMPLATE list (exported to Excel). Only rows where Record Type = Template are used.
// Only work orders whose Status column equals one of these values are listed (a "Show all" tick box overrides it)
const OPEN_RULE = { col: "Status", values: ["open"] };

// WHICH TEMPLATE A WORK ORDER USES: [status product, engine condition, template product name]
// Engine condition (read from the Aircraft/Engine column): "" = any engine, "leap" = LEAP1-A, "t700" = T700 / TRENT 700,
// "other-not-leap-t700" = anything except Leap and T700, "other-not-leap" = anything except Leap.
const TEMPLATE_RULES = [
  ["NACELLE- Air inlet", "", "Air Intake Cowl"],
  ["NACELLE- Fan Cowl", "", "Fan Cowl"],
  ["NACELLE- Thrust Reverser", "leap", "Thrust Reverser (A320neo)"],
  ["NACELLE- Thrust Reverser", "t700", "Thrust Reverser (T700)"],
  ["NACELLE- Thrust Reverser", "other-not-leap-t700", "Thrust Reverser (FEC|TR)"],
  ["NACELLE - Lower Pivoting Door", "", "Door"],
  ["NACELLE- Upper Pivoting Door", "", "Door"],
  ["NACELLE- Fan Exhaust Cowl", "", "Thrust Reverser (FEC|TR)"],
  ["NACELLE- Translating Cowl", "", "Thrust Reverser (FEC|TR)"],
  ["NACELLE- Exhaust Nozzle", "other-not-leap", "Exhaust Nozzle"],
  ["NACELLE- Exhaust Nozzle", "leap", "Exhaust Nozzle (A320neo)"],
  ["NACELLE - FWD Plug", "leap", "Plug Assy"],
  ["NACELLE - FWD Plug", "other-not-leap", "Exhaust AFT Plug"],
  ["NACELLE - AFT Plug", "leap", "Plug Assy"],
  ["NACELLE - AFT Plug", "other-not-leap", "Exhaust AFT Plug"],
  ["RADOME - Radome", "", "Radome"],
  ["FSM- Fan Stator Module", "", "Fan Stator Module"],
  ["FSM- Forward Fan Case", "", "Forward Fan Case"],
  ["FSM- Acoustic Panels", "", "Forward Fan Case"],
  ["LRU- Remote release system", "", "Check LRU - Remote Release System"],
  ["LRU- HOR", "", "Check LRU - Remote Release System"],
  ["LRU-Hinge Access Panel", "", "Check LRU - Remote Release System"],
  ["LRU- Closure device", "", "Check LRU - Remote Release System"]
];
const TCOLS = { title: "Title", product: "Product", type: "Record Type", seq: "Sequence", step: "Step Label", item: "Checklist Item", form: "Form", folder: "Folder" };

// Questions in the Maintenance Check popup. text:true = free text answer
const MAINT = [
  { t: "JC Check (sign, curing, stamp, date...)" },
  { t: "JC work performed vs unit status vs SFR" },
  { t: "Check Part List VS the JC (work performed). Part list (check that ARC if ap. is present). Check qty issued & expired part not used" },
  { t: "SFR (CA + doc used+)" },
  { t: "List document used - REF + REV", text: true },
  { t: "Check document revision - Impact or not?" }
];

const SCOPES = ["User.Read", "Files.ReadWrite.All", "Sites.Read.All"];
const G = "https://graph.microsoft.com/v1.0";

/* ---------- demo data (Try demo button) ---------- */
const U = (wo, item, st, loc, fnd, ca, jc, jcs, ndt, shop) => ({ WO: wo, Item: item, "Item Status": st, Location: loc, Finding: fnd, "Corrective Action": ca, "JC Reference": jc, "JC Status": jcs, NDT: ndt, "Shop Status": shop });
const DEMO_STATUS = [
  { Status:"OPEN", "Aircraft/Engine":"GE90-115", WO:"B24-804", Product:"Fan Stator Module", "Product code":"FSM", "S/N":"FHH00RC5", ...U("B24-804","IWS01","APPROVED","FSM","DISASSEMBLY","DISASSEMBLY OF THE FAN STATOR ASSEMBLY","CT-RC-72-24-00-STD-DIS-001-07 Rev 05","ARCHIVED","NA","DONE") },
  { Status:"OPEN", "Aircraft/Engine":"GE90-115", WO:"B24-804", Product:"Fan Stator Module", "Product code":"FSM", "S/N":"FHH00RC5", ...U("B24-804","IWS02","APPROVED","FSM","HANDLING","FSM TRANSFERS","RC-72-00-06-STD-STO-001 Rev 04","ARCHIVED","NA","DONE") },
  { Status:"OPEN", "Aircraft/Engine":"GE90-115", WO:"B24-804", Product:"Fan Stator Module", "Product code":"FSM", "S/N":"FHH00RC5", ...U("B24-804","IWS03","APPROVED","FSM","WORKSCOPE","FSM CLEANING","RC-70-21-24-STD-CLE-024-01 Rev 03","OPEN","NA","PENDING") },
  { Status:"OPEN", "Aircraft/Engine":"LEAP-1A", WO:"A24-847", Product:"NACELLE- Thrust Reverser", "Product code":"TR", "S/N":"ACL01003", ...U("A24-847","IWS01","APPROVED","TR","INSPECTION","INITIAL INSPECTION","GF90-115-FSM-GVI-001","ARCHIVED","NA","DONE") }
];
DEMO_STATUS.push({ Status:"OPEN", "Aircraft/Engine":"TRENT 700", WO:"A25-990", Product:"NACELLE- Thrust Reverser", "S/N":"TR791", Item:"IWS01" },
  { Status:"OPEN", "Aircraft/Engine":"GP7200", WO:"A25-931", Product:"NACELLE- Thrust Reverser", "S/N":"TR02024", Item:"IWS01" },
  { Status:"OPEN", "Aircraft/Engine":"GP7200", WO:"A25-016", Product:"NACELLE- Exhaust Nozzle", "S/N":"NOZ28547", Item:"IWS01" },
  { Status:"OPEN", "Aircraft/Engine":"LEAP-1A", WO:"A25-017", Product:"NACELLE - FWD Plug", "S/N":"P1", Item:"IWS01" });
DEMO_STATUS.push({ Status:"CLOSED", WO:"A20-001", Product:"Fan Stator Module", "S/N":"OLD1", Item:"IWS01" });
DEMO_STATUS.forEach((r, i) => r.Customer = ["UAE", "AFI (UAE)", "SNA (ETD)"][i % 3]);
const DEMO_TPL = [
  { Title:"TRT-01", Product:"Thrust Reverser (T700)", "Record Type":"Template", Sequence:1, "Step Label":"STEP 1", "Checklist Item":"PO customer T700", Form:"", Folder:"A" },
  { Title:"NZ-01", Product:"Exhaust Nozzle", "Record Type":"Template", Sequence:1, "Step Label":"STEP 1", "Checklist Item":"PO nozzle", Form:"", Folder:"A" },
  { Title:"PLG-01", Product:"Plug Assy", "Record Type":"Template", Sequence:1, "Step Label":"STEP 1", "Checklist Item":"PO plug", Form:"", Folder:"A" },
  { Title:"FSM-02", Product:"Fan Stator Module", "Record Type":"Template", Sequence:2, "Step Label":"STEP 1", "Checklist Item":"PO AFI", Form:"", Folder:"A" },
  { Title:"FSM-03", Product:"Fan Stator Module", "Record Type":"Template", Sequence:3, "Step Label":"STEP 1", "Checklist Item":"Technical Specifications AFI/ Worscope GE", Form:"", Folder:"A" },
  { Title:"FSM-04", Product:"Fan Stator Module", "Record Type":"Template", Sequence:4, "Step Label":"STEP 2", "Checklist Item":"PRELIMINARY INSPECTION JOB CARD /Initial Workpackage", Form:"BQG8010-Form9", Folder:"C" },
  { Title:"FSM-05", Product:"Fan Stator Module", "Record Type":"Template", Sequence:5, "Step Label":"STEP 2", "Checklist Item":"Incoming Inspection check list", Form:"BQG8009 Form2", Folder:"C" },
  { Title:"FSM-07", Product:"Fan Stator Module", "Record Type":"Template", Sequence:7, "Step Label":"STEP 2", "Checklist Item":"Unit Status", Form:"BQG8010-Form3", Folder:"B" },
  { Title:"FSM-99", Product:"Fan Stator Module", "Record Type":"Result", Sequence:99, "Step Label":"STEP 9", "Checklist Item":"Should not appear", Form:"", Folder:"" },
  { Title:"TR-01", Product:"Thrust Reverser (FEC|TR)", "Record Type":"Template", Sequence:1, "Step Label":"STEP 1", "Checklist Item":"PO customer", Form:"", Folder:"A" },
  { Title:"TRA-01", Product:"Thrust Reverser (A320neo)", "Record Type":"Template", Sequence:1, "Step Label":"STEP 1", "Checklist Item":"PO SNA", Form:"", Folder:"A" },
  { Title:"TR-02", Product:"Thrust Reverser (FEC|TR)", "Record Type":"Template", Sequence:2, "Step Label":"STEP 2", "Checklist Item":"Incoming Inspection check list", Form:"BQG8009 Form2", Folder:"C" }
];

/* ---------- helpers ---------- */
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[c]));
const norm = s => String(s ?? "").trim().toLowerCase();
const lower = rows => rows.map(r => Object.fromEntries(Object.entries(r).map(([k, v]) => [norm(k), String(v ?? "").trim()])));
const val = (r, c) => r[norm(c)] ?? "";
const fmt = iso => new Date(iso).toLocaleString();
const sid = id => String(id).replace(/['"\\]/g, "");

const S = {
  screen: "start", user: "", status: [], tpl: [], wo: "", sn: "", product: "",
  showAll: false, from: "", dash: null, engine: "", items: [], unit: [], unitState: "idle", unitErr: "", unitFile: "", cands: [], choice: "", row: null, tplProduct: "", data: null, saved: "", err: "", modal: null, dirty: false,
  cfg: Object.assign({ clientId:"", tenant:"common", statusUrl:"", tplUrl:"", folderUrl:"", statusSheet:"", tplSheet:"", unitUrl:"", unitSheet:"" },
        JSON.parse(localStorage.getItem("aprs_cfg") || "{}"))
};
let demoMode = false, msalApp = null, folderRef = null, saveTimer = null, toastTimer = null;
const photoCache = {};

function note(msg, ms = 0) {
  const t = $("#toast"); clearTimeout(toastTimer);
  if (!msg) { t.hidden = true; return; }
  t.textContent = msg; t.hidden = false;
  if (ms) toastTimer = setTimeout(() => t.hidden = true, ms);
}

/* ---------- Microsoft sign-in + Graph ---------- */
async function signIn() {
  const c = S.cfg;
  if (!c.clientId) throw new Error("Open Settings and add the Azure Client ID first.");
  if (!msalApp) {
    msalApp = new msal.PublicClientApplication({
      auth: { clientId: c.clientId, authority: "https://login.microsoftonline.com/" + (c.tenant || "common"), redirectUri: location.origin + location.pathname },
      cache: { cacheLocation: "localStorage" }
    });
    await msalApp.initialize?.();
  }
  let acc = msalApp.getAllAccounts()[0];
  if (!acc) acc = (await msalApp.loginPopup({ scopes: SCOPES })).account;
  S.user = acc.name || acc.username;
}
async function token() {
  await signIn();
  const account = msalApp.getAllAccounts()[0];
  try { return (await msalApp.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
  catch (e) { return (await msalApp.acquireTokenPopup({ scopes: SCOPES })).accessToken; }
}
async function gfetch(url, opt = {}) {
  const t = await token();
  return fetch(url, { ...opt, headers: { Authorization: "Bearer " + t, ...(opt.headers || {}) } });
}
const enc = u => "u!" + btoa(unescape(encodeURIComponent(u.trim()))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
async function readWb(url) {
  const r = await gfetch(G + "/shares/" + enc(url) + "/driveItem/content");
  if (!r.ok) throw new Error("Could not open the Excel file (" + r.status + "). Check the link and your access.");
  return XLSX.read(await r.arrayBuffer(), { type: "array" });
}
const sheetOf = (wb, name) => XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames.includes(name) ? name : wb.SheetNames[0]], { defval: "" });
async function refOf(url, what) {
  const r = await gfetch(G + "/shares/" + enc(url) + "/driveItem");
  if (!r.ok) throw new Error(`Could not open the ${what} (${r.status}). Check the link in Settings and your access.`);
  const j = await r.json();
  return { drive: j.parentReference.driveId, id: j.id };
}
async function folder() { return folderRef ||= await refOf(S.cfg.folderUrl, "results folder"); }
const pth = p => p.split("/").map(encodeURIComponent).join("/");
async function put(path, body, type) {
  const f = await folder();
  const r = await gfetch(`${G}/drives/${f.drive}/items/${f.id}:/${pth(path)}:/content`, { method: "PUT", headers: { "Content-Type": type }, body });
  if (!r.ok) throw new Error("Save failed (" + r.status + ")");
}
async function get(path) {
  const f = await folder();
  const r = await gfetch(`${G}/drives/${f.drive}/items/${f.id}:/${pth(path)}:/content`);
  if (r.status === 404) return null;
  if (!r.ok) throw new Error("Read failed (" + r.status + ")");
  return r;
}

/* ---------- storage (SharePoint folder, or memory in demo mode) ---------- */
const key = wo => String(wo).replace(/[\\/:*?"<>|#%]/g, "_");
const Store = {
  mem: new Map(),
  async loadJson(wo) {
    if (demoMode) return this.mem.get(wo) || null;
    const r = await get(key(wo) + "/checklist.json");
    return r ? await r.json() : null;
  },
  async saveJson(wo, d) {
    if (demoMode) { this.mem.set(wo, JSON.parse(JSON.stringify(d))); return; }
    await put(key(wo) + "/checklist.json", JSON.stringify(d), "application/json");
  },
  async putPhoto(wo, blob) {
    const name = "photos/" + Date.now() + "_" + Math.random().toString(36).slice(2, 6) + ".jpg";
    if (!demoMode) await put(key(wo) + "/" + name, blob, "image/jpeg");
    photoCache[name] = URL.createObjectURL(blob);
    return name;
  },
  async photoUrl(wo, name) {
    if (photoCache[name]) return photoCache[name];
    if (demoMode) return "";
    const r = await get(key(wo) + "/" + name);
    return r ? (photoCache[name] = URL.createObjectURL(await r.blob())) : "";
  }
};

/* ---------- load data / open work order ---------- */
async function loadAll(demo) {
  S.err = ""; demoMode = !!demo; render(); note("Loading data…");
  try {
    if (demoMode) { S.status = lower(DEMO_STATUS); S.tpl = lower(DEMO_TPL); S.user = "Demo user"; seedDemo(); }
    else {
      const c = S.cfg;
      if (!c.statusUrl || !c.tplUrl || !c.unitUrl || !c.folderUrl) throw new Error("Open Settings and fill in all four SharePoint links.");
      await signIn();
      const [a, b] = await Promise.all([readWb(c.statusUrl), readWb(c.tplUrl)]);
      S.status = lower(sheetOf(a, c.statusSheet));
      S.tpl = lower(sheetOf(b, c.tplSheet));
      unitIndexP = null; getIndex().catch(() => {});   // list unit status files in the background
    }
    note(""); go("wo");
  } catch (e) { note(""); S.err = e.message; render(); }
}
const isOpen = r => OPEN_RULE.values.includes(norm(val(r, OPEN_RULE.col)));
const hasStatusCol = () => S.status.some(r => val(r, OPEN_RULE.col));
const woList = () => {
  const rows = S.showAll || !hasStatusCol() ? S.status : S.status.filter(isOpen);
  return [...new Set(rows.map(r => val(r, COLS.wo)).filter(Boolean))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
};
function refreshWoList() {
  const l = woList();
  $("#wos").innerHTML = l.map(w => `<option value="${esc(w)}">`).join("");
  $("#wocount").textContent = l.length + (S.showAll || !hasStatusCol() ? " work orders" : " open work orders");
}
const pnorm = s => norm(s).replace(/\s+/g, " ");
const pstrip = s => pnorm(s).replace(/^[a-z]+\s*-\s*/, "");            // "FSM- Fan Stator Module" -> "fan stator module"
const templateRows = () => S.tpl.filter(t => !val(t, TCOLS.type) || norm(val(t, TCOLS.type)) === "template");
/* Which template does this work order use? Status product + Aircraft/Engine -> TEMPLATE_RULES. */
const engineKind = e => {
  const x = String(e || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  return x.includes("LEAP") ? "leap" : (x.includes("T700") || x.includes("TRENT700")) ? "t700" : "other";
};
const ENGINE_OK = {
  "": () => true, leap: k => k === "leap", t700: k => k === "t700",
  "other-not-leap-t700": k => k === "other", "other-not-leap": k => k !== "leap"
};
function templateNames(row) {
  const sp = pstrip(val(row, COLS.product)), kind = engineKind(val(row, COLS.engine));
  let names = TEMPLATE_RULES.filter(([p, e]) => pstrip(p) === sp && ENGINE_OK[e](kind)).map(r => r[2].trim());
  if (!names.length) names = [...new Set(templateRows().map(t => val(t, TCOLS.product)))].filter(n => pnorm(n) === pnorm(sp));
  return [...new Set(names)];
}
function tplBox() {
  const lab = "<small>Checklist template</small>";
  if (!S.row) return lab + "<b>-</b>";
  const eng = S.engine ? `<small style="margin-top:4px">Chosen from Aircraft/Engine: ${esc(S.engine)}</small>` : `<small style="margin-top:4px">Aircraft/Engine is empty in the status file</small>`;
  if (!S.cands.length) return lab + `<b style="color:var(--bad)">No template rule for "${esc(S.product)}"${S.engine ? " with engine " + esc(S.engine) : ""}.</b>`;
  if (S.cands.length === 1) return lab + `<b>${esc(S.cands[0])}</b>` + eng;
  return lab + `<select id="tplsel" onchange="S.choice=this.value;$('#go').disabled=!S.choice"><option value="">Select template…</option>${S.cands.map(n => `<option ${n === S.choice ? "selected" : ""}>${esc(n)}</option>`).join("")}</select>` + eng;
}
function selectWo(wo) {
  const r = S.status.find(r => val(r, COLS.wo) === String(wo).trim());
  S.row = r || null;
  S.wo = r ? String(wo).trim() : ""; S.sn = r ? val(r, COLS.sn) : ""; S.product = r ? val(r, COLS.product) : ""; S.engine = r ? val(r, COLS.engine) : "";
  S.cands = r ? templateNames(r) : []; S.choice = S.cands.length === 1 ? S.cands[0] : "";
}
function pickWo(v) {
  selectWo(v);
  $("#sn").textContent = S.sn || "-"; $("#prod").textContent = S.product || "-"; $("#tpl").innerHTML = tplBox();
  $("#go").disabled = !(S.wo && S.choice);
}
/* checklist lines of one template, in Sequence order */
const itemsFor = name => templateRows().filter(t => pnorm(val(t, TCOLS.product)) === pnorm(name))
  .sort((a, b) => (parseFloat(val(a, TCOLS.seq)) || 0) - (parseFloat(val(b, TCOLS.seq)) || 0));
const rk = i => val(S.items[i], TCOLS.title) || String(i);
async function openWo() {
  const items = itemsFor(S.choice);
  if (!items.length) { S.err = `No checklist lines found for template "${S.choice}". Check the Product name in the template list matches the Product list sheet.`; return render(); }
  S.err = ""; S.items = items; S.tplProduct = S.choice;
  S.unit = []; startUnit();                       // loads in the background while the checklist is open
  note("Opening " + S.wo + "…");
  try {
    const d = await Store.loadJson(S.wo);
    S.data = d || { wo: S.wo, sn: S.sn, product: S.product, rows: {}, maint: {} };
    S.data.rows ||= {}; S.data.maint ||= {}; S.data.template = S.choice; S.data.sn = S.sn; S.data.product = S.product;
    syncMaintTotal();
    S.saved = d?.updated ? "Last saved: " + fmt(d.updated) : "Not saved yet";
    note(""); go("check");
  } catch (e) { note(""); S.err = e.message; render(); }
}

/* ---------- unit status: one Excel file per work order in a SharePoint library ---------- */
let unitIndexP = null;
async function buildIndex() {
  const links = (S.cfg.unitUrl || "").split(/[\s,]+/).filter(u => /^https?:/i.test(u)), out = [];
  if (!links.length) throw new Error("Add the unit status folder links in Settings.");
  async function walk(drive, id, depth) {
    let url = `${G}/drives/${drive}/items/${id}/children?$top=999&$select=id,name,folder,file,lastModifiedDateTime`;
    while (url) {
      const r = await gfetch(url);
      if (!r.ok) throw new Error("Could not list the unit status files (" + r.status + ").");
      const j = await r.json();
      for (const it of j.value) {
        if (it.folder) { if (depth < 3) await walk(drive, it.id, depth + 1); }
        else if (/\.xls[xm]?$/i.test(it.name)) out.push({ id: it.id, drive, name: it.name, mod: it.lastModifiedDateTime || "" });
      }
      url = j["@odata.nextLink"];
    }
  }
  for (const u of links) { const ref = await refOf(u, "unit status folder"); await walk(ref.drive, ref.id, 0); }
  return out;
}
function getIndex() { return unitIndexP ||= buildIndex().catch(e => { unitIndexP = null; throw e; }); }
const UNIT_MAP = [
  ["Item", h => h === "item #" || h === "item" || h === "item#"],
  ["Item Status", h => h === "item status"],
  ["Location", h => h.startsWith("locat")],
  ["Finding", h => h.startsWith("findin")],
  ["Corrective Action", h => h.startsWith("corrective")],
  ["JC Reference", h => h.includes("jc reference")],
  ["JC Status", h => h.includes("jc status")],
  ["NDT", h => h === "ndt"],
  ["Shop Status", h => h.includes("shop status")]
];
function parseUnit(aoa) {
  const clean = c => String(c ?? "").replace(/\s+/g, " ").trim().toLowerCase();
  let hr = -1, cols = {};
  for (let r = 0; r < Math.min(aoa.length, 40) && hr < 0; r++) {
    const hs = aoa[r].map(clean), found = {};
    UNIT_MAP.forEach(([label, fn]) => { const c = hs.findIndex(fn); if (c >= 0) found[label] = c; });
    if ("Item" in found && "JC Status" in found) { hr = r; cols = found; }
  }
  if (hr < 0) throw new Error("Could not find the header row (Item # ... JC Status) in the unit status file.");
  const rows = [];
  for (let r = hr + 1; r < aoa.length; r++) {
    const row = {};
    UNIT_MAP.forEach(([l]) => row[norm(l)] = l in cols ? String(aoa[r][cols[l]] ?? "").trim() : "");
    if (row[norm("Item")]) rows.push(row);
  }
  return rows;
}
async function startUnit() {
  const wo = S.wo; S.unitState = "loading"; S.unitErr = ""; S.unitFile = "";
  try {
    let rows, file = "";
    if (demoMode) rows = S.status.filter(r => val(r, COLS.wo) === wo);
    else {
      const idx = await getIndex(), re = new RegExp("^" + wo.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "(?![0-9A-Za-z])", "i");
      const hits = idx.filter(f => re.test(f.name) && !f.name.startsWith("~$")).sort((a, b) => b.mod.localeCompare(a.mod));
      if (!hits.length) throw new Error(`No unit status file found for ${wo}. The file name should start with the work order, for example "${wo} Unit Status.xlsx".`);
      const f = hits[0], r = await gfetch(`${G}/drives/${f.drive}/items/${f.id}/content`);
      if (!r.ok) throw new Error("Could not open " + f.name + " (" + r.status + ").");
      const wb = XLSX.read(await r.arrayBuffer(), { type: "array" });
      const sn = wb.SheetNames.find(n => /unit status/i.test(n)) || wb.SheetNames[0];
      rows = parseUnit(XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "" }));
      file = f.name + (hits.length > 1 ? ` (latest of ${hits.length} matching files)` : "");
    }
    if (wo !== S.wo) return;
    S.unit = rows; S.unitFile = file; S.unitState = "ready"; syncMaintTotal();
  } catch (e) { if (wo !== S.wo) return; S.unitState = "error"; S.unitErr = e.message; }
  if (S.screen === "unit") render();
}
const reportBox = () => `<div class="rep"><div class="repbtns"><button class="btn" onclick="makePdf()">Generate PDF report</button>
  <button class="btn alt" onclick="makeReport()">Word report</button></div>
  <p class="prog">Creates a new file each time and saves it in the results folder under ${esc(S.wo)} / reports.</p></div>`;
function retryUnit() { unitIndexP = null; startUnit(); render(); }

/* ---------- editing + autosave ---------- */
function obj(scope, id) { const d = S.data[scope]; return d[id] = d[id] || {}; }
function touch() { S.dirty = true; setSaved("Saving…"); clearTimeout(saveTimer); saveTimer = setTimeout(save, 1200); }
function setSaved(t) { S.saved = t; const el = $("#saved"); if (el) el.textContent = t; }
async function save() {
  clearTimeout(saveTimer);
  if (!S.data || !S.dirty) return;
  S.dirty = false; S.data.updated = new Date().toISOString(); S.data.inspector = S.user;
  try { await Store.saveJson(S.wo, S.data); if (S.dash) S.dash.byKey[key(S.wo)] = JSON.parse(JSON.stringify(S.data)); setSaved("Last saved: " + fmt(S.data.updated) + (demoMode ? " (demo, not stored)" : "")); }
  catch (e) { S.dirty = true; setSaved("Not saved: " + e.message); }
}
function setF(scope, id, f, v) { obj(scope, id)[f] = v; touch(); }
function choose(scope, id, f, v) {
  const o = obj(scope, id); o[f] = o[f] === v ? "" : v; touch(); refresh(scope, id);
}
function refresh(scope, id) {
  if (scope === "rows") {
    const i = S.items.findIndex((_, x) => rk(x) === id), el = $("#row-" + i);
    if (el) { el.outerHTML = rowHTML(i); hydrate($("#row-" + i)); }
    const p = $("#prog"); if (p) p.textContent = progText();
  } else renderModal();
}

/* ---------- photos ---------- */
function compress(file) {
  return new Promise((res, rej) => {
    const im = new Image(), u = URL.createObjectURL(file);
    im.onload = () => {
      const k = Math.min(1, 1600 / Math.max(im.width, im.height)), c = document.createElement("canvas");
      c.width = Math.round(im.width * k); c.height = Math.round(im.height * k);
      c.getContext("2d").drawImage(im, 0, 0, c.width, c.height); URL.revokeObjectURL(u);
      c.toBlob(b => b ? res(b) : rej(new Error("Could not read photo")), "image/jpeg", .75);
    };
    im.onerror = () => rej(new Error("Could not read photo"));
    im.src = u;
  });
}
async function addPhotos(scope, id, f, inp) {
  const files = [...inp.files]; inp.value = "";
  if (!files.length) return;
  note("Uploading " + files.length + " photo" + (files.length > 1 ? "s" : "") + "…");
  try {
    const o = obj(scope, id), L = o[f + "P"] = o[f + "P"] || [];
    for (const file of files) L.push(await Store.putPhoto(S.wo, await compress(file)));
    touch(); note("");
  } catch (e) { note("Photo upload failed: " + e.message, 6000); }
  refresh(scope, id);
}
function delPhoto(scope, id, f, j) { obj(scope, id)[f + "P"].splice(j, 1); touch(); refresh(scope, id); }
function viewPhoto(img) { if (!img.src) return; const v = $("#viewer"); v.querySelector("img").src = img.src; v.hidden = false; }
function hydrate(root = document) {
  root.querySelectorAll("img[data-ph]").forEach(im => {
    if (im.getAttribute("src")) return;
    Store.photoUrl(S.wo, im.dataset.ph).then(u => { if (u) im.src = u; }).catch(() => {});
  });
}

/* ---------- html builders ---------- */
const GOOD = ["YES", "OK"];
function chips(scope, id, f, opts, cur) {
  const a = sid(id);
  return opts.map(v => `<button type="button" class="chip ${cur === v ? "on " + (GOOD.includes(v) ? "good" : "bad") : ""}" aria-pressed="${cur === v}" onclick="choose('${scope}','${a}','${f}','${v}')">${v}</button>`).join("");
}
function evid(scope, id, f, label) {
  const o = (S.data[scope] || {})[id] || {}, a = sid(id);
  return `<div class="evid"><textarea rows="2" placeholder="${label}" oninput="setF('${scope}','${a}','${f}C',this.value)">${esc(o[f + "C"] || "")}</textarea>
  <div class="phs">${(o[f + "P"] || []).map((n, j) => `<div class="ph"><img alt="Photo ${j + 1}" data-ph="${esc(n)}" onclick="viewPhoto(this)"><button type="button" aria-label="Remove photo" onclick="delPhoto('${scope}','${a}','${f}',${j})">×</button></div>`).join("")}
  <label class="add" tabindex="0">+ Add photos<input type="file" accept="image/*" multiple hidden onchange="addPhotos('${scope}','${a}','${f}',this)"></label></div></div>`;
}
function rowHTML(i) {
  const t = S.items[i], k = rk(i), d = S.data.rows[k] || {};
  const form = val(t, TCOLS.form), fo = val(t, TCOLS.folder), item = val(t, TCOLS.item);
  const step = val(t, TCOLS.step), prev = i > 0 ? val(S.items[i - 1], TCOLS.step) : null;
  const st = d.chk === "NOK" || d.ret === "NO" ? "bad" : (d.del && d.ret && d.chk === "OK" ? "done" : "");
  const isUnit = norm(item) === "unit status";
  return `<div class="row ${st}" id="row-${i}">
    <div class="step">${step !== prev ? esc(step) : ""}</div>
    <div class="what"><b>${esc(item)}</b>${form ? `<small>Form ${esc(form)}</small>` : ""}${fo ? `<small>Folder ${esc(fo)}</small>` : ""}
      ${isUnit ? `<button type="button" class="cb" style="margin-top:8px" onclick="go('unit')">Open unit status →</button>` : ""}</div>
    <div class="grp"><span>Delivered</span><div>${chips("rows", k, "del", ["YES", "N/A"], d.del)}</div></div>
    <div class="grp"><span>Returned</span><div>${chips("rows", k, "ret", ["YES", "NO"], d.ret)}</div></div>
    <div class="grp"><span>Checked</span><div>${chips("rows", k, "chk", ["OK", "NOK"], d.chk)}</div></div>
    ${d.ret === "NO" ? `<div class="ev"><h4>Not returned: add a comment or photos</h4>${evid("rows", k, "ret", "Comment")}</div>` : ""}
    ${d.chk === "NOK" ? `<div class="ev"><h4>Not OK: add a comment or photos</h4>${evid("rows", k, "chk", "Comment")}</div>` : ""}
  </div>`;
}
const progText = () => {
  const n = S.items.filter((_, i) => (S.data.rows[rk(i)] || {}).chk).length;
  return n + " of " + S.items.length + " lines checked";
};
function stClass(v) {
  v = norm(v);
  if (["archived", "done", "closed", "complete", "completed"].includes(v)) return "good";
  if (["na", "n/a"].includes(v)) return "na";
  if (["to do", "todo"].includes(v)) return "todo";
  if (v.startsWith("issued")) return "mid";
  if (["open", "pending", "in progress", "wip"].includes(v)) return "mid";
  return "";
}
function maintStateOf(o) {
  o = o || {};
  const qs = MAINT.map((q, i) => q.text ? null : o["q" + i]).filter((_, i) => !MAINT[i].text);
  const n = qs.filter(Boolean).length;
  if (qs.includes("NOK")) return ["NOK", "bad"];
  if (n === qs.length) return ["OK", "good"];
  return n ? ["In progress", "mid"] : ["Check", ""];
}
const maintState = id => maintStateOf(S.data.maint[id]);
/* remember how many unit status items this work order has, so the dashboard can count them without opening the file */
function syncMaintTotal() {
  if (S.data && S.unitState === "ready" && S.data.maintTotal !== S.unit.length) { S.data.maintTotal = S.unit.length; if (S.screen === "check" || S.screen === "unit") touch(); else S.dirty = true; }
}
function unitHTML() {
  if (S.unitState === "loading") return `<p class="sub">Loading unit status for ${esc(S.wo)}…</p>`;
  if (S.unitState === "error") return `<div class="err" role="alert">${esc(S.unitErr)}</div><button class="btn" onclick="retryUnit()">Try again</button>`;
  if (!S.unit.length) return `<p class="sub">No unit status rows found for ${esc(S.wo)}.</p>`;
  const th = COLS.status.map(c => `<th>${esc(c)}</th>`).join("") + "<th>Maintenance check</th>";
  const tr = S.unit.map((r, k) => {
    const id = val(r, COLS.status[0]), [lab, cl] = maintState(id);
    return "<tr>" + COLS.status.map((c, ci) => {
      const v = val(r, c);
      return `<td class="${ci >= 6 ? "st " + stClass(v) : ""}">${esc(v)}</td>`;
    }).join("") + `<td><button class="cb ${cl}" onclick="openMaint(${k})">${lab}</button></td></tr>`;
  }).join("");
  return `${S.unitFile ? `<p class="prog">Source: ${esc(S.unitFile)}</p>` : ""}<div class="tw"><table><thead><tr>${th}</tr></thead><tbody>${tr}</tbody></table></div>`;
}

/* ---------- maintenance check popup ---------- */
function openMaint(k) { S.modal = val(S.unit[k], COLS.status[0]); renderModal(); }
function closeMaint() { S.modal = null; renderModal(); save(); render(); }
function renderModal() {
  const m = $("#modal"), id = S.modal;
  if (id == null) { m.hidden = true; m.innerHTML = ""; return; }
  const o = S.data.maint[id] || {}, a = sid(id);
  const old = m.querySelector(".mbody"), top = old ? old.scrollTop : 0;
  m.innerHTML = `<div class="mcard" role="dialog" aria-modal="true"><div class="mh"><h2>Maintenance check: ${esc(id)}</h2><button class="btn" onclick="closeMaint()">Done</button></div>
  <div class="mbody">${MAINT.map((q, i) => `<div class="mq"><div>${i + 1}. ${esc(q.t)}</div>
    <div class="mc">${q.text ? `<input placeholder="REF + REV" value="${esc(o["q" + i + "T"] || "")}" oninput="setF('maint','${a}','q${i}T',this.value)">` : chips("maint", id, "q" + i, ["OK", "NOK"], o["q" + i])}</div>
    ${o["q" + i] === "NOK" ? `<div class="ev">${evid("maint", id, "q" + i, "Comment")}</div>` : ""}</div>`).join("")}</div></div>`;
  m.hidden = false; m.querySelector(".mbody").scrollTop = top; hydrate(m);
}

/* ---------- navigation + screens ---------- */
function go(s) { S.screen = s; S.err = ""; render(); scrollTo(0, 0); }
async function back() {
  if (S.screen === "check") { await save(); go(S.from === "dash" ? "dash" : "wo"); }
  else if (S.screen === "unit") go("check");
  else if (S.screen === "dash") go("wo");
  else if (S.screen === "settings") go(S.status.length ? "wo" : "start");
}
function saveCfg() {
  ["clientId", "tenant", "statusUrl", "statusSheet", "tplUrl", "tplSheet", "unitUrl", "folderUrl"].forEach(k => S.cfg[k] = $("#cf_" + k).value.trim());
  localStorage.setItem("aprs_cfg", JSON.stringify(S.cfg)); msalApp = null; folderRef = null; go("start");
}

function render() {
  const sc = S.screen;
  let t = "APRS CHECKLIST", left = "", right = "";
  if (["check", "unit", "settings", "dash"].includes(sc)) left = `<button class="bk" onclick="back()" aria-label="Back">←</button>`;
  if (sc === "check") t = (S.tplProduct || S.product).toUpperCase() + " CHECKLIST";
  if (sc === "unit") t = "UNIT STATUS : " + S.wo;
  if (sc === "settings") t = "Settings";
  if (sc === "dash") t = "DASHBOARD";
  if (S.user && (sc === "start" || sc === "wo")) right = `<div class="hello">Hello,<br>${esc(S.user)}</div>`;
  if (sc === "wo") right = `<button class="hbtn" onclick="openDash()">Dashboard</button>` + right;
  const hd = $("#hd"); hd.className = "hd"; hd.innerHTML = `${left}<h1>${esc(t)}</h1>${right}`;

  const err = S.err ? `<div class="err" role="alert">${esc(S.err)}</div>` : "";
  let h = "";
  if (sc === "start") {
    h = `<h2 class="big">Sign in to start</h2><p class="sub">Your work orders and checklist templates load from SharePoint.</p>${err}
    <div class="stack"><button class="btn" onclick="loadAll(false)">Sign in with Microsoft</button>
    <button class="btn alt" onclick="loadAll(true)">Try with demo data</button>
    <button class="link" onclick="go('settings')">Settings</button></div>`;
  } else if (sc === "wo") {
    h = `<h2 class="big">Select your work order</h2><p class="sub">Choose a WO to automatically retrieve the serial number and product details.</p>${err}
    <label class="f" for="wo">Work order</label>
    <input id="wo" list="wos" autocomplete="off" placeholder="Type or pick a work order" value="${esc(S.wo)}" oninput="pickWo(this.value)">
    <datalist id="wos">${woList().map(w => `<option value="${esc(w)}">`).join("")}</datalist>
    <p class="prog" style="margin:8px 0 0"><span id="wocount">${woList().length} ${S.showAll || !hasStatusCol() ? "work orders" : "open work orders"}</span>
      ${hasStatusCol() ? `&nbsp; <label><input type="checkbox" ${S.showAll ? "checked" : ""} onchange="S.showAll=this.checked;refreshWoList()"> Show all, including closed</label>` : ""}</p>
    <div class="two"><div class="box"><small>Serial number</small><b id="sn">${esc(S.sn) || "-"}</b></div><div class="box"><small>Product</small><b id="prod">${esc(S.product) || "-"}</b></div></div>
    <div class="box" id="tpl">${tplBox()}</div>
    <div class="cta"><button id="go" class="btn" ${S.wo && S.choice ? "" : "disabled"} onclick="S.from='';openWo()">Continue</button></div>`;
  } else if (sc === "check") {
    h = `<div class="band"><span>WO: <b>${esc(S.wo)}</b></span><span>Product: <b>${esc(S.product)}</b></span><span>SN: <b>${esc(S.sn)}</b></span></div>
    <p class="prog" id="prog">${progText()}</p>${S.items.map((_, i) => rowHTML(i)).join("")}
    ${S.items.some(t => norm(val(t, TCOLS.item)) === "unit status") ? "" : `<button class="unitlink" onclick="go('unit')"><span>Unit status</span><span aria-hidden="true">→</span></button>`}
    ${reportBox()}`;
  } else if (sc === "unit") {
    h = unitHTML() + reportBox();
  } else if (sc === "dash") {
    h = dashHTML();
  } else if (sc === "settings") {
    const f = (k, l, p = "") => `<label class="f" for="cf_${k}" style="margin-top:14px">${l}</label><input id="cf_${k}" value="${esc(S.cfg[k])}" placeholder="${p}">`;
    h = `<div style="max-width:640px"><p class="sub">One-time setup, saved on this device.</p>
    ${f("clientId", "Azure app Client ID")}${f("tenant", "Tenant ID (or common)")}
    ${f("statusUrl", "SharePoint link: Status Excel file")}${f("statusSheet", "Status sheet name (blank = first sheet)")}
    ${f("tplUrl", "SharePoint link: Checklist templates Excel file")}${f("tplSheet", "Templates sheet name (blank = first sheet)")}
    <label class="f" for="cf_unitUrl" style="margin-top:14px">SharePoint links: unit status folders (one per line: FSM, Nacelles, Radome)</label><textarea id="cf_unitUrl" rows="4">${esc(S.cfg.unitUrl)}</textarea>
    ${f("folderUrl", "SharePoint link: results folder (checklists and photos are saved here)")}
    <div class="cta" style="justify-content:flex-start"><button class="btn" onclick="saveCfg()">Save settings</button></div></div>`;
  }
  $("#app").innerHTML = h;
  const sv = $("#saved"); sv.hidden = !(sc === "check" || sc === "unit"); sv.textContent = S.saved;
  if (sc === "dash") afterDashRender();
  hydrate($("#app"));
}
addEventListener("beforeunload", e => { if (S.dirty) { e.preventDefault(); e.returnValue = ""; } });
render();
