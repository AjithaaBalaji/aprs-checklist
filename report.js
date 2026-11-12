/* =====================================================================
   WORD REPORT  (uses docx.js, runs after app.js)
   Layout lives in buildReport(): header table, checklist table,
   findings with photos, unit status table, maintenance checks.
   Each press of "Generate Word report" creates a NEW file with date and time.
   ===================================================================== */
const MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const R_GREEN = "5F8A2B", R_RED = "C0504D", R_OK = "4A7C1F", R_LIGHT = "E8F2DA";
const pad2 = n => String(n).padStart(2, "0");
const stampOf = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ${pad2(d.getHours())}-${pad2(d.getMinutes())}`;

async function waitUnit() {
  if (S.unitState === "idle") startUnit();
  for (let i = 0; i < 120 && S.unitState === "loading"; i++) await new Promise(r => setTimeout(r, 250));
}
async function photoData(name) {
  const u = await Store.photoUrl(S.wo, name);
  if (!u) return null;
  const buf = await (await fetch(u)).arrayBuffer();
  const dim = await new Promise(res => { const im = new Image(); im.onload = () => res([im.naturalWidth, im.naturalHeight]); im.onerror = () => res([4, 3]); im.src = u; });
  return { buf, w: dim[0], h: dim[1] };
}

async function buildReport() {
  const D = docx, d = S.data, now = new Date();
  const T = (text, o = {}) => new D.TextRun({ text: String(text ?? ""), bold: o.bold, color: o.color, size: o.size, italics: o.italics });
  const para = (children, o = {}) => new D.Paragraph({ children: Array.isArray(children) ? children : [children], ...o });
  const lines = (text, o) => String(text ?? "").split("\n").map(l => para(T(l, o)));
  const cell = (text, w, o = {}) => new D.TableCell({
    width: { size: w, type: D.WidthType.DXA },
    shading: o.fill ? { type: D.ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    verticalAlign: D.VerticalAlign.CENTER,
    children: lines(text, o)
  });
  const table = (widths, head, rows) => new D.Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: D.WidthType.DXA },
    columnWidths: widths,
    rows: [
      new D.TableRow({ tableHeader: true, children: head.map((h, i) => cell(h, widths[i], { bold: true, color: "FFFFFF", fill: R_GREEN })) }),
      ...rows.map(r => new D.TableRow({ cantSplit: true, children: r.map((c, i) => cell(c && c.t !== undefined ? c.t : c, widths[i], (c && c.o) || {})) }))
    ]
  });
  // coloured result text: OK/YES green, NO/NOK red
  const res = v => !v ? { t: "-", o: { color: "999999" } }
    : (v === "NO" || v === "NOK") ? { t: v, o: { bold: true, color: R_RED } }
    : (v === "OK" || v === "YES") ? { t: v, o: { bold: true, color: R_OK } } : { t: v };

  async function photoParas(names) {
    const runs = [];
    for (const n of names || []) {
      const p = await photoData(n);
      if (!p) continue;
      const k = Math.min(1, 260 / p.w, 300 / p.h);
      runs.push(new D.ImageRun({ type: "jpg", data: p.buf, transformation: { width: Math.round(p.w * k), height: Math.round(p.h * k) } }), T("   "));
    }
    return runs.length ? [new D.Paragraph({ children: runs, spacing: { before: 60, after: 120 } })] : [];
  }
  async function evidence(title, comment, names) {
    const out = [para(T(title, { bold: true }), { spacing: { before: 160, after: 40 }, keepNext: true })];
    if (comment) out.push(...lines(comment));
    out.push(...await photoParas(names));
    return out;
  }
  const h = (text, lvl, o = {}) => new D.Paragraph({ heading: lvl, children: [T(text)], ...o });

  const kids = [];
  kids.push(h((S.tplProduct || S.product) + " checklist report", D.HeadingLevel.HEADING_1));

  /* ---- header block ---- */
  const iw = [1300, 2550, 1300, 2550, 1300, 2549, 1300, 2549];
  const info = [
    ["WO", S.wo, "S/N", S.sn, "Product", S.product, "Template", S.tplProduct],
    ["Aircraft / Engine", S.engine, "Inspector", d.inspector || S.user, "Generated", now.toLocaleString(), "", ""]
  ];
  kids.push(new D.Table({
    width: { size: 15398, type: D.WidthType.DXA }, columnWidths: iw,
    rows: info.map(r => new D.TableRow({ children: r.map((v, i) => cell(v, iw[i], i % 2 === 0 ? { bold: true, fill: v ? R_LIGHT : undefined } : {})) }))
  }));

  /* ---- checklist ---- */
  kids.push(h("Checklist", D.HeadingLevel.HEADING_2, { spacing: { before: 280, after: 100 } }));
  const cw = [1000, 4300, 1700, 900, 1300, 1300, 1300, 3598];
  let nChecked = 0, nNok = 0, nNotRet = 0, nOpen = 0;
  const rows = S.items.map((t, i) => {
    const a = d.rows[rk(i)] || {}, step = val(t, TCOLS.step), prev = i ? val(S.items[i - 1], TCOLS.step) : null;
    if (a.chk) nChecked++; if (a.chk === "NOK") nNok++; if (a.ret === "NO") nNotRet++; if (!(a.del && a.ret && a.chk)) nOpen++;
    const ph = (a.retP || []).length + (a.chkP || []).length;
    const com = [a.retC && "Returned: " + a.retC, a.chkC && "Checked: " + a.chkC, ph && `${ph} photo(s), see Findings`].filter(Boolean).join("\n");
    return [{ t: step !== prev ? step : "", o: { bold: true } }, val(t, TCOLS.item), val(t, TCOLS.form), val(t, TCOLS.folder), res(a.del), res(a.ret), res(a.chk), com];
  });
  kids.push(table(cw, ["Step", "Checklist item", "Form", "Folder", "Delivered", "Returned", "Checked", "Comments"], rows));
  kids.push(para(T(`${nChecked} of ${S.items.length} lines checked.  NOK: ${nNok}.  Not returned: ${nNotRet}.  Lines not fully answered: ${nOpen}.`, { italics: true }), { spacing: { before: 100 } }));

  /* ---- findings with photos ---- */
  kids.push(h("Findings", D.HeadingLevel.HEADING_2, { spacing: { before: 280, after: 60 } }));
  let nf = 0;
  for (let i = 0; i < S.items.length; i++) {
    const t = S.items[i], a = d.rows[rk(i)] || {}, name = val(t, TCOLS.item) + (val(t, TCOLS.form) ? " (" + val(t, TCOLS.form) + ")" : "");
    if (a.ret === "NO") { nf++; kids.push(...await evidence(name + ": not returned", a.retC, a.retP)); }
    if (a.chk === "NOK") { nf++; kids.push(...await evidence(name + ": checked NOK", a.chkC, a.chkP)); }
  }
  if (!nf) kids.push(para(T("No lines were marked No or NOK.", { italics: true })));

  /* ---- unit status ---- */
  kids.push(h("Unit status", D.HeadingLevel.HEADING_2, { pageBreakBefore: true, spacing: { after: 60 } }));
  if (S.unitState === "ready" && S.unit.length) {
    if (S.unitFile) kids.push(para(T("Source: " + S.unitFile, { italics: true, color: "666666" }), { spacing: { after: 80 } }));
    const uw = [900, 1200, 1000, 1500, 3300, 2600, 1500, 800, 1300, 1298];
    const urows = S.unit.map(r => {
      const id = val(r, COLS.status[0]), ms = maintState(id), lab = ms[0] === "Check" ? "Not started" : ms[0];
      return [...COLS.status.map((c, ci) => ci === 0 ? { t: val(r, c), o: { bold: true } } : val(r, c)), lab === "Not started" ? { t: lab, o: { color: "999999" } } : res(lab)];
    });
    kids.push(table(uw, [...COLS.status, "Check"], urows));
  } else kids.push(para(T(S.unitState === "error" ? "Unit status could not be loaded: " + S.unitErr : "No unit status rows found for this work order.", { italics: true })));

  /* ---- maintenance checks ---- */
  kids.push(h("Maintenance checks", D.HeadingLevel.HEADING_2, { spacing: { before: 280, after: 60 } }));
  const mw = [600, 7500, 1400, 5898];
  let nm = 0;
  for (const r of S.unit) {
    const id = val(r, COLS.status[0]), o = d.maint[id];
    if (!o || !Object.keys(o).length) continue;
    nm++;
    kids.push(para(T(`Maintenance check: ${id}`, { bold: true, size: 24 }), { spacing: { before: 200, after: 60 }, keepNext: true }));
    kids.push(table(mw, ["#", "Question", "Result", "Comment / answer"],
      MAINT.map((q, i) => [String(i + 1), q.t, q.text ? { t: "-", o: { color: "999999" } } : res(o["q" + i]), q.text ? (o["q" + i + "T"] || "") : (o["q" + i + "C"] || "")])));
    for (let i = 0; i < MAINT.length; i++) {
      const ph = o["q" + i + "P"];
      if (ph && ph.length) kids.push(...await evidence(`${id}, question ${i + 1}: photos`, "", ph));
    }
  }
  if (!nm) kids.push(para(T("No maintenance checks have been filled in yet.", { italics: true })));

  const small = { size: 16, color: "666666" };
  const doc = new D.Document({
    creator: S.user || "APRS Checklist", title: `${S.wo} checklist report`,
    styles: {
      default: { document: { run: { font: "Calibri", size: 20 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 36, bold: true, color: R_GREEN }, paragraph: { spacing: { before: 0, after: 160 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 28, bold: true, color: R_GREEN }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } }
      ]
    },
    sections: [{
      properties: { page: { size: { orientation: D.PageOrientation.LANDSCAPE }, margin: { top: 720, bottom: 720, left: 720, right: 720 } } },
      footers: { default: new D.Footer({ children: [new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [
        T(`${S.wo}  |  ${S.tplProduct}  |  Page `, small), new D.TextRun({ children: [D.PageNumber.CURRENT], ...small }), T(" of ", small), new D.TextRun({ children: [D.PageNumber.TOTAL_PAGES], ...small })] })] }) },
      children: kids
    }]
  });
  return D.Packer.toBlob(doc);
}

/* ---- upload (large files use an upload session; the 4 MB simple upload is not enough with photos) ---- */
async function putLarge(path, blob, type) {
  if (blob.size <= 3 * 1024 * 1024) return put(path, blob, type);
  const f = await folder();
  const r = await gfetch(`${G}/drives/${f.drive}/items/${f.id}:/${pth(path)}:/createUploadSession`,
    { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }) });
  if (!r.ok) throw new Error("Could not start the upload (" + r.status + ")");
  const { uploadUrl } = await r.json(), CH = 10 * 327680;   // 3.1 MB chunks (multiple of 320 KiB)
  for (let start = 0; start < blob.size; start += CH) {
    const end = Math.min(start + CH, blob.size);
    const u = await fetch(uploadUrl, { method: "PUT", headers: { "Content-Range": `bytes ${start}-${end - 1}/${blob.size}` }, body: blob.slice(start, end) });
    if (!u.ok && u.status !== 202) throw new Error("Upload failed (" + u.status + ")");
  }
}
function download(blob, name) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 3000);
}

let reporting = false;
async function makeReport() {
  if (reporting) return;
  if (typeof docx === "undefined") { note("The report library did not load. Refresh the page and try again.", 7000); return; }
  reporting = true;
  try {
    note("Preparing report…");
    await save(); await waitUnit();
    const blob = await buildReport(), name = `${key(S.wo)} Report ${stampOf(new Date())}.docx`;
    if (!demoMode) { note("Saving to SharePoint…"); await putLarge(key(S.wo) + "/reports/" + name, blob, MIME_DOCX); }
    download(blob, name);
    note(demoMode ? "Report downloaded (demo mode)." : `Saved in the results folder: ${key(S.wo)} / reports / ${name}`, 9000);
  } catch (e) { note("Report failed: " + e.message, 9000); }
  reporting = false;
}
