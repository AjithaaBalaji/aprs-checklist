/* =====================================================================
   PDF REPORT  (uses pdf-lib.min.js, runs after report.js and app.js)
   Same content as the Word report: dashboard summary, checklist, findings with
   photos, unit status, maintenance checks. Each press creates a NEW file with
   date and time and saves it in <results folder>/<work order>/reports.
   ===================================================================== */
const MIME_PDF = "application/pdf";

async function buildPdf() {
  const { PDFDocument, StandardFonts, rgb } = PDFLib;
  const doc = await PDFDocument.create();
  const F = await doc.embedFont(StandardFonts.Helvetica), FB = await doc.embedFont(StandardFonts.HelveticaBold), FI = await doc.embedFont(StandardFonts.HelveticaOblique);
  const known = new Set(F.getCharacterSet());
  const clean = s => [...String(s ?? "").replace(/→/g, "->").replace(/\t/g, " ")].map(ch => ch === "\n" ? ch : known.has(ch.codePointAt(0)) ? ch : "?").join("");
  const col = h => rgb(parseInt(h.slice(0, 2), 16) / 255, parseInt(h.slice(2, 4), 16) / 255, parseInt(h.slice(4, 6), 16) / 255);
  const GREEN = "5F8A2B", RED = "C0504D", OKC = "4A7C1F", AMBER = "B7791F", GREY = "8F9488", LINE = "B9C3AA";
  const PW = 841.89, PH = 595.28, M = 36, CW = PW - 2 * M, FOOT = 14;
  let page, y;
  const newPage = () => { page = doc.addPage([PW, PH]); y = PH - M; };
  const room = h => y - h >= M + FOOT;
  const need = h => { if (!room(h)) newPage(); };
  newPage();

  function wrap(text, font, size, maxW) {
    const out = [];
    for (const para of clean(text).split("\n")) {
      if (!para) { out.push(""); continue; }
      let line = "";
      for (const word of para.split(" ")) {
        let w = word;
        const test = line ? line + " " + w : w;
        if (font.widthOfTextAtSize(test, size) <= maxW) { line = test; continue; }
        if (line) { out.push(line); line = ""; }
        while (font.widthOfTextAtSize(w, size) > maxW && w.length > 1) {
          let k = w.length; while (k > 1 && font.widthOfTextAtSize(w.slice(0, k), size) > maxW) k--;
          out.push(w.slice(0, k)); w = w.slice(k);
        }
        line = w;
      }
      out.push(line);
    }
    return out;
  }
  function para(text, o = {}) {
    const size = o.size || 9.5, font = o.font || F, lh = size * 1.3, lines = wrap(text, font, size, CW);
    lines.forEach(ln => { need(lh); page.drawText(ln, { x: M, y: y - size * 0.85, size, font, color: col(o.color || "222222") }); y -= lh; });
    y -= o.after ?? 4;
  }
  const h1 = t => { need(40); para(t, { size: 19, font: FB, color: GREEN, after: 8 }); };
  const h2 = (t, keep = 60) => { need(keep); y -= 6; para(t, { size: 13, font: FB, color: GREEN, after: 6 }); };

  /* coloured result cell */
  const res = v => !v ? { t: "-", color: "999999" } : (v === "NO" || v === "NOK") ? { t: v, bold: true, color: RED }
    : (v === "OK" || v === "YES") ? { t: v, bold: true, color: OKC } : { t: v };

  function table(widths, head, rows) {
    const sum = widths.reduce((a, b) => a + b, 0), W = widths.map(w => w * CW / sum), size = 8.5, lh = size * 1.3, px = 4, py = 3.5;
    const txt = c => c && typeof c === "object" ? c.t : c;
    const maxLines = Math.floor((PH - 2 * M - 30 - 2 * py) / lh);
    function drawRow(cells, isHead) {
      let lines = cells.map((c, i) => wrap(txt(c), isHead || (c && c.bold) ? FB : F, size, W[i] - 2 * px));
      lines = lines.map(l => l.length > maxLines ? [...l.slice(0, maxLines - 1), "..."] : l);
      const h = Math.max(1, ...lines.map(l => l.length)) * lh + 2 * py;
      if (!room(h)) { newPage(); if (!isHead && head) drawRow(head, true); }
      let x = M;
      cells.forEach((c, i) => {
        const spec = c && typeof c === "object" ? c : {};
        page.drawRectangle({ x, y: y - h, width: W[i], height: h, color: isHead ? col(GREEN) : spec.fill ? col(spec.fill) : undefined, borderColor: col(LINE), borderWidth: 0.5 });
        const font = isHead || spec.bold ? FB : F, color = isHead ? rgb(1, 1, 1) : col(spec.color || "222222");
        lines[i].forEach((ln, k) => page.drawText(ln, { x: x + px, y: y - py - size * 0.85 - k * lh, size, font, color }));
        x += W[i];
      });
      y -= h;
    }
    if (head) drawRow(head, true);
    rows.forEach(r => drawRow(r, false));
    y -= 8;
  }

  async function photoRow(names) {
    const imgs = [];
    for (const n of names || []) {
      const p = await photoData(n); if (!p) continue;
      try { const im = await doc.embedJpg(p.buf), k = Math.min(1, 200 / p.w, 170 / p.h); imgs.push({ im, w: p.w * k, h: p.h * k }); } catch (e) { /* skip unreadable photo */ }
    }
    return imgs;
  }
  async function evidence(title, comment, names) {
    const imgs = await photoRow(names), cl = comment ? wrap(comment, F, 9.5, CW).length * 12.4 : 0;
    need(18 + cl + (imgs.length ? imgs[0].h : 0));
    para(title, { font: FB, size: 10, after: 2 });
    if (comment) para(comment);
    let x = M, rowH = 0;
    for (const s of imgs) {
      if (x > M && x + s.w > M + CW) { y -= rowH + 8; x = M; rowH = 0; }
      if (!room(s.h)) { newPage(); x = M; rowH = 0; }
      page.drawImage(s.im, { x, y: y - s.h, width: s.w, height: s.h });
      x += s.w + 8; rowH = Math.max(rowH, s.h);
    }
    if (imgs.length) y -= rowH + 10; else y -= 4;
  }

  /* ---------------- numbers ---------------- */
  const d = S.data, now = new Date(), sm = summarize(d, S.items);
  const ms = { ok: 0, nok: 0, prog: 0, none: 0 }, unitReady = S.unitState === "ready";
  if (unitReady) S.unit.forEach(r => { const lab = maintStateOf(d.maint[val(r, COLS.status[0])])[0]; if (lab === "OK") ms.ok++; else if (lab === "NOK") ms.nok++; else if (lab === "In progress") ms.prog++; else ms.none++; });
  const pct = sm.n ? Math.round(100 * sm.full / sm.n) : 0, attention = sm.nok + sm.notRet + ms.nok;
  const banner = attention ? [RED, `ACTION NEEDED: ${sm.nok} line(s) NOK, ${sm.notRet} not returned, ${ms.nok} maintenance item(s) NOK`]
    : (sm.full < sm.n || (unitReady && ms.ok < S.unit.length)) ? [AMBER, `IN PROGRESS: ${sm.n - sm.full} checklist line(s) not fully answered` + (unitReady ? `, ${S.unit.length - ms.ok} maintenance item(s) not finished` : "")]
    : [OKC, "COMPLETE: all checklist lines and maintenance checks are OK"];

  /* ---------------- page 1 ---------------- */
  h1((S.tplProduct || S.product) + " checklist report");
  const L = t => ({ t, bold: true, fill: "E8F2DA" });
  table([75, 120, 55, 105, 55, 135, 60, 135], null, [
    [L("WO"), S.wo, L("S/N"), S.sn, L("Product"), S.product, L("Template"), S.tplProduct],
    [L("Aircraft / Engine"), S.engine, L("Inspector"), d.inspector || S.user, L("Generated"), now.toLocaleString(), "", ""]
  ]);
  need(30);
  page.drawRectangle({ x: M, y: y - 24, width: CW, height: 24, color: col(banner[0]) });
  page.drawText(clean(banner[1]), { x: M + 10, y: y - 16, size: 11, font: FB, color: rgb(1, 1, 1) });
  y -= 32;

  const good = ["E8F2DA", OKC], bad = ["FBE9E7", RED], neutral = ["F0F0EE", "444444"];
  const tiles = [
    [pct + "%", "Checklist complete", pct === 100 ? good : neutral], [`${sm.ok} / ${sm.n}`, "Lines checked OK", good],
    [String(sm.nok), "Lines NOK", sm.nok ? bad : good], [String(sm.notRet), "Not returned", sm.notRet ? bad : good],
    [unitReady ? `${ms.ok} / ${S.unit.length}` : "-", "Maintenance checks OK", unitReady && S.unit.length && ms.ok === S.unit.length ? good : neutral],
    [unitReady ? String(ms.nok) : "-", "Maintenance NOK", ms.nok ? bad : good], [String(sm.photos), "Photos attached", neutral]
  ];
  { const gap = 6, w = (CW - gap * 6) / 7, h = 54; need(h + 8);
    tiles.forEach(([v, lab, c], i) => {
      const x = M + i * (w + gap), vv = clean(v), ll = clean(lab);
      page.drawRectangle({ x, y: y - h, width: w, height: h, color: col(c[0]) });
      page.drawText(vv, { x: x + (w - FB.widthOfTextAtSize(vv, 22)) / 2, y: y - 29, size: 22, font: FB, color: col(c[1]) });
      page.drawText(ll, { x: x + (w - F.widthOfTextAtSize(ll, 7.5)) / 2, y: y - 45, size: 7.5, font: F, color: col("555555") });
    });
    y -= h + 10; }
  function bar(title, segs) {
    const tot = segs.reduce((a, s) => a + s.n, 0); if (!tot) return;
    need(34); page.drawText(clean(title), { x: M, y: y - 8, size: 8.5, font: FB, color: col("444444") }); y -= 13;
    let x = M;
    segs.filter(s => s.n > 0).forEach(s => {
      const w = CW * s.n / tot, label = clean(`${s.label}: ${s.n}`), lab = F.widthOfTextAtSize(label, 8.5) + 8 <= w ? label : (w > 14 ? String(s.n) : "");
      page.drawRectangle({ x, y: y - 17, width: w, height: 17, color: col(s.c), borderColor: rgb(1, 1, 1), borderWidth: 1 });
      if (lab) page.drawText(lab, { x: x + (w - FB.widthOfTextAtSize(lab, 8.5)) / 2, y: y - 12, size: 8.5, font: FB, color: rgb(1, 1, 1) });
      x += w;
    });
    y -= 25;
  }
  bar("Checklist lines", [{ n: sm.ok, label: "Checked OK", c: OKC }, { n: sm.nok, label: "NOK", c: RED }, { n: sm.notChk, label: "Not checked yet", c: GREY }]);
  if (unitReady) bar("Maintenance checks by unit status item", [{ n: ms.ok, label: "OK", c: OKC }, { n: ms.nok, label: "NOK", c: RED }, { n: ms.prog, label: "In progress", c: AMBER }, { n: ms.none, label: "Not started", c: GREY }]);

  /* ---------------- checklist ---------------- */
  h2("Checklist", 80);
  table([55, 235, 85, 38, 55, 55, 55, 190], ["Step", "Checklist item", "Form", "Folder", "Delivered", "Returned", "Checked", "Comments"],
    S.items.map((t, i) => {
      const a = d.rows[rk(i)] || {}, step = val(t, TCOLS.step), prev = i ? val(S.items[i - 1], TCOLS.step) : null, ph = (a.retP || []).length + (a.chkP || []).length;
      const com = [a.retC && "Returned: " + a.retC, a.chkC && "Checked: " + a.chkC, ph && `${ph} photo(s), see Findings`].filter(Boolean).join("\n");
      return [{ t: step !== prev ? step : "", bold: true }, val(t, TCOLS.item), val(t, TCOLS.form), val(t, TCOLS.folder), res(a.del), res(a.ret), res(a.chk), com];
    }));

  /* ---------------- findings ---------------- */
  h2("Findings", 60);
  let nf = 0;
  for (let i = 0; i < S.items.length; i++) {
    const t = S.items[i], a = d.rows[rk(i)] || {}, name = val(t, TCOLS.item) + (val(t, TCOLS.form) ? " (" + val(t, TCOLS.form) + ")" : "");
    if (a.ret === "NO") { nf++; await evidence(name + ": not returned", a.retC, a.retP); }
    if (a.chk === "NOK") { nf++; await evidence(name + ": checked NOK", a.chkC, a.chkP); }
  }
  if (!nf) para("No lines were marked No or NOK.", { font: FI });

  /* ---------------- unit status ---------------- */
  newPage(); para("Unit status", { size: 13, font: FB, color: GREEN, after: 6 });
  if (unitReady && S.unit.length) {
    if (S.unitFile) para("Source: " + S.unitFile, { font: FI, size: 8.5, color: "666666" });
    table([45, 62, 48, 78, 175, 140, 75, 38, 62, 55], [...COLS.status, "Check"], S.unit.map(r => {
      const id = val(r, COLS.status[0]), lab = maintStateOf(d.maint[id])[0];
      return [...COLS.status.map((c, ci) => ci === 0 ? { t: val(r, c), bold: true } : val(r, c)), lab === "Check" ? { t: "Not started", color: "999999" } : res(lab)];
    }));
  } else para(S.unitState === "error" ? "Unit status could not be loaded: " + S.unitErr : "No unit status rows found for this work order.", { font: FI });

  /* ---------------- maintenance checks ---------------- */
  h2("Maintenance checks", 90);
  let nm = 0;
  for (const r of S.unit) {
    const id = val(r, COLS.status[0]), o = d.maint[id];
    if (!o || !Object.keys(o).length) continue;
    nm++; need(110); para("Maintenance check: " + id, { font: FB, size: 11, after: 4 });
    table([28, 400, 60, 280], ["#", "Question", "Result", "Comment / answer"],
      MAINT.map((q, i) => [String(i + 1), q.t, q.text ? { t: "-", color: "999999" } : res(o["q" + i]), q.text ? (o["q" + i + "T"] || "") : (o["q" + i + "C"] || "")]));
    for (let i = 0; i < MAINT.length; i++) { const ph = o["q" + i + "P"]; if (ph && ph.length) await evidence(`${id}, question ${i + 1}: photos`, "", ph); }
  }
  if (!nm) para("No maintenance checks have been filled in yet.", { font: FI });

  /* ---------------- footer on every page ---------------- */
  const pages = doc.getPages();
  pages.forEach((p, i) => {
    const t = clean(`${S.wo}  |  ${S.tplProduct}  |  Page ${i + 1} of ${pages.length}`);
    p.drawText(t, { x: (PW - F.widthOfTextAtSize(t, 7.5)) / 2, y: 20, size: 7.5, font: F, color: col("666666") });
  });
  doc.setTitle(clean(`${S.wo} checklist report`)); doc.setCreator("APRS Checklist"); doc.setAuthor(clean(S.user || "APRS Checklist"));
  return doc.save();
}

let reportingPdf = false;
async function makePdf() {
  if (reportingPdf || reporting) return;
  if (typeof PDFLib === "undefined") { note("The PDF library did not load. Refresh the page (Ctrl+F5) and try again.", 7000); return; }
  reportingPdf = true;
  try {
    note("Preparing PDF…");
    await save(); await waitUnit();
    const bytes = await buildPdf(), blob = new Blob([bytes], { type: MIME_PDF }), name = `${key(S.wo)} Report ${stampOf(new Date())}.pdf`;
    if (!demoMode) { note("Saving to SharePoint…"); await putLarge(key(S.wo) + "/reports/" + name, blob, MIME_PDF); }
    download(blob, name);
    note(demoMode ? "PDF downloaded (demo mode)." : `Saved in the results folder: ${key(S.wo)} / reports / ${name}`, 9000);
  } catch (e) { note("PDF failed: " + e.message, 9000); }
  reportingPdf = false;
}
