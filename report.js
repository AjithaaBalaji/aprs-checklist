/* =====================================================================
   REPORT HELPERS  (shared by pdfreport.js): unit wait, photo loading, upload, download
   ===================================================================== */
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

/* ---- upload (large files use an upload session; the 4 MB simple upload is not enough with photos) ---- */
async function putLarge(path, blob, type, conflict = "rename") {
  if (blob.size <= 3 * 1024 * 1024) return put(path, blob, type);
  const f = await folder();
  const r = await gfetch(`${G}/drives/${f.drive}/items/${f.id}:/${pth(path)}:/createUploadSession`,
    { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": conflict } }) });
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

