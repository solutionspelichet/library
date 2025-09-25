// --- helpers UI ---

const resp = await fetch(gasUrl, {
  method: 'POST',
  mode: 'cors',
  redirect: 'follow',
  credentials: 'omit',
  headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // <-- évite le preflight
  body: JSON.stringify(payload) // on envoie quand même du JSON, mais comme texte
});

const $ = (id) => document.getElementById(id);
const logEl = $('log');
const log = (m) => { console.log(m); logEl.textContent += m + '\n'; };
const setBusy = (busy) => { $('runBtn').disabled = busy; $('runBtn').textContent = busy ? 'Traitement en cours…' : 'Lancer le traitement'; };

const excelLetterToIndex = (L) => {
  L = String(L || '').trim().toUpperCase();
  let idx = 0; for (const ch of L) idx = idx * 26 + (ch.charCodeAt(0) - 64);
  return idx - 1;
};
const parseDate = (v) => {
  const d = new Date(v);
  if (isNaN(d)) return null;
  const y = d.getFullYear(), m = String(d.getMonth()+1).padStart(2,'0'), dd = String(d.getDate()).padStart(2,'0');
  return `${y}-${m}-${dd}`; // AAAA-MM-JJ
};
const SHEET_ID = '1AptbV2NbY0WQZpe_Xt1K2iVlDpgKADElamKQCg3GcXQ';

// assure que le handler est attaché quand le DOM est prêt
document.addEventListener('DOMContentLoaded', () => {
  log('✅ App prête. Sélectionne les 2 fichiers et renseigne l’URL Apps Script.');
  $('runBtn').addEventListener('click', onRun);
});

async function onRun() {
  setBusy(true);
  try {
    log('--- Début ---');

    const sFile = $('suiviFile').files[0];
    const eFile = $('extractFile').files[0];
    const gasUrl = $('gasUrl').value.trim();
    const secret = $('secret').value.trim();

    // validations
    if (!sFile) { alert('Sélectionne le fichier de suivi (.xlsx)'); throw new Error('Suivi manquant'); }
    if (!eFile) { alert('Sélectionne le fichier d’extraction (.xlsx)'); throw new Error('Extraction manquante'); }
    if (!gasUrl) { alert('Renseigne l’URL Apps Script Web App (/exec)'); throw new Error('URL Apps Script absente'); }

    log('Lecture fichiers… (dans le navigateur)');
    const sWorkbook = await readWorkbook(sFile);
    const eWorkbook = await readWorkbook(eFile);

    log('Nettoyage & calcul tableaux (suivi / extraction)…');
    const sData = processWorkbook(sWorkbook, { key:'s_key', user:'s_user', sum:'s_sum', date:'s_date', head:'s_head' });
    const eData = processWorkbook(eWorkbook, { key:'e_key', user:'e_user', sum:'e_sum', date:'e_date', head:'e_head' });

    log('Calcul resultats (addition alignée) …');
    const resultats = mergeTablesByContactAndHeaders(sData.tableau, eData.tableau);
    const ml = multiplyValues(resultats, 0.35);

    const payload = {
      secret,
      sheetId: SHEET_ID,
      resultats: { headers: resultats.headers, rows: resultats.rows },
      ml: { headers: ml.headers, rows: ml.rows }
    };

    log('Envoi vers Google Sheets (Apps Script)…');
    const resp = await fetchWithTimeout(gasUrl, {
      method: 'POST',
      mode: 'cors',
      redirect: 'follow',
      credentials: 'omit',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }, 60000); // 60s

    let text = await resp.text();
    let json;
    try { json = JSON.parse(text); } catch { json = { raw: text }; }

    if (!resp.ok) {
      console.error('Réponse Apps Script non OK', resp.status, json);
      throw new Error(json?.error || `Apps Script HTTP ${resp.status}`);
    }

    log('✅ Écriture terminée: ' + JSON.stringify(json));
    alert('Terminé ! Vérifie le Google Sheet.');
  } catch (e) {
    console.error(e);
    log('❌ ' + e.message);
    alert('Erreur : ' + e.message);
  } finally {
    setBusy(false);
    log('--- Fin ---');
  }
}

// --- I/O & calculs ---
async function readWorkbook(file) {
  const data = await file.arrayBuffer();
  return XLSX.read(data, { type: 'array' });
}
function cleanSheetToAOA(workbook) {
  const firstName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstName];
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  if (aoa.length === 0) return aoa;
  const header = aoa[0].map(x => (x==null?'':String(x).trim()));
  if (aoa.length >= 2) {
    const firstDataRow = aoa[1].map(x => (x==null?'':String(x).trim()));
    if (arraysEqual(header, firstDataRow)) aoa.splice(1,1);
  }
  if (aoa.length >= 2) aoa.splice(aoa.length-1, 1);
  return aoa;
}
function arraysEqual(a,b){ return a.length===b.length && a.every((v,i)=>String(v).trim()===String(b[i]).trim()); }
function aoaToObjects(aoa) {
  if (!aoa || !aoa.length) return [];
  const headers = aoa[0].map(h => String(h||'').trim());
  return aoa.slice(1).map(row => {
    const o = {}; headers.forEach((h,i)=>o[h]=row[i]); return o;
  });
}
function processWorkbook(workbook, refs) {
  const aoa = cleanSheetToAOA(workbook);
  const rows = aoaToObjects(aoa);
  if (!rows.length) return { tableau: emptyTable() };

  const headers = aoa[0];
  const colByLetter = (L) => headers[excelLetterToIndex($(refs[L]).value)];
  const colKey  = colByLetter('key');
  const colUser = colByLetter('user');
  const colSum  = colByLetter('sum');
  const colDate = colByLetter('date');
  const colHead = colByLetter('head');

  const seen = new Set(); const dedupe = [];
  for (const r of rows) { const k = r[colKey]; if (!seen.has(k)) { seen.add(k); dedupe.push(r); } }

  // per-day
  const perDayMap = new Map(); const usersSet = new Set(); const daysSet = new Set();
  for (const r of dedupe) {
    const u = r[colUser]; usersSet.add(u);
    const d = parseDate(r[colDate]); if (!d) continue; daysSet.add(d);
    const val = Number(r[colSum] ?? 0) || 0; const key = `${u}||${d}`;
    perDayMap.set(key, (perDayMap.get(key)||0) + val);
  }
  const days = Array.from(daysSet).sort();
  const headersOut = ['Contact', ...days.map(d => `nombre colonne carton ${d}`)];
  const rowsOut = [];
  for (const u of Array.from(usersSet).sort()) {
    const row = [u]; for (const d of days) row.push(perDayMap.get(`${u}||${d}`) || 0);
    rowsOut.push(row);
  }
  return { tableau: { headers: headersOut, rows: rowsOut } };
}
function emptyTable(){ return { headers:['Contact'], rows:[] } }
function indexMap(h){ const m={}; (h||[]).forEach((name,i)=>m[name]=i); return m; }
function sumRowInto(targetRow, srcRow, headers, srcIdxMap){
  for (let i=1;i<headers.length;i++){
    const name = headers[i];
    const si = srcIdxMap[name];
    const v = (si==null) ? 0 : Number(srcRow[si]||0);
    targetRow[i] = Number(targetRow[i]||0) + (isNaN(v)?0:v);
  }
}
function mergeTablesByContactAndHeaders(A, B) {
  const headers = Array.from(new Set([...(A.headers||[]), ...(B.headers||[]) ]));
  let ci = headers.indexOf('Contact'); if (ci > 0){ headers.splice(ci,1); headers.unshift('Contact'); }
  const idxA = indexMap(A.headers||[]), idxB = indexMap(B.headers||[]);
  const contacts = new Set([...(A.rows||[]).map(r=>r[0]), ...(B.rows||[]).map(r=>r[0])]);
  const outRows = [];
  for (const c of Array.from(contacts).sort()) {
    const row = Array(headers.length).fill(0); row[0] = c;
    const ra = (A.rows||[]).find(r => r[0]===c); if (ra) sumRowInto(row, ra, headers, idxA);
    const rb = (B.rows||[]).find(r => r[0]===c); if (rb) sumRowInto(row, rb, headers, idxB);
    for (let i=1;i<row.length;i++) row[i] = Number(row[i]||0);
    outRows.push(row);
  }
  return { headers, rows: outRows };
}
function multiplyValues(table, k) {
  const headers = table.headers.slice();
  const rows = table.rows.map(r=>{
    const out = r.slice(); for (let i=1;i<out.length;i++) out[i] = Number(out[i]||0)*k; return out;
  });
  return { headers, rows };
}
async function fetchWithTimeout(url, options={}, timeoutMs=60000){
  const ctrl = new AbortController();
  const id = setTimeout(()=>ctrl.abort(new Error('Timeout')), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: ctrl.signal });
  } finally { clearTimeout(id); }
}
