// =====================
// app.js (PWA Frontend)
// =====================

// --- Helpers UI & logs ---
const $ = (id) => document.getElementById(id);
const logEl = $('log');
const log = (m) => { try { console.log(m); } catch(_){} if (logEl) logEl.textContent += m + '\n'; };
const setBusy = (busy) => { const b = $('runBtn'); if (b) { b.disabled = busy; b.textContent = busy ? 'Traitement en cours…' : 'Lancer le traitement'; } };

// Log erreurs visibles
window.addEventListener('error', (e) => log('⛔ JS error: ' + (e?.error?.message || e.message || e.toString())));
window.addEventListener('unhandledrejection', (e) => log('⛔ Promise rejection: ' + (e?.reason?.message || e.reason || e.toString())));

// --- Config fixe ---
const SHEET_ID = '1AptbV2NbY0WQZpe_Xt1K2iVlDpgKADElamKQCg3GcXQ';
const GAS_URL  = 'https://script.google.com/macros/s/AKfycbwO0P3Yo5kw9PPriJPXzUMipBrzlGTR_r-Ff6OyEUnsNu-I9q-rESbBq7l2m6KLA3RJ/exec'; // <— remplace

// (optionnel) mémoriser le secret
document.addEventListener('DOMContentLoaded', () => {
  log('✅ App prête. Sélectionne les 2 fichiers.');
  const saved = localStorage.getItem('PWA_SECRET');
  if (saved && $('secret')) $('secret').value = saved;

  const btn = $('runBtn'); if (btn && !btn.onclick) btn.addEventListener('click', onRun);
  const tbtn = $('testBtn'); if (tbtn && !tbtn.onclick) tbtn.addEventListener('click', testConnexion);

  if ($('secret')) $('secret').addEventListener('change', e => localStorage.setItem('PWA_SECRET', e.target.value));
});

// Test GET (peut échouer à cause de CORS ; juste indicatif)
async function testConnexion(){
  try{
    if(!GAS_URL) { alert('Définis GAS_URL dans app.js'); return; }
    log('Ping Apps Script…');
    const r = await fetch(GAS_URL, { method:'GET', mode:'no-cors' });
    log('GET envoyé (no-cors). Si besoin, vérifie dans Apps Script > Exécutions.');
  }catch(e){ log('❌ Test: ' + e.message); alert(e.message); }
}

// Lancer le traitement
async function onRun() {
  setBusy(true);
  try {
    log('--- Début ---');

    const sFile = $('suiviFile')?.files?.[0];
    const eFile = $('extractFile')?.files?.[0];
    const secret = $('secret') ? $('secret').value.trim() : '';

    if (!sFile) { alert('Sélectionne le fichier de suivi (.xlsx)'); throw new Error('Suivi manquant'); }
    if (!eFile) { alert('Sélectionne le fichier d’extraction (.xlsx)'); throw new Error('Extraction manquante'); }
    if (!GAS_URL) { alert('Définis GAS_URL dans app.js'); throw new Error('URL Apps Script absente'); }

    log('Lecture fichiers… (dans le navigateur)');
    const sWorkbook = await readWorkbook(sFile);
    const eWorkbook = await readWorkbook(eFile);

    log('Nettoyage & calcul tableaux (suivi / extraction)…');
    const sData = processWorkbook(sWorkbook, { key:'s_key', user:'s_user', sum:'s_sum', date:'s_date', head:'s_head' });
    const eData = processWorkbook(eWorkbook, { key:'e_key', user:'e_user', sum:'e_sum', date:'e_date', head:'e_head' });

    log('Calcul resultats (addition alignée)…');
    const resultats = mergeTablesByContactAndHeaders(sData.tableau, eData.tableau);
    const ml = multiplyValues(resultats, 0.35);

    const payload = {
      secret,
      sheetId: SHEET_ID,
      resultats: { headers: resultats.headers, rows: resultats.rows },
      ml: { headers: ml.headers, rows: ml.rows }
    };

    log('Envoi vers Google Sheets (Apps Script)…');
    // IMPORTANT: no-cors => on n’essaie pas de lire la réponse
    await fetch(GAS_URL, {
      method: 'POST',
      mode: 'no-cors',             // <-- évite CORS; la réponse est "opaque"
      redirect: 'follow',
      credentials: 'omit',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload)
    });

    log('✅ Requête envoyée (no-cors). Vérifie le Google Sheet (onglets "resultats" et "ML").');
    alert('Envoi effectué. Ouvre le Google Sheet pour vérifier.');
  } catch (e) {
    log('❌ ' + e.message);
    alert('Erreur : ' + e.message);
  } finally {
    setBusy(false);
    log('--- Fin ---');
  }
}

// ==================
// Lecture / parsing
// ==================
async function readWorkbook(file) {
  const data = await file.arrayBuffer();
  return XLSX.read(data, { type: 'array' });
}

// 1ère feuille → AOA ; supprime entête dupliquée ; supprime dernière ligne UNIQUEMENT si c’est un TOTAL
function cleanSheetToAOA(workbook) {
  const firstName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstName];
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  if (!aoa.length) return aoa;

  const header = aoa[0].map(x => (x==null?'':String(x).trim()));

  // entête dupliquée (ligne 2 == entête)
  if (aoa.length >= 2) {
    const firstDataRow = aoa[1].map(x => (x==null?'':String(x).trim()));
    if (arraysEqual(header, firstDataRow)) aoa.splice(1, 1);
  }

  // détecter si la dernière ligne est un TOTAL
  if (aoa.length >= 2) {
    const last = aoa[aoa.length - 1];
    const firstCell = (last[0] == null ? '' : String(last[0])).trim().toLowerCase();
    const looksLikeTotalLabel = /total|totaux|somme|sum|grand total|subtotal/.test(firstCell);

    let numericCount = 0, nonEmptyCount = 0;
    for (let i = 1; i < last.length; i++) {
      const v = last[i];
      if (v !== null && v !== '') nonEmptyCount++;
      const n = typeof v === 'number' ? v : Number(String(v).replace(/\s/g,'').replace(',', '.'));
      if (!isNaN(n)) numericCount++;
    }
    const looksLikeTotalsRow = looksLikeTotalLabel || (numericCount >= Math.max(1, nonEmptyCount - 1) && nonEmptyCount > 0);
    if (looksLikeTotalsRow) aoa.splice(aoa.length - 1, 1);
  }
  return aoa;
}
function arraysEqual(a,b){ return a.length===b.length && a.every((v,i)=>String(v).trim()===String(b[i]).trim()); }
function aoaToObjects(aoa){
  if(!aoa || !aoa.length) return [];
  const headers = aoa[0].map(h => String(h||'').trim());
  return aoa.slice(1).map(row => { const o={}; headers.forEach((h,i)=>o[h]=row[i]); return o; });
}

// ======================
// Traitements principaux
// ======================
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

  // Dédoublonnage (garde le 1er)
  const seen = new Set(); const dedupe = [];
  for (const r of rows) {
    const k = r[colKey];
    if (!seen.has(k)) { seen.add(k); dedupe.push(r); }
  }

  // Agrégation par jour (date = colDate tronquée AAAA-MM-JJ)
  const perDayMap = new Map(); const usersSet = new Set(); const daysSet = new Set();
  for (const r of dedupe) {
    const u = (r[colUser]==null ? '' : String(r[colUser])).trim(); // normalisation Contact
    usersSet.add(u);
    const d = parseDate(r[colDate]); if (!d) continue; daysSet.add(d);
    const val = parseNumber(r[colSum]); // parsing FR robuste
    const key = `${u}||${d}`;
    perDayMap.set(key, (perDayMap.get(key)||0) + val);
  }

  const days = Array.from(daysSet).sort();
  const headersOut = ['Contact', ...days.map(d => `nombre colonne carton ${d}`)];
  console.log('[DEBUG] dates détectées', days);

  const rowsOut = [];
  for (const u of Array.from(usersSet).sort()) {
    const row = [u];
    for (const d of days) row.push(perDayMap.get(`${u}||${d}`) || 0);
    rowsOut.push(row);
  }

  return { tableau: { headers: headersOut, rows: rowsOut } };
}

function emptyTable(){ return { headers:['Contact'], rows:[] } }

// =====================
// Outils de fusion/pivot
// =====================
function excelLetterToIndex(L) {
  L = String(L || '').trim().toUpperCase();
  let idx = 0;
  for (const ch of L) idx = idx * 26 + (ch.charCodeAt(0) - 64);
  return idx - 1;
}

function mergeTablesByContactAndHeaders(A, B) {
  const headers = Array.from(new Set([...(A.headers||[]), ...(B.headers||[])]));
  const ci = headers.indexOf('Contact'); if (ci>0){ headers.splice(ci,1); headers.unshift('Contact'); }

  const idxA = indexMap(A.headers||[]), idxB = indexMap(B.headers||[]);
  const norm = (s) => (s==null ? '' : String(s)).trim();

  const contacts = new Set([...(A.rows||[]).map(r=>norm(r[0])), ...(B.rows||[]).map(r=>norm(r[0]))]);
  const mapA = new Map(); (A.rows||[]).forEach(r => mapA.set(norm(r[0]), r));
  const mapB = new Map(); (B.rows||[]).forEach(r => mapB.set(norm(r[0]), r));

  const outRows = [];
  for (const c of Array.from(contacts).sort()) {
    const row = Array(headers.length).fill(0); row[0] = c;
    const ra = mapA.get(c), rb = mapB.get(c);
    if (ra) sumRowIntoParsed(row, ra, headers, idxA);
    if (rb) sumRowIntoParsed(row, rb, headers, idxB);
    for (let i=1;i<row.length;i++) row[i] = parseNumber(row[i]);
    outRows.push(row);
  }
  return { headers, rows: outRows };
}

function indexMap(h){ const m={}; (h||[]).forEach((name,i)=>m[name]=i); return m; }
function sumRowIntoParsed(targetRow, srcRow, headers, srcIdxMap){
  for (let i=1;i<headers.length;i++){
    const name = headers[i];
    const si = srcIdxMap[name];
    const v = (si==null) ? 0 : parseNumber(srcRow[si]);
    targetRow[i] = parseNumber(targetRow[i]) + v;
  }
}

function multiplyValues(table, k){
  const headers = table.headers.slice();
  const rows = table.rows.map(r=>{
    const out = r.slice();
    for (let i=1;i<out.length;i++) out[i] = parseNumber(out[i]) * k;
    return out;
  });
  return { headers, rows };
}

// ===================
// Parsing nombres/dates
// ===================
function parseNumber(v){
  if (v == null || v === '') return 0;
  if (typeof v === 'number' && isFinite(v)) return v;
  let s = String(v).trim();
  s = s.replace(/[\u202F\u00A0\s']/g, ''); // espaces fines, insécables, apostrophes
  const lastComma = s.lastIndexOf(',');
  const lastDot = s.lastIndexOf('.');
  if (lastComma > lastDot) { s = s.replace(/\./g, '').replace(',', '.'); }
  else if (lastDot > lastComma) { s = s.replace(/,/g, ''); }
  else { if (s.includes(',')) s = s.replace(',', '.'); }
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function parseDate(v) {
  if (v == null || v === '') return null;

  // Numéro Excel
  if (typeof v === 'number' && isFinite(v)) {
    const ms = Math.round((v - 25569) * 86400 * 1000); // base Excel 1899-12-30
    const d = new Date(ms);
    if (!isNaN(d)) return fmtYMD(d);
  }

  // ISO-like en local
  if (typeof v === 'string') {
    let m = v.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2})(?::(\d{2})(?::(\d{2}))?)?)?/);
    if (m) {
      const [_, YYYY, MM, DD, hh='0', mm='0', ss='0'] = m;
      const d = new Date(Number(YYYY), Number(MM)-1, Number(DD), Number(hh), Number(mm), Number(ss));
      if (!isNaN(d)) return fmtYMD(d);
    }
    // JJ/MM/AAAA
    m = v.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) {
      const [_, DD, MM, YYYY, hh='0', mm='0', ss='0'] = m;
      const d = new Date(Number(YYYY), Number(MM)-1, Number(DD), Number(hh), Number(mm), Number(ss));
      if (!isNaN(d)) return fmtYMD(d);
    }
  }

  // fallback
  const d = new Date(v);
  return isNaN(d) ? null : fmtYMD(d);
}
function fmtYMD(d){ const y=d.getFullYear(), m=String(d.getMonth()+1).padStart(2,'0'), day=String(d.getDate()).padStart(2,'0'); return `${y}-${m}-${day}`; }

// Expose global (fallback onclick dans index.html)
window.onRun = onRun;
window.testConnexion = testConnexion;
