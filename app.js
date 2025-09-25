// Utils
const logEl = document.getElementById('log');
const log = (m) => { logEl.textContent += m + '\n'; };
const excelLetterToIndex = (L) => {
  L = String(L || '').trim().toUpperCase();
  let idx = 0; for (const ch of L) idx = idx * 26 + (ch.charCodeAt(0) - 64);
  return idx - 1;
};
const parseDate = (v) => {
  // force AAAA-MM-JJ
  const d = new Date(v);
  if (isNaN(d)) return null;
  const y = d.getFullYear(), m = (''+(d.getMonth()+1)).padStart(2,'0'), dd=(''+d.getDate()).padStart(2,'0');
  return `${y}-${m}-${dd}`;
};

const SHEET_ID = '1AptbV2NbY0WQZpe_Xt1K2iVlDpgKADElamKQCg3GcXQ';

document.getElementById('runBtn').addEventListener('click', async () => {
  try {
    const sFile = document.getElementById('suiviFile').files[0];
    const eFile = document.getElementById('extractFile').files[0];
    const gasUrl = document.getElementById('gasUrl').value.trim();
    const secret = document.getElementById('secret').value.trim();

    if (!sFile || !eFile) return alert('Sélectionne les deux fichiers .xlsx');
    if (!gasUrl) return alert('Renseigne l’URL du déploiement Apps Script (Web App).');

    log('Lecture fichiers… (dans le navigateur)');
    const sWorkbook = await readWorkbook(sFile);
    const eWorkbook = await readWorkbook(eFile);

    const sData = processWorkbook(sWorkbook, {
      key: 's_key', user: 's_user', sum: 's_sum', date: 's_date', head: 's_head'
    });
    const eData = processWorkbook(eWorkbook, {
      key: 'e_key', user: 'e_user', sum: 'e_sum', date: 'e_date', head: 'e_head'
    });

    log('Calcul resultats (suivi + extraction)…');
    const resultats = mergeTablesByContactAndHeaders(sData.tableau, eData.tableau);   // addition
    const ml = multiplyValues(resultats, 0.35);

    // Envoi vers Google Apps Script
    log('Envoi vers Google Sheets (Apps Script)…');
    const payload = {
      secret,
      sheetId: SHEET_ID,
      // nous écrirons dans 2 onglets: 'resultats' et 'ML'
      resultats: { headers: resultats.headers, rows: resultats.rows },
      ml: { headers: ml.headers, rows: ml.rows }
    };
    const r = await fetch(gasUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const j = await r.json();
    if (!r.ok) throw new Error(j.error || r.statusText);
    log('✅ Écriture terminée : ' + JSON.stringify(j));
    alert('Terminé ! Vérifie le Google Sheet.');
  } catch (e) {
    console.error(e);
    log('❌ ' + e.message);
    alert('Erreur : ' + e.message);
  }
});

// ---- Lecture / nettoyage / calculs ----
async function readWorkbook(file) {
  const data = await file.arrayBuffer();
  return XLSX.read(data, { type: 'array' });
}

// Prend la 1ère feuille, supprime (optionnel) ligne d’entête dupliquée, supprime la dernière ligne
function cleanSheetToAOA(workbook) {
  const firstName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstName];
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  if (aoa.length === 0) return aoa;
  const header = aoa[0].map(x => (x==null?'':String(x).trim()));
  // si la 1ère ligne de données == entête, supprime-la
  if (aoa.length >= 2) {
    const firstDataRow = aoa[1].map(x => (x==null?'':String(x).trim()));
    if (arraysEqual(header, firstDataRow)) aoa.splice(1,1);
  }
  // supprime la dernière ligne
  if (aoa.length >= 2) aoa.splice(aoa.length-1, 1);
  return aoa;
}
const arraysEqual = (a,b) => a.length===b.length && a.every((v,i)=>String(v).trim()===String(b[i]).trim());

// Convertit AOA → tableau d’objets avec entêtes
function aoaToObjects(aoa) {
  if (!aoa || !aoa.length) return [];
  const headers = aoa[0].map(h => String(h||'').trim());
  return aoa.slice(1).map(row => {
    const o = {};
    headers.forEach((h,i)=>o[h]=row[i]);
    return o;
  });
}

function processWorkbook(workbook, refs) {
  const aoa = cleanSheetToAOA(workbook);
  const rows = aoaToObjects(aoa);
  if (!rows.length) return { tableau: emptyTable() };

  const headers = aoa[0];
  const colByLetter = (L) => headers[excelLetterToIndex(document.getElementById(refs[L]).value)];
  const colKey  = colByLetter('key');
  const colUser = colByLetter('user');
  const colSum  = colByLetter('sum');
  const colDate = colByLetter('date');
  const colHead = colByLetter('head'); // pour le label "nombre colonne carton <date>"

  // dédoublonnage par clé (garde 1er)
  const seen = new Set();
  const dedupe = [];
  for (const r of rows) {
    const k = r[colKey];
    if (!seen.has(k)) { seen.add(k); dedupe.push(r); }
  }

  // somme par utilisateur
  // (on ne l’envoie pas, mais on garde la logique identique au script)
  // label de la colonne somme avec date(s) issues de colHead
  const datesHead = uniqueDates(dedupe.map(r => r[colHead]));
  const labelSum = 'nombre colonne carton' + (datesHead.length===1 ? ` ${datesHead[0]}` :
                     datesHead.length>1 ? ` ${datesHead[0]}–${datesHead[datesHead.length-1]}` : '');

  // synthèse par jour (date = colDate tronquée au jour)
  const perDayMap = new Map(); // key = user|date → sum
  const usersSet = new Set();
  const daysSet = new Set();
  for (const r of dedupe) {
    const u = r[colUser];
    usersSet.add(u);
    const d = parseDate(r[colDate]);
    if (!d) continue;
    daysSet.add(d);
    const val = Number(r[colSum] ?? 0) || 0;
    const key = `${u}||${d}`;
    perDayMap.set(key, (perDayMap.get(key)||0) + val);
  }
  const days = Array.from(daysSet).sort();
  // pivot Contact × dates
  const headersOut = ['Contact', ...days.map(d => `nombre colonne carton ${d}`)];
  const rowsOut = [];
  for (const u of Array.from(usersSet).sort()) {
    const row = [u];
    for (const d of days) {
      const key = `${u}||${d}`;
      row.push(perDayMap.get(key) || 0);
    }
    rowsOut.push(row);
  }
  return { tableau: { headers: headersOut, rows: rowsOut } };
}

function uniqueDates(values) {
  const s = new Set();
  for (const v of values) {
    const d = parseDate(v);
    if (d) s.add(d);
  }
  return Array.from(s).sort();
}

function emptyTable(){ return { headers:['Contact'], rows:[] } }

// addition au nom de colonne & contact
function mergeTablesByContactAndHeaders(A, B) {
  const headers = Array.from(new Set([...(A.headers||[]), ...(B.headers||[])]));
  // assure "Contact" en première position
  const contactIdx = headers.indexOf('Contact');
  if (contactIdx > 0) { headers.splice(contactIdx,1); headers.unshift('Contact'); }

  const idxMapA = indexMap(A.headers||[]);
  const idxMapB = indexMap(B.headers||[]);

  const contacts = new Set();
  for (const r of (A.rows||[])) contacts.add(r[0]);
  for (const r of (B.rows||[])) contacts.add(r[0]);
  const outRows = [];
  for (const c of Array.from(contacts).sort()) {
    const row = Array(headers.length).fill(0);
    row[0] = c;
    // A
    const ra = (A.rows||[]).find(r => r[0]===c);
    if (ra) sumRowInto(row, ra, headers, idxMapA);
    // B
    const rb = (B.rows||[]).find(r => r[0]===c);
    if (rb) sumRowInto(row, rb, headers, idxMapB);
    outRows.push(row);
  }
  // remplace NaN par 0
  for (const r of outRows) for (let i=1;i<r.length;i++) r[i] = Number(r[i]||0);
  return { headers, rows: outRows };
}
function indexMap(h){ const m={}; (h||[]).forEach((name,i)=>m[name]=i); return m; }
function sumRowInto(targetRow, srcRow, headers, srcIdxMap){
  for (let i=1;i<headers.length;i++){
    const name = headers[i];
    const si = srcIdxMap[name];
    const v = (si==null) ? 0 : Number(srcRow[si]||0);
    targetRow[i] = Number(targetRow[i]||0) + (isNaN(v)?0:v);
  }
}
function multiplyValues(table, k){
  const headers = table.headers.slice();
  const rows = table.rows.map(r=>{
    const out = r.slice();
    for (let i=1;i<out.length;i++) out[i] = Number(out[i]||0)*k;
    return out;
  });
  return { headers, rows };
}
