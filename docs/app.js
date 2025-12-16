/*
  client-side app.js - Router + page logic + Excel generation
  This file ports the original Flask + server-side processing into a client-only implementation.
  It preserves HTML structure and CSS from the original app and reproduces reconciliation logic in JS.
*/

// Simple hash router that loads partials from /pages
const routes = {
  '/': 'pages/start.html',
  '/upload': 'pages/reconciliation.html',
  '/cats-edits': 'pages/cats-edits.html'
};

function loadRoute() {
  const path = location.hash.replace('#', '') || '/';
  const route = routes[path] || routes['/'];
  fetch(route).then(r => r.text()).then(html => {
    document.getElementById('app').innerHTML = html;
    // After injecting page, run page-specific init
    if (path === '/') initStartPage();
    if (path === '/upload') initReconciliationPage();
    if (path === '/cats-edits') initCatsEditsPage();
  }).catch(err => {
    document.getElementById('app').innerHTML = '<p>Error loading page.</p>';
    console.error(err);
  });
}

window.addEventListener('hashchange', loadRoute);
window.addEventListener('DOMContentLoaded', loadRoute);

// Shared state (stored in sessionStorage so it persists across reloads in browser tab)
let drmisFileData = null; // array of rows (arrays)
let oracleFileData = null;
let resultData = null; // reconciliation result (headers,data,message,count)
let catsEditsData = null; // generated cats edits
// Optional whitelist of leave codes loaded from a local Excel file (if present)
let leaveCodeWhitelist = null; // Set of strings
// Lookup map built from the raw DRMIS sheet rows: key = YYYY-MM-DD|PersNo -> array of entries
let drmisLookup = new Map();

function saveStateToSession() {
  const state = {
    drmisLoaded: !!drmisFileData,
    oracleLoaded: !!oracleFileData,
    resultData: resultData,
    catsEditsData: catsEditsData
  };
  sessionStorage.setItem('leave_reconcile_state', JSON.stringify(state));
}

function loadStateFromSession() {
  const raw = sessionStorage.getItem('leave_reconcile_state');
  if (!raw) return;
  try {
    const s = JSON.parse(raw);
    if (s && s.resultData) resultData = s.resultData;
    if (s && s.catsEditsData) catsEditsData = s.catsEditsData;
  } catch (e) {
    console.warn('Failed to parse session state', e);
  }
}

/* -------------------- Utilities: reading Excel and parsing -------------------- */
function readExcelFile(file, type) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const range = firstSheet['!ref'] ? XLSX.utils.decode_range(firstSheet['!ref']) : {s:{r:0,c:0},e:{r:0,c:0}};
        const jsonData = [];
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const row = [];
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({r: R, c: C});
                const cell = firstSheet[cellAddress];
                row.push(cell ? cell.v : null);
            }
            jsonData.push(row);
        }
        resolve(jsonData);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function sheetRowsToObjects(rows) {
  if (!rows || rows.length === 0) return { headers: [], rows: [] };
  // Robust header detection:
  // Some exported sheets include top metadata rows, blank columns, or title rows before the real header.
  // Strategy: examine the first N rows, score each row by number of non-empty cells and presence of header keywords,
  // and pick the best candidate as the header row. Then trim leading/trailing mostly-empty columns.
  const maxScan = Math.min(30, rows.length);
  const headerKeywords = ['pers','pers.no','pers no','pers_no','date','hours','a/atype','aatype','a a type','a a','leave','leave code','a/a type','a/atype'];
  let bestIdx = 0;
  let bestScore = -1;
  for (let i = 0; i < maxScan; i++) {
    const r = rows[i] || [];
    const norm = r.map(c => (c === null || c === undefined) ? '' : String(c).trim().toLowerCase());
    const nonEmpty = norm.filter(x => x !== '');
    const hasKeyword = norm.some(x => headerKeywords.some(k => x.includes(k)));
    const score = nonEmpty.length + (hasKeyword ? 50 : 0);
    if (score > bestScore) { bestScore = score; bestIdx = i; }
  }

  // Determine column bounds: skip leading columns that are mostly empty
  const totalRows = rows.length - (bestIdx + 1);
  const colCount = rows[bestIdx] ? rows[bestIdx].length : 0;
  const colNonEmpty = new Array(colCount).fill(0);
  for (let r = bestIdx + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < colCount; c++) {
      if (row[c] !== null && row[c] !== undefined && String(row[c]).toString().trim() !== '') colNonEmpty[c]++;
    }
  }
  // threshold: consider a column present if it has at least max(3, 15% of data rows) non-empty values
  const threshold = Math.max(3, Math.ceil((totalRows || 0) * 0.15));
  let firstCol = 0; while (firstCol < colCount && colNonEmpty[firstCol] < threshold) firstCol++;
  let lastCol = colCount - 1; while (lastCol >= firstCol && colNonEmpty[lastCol] < threshold) lastCol--;
  if (firstCol > lastCol) { firstCol = 0; lastCol = colCount - 1; }

  // build headers from detected header row trimmed to bounds
  const rawHeaderRow = rows[bestIdx] || [];
  const headers = [];
  for (let c = firstCol; c <= lastCol; c++) {
    const h = rawHeaderRow[c];
    headers.push(h !== null && h !== undefined ? String(h).trim() : '');
  }

  // build objects from subsequent rows (rows after header)
  const objs = [];
  for (let r = bestIdx + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    // skip rows that are completely blank in the header columns
    let any = false;
    for (let c = firstCol; c <= lastCol; c++) {
      const v = row[c]; if (v !== null && v !== undefined && String(v).toString().trim() !== '') { any = true; break; }
    }
    if (!any) continue;
    const obj = {};
    for (let ci = firstCol, hi = 0; ci <= lastCol; ci++, hi++) {
      obj[headers[hi]] = row[ci];
    }
    objs.push(obj);
  }
  return { headers, rows: objs };
}

// Build a lookup map from raw DRMIS sheet objects so we can prefill CATs editable rows.
function buildDrmisLookup(rowsObjects) {
  drmisLookup = new Map();
  if (!rowsObjects || rowsObjects.length === 0) return;
  const keys = Object.keys(rowsObjects[0] || {});
  const lowerMap = {};
  keys.forEach(k => { lowerMap[String(k).toLowerCase().trim()] = k; });
  // heuristics for common column names
  const persKey = Object.keys(lowerMap).find(k => k.includes('pers') && k.includes('no')) ? lowerMap[Object.keys(lowerMap).find(k => k.includes('pers') && k.includes('no'))] : (lowerMap['pers.no']||lowerMap['pers']||null);
  const dateKey = lowerMap['date'] || Object.keys(lowerMap).find(k => k.includes('date')) ? lowerMap[Object.keys(lowerMap).find(k => k.includes('date'))] : null;
  const recKey = Object.keys(lowerMap).find(k => k.includes('rec') && k.includes('order')) ? lowerMap[Object.keys(lowerMap).find(k => k.includes('rec') && k.includes('order'))] : (lowerMap['rec.order']||null);
  const actKey = Object.keys(lowerMap).find(k => k.includes('act')) ? lowerMap[Object.keys(lowerMap).find(k => k.includes('act'))] : null;
  const aatypeKey = Object.keys(lowerMap).find(k => k.includes('a/a') && k.includes('type')) ? lowerMap[Object.keys(lowerMap).find(k => k.includes('a/a') && k.includes('type'))] : (lowerMap['a/atype']||null);
  const hoursKey = Object.keys(lowerMap).find(k => k==='hours') ? 'Hours' : (lowerMap['hours']||null);

  rowsObjects.forEach(r => {
    const dateVal = r[dateKey];
    const d = normalizeDate(dateVal);
    if (!d) return;
    const dateKeyStr = d.toISOString().split('T')[0];
    const pers = (persKey && r[persKey]) ? String(r[persKey]).trim() : '';
    const recOrder = recKey && r[recKey] ? String(r[recKey]).trim() : '';
    const act = actKey && r[actKey] ? String(r[actKey]).trim() : '';
    let aatype = aatypeKey && r[aatypeKey] ? String(r[aatypeKey]).trim() : '';
    if (aatype === '' || aatype === '-') aatype = '0';
    const hours = hoursKey && r[hoursKey] !== undefined && r[hoursKey] !== null ? (Number(r[hoursKey]) || 0) : 0;
    const entry = { recOrder, act, aatype, hours };
    const mapKey = dateKeyStr + '|' + String(pers || '');
    if (!drmisLookup.has(mapKey)) drmisLookup.set(mapKey, []);
    drmisLookup.get(mapKey).push(entry);
  });
}

// Normalize date similar to Python normalize_date
function normalizeDate(value) {
  if (value === null || value === undefined || value === '') return null;
  if (value instanceof Date) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  // Excel date serial (number)
  if (typeof value === 'number') {
    // Excel uses 1900-date system by default (offset from 1899-12-30)
    const epoch = new Date(Date.UTC(1899,11,30));
    const ms = (value - 0) * 24*60*60*1000; // treat as days
    const d = new Date(epoch.getTime() + ms);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  if (typeof value === 'string') {
    const s = value.trim();
    // try known formats
    const fmts = [/^(\d{1,2})[.](\d{1,2})[.](\d{4})$/, /^(\d{1,2})-(\d{1,2})-(\d{4})$/, /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, /^(\d{4})-(\d{1,2})-(\d{1,2})$/];
    for (let f of fmts) {
      const m = s.match(f);
      if (m) {
        let y,mn,day;
        if (f === fmts[0] || f === fmts[1]) { day = parseInt(m[1]); mn = parseInt(m[2]); y = parseInt(m[3]); }
        else if (f === fmts[2]) { mn = parseInt(m[1]); day = parseInt(m[2]); y = parseInt(m[3]); }
        else { y = parseInt(m[1]); mn = parseInt(m[2]); day = parseInt(m[3]); }
        return new Date(y, mn-1, day);
      }
    }
    // fallback to Date.parse
    const parsed = Date.parse(s);
    if (!isNaN(parsed)) return new Date(new Date(parsed).getFullYear(), new Date(parsed).getMonth(), new Date(parsed).getDate());
  }
  return null;
}

function isBusinessDay(dateObj) {
  if (!dateObj) return false;
  // Canadian Federal holidays for 2025 as in original app
  const holidays2025 = [
    new Date(2025,0,1).toDateString(),
    new Date(2025,3,18).toDateString(),
    new Date(2025,4,19).toDateString(),
    new Date(2025,6,1).toDateString(),
    new Date(2025,7,4).toDateString(),
    new Date(2025,8,1).toDateString(),
    new Date(2025,8,30).toDateString(),
    new Date(2025,9,13).toDateString(),
    new Date(2025,10,11).toDateString(),
    new Date(2025,11,25).toDateString(),
    new Date(2025,11,26).toDateString()
  ];
  if (holidays2025.includes(dateObj.toDateString())) return false;
  const day = dateObj.getDay(); // 0=Sun .. 6=Sat
  return day >=1 && day <=5;
}

function expandOracleEntries(oracleRows) {
  const expanded = [];
  oracleRows.forEach(row => {
    let hours = Number(row['Oracle Hours']) || 0;
    let date = row['Date'];
    const leaveCode = row['Oracle Leave Code'];
    if (!date) return;
    if (hours <= 8) {
      expanded.push({...row, 'Oracle Hours': hours, 'Date': date});
    } else {
      let remaining = hours;
      let currentDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      while (remaining > 0) {
        // find next business day
        while (!isBusinessDay(currentDate)) {
          currentDate = new Date(currentDate.getTime() + 24*60*60*1000);
        }
        const dayHours = Math.min(8, remaining);
        expanded.push({'Date': new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate()), 'Oracle Hours': dayHours, 'Oracle Leave Code': leaveCode});
        remaining -= dayHours;
        currentDate = new Date(currentDate.getTime() + 24*60*60*1000);
      }
    }
  });
  return expanded;
}

/* -------------------- Leave Codes whitelist loader -------------------- */
async function loadLeaveCodeWhitelist() {
  // Try several candidate paths relative to the server root where the app is served.
  const candidates = [
    './sample_data/Leave_Codes - Actual Leave.xlsx'
  ];
  for (let p of candidates) {
    try {
      const url = encodeURI(p);
      const resp = await fetch(url, { method: 'GET' });
      if (!resp.ok) continue;
      const ab = await resp.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(ab), { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      if (!sheet) continue;
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      if (!rows || rows.length === 0) continue;
      const headers = rows[0].map(h => h === null || h === undefined ? '' : String(h).trim());
      // find a header matching A/AType variants (case-insensitive)
      const idx = headers.findIndex(h => /a\/?a?type/i.test(h.replace(/\s+/g,'')) || /a\/atype/i.test(h.toLowerCase()));
      const colIndex = idx >= 0 ? idx : headers.findIndex(h => /a\/?atype/i.test(h.toLowerCase()));
      const codes = new Set();
      if (colIndex >= 0) {
        for (let i = 1; i < rows.length; i++) {
          const v = rows[i][colIndex];
          if (v === null || v === undefined) continue;
          const s = String(v).trim();
          if (s === '') continue;
          // normalize to digits-only string for consistent matching
          const norm = String(s).replace(/\D+/g, '');
          if (norm === '') continue;
          codes.add(norm);
        }
      } else {
        // fallback: try first column
        for (let i = 1; i < rows.length; i++) {
          const v = rows[i][0];
          if (v === null || v === undefined) continue;
          const s = String(v).trim();
          if (s === '') continue;
          const norm = String(s).replace(/\D+/g, '');
          if (norm === '') continue;
          codes.add(norm);
        }
      }
      if (codes.size > 0) {
        leaveCodeWhitelist = codes;
        console.info('Loaded leave-code whitelist from', p, 'entries:', codes.size);
        return;
      }
    } catch (e) {
      // ignore and try next candidate
      // console.debug('leave whitelist load failed for', p, e);
    }
  }
  // if none found, leave as null
  leaveCodeWhitelist = null;
}

function applyColumnMappingToObjects(objects, mappings) {
  // mappings: { 'Pers No': index, 'Date': index, ... } mapping to column index in original header
  // If mappings not provided, return objects unchanged
  if (!mappings) return objects;
  // build array of keys from first object
  const keys = Object.keys(objects[0] || {});
  const renamed = objects.map(obj => {
    const out = {};
    Object.entries(mappings).forEach(([requiredName, idx]) => {
      if (idx < keys.length) {
        const actualKey = keys[idx];
        if (!actualKey) return;
        if (requiredName === 'Pers No') out['Pers No'] = obj[actualKey];
        else if (requiredName === 'Date') out['Date'] = obj[actualKey];
        else if (requiredName === 'Hours') out['Hours'] = obj[actualKey];
        else if (requiredName === 'A/A Type') out['A/AType'] = obj[actualKey];
        else if (requiredName === 'From Date') out['Date'] = obj[actualKey];
        else if (requiredName === 'Hours Recorded') out['Hours Recorded'] = obj[actualKey];
        else if (requiredName === 'Leave Code') out['Leave Code'] = obj[actualKey];
        else out[actualKey] = obj[actualKey];
      }
    });
    return out;
  });
  return renamed;
}

/* -------------------- Extraction functions ported from Python -------------------- */
function extractDrmisData(rowsObjects) {
  // rowsObjects: array of objects keyed by header strings
  if (!rowsObjects || rowsObjects.length === 0) return [];
  // keys lower map
  const first = rowsObjects[0];
  const keys = Object.keys(first);
  const lowerMap = {};
  keys.forEach(k => lowerMap[k.toLowerCase().trim()] = k);

  const persNoKey = Object.keys(lowerMap).find(k => k.includes('pers') && k.includes('no')) ? lowerMap[Object.keys(lowerMap).find(k => k.includes('pers') && k.includes('no'))] : null;
  const dateKey = lowerMap['date'] || null;
  const hoursKey = lowerMap['hours'] || null;
  const atypeKey = Object.keys(lowerMap).find(k => k.includes('a/a') && k.includes('type') && !k.includes('text')) ? lowerMap[Object.keys(lowerMap).find(k => k.includes('a/a') && k.includes('type') && !k.includes('text'))] : null;

  const missing = [];
  if (!persNoKey) missing.push('Pers No');
  if (!dateKey) missing.push('Date');
  if (!hoursKey) missing.push('Hours');
  if (!atypeKey) missing.push('A/A Type');
  if (missing.length>0) throw new Error('DRMIS file is missing required columns: ' + missing.join(', ') + '. Available columns: ' + keys.join(', '));

  const out = [];
  function normalizeLeaveCodeRaw(leaveRaw) {
    if (leaveRaw === null || leaveRaw === undefined || String(leaveRaw).trim() === '' || String(leaveRaw).trim() === '-') {
      return '0';
    }
    // remove non-digits
    const s = String(leaveRaw).trim();
    const digits = s.replace(/\D+/g, '');
    return digits === '' ? String(s) : digits;
  }
  rowsObjects.forEach(r => {
    const date = normalizeDate(r[dateKey]);
    if (!date) return; // skip rows with no date
    const pers = r[persNoKey];
    const drmisHours = Number(r[hoursKey]) || 0;
    let leaveRaw = r[atypeKey];
    const leaveStr = normalizeLeaveCodeRaw(leaveRaw);
    // filter out codes starting with 30
    if (String(leaveStr).startsWith('30')) return;
    // If a whitelist is loaded, only keep rows whose normalized code is in the whitelist
    if (leaveCodeWhitelist && leaveCodeWhitelist.size>0) {
      const matchCode = (leaveStr === '0') ? '0' : String(leaveStr).replace(/\D+/g,'');
      if (!leaveCodeWhitelist.has(matchCode)) return;
    }
    out.push({ 'Pers No': pers, 'Date': date, 'Drmis Hours': drmisHours, 'Drmis Leave Code': leaveStr });
  });
  return out;
}

function extractOracleData(rowsObjects) {
  if (!rowsObjects || rowsObjects.length === 0) return [];
  const keys = Object.keys(rowsObjects[0]);
  const lowerMap = {};
  keys.forEach(k => lowerMap[k.toLowerCase().trim()] = k);
  const fromDateKey = Object.keys(lowerMap).find(k => k.includes('from') && k.includes('date')) ? lowerMap[Object.keys(lowerMap).find(k => k.includes('from') && k.includes('date'))] : (lowerMap['date'] || null);
  const hoursKey = Object.keys(lowerMap).find(k => k.includes('hours') && k.includes('recorded')) ? lowerMap[Object.keys(lowerMap).find(k => k.includes('hours') && k.includes('recorded'))] : (lowerMap['hours']||null);
  const leaveCodeKey = Object.keys(lowerMap).find(k => k.includes('leave') && k.includes('code')) ? lowerMap[Object.keys(lowerMap).find(k => k.includes('leave') && k.includes('code'))] : null;

  const missing = [];
  if (!fromDateKey) missing.push('From Date');
  if (!hoursKey) missing.push('Hours Recorded');
  if (!leaveCodeKey) missing.push('Leave Code');
  if (missing.length>0) throw new Error('ORACLE file is missing required columns: ' + missing.join(', ') + '. Available columns: ' + keys.join(', '));

  const out = [];
  function normalizeLeaveCodeRawOracle(codeRaw) {
    if (codeRaw === null || codeRaw === undefined || String(codeRaw).trim() === '' || String(codeRaw).trim() === '-') {
      return '0';
    }
    const s = String(codeRaw).trim();
    const digits = s.replace(/\D+/g, '');
    return digits === '' ? String(s) : digits;
  }
  rowsObjects.forEach(r => {
    const d = normalizeDate(r[fromDateKey]);
    if (!d) return;
    const hours = Number(r[hoursKey]) || 0;
    let code = r[leaveCodeKey];
    const norm = normalizeLeaveCodeRawOracle(code);
    // Normalize Oracle code to DRMiS comparable form:
    // Oracle exports use 3-digit codes (e.g., '110'); for comparison we convert to '1' + code -> '1110'
    let oracleKey = '0';
    if (norm === '0') {
      oracleKey = '0';
    } else if (String(norm).length === 3) {
      oracleKey = '1' + String(norm);
    } else if (String(norm).length === 4) {
      oracleKey = String(norm);
    } else {
      // fallback: if it's other length, keep digits as-is
      oracleKey = String(norm);
    }
    // filter codes (retain previous hard-coded exclusions for compatibility)
    const filtered = ['1200','1260','1261','1660'];
    if (filtered.includes(String(oracleKey))) return;
    // whitelist filtering: if present, only keep rows whose normalized transformed code is in whitelist
    if (leaveCodeWhitelist && leaveCodeWhitelist.size>0) {
      if (oracleKey !== '0' && !leaveCodeWhitelist.has(String(oracleKey))) return;
      if (oracleKey === '0') return; // Oracle rows with missing code are not part of whitelist
    }
    out.push({ 'Date': d, 'Oracle Hours': hours, 'Oracle Leave Code': String(oracleKey) });
  });

  // expand multi-day
  const expanded = expandOracleEntries(out);
  return expanded;
}

function reconcileData(drmisData, oracleData) {
  // drmisData: array of objects with Pers No, Date (Date obj), Drmis Hours, Drmis Leave Code
  // oracleData: array of objects with Date (Date obj), Oracle Hours, Oracle Leave Code
  // Build a key by date string and pers no
  const map = new Map();

  // Determine a default Pers No to apply to Oracle-only rows. The app reconciles
  // leave for one person at a time: use the first non-empty Pers No from the DRMIS
  // data as the default for Oracle rows that don't include Pers No.
  let defaultPers = '';
  if (Array.isArray(drmisData) && drmisData.length > 0) {
    const firstWithPers = drmisData.find(r => r && r['Pers No'] !== undefined && r['Pers No'] !== null && String(r['Pers No']).toString().trim() !== '');
    if (firstWithPers) defaultPers = String(firstWithPers['Pers No']);
  }

  function keyFor(date, pers) {
    const d = date instanceof Date ? date.toISOString().split('T')[0] : String(date);
    return d + '|' + String(pers || '');
  }

  drmisData.forEach(r => {
    const k = keyFor(r.Date, r['Pers No']);
    if (!map.has(k)) map.set(k, { Date: r.Date, 'Pers No': r['Pers No'], 'Drmis Hours': r['Drmis Hours'], 'Drmis Leave Code': r['Drmis Leave Code'], 'Oracle Hours': 0, 'Oracle Leave Code': '' });
    else {
      // keep first Drmis Hours and Leave Code if duplicates
    }
  });

  oracleData.forEach(r => {
    // match by date; if multiple pers, fill empty pers
    // We'll find existing keys with same date
    // For simplicity place Oracle rows by date and if there's an existing pers fill that, otherwise create new row with empty pers
    const dateStr = r.Date.toISOString().split('T')[0];
    // find any key starting with dateStr + '|'
    const foundKey = Array.from(map.keys()).find(k => k.startsWith(dateStr + '|'));
    if (foundKey) {
      const entry = map.get(foundKey);
      entry['Oracle Hours'] = r['Oracle Hours'];
      entry['Oracle Leave Code'] = r['Oracle Leave Code'];
      map.set(foundKey, entry);
    } else {
      // Use defaultPers if available; otherwise leave Pers No empty
      const persForNew = defaultPers || '';
      const k = dateStr + '|' + persForNew;
      map.set(k, { Date: r.Date, 'Pers No': persForNew, 'Drmis Hours': 0, 'Drmis Leave Code': '0', 'Oracle Hours': r['Oracle Hours'], 'Oracle Leave Code': r['Oracle Leave Code'] });
    }
  });

  // Build merged array from map
  const merged = Array.from(map.values()).map(m => ({
    Date: m.Date,
    // Ensure Pers No falls back to defaultPers when empty
    'Pers No': (m['Pers No'] === '' || m['Pers No'] === undefined || m['Pers No'] === null) ? (defaultPers || '') : (Number(m['Pers No']) || m['Pers No']),
    'Oracle Leave Code': m['Oracle Leave Code'] || '',
    'Oracle Hours': m['Oracle Hours'] || 0,
    'Drmis Leave Code': m['Drmis Leave Code'] || '0',
    'Drmis Hours': m['Drmis Hours'] || 0
  }));

  // convert Drmis leave code to 4-digit string if numeric
  merged.forEach(row => {
    // Represent missing/empty Drmis leave codes as '0'. If numeric, keep numeric string (no zero-padding).
    if (row['Drmis Leave Code'] === null || row['Drmis Leave Code'] === undefined || String(row['Drmis Leave Code']).trim() === '') {
      row['Drmis Leave Code'] = '0';
    } else if (row['Drmis Leave Code'] !== 0) {
      const num = parseInt(row['Drmis Leave Code']);
      if (!isNaN(num)) row['Drmis Leave Code'] = String(num);
    }
    // OracleLeaveCode leave as string
  });

  // Find mismatches where codes or hours differ
  // If a leave-code whitelist is present, filter merged rows to only those codes
  let filteredMerged = merged;
  if (leaveCodeWhitelist && leaveCodeWhitelist.size>0) {
    filteredMerged = merged.filter(r => {
      const dr = String(r['Drmis Leave Code'] || '').trim();
      const orc = String(r['Oracle Leave Code'] || '').trim();
      return leaveCodeWhitelist.has(dr) || leaveCodeWhitelist.has(orc);
    });
  }

  const mismatches = filteredMerged.filter(r => (String(r['Drmis Leave Code']) !== String(r['Oracle Leave Code'])) || (Number(r['Drmis Hours']) !== Number(r['Oracle Hours'])) ).map(r => ({
    Date: r.Date,
    'Pers No': r['Pers No'],
    'Oracle Leave Code': r['Oracle Leave Code'],
    'Oracle Hours': r['Oracle Hours'],
    'Drmis Leave Code': r['Drmis Leave Code'],
    'Drmis Hours': r['Drmis Hours'],
    'Add to CATs Edits': true
  }));

  // Sort by date
  mismatches.sort((a,b)=> new Date(a.Date) - new Date(b.Date));
  return mismatches;
}

function generateCatsEdits(mismatches) {
  const cats = [];
  for (let row of mismatches) {
    const reasons = [];
    if (String(row['Oracle Leave Code']) !== String(row['Drmis Leave Code'])) reasons.push('Code Mismatch');
    if (Number(row['Oracle Hours']) !== Number(row['Drmis Hours'])) reasons.push('Hours Mismatch');
    const discrepancy = reasons.join(', ');
    const dateStr = row.Date instanceof Date ? row.Date.toDateString() : String(row.Date);

    // Determine aa_code value
    let aa_code_value = '';
    // New rule: if Oracle Leave Code is empty/0 AND Oracle Hours is empty/0,
    // then use Drmis Leave Code for AA Code and set hours to 0. This signals
    // Drmis to remove the existing leave by setting its hours to 0.
    const oracleCode = row['Oracle Leave Code'];
    const oracleHoursRaw = row['Oracle Hours'];
    const oracleHoursNum = Number(oracleHoursRaw) || 0;
    let hoursValue = oracleHoursNum;
    if ((oracleCode === undefined || oracleCode === null || String(oracleCode).trim() === '' || String(oracleCode) === '0') && (oracleHoursRaw === undefined || oracleHoursRaw === null || Number(oracleHoursRaw) === 0)) {
      // Use Drmis code (if present) and set hours to 0
      if (row['Drmis Leave Code'] && String(row['Drmis Leave Code']) !== '0') {
        const num = parseInt(row['Drmis Leave Code']);
        aa_code_value = isNaN(num) ? row['Drmis Leave Code'] : num;
      } else {
        aa_code_value = '';
      }
      hoursValue = 0;
    } else {
      // Default behaviour: prefer Oracle code when present, otherwise fallback to Drmis
      if (oracleCode && String(oracleCode) !== '0') {
        const num = parseInt(oracleCode);
        aa_code_value = isNaN(num) ? oracleCode : num;
      } else if (row['Drmis Leave Code'] && String(row['Drmis Leave Code']) !== '0') {
        const num = parseInt(row['Drmis Leave Code']);
        aa_code_value = isNaN(num) ? row['Drmis Leave Code'] : num;
      }
      hoursValue = oracleHoursNum;
    }

    // include original DRMIS leave metadata and a replacement flag so prefill logic can act accordingly
    const originalDrmisCode = row['Drmis Leave Code'] !== undefined && row['Drmis Leave Code'] !== null ? String(row['Drmis Leave Code']).trim() : '';
    const originalDrmisHours = Number(row['Drmis Hours']) || 0;
    const replaced = (oracleCode && String(oracleCode) !== '0' && originalDrmisCode && String(originalDrmisCode) !== '0' && String(oracleCode) !== String(originalDrmisCode));
    cats.push({ pers_no: row['Pers No'] ? Number(row['Pers No']) : '', date: formattedDateForDisplay(row.Date), work_order: '', act: '', aa_code: aa_code_value, hours: hoursValue, discrepancy_reason: discrepancy, row_type: 'data-row', editable: false, original_drmis_code: originalDrmisCode, original_drmis_hours: originalDrmisHours, replaced: replaced });
    cats.push({ pers_no: '', date: '', work_order: '', act: '', aa_code: 'Total Hours', hours: Number(row['Oracle Hours']) || '', discrepancy_reason: '', row_type: 'total-row', editable: false });
  }
  return cats;
}

function formattedDateForDisplay(d) {
  if (!d) return '';
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${months[d.getMonth()]} ${String(d.getDate()).padStart(2,'0')}, ${d.getFullYear()}`;
}

/* -------------------- Page initializers and wiring -------------------- */
function initStartPage(){
  // Nothing dynamic for start page
}

function initReconciliationPage(){
  loadStateFromSession();
  // attempt to load leave-code whitelist (non-blocking)
  loadLeaveCodeWhitelist().catch(e=>console.warn('Failed to auto-load leave-code whitelist', e));
  // hook up file inputs
  const drmisInput = document.getElementById('drmis_file');
  const oracleInput = document.getElementById('oracle_file');
  const drmisName = document.getElementById('drmis_name');
  const oracleName = document.getElementById('oracle_name');
  const drmisPreviewBtn = document.getElementById('drmis_preview_btn');
  const oraclePreviewBtn = document.getElementById('oracle_preview_btn');
  const loadingEl = document.getElementById('loading');
  const errorEl = document.getElementById('error');
  const resultsEl = document.getElementById('results');
  const resultMessage = document.getElementById('resultMessage');

  if (resultData && resultData.data && resultData.data.length>0) {
    displayResults(resultData);
    resultsEl.classList.remove('hidden');
  }

  drmisInput.addEventListener('change', async function(e){
    const file = e.target.files[0];
    drmisName.textContent = file ? file.name : '';
    if (file) {
      drmisFileData = await readExcelFile(file, 'drmis');
      // build raw DRMiS lookup for prefill use
      try { const dr = sheetRowsToObjects(drmisFileData); buildDrmisLookup(dr.rows); } catch(e){ console.warn('Failed to build DRMiS lookup', e); }
      // show preview button and record count
      drmisPreviewBtn.style.display = 'inline-block';
      try { const dr = sheetRowsToObjects(drmisFileData); const drCount = (dr && dr.rows) ? dr.rows.length : 0; document.getElementById('drmis_count').textContent = drCount + (drCount === 1 ? ' record found' : ' records found'); } catch(e){ document.getElementById('drmis_count').textContent = ''; }
      document.getElementById('sessionInfoBanner').style.display = 'none';
      // Validate columns immediately and show mapping dialog if needed
      try {
        const dr = sheetRowsToObjects(drmisFileData);
        // attempt extraction to validate presence of required columns
        extractDrmisData(dr.rows);
      } catch (err) {
        // show mapping dialog to user with header info
        const dr = sheetRowsToObjects(drmisFileData);
        showColumnMappingDialogClient('DRMIS', err.message, dr);
      }
    } else {
      drmisFileData = null; drmisPreviewBtn.style.display='none'; document.getElementById('drmis_count').textContent = '';
    }
  });

  oracleInput.addEventListener('change', async function(e){
    const file = e.target.files[0];
    oracleName.textContent = file ? file.name : '';
    if (file) {
      oracleFileData = await readExcelFile(file, 'oracle');
      oraclePreviewBtn.style.display = 'inline-block';
      try { const or = sheetRowsToObjects(oracleFileData); const orCount = (or && or.rows) ? or.rows.length : 0; document.getElementById('oracle_count').textContent = orCount + (orCount === 1 ? ' record found' : ' records found'); } catch(e){ document.getElementById('oracle_count').textContent = ''; }
      document.getElementById('sessionInfoBanner').style.display = 'none';
      // Validate columns immediately and show mapping dialog if needed
      try {
        const or = sheetRowsToObjects(oracleFileData);
        extractOracleData(or.rows);
      } catch (err) {
        const or = sheetRowsToObjects(oracleFileData);
        showColumnMappingDialogClient('ORACLE', err.message, or);
      }
    } else { oracleFileData = null; oraclePreviewBtn.style.display='none'; }
  });

  document.getElementById('uploadForm').addEventListener('submit', async function(e){
    e.preventDefault();
    errorEl.classList.add('hidden'); resultsEl.classList.add('hidden');
    if (!drmisFileData || !oracleFileData) { errorEl.textContent = '❌ Please select both files before submitting.'; errorEl.classList.remove('hidden'); return; }
    loadingEl.classList.remove('hidden');
    try {
      // convert sheet arrays to objects
      const dr = sheetRowsToObjects(drmisFileData);
      const or = sheetRowsToObjects(oracleFileData);
      // Use extracted objects
      const drObjs = dr.rows;
      const orObjs = or.rows;
      // Attempt to extract using auto-detection; if missing columns, show mapping UI
      let drmisExtracted, oracleExtracted;
      try {
        drmisExtracted = extractDrmisData(drObjs);
      } catch (err) {
        loadingEl.classList.add('hidden');
        // show mapping modal to user: reuse mapping UI idea - but for brevity alert and throw
        showColumnMappingDialogClient('DRMIS', err.message, dr);
        return;
      }
      try {
        oracleExtracted = extractOracleData(orObjs);
      } catch (err) {
        loadingEl.classList.add('hidden');
        showColumnMappingDialogClient('ORACLE', err.message, or);
        return;
      }
      // reconcile
      const mismatches = reconcileData(drmisExtracted, oracleExtracted);
      // format Date for display
      mismatches.forEach(m => m.Date = formattedDateForDisplay(new Date(m.Date)));
      // build resultData compatible with previous UI
      const headers = ['Date', 'Pers No', 'Oracle Leave Code', 'Oracle Hours', 'Drmis Leave Code', 'Drmis Hours', 'Add to CATs Edits'];
      const dataRows = mismatches.map(m => [m.Date, m['Pers No'], m['Oracle Leave Code'], m['Oracle Hours'], m['Drmis Leave Code'], m['Drmis Hours'], m['Add to CATs Edits']]);
      resultData = { headers: headers, data: dataRows, count: mismatches.length, message: 'Found ' + mismatches.length + ' mismatches' };
      // save cats edits for later
      // regenerate cats edits using original (unformatted) Date objects: need to re-run reconcile with preserved Date objects; recreate mismatches2
      // We'll reconstruct mismatches2 from reconcilation function using Date objects
      const drmisExtracted2 = drmisExtracted;
      const oracleExtracted2 = oracleExtracted;
      const mismatches2 = reconcileData(drmisExtracted2, oracleExtracted2);
      catsEditsData = generateCatsEdits(mismatches2);

      saveStateToSession();
      displayResults(resultData);
      resultsEl.classList.remove('hidden');
    } catch (err) {
      console.error(err);
      errorEl.textContent = '❌ ' + (err.message || err);
      errorEl.classList.remove('hidden');
    } finally {
      loadingEl.classList.add('hidden');
    }
  });

  document.getElementById('downloadResultsBtn').addEventListener('click', function(){
    // Export the live results table as a styled .xlsx (includes any UI changes)
    downloadReconciliationXlsxExcelJS();
  });

  document.getElementById('resetBtn').addEventListener('click', function(){
    clearClientSideData();
  });

  document.getElementById('downloadReconciliationBtn').addEventListener('click', function(){ downloadReconciliationXlsxExcelJS(); });
  document.getElementById('resetFormBtn').addEventListener('click', clearClientSideData);
}

function showColumnMappingDialogClient(fileType, message, sheetObj) {
  // Interactive mapping dialog: present required fields and allow user to map them to available headers
  const headers = sheetObj.headers || [];
  const required = (fileType === 'DRMIS') ? ['Pers No', 'Date', 'Hours', 'A/A Type'] : ['From Date', 'Hours Recorded', 'Leave Code'];
  const modal = document.createElement('div'); modal.className = 'modal'; modal.style.display = 'flex';
  const options = ['-- Select column --'].concat(headers);
  // Try to auto-detect existing matching headers using the same heuristics as extractors
  function detectDrmisMap(headersArr) {
    const lowerMap = {}; headersArr.forEach(h=> lowerMap[String(h).toLowerCase().trim()] = h);
    const pers = Object.keys(lowerMap).find(k => k.includes('pers') && k.includes('no'));
    const date = Object.keys(lowerMap).find(k => k === 'date');
    const hours = Object.keys(lowerMap).find(k => k === 'hours');
    const atype = Object.keys(lowerMap).find(k => (k.includes('a/a') && k.includes('type') && !k.includes('text')));
    const out = {};
    if (pers) out['Pers No'] = lowerMap[pers];
    if (date) out['Date'] = lowerMap[date];
    if (hours) out['Hours'] = lowerMap[hours];
    if (atype) out['A/A Type'] = lowerMap[atype];
    return out;
  }
  function detectOracleMap(headersArr) {
    const lowerMap = {}; headersArr.forEach(h=> lowerMap[String(h).toLowerCase().trim()] = h);
    const fromDate = Object.keys(lowerMap).find(k => (k.includes('from') && k.includes('date'))) || (Object.keys(lowerMap).find(k=>k==='date'));
    const hours = Object.keys(lowerMap).find(k => k.includes('hours') && k.includes('recorded')) || (Object.keys(lowerMap).find(k=>k==='hours'));
    const leave = Object.keys(lowerMap).find(k => k.includes('leave') && k.includes('code'));
    const out = {};
    if (fromDate) out['From Date'] = lowerMap[fromDate];
    if (hours) out['Hours Recorded'] = lowerMap[hours];
    if (leave) out['Leave Code'] = lowerMap[leave];
    return out;
  }
  const detected = (fileType === 'DRMIS') ? detectDrmisMap(headers) : detectOracleMap(headers);
  let selectsHtml = '';
  // Only show selects for fields that were NOT auto-detected. Detected fields are applied automatically.
  required.forEach((r) => {
    if (detected[r]) return; // skip showing detected fields
    const opts = options.map(o => `<option value="${o==='-- Select column --' ? '' : o}">${o}</option>`).join('');
    selectsHtml += `<div style="margin-bottom:10px;"><label style="display:block;font-weight:600;margin-bottom:4px;">${r} <small style="color:#6c757d;font-weight:400;margin-left:6px;">(please choose)</small></label><select data-required="${r}" class="mapping-select" style="width:100%;padding:6px;">${opts}</select></div>`;
  });

  // remove in-message "Available columns: ..." chunk so we only show the columns once (at the bottom)
  let cleanMessage = '';
  try {
    if (message && typeof message === 'string') {
      cleanMessage = message.replace(/Available columns:\s*.*$/i, '').trim();
      // remove trailing punctuation
      cleanMessage = cleanMessage.replace(/[\:\-\s]*$/,'');
    } else {
      cleanMessage = message || '';
    }
  } catch (e) { cleanMessage = message; }

  const html = `
    <div class="modal-content" style="max-width: 780px;" onclick="event.stopPropagation()">
      <div class="modal-header" style="border-left: 4px solid #af3c43;">
        <h2 style="color: #af3c43;">Map Columns for ${fileType} File</h2>
        <button class="modal-close" id="mappingCloseBtn">&times;</button>
      </div>
      <div class="modal-body">
        <p style="font-size:14px;">${cleanMessage}</p>
        <p style="font-size:13px; color:#6c757d;">Select which column from your uploaded file corresponds to each required field. After mapping, the extraction will re-run automatically.</p>
        <div style="margin-top:12px;">
          ${selectsHtml}
        </div>
        <div style="margin-top:12px; font-size:13px; color:#6c757d;">Available columns: <strong>${headers.join(', ')}</strong></div>
        <div style="text-align:right; margin-top:18px;"><button id="mappingCancelBtn" style="margin-right:8px; background:#6c757d;color:white;border:none;padding:8px 14px;">Cancel</button><button id="mappingApplyBtn" style="background:#284162;color:white;border:none;padding:8px 14px;">Apply Mapping</button></div>
      </div>
    </div>
  `;
  modal.innerHTML = html;
  document.body.appendChild(modal);

  // wire up buttons
  document.getElementById('mappingCloseBtn').addEventListener('click', ()=> modal.remove());
  document.getElementById('mappingCancelBtn').addEventListener('click', ()=> modal.remove());
  document.getElementById('mappingApplyBtn').addEventListener('click', async ()=>{
    // Collect only user-provided mappings (for fields shown). Combine with detected mappings.
    const selects = Array.from(modal.querySelectorAll('.mapping-select'));
    const userMapping = {};
    let ok = true;
    selects.forEach(s => {
      const key = s.dataset.required;
      const val = s.value;
      if (!val) ok = false;
      userMapping[key] = val;
    });
    if (!ok) { alert('Please select a column for each required field shown, or cancel and re-upload a file with matching headers.'); return; }

    const combinedMapping = Object.assign({}, detected, userMapping);

    // Apply mapping to sheetObj.rows (array of objects). We'll add canonical keys expected by extractors and re-run extraction.
    try {
      const mappedRows = sheetObj.rows.map(r => { const rr = Object.assign({}, r); Object.entries(combinedMapping).forEach(([req, hdr]) => { rr[req] = r[hdr]; }); return rr; });
      if (fileType === 'DRMIS') {
        // attempt extraction
        const extracted = extractDrmisData(mappedRows);
        // if success, update in-memory drmisFileData to a rows-array form so future submits will use it
        const newHeaders = Object.keys(mappedRows[0] || {});
        const newRows = [newHeaders].concat(mappedRows.map(o => newHeaders.map(h => o[h])));
        drmisFileData = newRows;
          // rebuild raw DRMiS lookup from the mapped objects so prefill works
          try { buildDrmisLookup(mappedRows); } catch(e){ console.warn('Failed to build DRMiS lookup after mapping', e); }
        modal.remove();
        // Show mapping summary modal (detected + user mappings)
        const userMap = {};
        Object.keys(detected || {}).forEach(k=>{});
        // build userMap from combinedMapping only for fields the user provided (not auto-detected)
        const provided = {};
        Object.keys(combinedMapping).forEach(k => {
          if (!detected[k]) provided[k] = combinedMapping[k];
        });
        showMappingResultModal('DRMIS', detected, provided);
        if (oracleFileData) {
          try { document.getElementById('uploadForm').dispatchEvent(new Event('submit', {cancelable:true})); } catch(e) { console.warn('Could not auto-submit after mapping', e); }
        }
      } else {
        const extracted = extractOracleData(mappedRows);
        const newHeaders = Object.keys(mappedRows[0] || {});
        const newRows = [newHeaders].concat(mappedRows.map(o => newHeaders.map(h => o[h])));
        oracleFileData = newRows;
        modal.remove();
        const provided = {};
        Object.keys(combinedMapping).forEach(k => { if (!detected[k]) provided[k] = combinedMapping[k]; });
        showMappingResultModal('ORACLE', detected, provided);
        if (drmisFileData) {
          try { document.getElementById('uploadForm').dispatchEvent(new Event('submit', {cancelable:true})); } catch(e){ console.warn('Could not auto-submit after mapping', e); }
        }
      }
    } catch (err) {
      alert('Extraction failed after applying mapping: ' + (err.message || err));
    }
  });
}

// Show a mapping result modal summarizing which fields were auto-detected and which the user mapped
function showMappingResultModal(fileType, detectedMap, userMap) {
  const modal = document.createElement('div'); modal.className='modal'; modal.style.display='flex';
  const detectedEntries = Object.entries(detectedMap || {});
  const userEntries = Object.entries(userMap || {});
  let detectedHtml = '';
  if (detectedEntries.length>0) {
    detectedHtml = '<div style="margin-bottom:10px; color:#3a3a3a; font-size:13px;"><strong>Auto-detected:</strong><ul style="margin:6px 0 0 18px;">' + detectedEntries.map(([k,v])=>`<li>${k} <span style="color:#6c757d;">➜</span> ${v}</li>`).join('') + '</ul></div>';
  }
  let userHtml = '';
  if (userEntries.length>0) {
    userHtml = '<div style="margin-bottom:10px; font-size:14px;"><strong>Mapped by you:</strong><ul style="margin:6px 0 0 18px;">' + userEntries.map(([k,v])=>`<li><strong style="color:#284162">${k}</strong> <span style="color:#284162; font-weight:600; margin:0 8px;">➜</span> <strong style="color:#af3c43">${v}</strong></li>`).join('') + '</ul></div>';
  }
  const nextText = (fileType==='DRMIS') ? 'Please map ORACLE file if needed, or re-upload ORACLE file and submit.' : 'Please map DRMIS file if needed, or re-upload DRMIS file and submit.';
  const html = `
    <div class="modal-content" style="max-width:640px;" onclick="event.stopPropagation()">
      <div class="modal-header" style="border-left: 4px solid #2e8b57;">
        <h2 style="color:#2e8b57;">${fileType} columns mapped successfully</h2>
        <button class="modal-close" id="mappingResultClose">&times;</button>
      </div>
      <div class="modal-body" style="font-size:13px;color:#222;">
        <p style="margin-top:6px;">The Following Column(s) linking was applied to your ${fileType} file.</p>
        ${userHtml}
        ${detectedHtml}
        <p style="margin-top:8px;color:#444;">${nextText}</p>
        <div style="text-align:right;margin-top:14px;"><button id="mappingResultOk" style="background:#284162;color:white;border:none;padding:8px 14px;">OK</button></div>
      </div>
    </div>
  `;
  modal.innerHTML = html; document.body.appendChild(modal);
  document.getElementById('mappingResultClose').addEventListener('click', ()=> modal.remove());
  document.getElementById('mappingResultOk').addEventListener('click', ()=> modal.remove());
}

function clearClientSideData(){
  try { sessionStorage.removeItem('leave_reconcile_state'); } catch(e){}
  drmisFileData = null; oracleFileData = null; resultData = null; catsEditsData = null;
  // reload page to clear UI
  location.hash = '#/upload';
  loadRoute();
}

function displayResults(data){
  const headerRow = document.getElementById('tableHeader'); headerRow.innerHTML='';
  const checkboxTh = document.createElement('th'); checkboxTh.textContent='Add to CATs Edits'; checkboxTh.style.textAlign='center'; headerRow.appendChild(checkboxTh);
  data.headers.forEach(h=>{ if(h==='Add to CATs Edits') return; const th=document.createElement('th'); th.textContent=h; headerRow.appendChild(th); });

  const tableBody = document.getElementById('tableBody'); tableBody.innerHTML='';
  const addToCatsIndex = data.headers.indexOf('Add to CATs Edits');

  data.data.forEach((row,rowIndex)=>{
    const tr = document.createElement('tr');
    const checkboxTd = document.createElement('td'); checkboxTd.style.textAlign='center'; const isChecked = addToCatsIndex>=0 ? row[addToCatsIndex] : true; checkboxTd.innerHTML = `<input type="checkbox" class="cats-checkbox" data-row-index="${rowIndex}" ${isChecked? 'checked':''} onchange="updateCatsSelection()">`; tr.appendChild(checkboxTd);
    row.forEach((cell,cellIndex)=>{
      if(cellIndex===addToCatsIndex) return;
      const td = document.createElement('td');
      const header = data.headers[cellIndex];
      if((header==='Oracle Hours' || header==='Drmis Hours') && cell !== null && cell !== ''){
        const hours = parseFloat(cell);
        td.textContent = !isNaN(hours) ? hours.toFixed(2) : (cell===0 ? '0.00' : cell);
      } else if(header==='Oracle Leave Code' || header==='Drmis Leave Code'){
        const c = (cell === null || cell === undefined) ? '' : String(cell).trim();
        td.textContent = (c === '' || c === '-') ? '0' : c;
      } else {
        td.textContent = (cell === null || cell === '') ? '-' : cell;
      }
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });

  // select all row
  const selectAllRow = document.createElement('tr'); selectAllRow.style.backgroundColor='#f8f9fa'; selectAllRow.style.borderTop='2px solid #284162';
  const selectAllCell = document.createElement('td'); selectAllCell.colSpan = data.headers.length; selectAllCell.style.textAlign='center'; selectAllCell.style.padding='10px'; selectAllCell.innerHTML = `<label style="cursor:pointer; font-weight:600; color:#284162;"><input type="checkbox" id="selectAllCats" checked onchange="toggleAllCatsSelection(this.checked)" style="margin-right:8px;">Select All</label>`;
  selectAllRow.appendChild(selectAllCell); tableBody.appendChild(selectAllRow);

  // store and message
  window.resultData = data;
  try{
    const msgEl = document.getElementById('resultMessage');
    if(msgEl){ const count = Number(data.count) || (data.data? data.data.length:0); msgEl.textContent = count>0? `${count} mismatch record${count===1?'':'s'} found.` : 'No mismatch records found.'; msgEl.style.margin='6px 0 0 0'; msgEl.style.color='#6c757d'; msgEl.setAttribute('aria-live','polite'); }
  }catch(e){ console.warn('Could not update result message:', e); }
}

function toggleAllCatsSelection(checked){ document.querySelectorAll('.cats-checkbox').forEach(cb=>cb.checked=checked); updateCatsSelection(); }

function updateCatsSelection(){
  const selectedIndices = []; document.querySelectorAll('.cats-checkbox').forEach(cb=>{ if(cb.checked) selectedIndices.push(parseInt(cb.dataset.rowIndex)); });
  // Update catsEditsData accordingly: regenerate catsEditsData from resultData using selectedIndices
  if (!window.resultData) return;
  // Reconstruct mismatches from resultData into an array of objects
  const headers = window.resultData.headers;
  const dataRows = window.resultData.data;
  const mismatches = dataRows.map(r=>{
    const obj = {};
    headers.forEach((h, i)=>{ obj[h]=r[i]; });
    return obj;
  });
  const selected = selectedIndices.map(i=>mismatches[i]).filter(Boolean);
  // Need to convert back to the internal mismatches used by generateCatsEdits (with Date objects)
  // We'll attempt to parse Date strings in 'Apr 10, 2025' format
  const converted = selected.map(s=>({ Date: parseDisplayDate(s['Date']), 'Pers No': s['Pers No'], 'Oracle Leave Code': s['Oracle Leave Code'], 'Oracle Hours': Number(s['Oracle Hours']) || 0, 'Drmis Leave Code': s['Drmis Leave Code'], 'Drmis Hours': Number(s['Drmis Hours']) || 0 }));
  catsEditsData = generateCatsEdits(converted);
  saveStateToSession();
}

function parseDisplayDate(str) {
  // expect 'Apr 10, 2025'
  if (!str) return null;
  const parts = str.split(' ');
  if (parts.length<3) return new Date(str);
  const mon = parts[0].replace(',',''); const day = parseInt(parts[1].replace(',','')); const year = parseInt(parts[2]);
  const months = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
  return new Date(year, months[mon], day);
}

/* -------------------- File preview modal handlers -------------------- */
function showFilePreview(which) {
  // which: 'drmis' or 'oracle'
  const modal = document.getElementById('filePreviewModal');
  const title = document.getElementById('modalTitle');
  // Robustly locate/create preview header/body elements (they may be missing if the page was re-rendered)
  let headerRow = document.getElementById('previewHeader');
  let body = document.getElementById('previewBody');
  const modalBody = modal ? modal.querySelector('.modal-body') : null;
  if ((!headerRow || !body) && modalBody) {
    // try to (re)create a table wrapper and thead/tbody with the expected ids
    let tableWrapper = modalBody.querySelector('.table-wrapper');
    if (!tableWrapper) {
      tableWrapper = document.createElement('div');
      tableWrapper.className = 'table-wrapper';
      modalBody.appendChild(tableWrapper);
    }
    let tableEl = tableWrapper.querySelector('table');
    if (!tableEl) {
      tableEl = document.createElement('table');
      tableEl.appendChild(document.createElement('thead'));
      tableEl.appendChild(document.createElement('tbody'));
      tableWrapper.appendChild(tableEl);
    }
    const thead = tableEl.querySelector('thead');
    let tr = thead.querySelector('tr');
    if (!tr) { tr = document.createElement('tr'); thead.appendChild(tr); }
    tr.id = 'previewHeader';
    const tbody = tableEl.querySelector('tbody');
    tbody.id = 'previewBody';
    headerRow = document.getElementById('previewHeader');
    body = document.getElementById('previewBody');
  }
  if (headerRow) headerRow.innerHTML = '';
  if (body) body.innerHTML = '';
  let data = null;
  if (which === 'drmis') data = drmisFileData; else if (which === 'oracle') data = oracleFileData;
  if (!data || !Array.isArray(data) || data.length === 0) {
    title.textContent = 'No file loaded';
    modal.style.display = 'flex';
    return;
  }
  title.textContent = (which === 'drmis' ? 'DRMIS File Preview' : 'ORACLE File Preview');
  // Use robust header detection to show the original table only (pre-filter)
  const detected = sheetRowsToObjects(data);
  const headers = detected.headers || [];
  headers.forEach(h => { const th = document.createElement('th'); th.textContent = h === null || h === undefined || h === '' ? '-' : String(h); headerRow.appendChild(th); });
  // show up to first 1000 data rows and display a short note
  const totalRows = Math.max(0, detected.rows.length);
  const maxRows = Math.min(1000, totalRows);
  // insert a short note showing how many rows are displayed
  if (modalBody) {
    const existingNote = modal.querySelector('.preview-note');
    if (!existingNote) {
      const noteRow = document.createElement('div');
      noteRow.className = 'preview-note';
      noteRow.style.margin = '6px 0 8px 0';
      noteRow.style.color = '#444';
      noteRow.style.fontSize = '13px';
      noteRow.textContent = `Showing ${maxRows} of ${totalRows} data row${totalRows===1? '' : 's'}`;
      // try to insert before table wrapper if present
      const tableWrapper = modalBody.querySelector('.table-wrapper');
      if (tableWrapper) modalBody.insertBefore(noteRow, tableWrapper);
      else modalBody.insertBefore(noteRow, modalBody.firstChild);
    }
  }
  // allow horizontal scrolling for very wide tables
  const tableWrapper = modal.querySelector('.table-wrapper') || document.createElement('div');
  tableWrapper.className = 'table-wrapper';
  tableWrapper.style.overflow = 'auto';
  tableWrapper.style.maxWidth = '100%';
  tableWrapper.style.marginTop = '8px';
  // ensure header/body are within the wrapper
  let tableEl = tableWrapper.querySelector('table');
  if (!tableEl) {
    tableEl = document.createElement('table');
    tableEl.style.borderCollapse = 'collapse';
    tableEl.style.minWidth = '600px';
    tableEl.appendChild(document.createElement('thead'));
    tableEl.appendChild(document.createElement('tbody'));
    tableWrapper.appendChild(tableEl);
  }
  // replace thead/tbody contents
  const thead = tableEl.querySelector('thead'); thead.innerHTML = '';
  const headerTr = document.createElement('tr');
  headers.forEach(h => { const th = document.createElement('th'); th.style.padding = '6px 8px'; th.style.border = '1px solid #ddd'; th.textContent = h === null || h === undefined || h === '' ? '-' : String(h); headerTr.appendChild(th); });
  thead.appendChild(headerTr);
  const tbodyEl = tableEl.querySelector('tbody'); tbodyEl.innerHTML = '';
  for (let i = 0; i < maxRows; i++) {
    const rowObj = detected.rows[i];
    if (!rowObj) continue;
    const tr = document.createElement('tr');
    headers.forEach(h => { const td = document.createElement('td'); td.style.padding = '4px 8px'; td.style.border = '1px solid #eee'; const val = rowObj[h]; td.textContent = (val === null || val === undefined) ? '' : String(val); tr.appendChild(td); });
    tbodyEl.appendChild(tr);
  }
  // attach wrapper into modal (prefer placing before any existing table wrapper)
  const existing = modalBody.querySelector('.table-wrapper');
  if (existing) existing.replaceWith(tableWrapper); else modalBody.appendChild(tableWrapper);
  modal.style.display = 'flex';
}

function closeFilePreview() {
  const modal = document.getElementById('filePreviewModal');
  if (modal) modal.style.display = 'none';
}

/* -------------------- Excel generation (SheetJS) -------------------- */
function downloadReconciliationExcel(resultData) {
  // We'll build a simple workbook with headers and rows similar to server output
  const ws_data = [];
  // headers without 'Add to CATs Edits'
  const headers = resultData.headers.filter(h=>h!=='Add to CATs Edits');
  ws_data.push(headers);
  resultData.data.forEach(row => {
    const filtered = [];
    resultData.headers.forEach((h,i)=>{ if (h!=='Add to CATs Edits') filtered.push(row[i]); });
    ws_data.push(filtered);
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(wb, ws, 'Mismatches');
  const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
  const blob = new Blob([wbout], {type:'application/octet-stream'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = `Leave_Reconciliation_${(new Date()).toISOString().split('T')[0]}.xlsx`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}

function downloadCatsEditsExcelInternal() {
  if (!catsEditsData) { alert('No CATs Edits available'); return; }
  // Build workbook with styled header rows where possible
  const wb = XLSX.utils.book_new();
  const ws_data = [];
  ws_data.push(['Modifications to CATs Entries']);
  ws_data.push([]);
  ws_data.push(['These are the changes that need to be made in CATs on DRMIS']);
  ws_data.push([]);
  ws_data.push(['Pers No','Date','Work Order','Act','AA Code','Hours','Discrepancy Reason']);
  catsEditsData.forEach(entry => {
    ws_data.push([entry.pers_no, entry.date, entry.work_order, entry.act, entry.aa_code, entry.hours, entry.discrepancy_reason]);
  });
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(wb, ws, 'CATs Edits');
  const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
  const blob = new Blob([wbout], {type:'application/octet-stream'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = `CATs_Edits_${(new Date()).toISOString().split('T')[0]}.xlsx`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}

/* -------------------- Cats Edits page rendering and actions -------------------- */
function initCatsEditsPage(){
  loadStateFromSession();
  const container = document.getElementById('catsContent');
  container.innerHTML = '';
  // if catsEditsData not present, try to generate from resultData
  if (!catsEditsData && window.resultData) {
    // generate from current resultData
    // reconstruct mismatches into internal format
    const headers = window.resultData.headers;
    const rows = window.resultData.data;
    const mismatches = rows.map(r => {
      const obj = {};
      headers.forEach((h,i)=>{ obj[h]=r[i]; });
      return obj;
    });
    const converted = mismatches.map(s=>({ Date: parseDisplayDate(s['Date']), 'Pers No': s['Pers No'], 'Oracle Leave Code': s['Oracle Leave Code'], 'Oracle Hours': Number(s['Oracle Hours'])||0, 'Drmis Leave Code': s['Drmis Leave Code'], 'Drmis Hours': Number(s['Drmis Hours'])||0 }));
    catsEditsData = generateCatsEdits(converted);
    saveStateToSession();
  }
  if (!catsEditsData || catsEditsData.length===0) {
    container.innerHTML = `<div class="no-data"><p>No CATs edits to display. Please upload files first.</p><a href="#/upload" class="nav-button">Go to Reconcile Page</a></div>`;
    // wire buttons
    document.getElementById('downloadCatsBtn').addEventListener('click', ()=>{ alert('No data to download'); });
    document.getElementById('copyClipboardBtn').addEventListener('click', ()=>{ alert('No data to copy'); });
    document.getElementById('emailTableBtn').addEventListener('click', ()=>{ alert('No data to email'); });
    return;
  }

  // build table similar to original
  const tableWrapper = document.createElement('div'); tableWrapper.className='cats-table-wrapper';
  // Note: removed datalist to avoid native browser dropdown icon; using a plain input with a small custom arrow instead
  const table = document.createElement('table'); table.className='cats-table'; table.id='catsTable';
  const thead = document.createElement('thead'); thead.innerHTML = `<tr><th>Pers No</th><th>Date</th><th>Work Order</th><th>Act</th><th>AA Code</th><th>Hours</th><th>Discrepancy Reason</th><th></th></tr>`;
  const tbody = document.createElement('tbody');
  catsEditsData.forEach(entry => {
    const tr = document.createElement('tr'); tr.className = entry.row_type; if (entry.editable) tr.setAttribute('data-editable','true');
    const td1 = document.createElement('td'); td1.className = (entry.row_type=='data-row' ? 'non-editable' : ''); td1.textContent = entry.pers_no || '';
    const td2 = document.createElement('td'); td2.className = (entry.row_type=='data-row' ? 'non-editable' : ''); td2.textContent = entry.date || '';
    const td3 = document.createElement('td'); td3.className = (entry.row_type=='data-row' ? 'non-editable' : 'editable'); if (entry.row_type!=='data-row' && entry.row_type!=='total-row') td3.contentEditable='true'; td3.textContent = entry.work_order || '';
    const td4 = document.createElement('td'); td4.className = (entry.row_type=='data-row' ? 'non-editable' : (entry.row_type=='editable-row' ? 'act-dropdown':'editable'));
    if (entry.row_type=='editable-row') {
      td4.innerHTML = `<input class="act-input act-select" />`;
      const inp = td4.querySelector('.act-input'); if (inp) inp.oninput = updateTotals;
    } else td4.textContent = entry.act || '';
    const td5 = document.createElement('td'); td5.className = (entry.row_type=='data-row' || entry.row_type=='total-row' ? 'non-editable' : 'editable'); if (entry.row_type!=='data-row' && entry.row_type!=='total-row') { td5.contentEditable='true'; td5.oninput = updateTotals; } td5.textContent = entry.aa_code || '';
    const td6 = document.createElement('td');
    td6.className = (entry.row_type=='data-row' || entry.row_type=='total-row' ? 'non-editable' + (entry.row_type=='total-row' ? ' total-hours' : '') : 'editable');
    if (entry.row_type!=='data-row' && entry.row_type!=='total-row') {
      td6.contentEditable='true';
      td6.oninput = updateTotals;
      td6.onblur = function(){ formatHours(this); };
    }
    // Always display numeric hours (including 0) when present. Format to two decimals.
    let hoursText = '';
    if (entry.hours !== undefined && entry.hours !== null && entry.hours !== '') {
      const num = Number(entry.hours);
      if (!isNaN(num)) hoursText = num.toFixed(2);
    }
    td6.textContent = hoursText;
    const td7 = document.createElement('td'); td7.className = (entry.row_type=='data-row' || entry.row_type=='total-row' ? 'non-editable' : ''); td7.textContent = entry.discrepancy_reason || '';
    const td8 = document.createElement('td'); td8.className='action-cell';
    if (entry.row_type=='data-row') td8.innerHTML = `<button class="btn-icon" onclick="addRowAfter(this)" title="Add additional entry">+</button>`;
    else if (entry.row_type=='editable-row') td8.innerHTML = `<button class="btn-icon btn-icon-delete" onclick="deleteRow(this)" title="Delete row">×</button>`;

    tr.append(td1,td2,td3,td4,td5,td6,td7,td8);
    // attach original DRMiS metadata if present
    try {
      if (entry.original_drmis_code) tr.dataset.originalDrmis = entry.original_drmis_code;
      if (entry.original_drmis_hours !== undefined) tr.dataset.originalDrmisHours = String(entry.original_drmis_hours);
      if (entry.replaced) tr.dataset.replaced = 'true';
    } catch(e){}
    tbody.appendChild(tr);
    // If this is a data-row, automatically insert a prefilled editable-row (from DRMIS) right after it
    if (entry.row_type === 'data-row') {
      const preRow = document.createElement('tr'); preRow.className = 'editable-row'; preRow.setAttribute('data-editable','true');
      preRow.innerHTML = `
        <td class="non-editable"></td>
        <td class="non-editable"></td>
        <td class="editable" contenteditable="true"></td>
        <td class="act-dropdown">
          <input class="act-input act-select" />
        </td>
        <td class="editable" contenteditable="true" oninput="updateTotals()"></td>
        <td class="editable" contenteditable="true" oninput="updateTotals()" onblur="formatHours(this)"></td>
        <td class="non-editable"></td>
        <td class="action-cell"><button class="btn-icon btn-icon-delete" onclick="deleteRow(this)" title="Delete row">×</button></td>
      `;
      // attempt prefill using the data-row we just created (catch errors)
      let filled = false;
      try { filled = prefillEditableRow(preRow, tr); } catch (e) { console.warn('Auto-prefill failed for data-row', e); }
      // Only append the prefilled editable-row if it actually received data
      if (filled) tbody.appendChild(preRow);
    }
        // ensure totals are correct after building each pair
        try { updateTotals(); } catch(e) { console.warn('updateTotals failed during table build', e); }
  });
  table.appendChild(thead); table.appendChild(tbody); tableWrapper.appendChild(table);
  container.appendChild(tableWrapper);

  // Footer actions
  const footer = document.createElement('div'); footer.className='results-footer'; footer.innerHTML = `<button type="button" class="btn-download" id="downloadCatsEditsBtn">Download as Excel</button><button type="button" class="btn-download" id="copyCatsBtn">Copy to Clipboard</button><button type="button" class="btn-download" id="emailCatsBtn">Email Table</button><button type="button" class="btn-reset" onclick="location.hash='#/upload'">Back to Reconcile</button>`;
  container.appendChild(footer);

  // wire actions
  document.getElementById('downloadCatsEditsBtn').addEventListener('click', downloadCatsEditsXlsxExcelJS);
  document.getElementById('copyCatsBtn').addEventListener('click', copyCatsToClipboard);
  document.getElementById('emailCatsBtn').addEventListener('click', emailCatsTable);

  // Final totals pass after table constructed
  try { updateTotals(); } catch(e) { console.warn('Final updateTotals failed', e); }

  // Also wire the top-right icon buttons (if present) so they perform the same actions
  try {
    const topDownload = document.getElementById('downloadCatsBtn');
    if (topDownload) topDownload.addEventListener('click', () => { downloadCatsEditsXlsxExcelJS(); });
    const topCopy = document.getElementById('copyClipboardBtn');
    if (topCopy) topCopy.addEventListener('click', () => { copyCatsToClipboard(); });
    const topEmail = document.getElementById('emailTableBtn');
    if (topEmail) topEmail.addEventListener('click', () => { emailCatsTable(); });
  } catch (e) { console.warn('Top icon button wiring failed', e); }
}

async function downloadCatsEditsXlsxExcelJS(){
  if (!validateAllRows()) { alert('Please fill in AA Code and Hours for all added rows before exporting.'); return; }
  if (typeof ExcelJS === 'undefined') { alert('ExcelJS library not loaded. Please ensure the app includes ExcelJS.'); return; }
  const table = document.getElementById('catsTable');
  if (!table) { alert('No CATs table found'); return; }

  // Build live rows from the current DOM so user-added/edited rows are included
  const rows = Array.from(table.querySelectorAll('tbody tr'));
  if (rows.length === 0) { alert('No CATs rows to export'); return; }
  const liveData = rows.map(tr => {
    const cells = tr.querySelectorAll('td');
    const pers_no = (cells[0] && cells[0].textContent) ? cells[0].textContent.trim() : '';
    const date = (cells[1] && cells[1].textContent) ? cells[1].textContent.trim() : '';
    const work_order = (cells[2] && cells[2].textContent) ? cells[2].textContent.trim() : '';
    let act = '';
    if (cells[3]) {
      const sel = (cells[3].querySelector('.act-input') || cells[3].querySelector('.act-select'));
      act = sel ? (sel.value !== undefined ? String(sel.value).trim() : cells[3].textContent.trim()) : cells[3].textContent.trim();
    }
    const aa_code = (cells[4] && cells[4].textContent) ? cells[4].textContent.trim() : '';
    let hoursVal = '';
    if (cells[5]) {
      const htxt = cells[5].textContent.trim(); const hn = parseFloat(htxt); hoursVal = (!isNaN(hn) ? hn : (htxt === '' ? '' : htxt));
    }
    const discrepancy_reason = (cells[6] && cells[6].textContent) ? cells[6].textContent.trim() : '';
    const isTotal = tr.classList.contains('total-row');
    return { pers_no, date, work_order, act, aa_code, hours: hoursVal, discrepancy_reason, isTotal };
  });

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Leave Reconcile';
  wb.created = new Date();
  const ws = wb.addWorksheet('CATs Edits');

  // Title rows with merges and basic styling
  ws.mergeCells('A1:G1'); ws.getCell('A1').value = 'Modifications to CATs Entries';
  ws.getCell('A1').font = {name:'Calibri', size:14, bold:true};
  ws.getCell('A1').alignment = {vertical:'middle', horizontal:'center'};
  ws.getRow(1).height = 22;

  ws.mergeCells('A3:G3'); ws.getCell('A3').value = 'These are the changes that need to be made in CATs on DRMIS';
  ws.getCell('A3').font = {name:'Calibri', size:11};
  ws.getCell('A3').alignment = {vertical:'middle', horizontal:'left'};

  // Header row
  const headerRowIndex = 5;
  const headers = ['Pers No','Date','Work Order','Act','AA Code','Hours','Discrepancy Reason'];
  const headerRow = ws.getRow(headerRowIndex);
  headers.forEach((h, i)=>{
    const cell = headerRow.getCell(i+1);
    cell.value = h;
    cell.font = {name:'Calibri', bold:true, color:{argb:'FF000000'}};
    cell.fill = {type:'pattern', pattern:'solid', fgColor:{argb:'FFD4EDDA'}};
    cell.border = {top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'}};
    cell.alignment = {vertical:'middle', horizontal:'center'};
  });
  ws.getRow(headerRowIndex).height = 18;

  // Column widths
  ws.columns = [
    {key:'pers_no', width:12},
    {key:'date', width:12},
    {key:'work_order', width:12},
    {key:'act', width:8},
    {key:'aa_code', width:10},
    {key:'hours', width:8},
    {key:'discrepancy_reason', width:30}
  ];

  // Data rows (use liveData from DOM)
  let r = headerRowIndex + 1;
  liveData.forEach(entry => {
    const row = ws.getRow(r);
    row.getCell(1).value = entry.pers_no || '';
    row.getCell(2).value = entry.date || '';
    row.getCell(3).value = entry.work_order || '';
    row.getCell(4).value = entry.act || '';
    row.getCell(5).value = entry.aa_code || '';
    const hoursCell = row.getCell(6);
    hoursCell.value = (entry.hours !== undefined && entry.hours !== null && entry.hours !== '') ? Number(entry.hours) : '';
    hoursCell.numFmt = '0.00';
    row.getCell(7).value = entry.discrepancy_reason || '';
    // apply borders for data; if this is a total-row, draw a thicker/separating bottom border and bold the row
    const isTotal = entry.isTotal === true;
    for (let c=1;c<=7;c++){
      const bottomStyle = isTotal ? 'thick' : 'thin';
      const border = {top:{style:'thin'}, left:{style:'thin'}, bottom:{style:bottomStyle}, right:{style:'thin'}};
      if (isTotal) border.bottom.color = {argb:'FF284162'};
      row.getCell(c).border = border;
      if (isTotal) {
        // emphasize totals
        const cell = row.getCell(c);
        cell.font = Object.assign({}, cell.font, {bold:true});
        // set light grey background for total rows
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } };
      }
    }
    r++;
  });

  // Auto-filter
  ws.autoFilter = {from: {row: headerRowIndex, column:1}, to: {row: headerRowIndex, column:7}};

  try {
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = `CATs_Edits_${(new Date()).toISOString().split('T')[0]}.xlsx`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  } catch (e) {
    console.error('ExcelJS export failed', e); alert('Excel export failed: ' + e.message);
  }
}

/* Reuse functions from original cats_edits.js logic (addRowAfter, deleteRow, formatHours, updateTotals, validation, copyToClipboard/email) */
function addRowAfter(button){
  const currentRow = button.closest('tr');
  const newRow = document.createElement('tr'); newRow.className='editable-row'; newRow.setAttribute('data-editable','true');
    newRow.innerHTML = `
    <td class="non-editable"></td>
    <td class="non-editable"></td>
    <td class="editable" contenteditable="true"></td>
    <td class="act-dropdown">
      <input class="act-input act-select" />
    </td>
    <td class="editable" contenteditable="true" oninput="updateTotals()"></td>
    <td class="editable" contenteditable="true" oninput="updateTotals()" onblur="formatHours(this)"></td>
    <td class="non-editable"></td>
    <td class="action-cell"><button class="btn-icon btn-icon-delete" onclick="deleteRow(this)">×</button></td>
  `;
  let insertAfter = currentRow; let nextRow = currentRow.nextElementSibling;
  while (nextRow && nextRow.classList.contains('editable-row')) { insertAfter = nextRow; nextRow = nextRow.nextElementSibling; }
  insertAfter.after(newRow);
  // wire newly created act input to update totals when changed
  try { const ai = newRow.querySelector('.act-input'); if (ai) ai.oninput = updateTotals; } catch(e){}
}
function prefillEditableRow(newRow, dataRow){
  if (!drmisLookup || drmisLookup.size===0) return false;
  // Extract date, pers no, leave code and leave hours from the dataRow (data-row)
  const persNo = (dataRow.querySelector('td:nth-child(1)') && dataRow.querySelector('td:nth-child(1)').textContent) ? dataRow.querySelector('td:nth-child(1)').textContent.trim() : '';
  const dateStr = (dataRow.querySelector('td:nth-child(2)') && dataRow.querySelector('td:nth-child(2)').textContent) ? dataRow.querySelector('td:nth-child(2)').textContent.trim() : '';
  const leaveCode = (dataRow.querySelector('td:nth-child(5)') && dataRow.querySelector('td:nth-child(5)').textContent) ? dataRow.querySelector('td:nth-child(5)').textContent.trim() : '';
  const leaveHours = parseFloat((dataRow.querySelector('td:nth-child(6)') && dataRow.querySelector('td:nth-child(6)').textContent) ? dataRow.querySelector('td:nth-child(6)').textContent.trim() : '') || 0;
  const dateObj = parseDisplayDate(dateStr);
  if (!dateObj) return false;
  const key = dateObj.toISOString().split('T')[0] + '|' + String(persNo || '');
  const candidates = drmisLookup.get(key) || [];
  if (!candidates || candidates.length===0) return false;
  // determine if this data-row indicates a replacement (Oracle replacing DRMiS leave)
  const originalDrmis = dataRow.dataset && dataRow.dataset.originalDrmis ? String(dataRow.dataset.originalDrmis).trim() : '';
  const isReplaced = dataRow.dataset && dataRow.dataset.replaced === 'true';
  // filter out DRMIS entries that represent leave rows we should not use as work entries
  let filtered;
  if (isReplaced && originalDrmis) {
    // If replaced, filter out both the original DRMiS leave and any entry matching the Oracle leave code
    filtered = candidates.filter(c => {
      try {
        const a = String(c.aatype || '').replace(/\D+/g,'');
        const lc = String(leaveCode || '').replace(/\D+/g,'');
        const od = String(originalDrmis || '').replace(/\D+/g,'');
        return a !== lc && a !== od;
      } catch(e){ return true; }
    });
  } else {
    filtered = candidates.filter(c => { try { return String(c.aatype || '').replace(/\D+/g,'') !== String(leaveCode).replace(/\D+/g,''); } catch(e){ return true; } });
  }
  if (filtered.length === 0) return false;

  // compute existing editable rows between dataRow and the following total-row (they consume some of the subtraction already)
  let existingEditableHours = 0;
  let next = dataRow.nextElementSibling;
  while (next && !next.classList.contains('total-row')){
    if (next.classList.contains('editable-row')){
      const h = parseFloat((next.querySelector('td:nth-child(6)') && next.querySelector('td:nth-child(6)').textContent) ? next.querySelector('td:nth-child(6)').textContent.trim() : '') || 0;
      existingEditableHours += (isNaN(h) ? 0 : h);
    }
    next = next.nextElementSibling;
  }

  // Build residual hours for each filtered candidate
  const residuals = filtered.map(c => ({ recOrder: c.recOrder, act: c.act, aatype: c.aatype, hours: Number(c.hours) || 0 }));
  // consume existingEditableHours across residuals
  let toConsume = existingEditableHours;
  for (let i=0;i<residuals.length && toConsume>0;i++){
    const take = Math.min(residuals[i].hours, toConsume);
    residuals[i].hours = Math.max(0, residuals[i].hours - take);
    toConsume -= take;
  }
  // now handle reassignment/subtraction of the leave hours across residuals
  // If this data-row represents an Oracle->DRMiS replacement, we should NOT subtract Oracle leave from DRMIS work (leave was replacement)
  // Additionally, if this reconciliation row indicates Oracle removed the DRMiS leave (Oracle code empty/0 and Oracle hours 0)
  // then we should take the original DRMiS hours (if present) and reassign them to the remaining DRMiS residuals so prefilled editable rows receive those hours.
  const originalDrmisHours = Number(dataRow.dataset && dataRow.dataset.originalDrmisHours ? Number(dataRow.dataset.originalDrmisHours) : 0) || 0;
  const isOracleRemoved = (!isReplaced && originalDrmisHours > 0 && Number(leaveHours) === 0);

  if (isOracleRemoved) {
    // Distribute the original DRMiS hours across residuals proportionally to their current hours when possible,
    // otherwise add to the first residual.
    let toAdd = originalDrmisHours;
    const sumAvail = residuals.reduce((s,r)=>s + (Number(r.hours)||0), 0);
    if (sumAvail > 0) {
      // distribute proportionally and correct rounding by assigning remainder to first
      let allocated = 0;
      for (let i=0;i<residuals.length;i++){
        const share = Math.floor(((residuals[i].hours / sumAvail) * toAdd) * 100) / 100; // two-decimal precision
        residuals[i].hours = Number(residuals[i].hours) + share;
        allocated += share;
      }
      const remainder = Math.round((toAdd - allocated) * 100) / 100;
      if (remainder > 0 && residuals.length>0) residuals[0].hours = Number(residuals[0].hours) + remainder;
    } else {
      // no available hours to proportion against; add to first residual
      if (residuals.length>0) residuals[0].hours = Number(residuals[0].hours) + toAdd;
    }
    // do not subtract leaveHours in this case (we're reallocating)
  } else {
    // default behaviour: subtract the remaining leave hours (leaveHours) across residuals
    let remainingLeave = Math.max(0, leaveHours - existingEditableHours);
    if (isReplaced) remainingLeave = 0;
    for (let i=0;i<residuals.length && remainingLeave>0;i++){
      const take = Math.min(residuals[i].hours, remainingLeave);
      residuals[i].hours = Math.max(0, residuals[i].hours - take);
      remainingLeave -= take;
    }
  }

  // find first residual with hours > 0
  const pick = residuals.find(r => r.hours > 0);
  if (!pick) return false; // nothing left to prefill

  // Fill fields in newRow: td3 Work Order, td4 Act (select), td5 AA Code, td6 Hours
  const cells = newRow.querySelectorAll('td');
  if (cells[2]) cells[2].textContent = pick.recOrder || '';
  if (cells[3]) {
    const actEl = cells[3].querySelector('.act-input') || cells[3].querySelector('.act-select');
    const desired = (pick.act || '').toString().trim();
    if (actEl) {
      try {
        if (actEl.tagName && actEl.tagName.toUpperCase() === 'SELECT') {
          // existing select logic
          let matched = Array.from(actEl.options).some(o => o.value === desired || o.text === desired);
          if (matched) actEl.value = desired;
          else {
            const numeric = desired.replace(/\D+/g,'');
            if (numeric) {
              const padded = numeric.padStart(4,'0');
              matched = Array.from(actEl.options).some(o => o.value === padded || o.text === padded);
              if (matched) actEl.value = padded;
              else {
                // replace with plain text
                cells[3].innerHTML = '';
                cells[3].textContent = desired;
              }
            } else {
              cells[3].innerHTML = '';
              cells[3].textContent = desired;
            }
          }
        } else {
          // input element: set value directly so user can edit further
          try { actEl.value = desired; } catch(e) { cells[3].textContent = desired; }
        }
      } catch(e) { cells[3].textContent = desired; }
    } else {
      cells[3].textContent = desired || '';
    }
  }
  if (cells[4]) cells[4].textContent = pick.aatype || '';
  if (cells[5]) {
    const num = Number(pick.hours) || 0;
    cells[5].textContent = num>0 ? num.toFixed(2) : (pick.hours===''? '': '0.00');
  }
  // update totals styling/logic
  updateTotals();
  return true;
}

function deleteRow(button) { const row = button.closest('tr'); if(row && row.getAttribute('data-editable')==='true'){ row.remove(); updateTotals(); } }
function formatHours(cell){ const value = cell.textContent.trim(); if(value!==''){ const num = parseFloat(value); if(!isNaN(num)) cell.textContent = num.toFixed(2); } }

function updateTotals(){
  const totalRows = document.querySelectorAll('.total-row');
  totalRows.forEach(totalRow => {
    // Sum any editable rows and the associated data-row above the total row
    let sum = 0;
    let node = totalRow.previousElementSibling;
    // walk backwards until we reach a data-row or the table start
    while (node) {
      // only consider rows that have a hours cell
      const hoursCell = node.querySelector && node.querySelector('td:nth-child(6)');
      const txt = hoursCell ? hoursCell.textContent.trim() : '';
      const n = parseFloat(txt);
      if (!isNaN(n)) sum += n;
      // stop once we have processed the data-row (which should be the first non-editable row before editable ones)
      if (node.classList && node.classList.contains('data-row')) break;
      node = node.previousElementSibling;
    }
    const totalHoursCell = totalRow.querySelector('.total-hours');
    if (totalHoursCell) {
      const previousTotal = parseFloat(totalHoursCell.textContent)||0;
      totalHoursCell.textContent = sum.toFixed(2);
      totalHoursCell.classList.remove('equals-eight','over-eight');
      if (sum===8) totalHoursCell.classList.add('equals-eight');
      else if (sum>8) {
        totalHoursCell.classList.add('over-eight');
        if (previousTotal<=8 && !totalHoursCell.dataset.warned) {
          showCustomAlert('Employee Regular Work Hours Exceeded','Total hours exceed 8 hours for this day.');
          totalHoursCell.dataset.warned='true';
        }
      } else delete totalHoursCell.dataset.warned;
    }
  });
}

function validateRow(row){ if (!row.classList.contains('editable-row')) return true; const aa = row.querySelector('td:nth-child(5)').textContent.trim(); const hrs = row.querySelector('td:nth-child(6)').textContent.trim(); return aa!=='' && hrs!==''; }
function validateAllRows(){ const editableRows = document.querySelectorAll('.editable-row'); let all=true; editableRows.forEach(r=>{ if(!validateRow(r)) all=false; }); return all; }

function downloadCatsEditsExcel(){ if (!validateAllRows()) { alert('Please fill in AA Code and Hours for all added rows before exporting.'); return; } downloadCatsEditsExcelInternal(); }

function copyCatsToClipboard(){ if (!validateAllRows()) { showCustomAlert('Validation Required','Please fill in AA Code and Hours for all added rows before copying.'); return; }
  const table = document.getElementById('catsTable');
  // build HTML and plain text identical to original implementation
  let html = '<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11pt;">';
  html += '<tr style="height: 20px;"><td colspan="7" style="background-color: #ffffff; font-weight: bold; font-size: 14pt; padding: 2px 4px; border: none; line-height: 1.1;">Modifications to CATs Entries</td></tr>';
  html += '<tr style="height: 5px;"><td colspan="7" style="border: none; padding: 0;"></td></tr>';
  html += '<tr style="height: 18px;"><td colspan="7" style="background-color: #ffffff; font-size: 11pt; padding: 2px 4px; border: none; line-height: 1.1;">These are the changes that need to be made in CATs on DRMIS</td></tr>';
  html += '<tr style="height: 5px;"><td colspan="7" style="border: none; padding: 0;"></td></tr>';
  html += '<tr style="height: 20px;">';
  const headers = Array.from(table.querySelectorAll('thead th')).slice(0,-1);
  headers.forEach(th=> html += `<td style="background-color: #d4edda; color: #000000; font-weight: bold; padding: 2px 4px; border: 1px solid #000000; line-height: 1.1;">${th.textContent}</td>`);
  html += '</tr>';
  const rows = table.querySelectorAll('tbody tr');
  rows.forEach(row=>{
    const cells = Array.from(row.querySelectorAll('td')).slice(0,-1);
    html += '<tr style="height: 18px;">';
    cells.forEach((td, idx)=>{ let cellValue=''; if (idx===3) { const select = (td.querySelector('.act-input') || td.querySelector('.act-select')); cellValue = select ? (select.value !== undefined ? String(select.value).trim() : td.textContent) : td.textContent; } else cellValue = td.textContent; if (idx===5 && cellValue !== '' && cellValue !== 'Total Hours') { const hours = parseFloat(cellValue); if (!isNaN(hours) && hours !== 0) cellValue = hours.toFixed(2); }
      let cellStyle = 'padding: 2px 4px; border: 1px solid #000000; line-height: 1.1;';
      if (row.classList.contains('total-row')) { cellStyle += ' background-color: #f0f0f0; font-weight: bold; border-bottom: 3px solid #284162;'; if (idx===5) { const hours = parseFloat(cellValue); if (!isNaN(hours)) { if (hours===8) cellStyle += ' color: #008000;'; else if (hours>8) cellStyle += ' color: #FF0000;'; } } } else { cellStyle += ' background-color: #ffffff;'; }
      html += `<td style="${cellStyle}">${cellValue}</td>`;
    });
    html += '</tr>';
  });
  html += '</table>';
  const plainTextParts = [];
  plainTextParts.push('Modifications to CATs Entries\n\n'); plainTextParts.push('These are the changes that need to be made in CATs on DRMIS\n\n'); plainTextParts.push(headers.map(h => h.textContent).join('\t') + '\n');
  rows.forEach(row=>{ const cells = Array.from(row.querySelectorAll('td')).slice(0,-1); plainTextParts.push(cells.map((td,idx)=>{ if (idx===3) { const sel = (td.querySelector('.act-input') || td.querySelector('.act-select')); return sel ? (sel.value !== undefined ? String(sel.value).trim() : td.textContent) : td.textContent; } return td.textContent; }).join('\t') + '\n'); });
  const htmlBlob = new Blob([html], {type:'text/html'}); const txtBlob = new Blob([plainTextParts.join('')], {type:'text/plain'});
  const item = new ClipboardItem({'text/html': htmlBlob, 'text/plain': txtBlob});
  navigator.clipboard.write([item]).then(()=> showCustomAlert('Copied to Clipboard','You can now paste it into Excel or email.')).catch(err=>{ navigator.clipboard.writeText(plainTextParts.join('')).then(()=> showCustomAlert('Copied to Clipboard','Copied as plain text. You can now paste it into Excel or email.')).catch(e=> showCustomAlert('Copy Failed','Failed to copy: ' + e)); });
}

function downloadCatsEditsExcelFormatted(){
  if (!validateAllRows()) { alert('Please fill in AA Code and Hours for all added rows before exporting.'); return; }
  const table = document.getElementById('catsTable');
  if (!table) { alert('No CATs table found'); return; }
  // Reuse the same HTML generation as copyCatsToClipboard but produce a downloadable .xls file
  let html = '<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11pt;">';
  html += '<tr style="height: 20px;"><td colspan="7" style="background-color: #ffffff; font-weight: bold; font-size: 14pt; padding: 2px 4px; border: none; line-height: 1.1;">Modifications to CATs Entries</td></tr>';
  html += '<tr style="height: 5px;"><td colspan="7" style="border: none; padding: 0;"></td></tr>';
  html += '<tr style="height: 18px;"><td colspan="7" style="background-color: #ffffff; font-size: 11pt; padding: 2px 4px; border: none; line-height: 1.1;">These are the changes that need to be made in CATs on DRMIS</td></tr>';
  html += '<tr style="height: 5px;"><td colspan="7" style="border: none; padding: 0;"></td></tr>';
  html += '<tr style="height: 20px;">';
  const headers = Array.from(table.querySelectorAll('thead th')).slice(0,-1);
  headers.forEach(th=> html += `<td style="background-color: #d4edda; color: #000000; font-weight: bold; padding: 2px 4px; border: 1px solid #000000; line-height: 1.1;">${th.textContent}</td>`);
  html += '</tr>';
  const rows = table.querySelectorAll('tbody tr');
  rows.forEach(row=>{
    const cells = Array.from(row.querySelectorAll('td')).slice(0,-1);
    html += '<tr style="height: 18px;">';
    cells.forEach((td, idx)=>{ let cellValue=''; if (idx===3) { const select = (td.querySelector('.act-input') || td.querySelector('.act-select')); cellValue = select ? (select.value !== undefined ? String(select.value).trim() : td.textContent) : td.textContent; } else cellValue = td.textContent; if (idx===5 && cellValue !== '' && cellValue !== 'Total Hours') { const hours = parseFloat(cellValue); if (!isNaN(hours) && hours !== 0) cellValue = hours.toFixed(2); }
      let cellStyle = 'padding: 2px 4px; border: 1px solid #000000; line-height: 1.1;';
      if (row.classList.contains('total-row')) { cellStyle += ' background-color: #f0f0f0; font-weight: bold; border-bottom: 3px solid #284162;'; if (idx===5) { const hours = parseFloat(cellValue); if (!isNaN(hours)) { if (hours===8) cellStyle += ' color: #008000;'; else if (hours>8) cellStyle += ' color: #FF0000;'; } } } else { cellStyle += ' background-color: #ffffff;'; }
      html += `<td style="${cellStyle}">${cellValue}</td>`;
    });
    html += '</tr>';
  });
  html += '</table>';

  const blob = new Blob([html], {type: 'application/vnd.ms-excel'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = `CATs_Edits_${(new Date()).toISOString().split('T')[0]}.xls`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}

function emailCatsTable(){ if (!validateAllRows()) { showCustomAlert('Validation Required','Please fill in AA Code and Hours for all added rows before emailing.'); return; } const table = document.getElementById('catsTable'); let persNo=''; const rows = table.querySelectorAll('tbody tr'); for (let row of rows){ const cells = row.querySelectorAll('td'); if (cells[0] && cells[0].textContent.trim()!=='') { persNo = cells[0].textContent.trim(); break; }}
  // Build html body
  let html = '<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11pt;">';
  html += '<tr style="height: 20px;">'; const headers = Array.from(table.querySelectorAll('thead th')).slice(0,-1); headers.forEach(th=> html += `<td style="background-color: #d4edda; color: #000000; font-weight: bold; padding: 2px 4px; border: 1px solid #000000; line-height: 1.1;">${th.textContent}</td>`); html += '</tr>';
  rows.forEach(row=>{ const cells = Array.from(row.querySelectorAll('td')).slice(0,-1); html += '<tr style="height: 18px;">'; cells.forEach((td,idx)=>{ let cellValue = (idx===3 ? ((td.querySelector('.act-input') || td.querySelector('.act-select')) ? ((td.querySelector('.act-input') || td.querySelector('.act-select')).value !== undefined ? String((td.querySelector('.act-input') || td.querySelector('.act-select')).value).trim() : td.textContent) : td.textContent) : td.textContent); if (idx===5 && cellValue && cellValue.trim()!=='') { const hv = parseFloat(cellValue); if (!isNaN(hv)) cellValue = hv===0 ? '0' : hv.toFixed(2); if (row.classList.contains('total-row')) { if (hv===8) {} else if (hv>8) {} } } let cellStyle = 'padding:2px 4px; border:1px solid #000000; line-height:1.1;'; if (row.classList.contains('total-row')) cellStyle += ' background-color:#f0f0f0; font-weight:bold; border-bottom:3px solid #000000;'; html += `<td style="${cellStyle}">${cellValue}</td>`; }); html += '</tr>'; }); html += '</table>';
  const emailText = `Good Day,\n\nI have just audited an employee's leave and found the following discrepancies that need to be updated in DRMIS.\n\nEmployee Number: ${persNo}\n\n`;
  const plainText = emailText + headers.map(h=>h.textContent).join('\t') + '\n' + Array.from(rows).map(r=> Array.from(r.querySelectorAll('td')).slice(0,-1).map((td,idx)=> idx===3 ? ((td.querySelector('.act-input') || td.querySelector('.act-select')) ? ((td.querySelector('.act-input') || td.querySelector('.act-select')).value !== undefined ? String((td.querySelector('.act-input') || td.querySelector('.act-select')).value).trim() : td.textContent) : td.textContent) : td.textContent).join('\t')).join('\n') + '\n\nThanks';
  const fullHtmlBody = `<p>Good Day,</p><p>I have just audited an employee's leave and found the following discrepancies that need to be updated in DRMIS.</p><p>Employee Number: ${persNo}</p><br>${html}<br><p>Thanks</p>`;
  const htmlBlob = new Blob([fullHtmlBody], {type:'text/html'}); const txtBlob = new Blob([plainText], {type:'text/plain'});
  const item = new ClipboardItem({'text/html': htmlBlob, 'text/plain': txtBlob});
  navigator.clipboard.write([item]).then(()=>{ const subject = encodeURIComponent(`DRMIS Leave Discrepancies - Employee ${persNo}`); showCustomAlert('Email Content Copied','The email content with the table has been copied to your clipboard. Click OK to open your email client, then paste (CTRL + V) the content into the message body.', ()=> { window.location.href = `mailto:?subject=${subject}`; }); }).catch(err=>{ navigator.clipboard.writeText(plainText).then(()=>{ const subject = encodeURIComponent(`DRMIS Leave Discrepancies - Employee ${persNo}`); showCustomAlert('Email Content Copied','The email content has been copied as plain text. Click OK to open your email client, then paste (CTRL + V) the content into the message body.', ()=> { window.location.href = `mailto:?subject=${subject}`; }); }).catch(e=> showCustomAlert('Copy Failed','Failed to copy: ' + e)); });
}

// Export the current Reconciliation results table (live DOM) as a styled .xlsx using ExcelJS
async function downloadReconciliationXlsxExcelJS(){
  const table = document.getElementById('resultTable');
  if (!table) { alert('No results table found'); return; }
  // Ensure there is data
  const tbody = table.querySelector('tbody');
  if (!tbody) { alert('No results to export'); return; }
  // Collect header names (skip the first checkbox column)
  const thead = table.querySelector('thead');
  const headerThs = thead ? Array.from(thead.querySelectorAll('th')) : [];
  const headers = headerThs.slice(1).map(th => th.textContent.trim());

  // Build live rows, skipping any select-all row (which contains #selectAllCats)
  const rows = Array.from(tbody.querySelectorAll('tr')).filter(tr => !tr.querySelector('#selectAllCats'));
  if (rows.length === 0) { alert('No mismatch rows to export'); return; }
  const liveData = rows.map(tr => {
    const cells = Array.from(tr.querySelectorAll('td'));
    // skip first cell (checkbox)
    const dataCells = cells.slice(1);
    return dataCells.map(td => td.textContent.trim());
  });

  if (typeof ExcelJS === 'undefined') { alert('ExcelJS not available'); return; }
  const wb = new ExcelJS.Workbook(); wb.creator='Leave Reconcile'; wb.created = new Date();
  const ws = wb.addWorksheet('Reconciliation');

  // Title row
  ws.mergeCells('A1:F1'); ws.getCell('A1').value = 'Reconciliation Results'; ws.getCell('A1').font = {name:'Calibri', size:14, bold:true}; ws.getCell('A1').alignment = {horizontal:'center', vertical:'middle'}; ws.getRow(1).height = 20;
  // Subtitle row
  ws.mergeCells('A2:F2'); ws.getCell('A2').value = `Generated: ${(new Date()).toLocaleString()}`; ws.getCell('A2').font = {name:'Calibri', size:10}; ws.getCell('A2').alignment = {horizontal:'left', vertical:'middle'};

  const headerRowIndex = 4;
  const headerRow = ws.getRow(headerRowIndex);
  headers.forEach((h,i)=>{
    const cell = headerRow.getCell(i+1);
    cell.value = h;
    cell.font = {name:'Calibri', bold:true};
    cell.fill = {type:'pattern', pattern:'solid', fgColor:{argb:'FFD4EDDA'}};
    cell.border = {top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'}};
    cell.alignment = {vertical:'middle', horizontal:'center'};
  });
  ws.getRow(headerRowIndex).height = 18;

  // Columns widths heuristic
  ws.columns = headers.map(h => ({ key: h, width: Math.max(10, Math.min(30, h.length + 6)) }));

  // Data rows
  let r = headerRowIndex + 1;
  liveData.forEach(rowArr => {
    const row = ws.getRow(r);
    rowArr.forEach((val, idx) => {
      const cell = row.getCell(idx+1);
      // Format hours columns if header contains 'Hours'
      if (/hours/i.test(headers[idx])) {
        const n = parseFloat(val);
        cell.value = isNaN(n) ? (val === '' ? '' : val) : n;
        cell.numFmt = '0.00';
        cell.alignment = {horizontal:'right'};
      } else {
        cell.value = val;
      }
      cell.border = {top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'}};
    });
    r++;
  });

  // Autofilter
  ws.autoFilter = { from: { row: headerRowIndex, column: 1 }, to: { row: headerRowIndex, column: headers.length } };

  try {
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = `Reconciliation_${(new Date()).toISOString().split('T')[0]}.xlsx`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  } catch (e) {
    console.error('Export failed', e); alert('Export failed: ' + e.message);
  }
}

/* Small UI helpers from original */
function showCustomAlert(title, message, onOk){ const modal = document.createElement('div'); modal.style.cssText='position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);display:flex;align-items:center;justify-content:center;z-index:10000;'; const alertBox = document.createElement('div'); alertBox.style.cssText='background:white;border-left:4px solid #d3080c;padding:0;max-width:500px;box-shadow:0 2px 8px rgba(0,0,0,0.2);font-family:"Noto Sans",sans-serif;'; const header = document.createElement('div'); header.style.cssText='background:#f8f8f8;padding:15px 20px;border-bottom:1px solid #ddd;'; const titleEl = document.createElement('h3'); titleEl.textContent = title; titleEl.style.cssText='margin:0;color:#333;font-size:1.1rem;font-weight:700;'; const content = document.createElement('div'); content.style.cssText='padding:20px;'; const messageEl = document.createElement('p'); messageEl.textContent = message; messageEl.style.cssText='margin:0 0 20px 0;color:#333;line-height:1.5;'; const buttonContainer = document.createElement('div'); buttonContainer.style.cssText='text-align:right;'; const okButton = document.createElement('button'); okButton.textContent='OK'; okButton.style.cssText='background:#284162;color:white;border:none;padding:10px 24px;cursor:pointer;font-size:1rem;font-weight:400;min-width:80px;'; okButton.onmouseover = ()=> okButton.style.background='#1c578a'; okButton.onmouseout = ()=> okButton.style.background='#284162'; okButton.onclick = ()=>{ document.body.removeChild(modal); if (onOk && typeof onOk==='function') onOk(); }; header.appendChild(titleEl); buttonContainer.appendChild(okButton); content.appendChild(messageEl); content.appendChild(buttonContainer); alertBox.appendChild(header); alertBox.appendChild(content); modal.appendChild(alertBox); document.body.appendChild(modal); okButton.focus(); }

/* End of app.js */