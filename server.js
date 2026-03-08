const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

// Serve static files (the portal)
app.use(express.static(path.join(__dirname)));

const DATA_DIR = path.join(__dirname, 'data');
const JSON_FILE = path.join(DATA_DIR, 'roster.json');
const CSV_FILE = path.join(__dirname, 'roster.csv');

function parseCSV(text) {
  const rows = [];
  let cur = '';
  let row = [];
  let inQuotes = false;
  const normalized = String(text || '').replace(/\uFEFF/g, '');

  for (let i = 0; i < normalized.length; i++) {
    const ch = normalized[i];
    const next = normalized[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === ',' && !inQuotes) {
      row.push(cur);
      cur = '';
    } else if ((ch === '\n' || ch === '\r') && !inQuotes) {
      if (ch === '\r' && next === '\n') i++;
      row.push(cur);
      rows.push(row);
      row = [];
      cur = '';
    } else {
      cur += ch;
    }
  }

  if (cur !== '' || row.length) {
    row.push(cur);
    rows.push(row);
  }

  return rows.filter(r => r.some(c => String(c || '').trim() !== ''));
}

function scoreHeaderRow(row) {
  const cells = (row || []).map(c => normalizeKey(c));
  let score = 0;
  const known = [
    'id', 'officerid', 'officer_id', 'badgeid', 'badgenumber', 'badge_number',
    'name', 'rpname', 'rp_name', 'officername', 'officer_name',
    'callsign', 'call_sign', 'rank', 'division', 'unit'
  ].map(normalizeKey);

  cells.forEach((cell) => {
    if (known.includes(cell)) score += 2;
    if (/id|name|callsign|rank|division|unit/.test(cell)) score += 1;
  });
  return score;
}

function detectHeaderRowIndex(rows) {
  const maxScan = Math.min(rows.length, 12);
  let bestIdx = 0;
  let bestScore = -1;

  for (let i = 0; i < maxScan; i++) {
    const s = scoreHeaderRow(rows[i]);
    if (s > bestScore) {
      bestScore = s;
      bestIdx = i;
    }
  }

  // If no meaningful header signal is found, keep first row behavior.
  return bestScore <= 0 ? 0 : bestIdx;
}

function toObjectsFromCSV(csvText) {
  const rows = parseCSV(csvText);
  if (!rows.length) return [];

  const headerRowIndex = detectHeaderRowIndex(rows);
  const header = rows[headerRowIndex].map((h, idx) => String(h || '').trim() || ('col' + idx));
  const dataRows = rows.slice(headerRowIndex + 1);

  return dataRows.map(r => {
    const obj = {};
    header.forEach((h, idx) => {
      obj[h] = String(r[idx] || '').trim();
    });
    return obj;
  }).filter(obj => Object.values(obj).some(v => String(v).trim() !== ''));
}

function sanitizeName(name) {
  return String(name || 'tab').toLowerCase().replace(/[^a-z0-9_-]/g, '_');
}

function normalizeGoogleCsvUrl(rawUrl) {
  const input = String(rawUrl || '').trim();
  if (!input) return '';

  let parsed;
  try {
    parsed = new URL(input);
  } catch (e) {
    return input;
  }

  const host = String(parsed.hostname || '').toLowerCase();
  const isGoogleSheet = host.includes('docs.google.com') && parsed.pathname.includes('/spreadsheets/');
  if (!isGoogleSheet) return input;

  const tqx = String(parsed.searchParams.get('tqx') || '').toLowerCase();
  const output = String(parsed.searchParams.get('output') || '').toLowerCase();
  if (tqx.includes('out:csv') || output === 'csv') return input;

  const idMatch = parsed.pathname.match(/\/d\/([^/]+)/);
  if (!idMatch || !idMatch[1]) return input;

  let gid = parsed.searchParams.get('gid');
  if (!gid && parsed.hash) {
    const hash = parsed.hash.replace(/^#/, '');
    const hashParams = new URLSearchParams(hash);
    gid = hashParams.get('gid');
  }

  const csvUrl = new URL('https://docs.google.com/spreadsheets/d/' + idMatch[1] + '/gviz/tq');
  csvUrl.searchParams.set('tqx', 'out:csv');
  csvUrl.searchParams.set('gid', gid || '0');
  return csvUrl.toString();
}

function getGoogleSheetId(urlText) {
  const input = String(urlText || '').trim();
  const m1 = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (m1 && m1[1]) return m1[1];
  const m2 = input.match(/\/spreadsheets\/d\/e\/([a-zA-Z0-9-_]+)/);
  if (m2 && m2[1]) return m2[1];
  return '';
}

function getGoogleGid(urlText) {
  const input = String(urlText || '').trim();
  try {
    const u = new URL(input);
    const gid = u.searchParams.get('gid');
    if (gid) return gid;
    if (u.hash) {
      const hash = u.hash.replace(/^#/, '');
      const params = new URLSearchParams(hash);
      if (params.get('gid')) return params.get('gid');
    }
  } catch (e) {
    // Ignore parse errors and fall back to default gid.
  }
  return '0';
}

function buildGoogleCsvCandidates(rawUrl) {
  const original = String(rawUrl || '').trim();
  const candidates = [];
  if (!original) return candidates;
  candidates.push(original);

  const normalized = normalizeGoogleCsvUrl(original);
  if (normalized && !candidates.includes(normalized)) candidates.push(normalized);

  const sheetId = getGoogleSheetId(original);
  const gid = getGoogleGid(original);
  if (sheetId) {
    const exportUrl = 'https://docs.google.com/spreadsheets/d/' + sheetId + '/export?format=csv&gid=' + encodeURIComponent(gid);
    if (!candidates.includes(exportUrl)) candidates.push(exportUrl);

    const gvizUrl = 'https://docs.google.com/spreadsheets/d/' + sheetId + '/gviz/tq?tqx=out:csv&gid=' + encodeURIComponent(gid);
    if (!candidates.includes(gvizUrl)) candidates.push(gvizUrl);
  }

  return candidates;
}

async function fetchCsvWithFallback(rawUrl) {
  const candidates = buildGoogleCsvCandidates(rawUrl);
  let lastStatus = 0;
  let lastStatusText = 'Unknown error';
  let lastUrl = String(rawUrl || '');

  for (const candidate of candidates) {
    lastUrl = candidate;
    const response = await fetch(candidate);
    if (response.ok) {
      const text = await response.text();
      return { ok: true, csvText: text, urlUsed: candidate };
    }
    lastStatus = response.status;
    lastStatusText = response.statusText;
  }

  const guidance = 'Google returned HTTP ' + lastStatus + ' for all link formats. ' +
    'Verify the sheet/tab is published to web as CSV and accessible to anyone with the link.';

  return {
    ok: false,
    error: guidance,
    status: lastStatus,
    statusText: lastStatusText,
    urlTried: lastUrl,
    candidates
  };
}

function saveTabJson(tabName, records) {
  ensureDataDir();
  const safe = sanitizeName(tabName);
  const filePath = path.join(DATA_DIR, safe + '.json');
  fs.writeFileSync(filePath, JSON.stringify(records, null, 2));
  return filePath;
}

function ensureDataDir(){
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);
}

function loadJson(){
  ensureDataDir();
  if (!fs.existsSync(JSON_FILE)) {
    // If CSV exists, try to load from it
    if (fs.existsSync(CSV_FILE)) {
      const csv = fs.readFileSync(CSV_FILE, 'utf8');
      const rows = csv.replace(/\r/g,'').split('\n').filter(Boolean);
      const header = rows.shift().split(',');
      const out = rows.map(r=>{
        const cols = r.split(',');
        const obj = {};
        header.forEach((h,i)=> obj[h.trim()]= (cols[i]||'').trim());
        return obj;
      });
      fs.writeFileSync(JSON_FILE, JSON.stringify(out, null, 2));
      return out;
    }
    fs.writeFileSync(JSON_FILE, '[]');
    return [];
  }
  try {
    return JSON.parse(fs.readFileSync(JSON_FILE,'utf8')||'[]');
  } catch(e){
    console.error('Failed to parse roster.json', e);
    return [];
  }
}

function saveJson(data){
  ensureDataDir();
  fs.writeFileSync(JSON_FILE, JSON.stringify(data,null,2));
}

function normalizeKey(key) {
  return String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function getByAliases(record, aliases) {
  const keys = Object.keys(record || {});
  for (const alias of aliases) {
    const wanted = normalizeKey(alias);
    const foundKey = keys.find(k => normalizeKey(k) === wanted);
    if (foundKey) {
      const value = String(record[foundKey] || '').trim();
      if (value) return value;
    }
  }
  return '';
}

function firstNonEmpty(values) {
  for (const v of values) {
    const s = String(v || '').trim();
    if (s) return s;
  }
  return '';
}

function cleanField(value) {
  const v = String(value || '').trim();
  if (!v) return '';
  const lowered = v.toLowerCase();
  if (['n/a', 'na', 'none', 'null', '-', '--'].includes(lowered)) return '';
  return v;
}

function isMeaningful(value) {
  return cleanField(value) !== '';
}

function mapRosterRecords(records) {
  return records.map((r, idx) => {
    const kv = {};
    Object.keys(r || {}).forEach((k) => {
      kv[normalizeKey(k)] = cleanField(r[k]);
    });

    const pick = (...keys) => firstNonEmpty(keys.map((k) => kv[normalizeKey(k)]));

    const id = pick('officer_id', 'officerid', 'badge_number', 'badgenumber', 'badgeid', 'employeeid', 'memberid', 'id');
    let name = pick('rp_name', 'officer_name', 'officername', 'character_name', 'charactername', 'fullname', 'name');
    let callsign = pick('callsign', 'call_sign', 'callsigns', 'radioid', 'radio');
    const rank = pick('rank', 'officer_rank', 'officerrank', 'grade');
    const division = pick('division', 'department_division', 'departmentdivision', 'unit');

    // Recover from imports where a text name landed in Callsign.
    if (!name && callsign && /[a-z]/i.test(callsign) && !/\d/.test(callsign)) {
      name = callsign;
      callsign = '';
    }

    const hasIdentity = [id, name, callsign].some(isMeaningful);
    if (!hasIdentity) return null;

    return {
      ID: id || ('IMP-' + Date.now() + '-' + idx),
      Name: name,
      Callsign: callsign,
      Rank: rank,
      Division: division
    };
  }).filter(Boolean);
}

// API: list roster
app.get('/api/roster', (req,res)=>{
  const data = loadJson();
  res.json(data);
});

// API: add roster item
app.post('/api/roster', (req,res)=>{
  const data = loadJson();
  const item = req.body || {};
  // Ensure an ID
  item.ID = item.ID || String(Date.now());
  data.push(item);
  saveJson(data);
  res.status(201).json(item);
});

// API: update by ID
app.put('/api/roster/:id', (req,res)=>{
  const id = req.params.id;
  const data = loadJson();
  const idx = data.findIndex(x=>String(x.ID) === String(id));
  if (idx === -1) return res.status(404).json({error:'Not found'});
  data[idx] = Object.assign({}, data[idx], req.body);
  saveJson(data);
  res.json(data[idx]);
});

// API: delete by ID
app.delete('/api/roster/:id', (req,res)=>{
  const id = req.params.id;
  let data = loadJson();
  const before = data.length;
  data = data.filter(x=>String(x.ID) !== String(id));
  saveJson(data);
  res.json({deleted: before - data.length});
});

// Import one or many Google Sheet CSV tabs server-side (avoids browser CORS issues)
// Body format:
// {
//   "tabs": [
//     {"name":"roster","url":"https://...output=csv"},
//     {"name":"divisions","url":"https://...output=csv"}
//   ]
// }
app.post('/api/sheets/import', async (req, res) => {
  try {
    const tabs = Array.isArray(req.body && req.body.tabs) ? req.body.tabs : [];
    if (!tabs.length) {
      return res.status(400).json({ error: 'No tabs provided. Send body: { tabs: [{name,url}] }' });
    }

    const result = [];

    for (const tab of tabs) {
      const name = sanitizeName(tab && tab.name);
      const url = String(tab && tab.url || '').trim();
      if (!name || !url) {
        result.push({ name: name || 'unknown', ok: false, error: 'Missing name or url' });
        continue;
      }

      try {
        const fetchResult = await fetchCsvWithFallback(url);
        if (!fetchResult.ok) {
          result.push({
            name,
            ok: false,
            error: fetchResult.error,
            status: fetchResult.status,
            statusText: fetchResult.statusText
          });
          continue;
        }

        const csvText = fetchResult.csvText;
        const records = toObjectsFromCSV(csvText);
        const sampleHeaders = records[0] ? Object.keys(records[0]) : [];

        if (name === 'roster') {
          const roster = mapRosterRecords(records);
          saveJson(roster);
          saveTabJson(name, records);
          result.push({
            name,
            ok: true,
            rows: roster.length,
            parsedRows: records.length,
            headers: sampleHeaders,
            sourceUrl: fetchResult.urlUsed,
            note: 'Roster replaced from sheet'
          });
        } else {
          saveTabJson(name, records);
          result.push({ name, ok: true, rows: records.length, headers: sampleHeaders, sourceUrl: fetchResult.urlUsed });
        }
      } catch (err) {
        result.push({ name, ok: false, error: err.message || String(err) });
      }
    }

    return res.json({ ok: true, result });
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

app.get('/api/sheets/tabs', (req, res) => {
  ensureDataDir();
  const files = fs.readdirSync(DATA_DIR)
    .filter(f => f.endsWith('.json'))
    .map(f => f.replace(/\.json$/, ''));
  res.json(files);
});

app.get('/api/sheets/tab/:name', (req, res) => {
  const safe = sanitizeName(req.params.name);
  const filePath = path.join(DATA_DIR, safe + '.json');
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'Tab not found' });
  }
  try {
    const data = JSON.parse(fs.readFileSync(filePath, 'utf8') || '[]');
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.listen(PORT, ()=>{
  console.log('Server running on http://localhost:'+PORT);
});
