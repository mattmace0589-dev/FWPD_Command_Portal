const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const crypto = require('crypto');
let PgPool = null;
try {
  ({ Pool: PgPool } = require('pg'));
} catch (e) {
  PgPool = null;
}

const app = express();
const PORT = process.env.PORT || 3000;
const BOOTED_AT = new Date().toISOString();
const AUTH_SECRET = String(process.env.AUTH_SECRET || 'fwpd-default-auth-secret-change-me');
const AUTH_TOKEN_TTL_SECONDS = Number(process.env.AUTH_TOKEN_TTL_SECONDS || (60 * 60 * 24 * 30));
const DATABASE_URL = String(process.env.DATABASE_URL || '').trim();
const DB_ENABLED = !!(DATABASE_URL && PgPool);
let dbPool = null;

app.use(cors());
app.use(express.json());

// Serve static files (the portal)
app.use(express.static(path.join(__dirname)));

const DATA_DIR = path.join(__dirname, 'data');
const JSON_FILE = path.join(DATA_DIR, 'roster.json');
const CSV_FILE = path.join(__dirname, 'roster.csv');
const SHEETS_CONFIG_FILE = path.join(DATA_DIR, 'sheets-config.json');
const USERS_FILE = path.join(DATA_DIR, 'users.json');
const SESSIONS_FILE = path.join(DATA_DIR, 'sessions.json');
const REPORT_APPROVALS_FILE = path.join(DATA_DIR, 'report_approvals.json');
const REPORTS_CONFIG_FILE = path.join(DATA_DIR, 'reports-config.json');
const INTERNAL_MESSAGES_FILE = path.join(DATA_DIR, 'internal_mailbox.json');
const ROLE_OVERRIDES_FILE = path.join(DATA_DIR, 'role_overrides.json');
const FTO_FILE = path.join(DATA_DIR, 'fto.json');
const HARDSET_DEFAULT_ROSTER_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vR6_40O35zd-9GMo_nTg5KS76Svzt1P8ZKrfBQwPAtLloGFtpE1r4JBP3t-F-meLlDKCpvWzZkhMlOb/pub?output=csv&gid=757275616';
const PORTAL_OWNER_EMAILS = String(process.env.PORTAL_OWNER_EMAILS || 'mattprz89@gmail.com')
  .split(',')
  .map((s) => String(s || '').trim().toLowerCase())
  .filter(Boolean);
const PORTAL_DEFAULT_ADMIN_EMAILS = String(process.env.PORTAL_DEFAULT_ADMIN_EMAILS || 'mattprz89@gmail.com')
  .split(',')
  .map((s) => String(s || '').trim().toLowerCase())
  .filter(Boolean);
const INTERNAL_DATA_FILES = new Set([
  'sheets-config.json',
  'reports-config.json',
  'users.json',
  'sessions.json',
  'report_approvals.json',
  'internal_mailbox.json',
  'role_overrides.json',
  'fto.json'
]);

function getDbPool() {
  if (!DB_ENABLED) return null;
  if (!dbPool) {
    const useSsl = !/localhost|127\.0\.0\.1/i.test(DATABASE_URL);
    dbPool = new PgPool({
      connectionString: DATABASE_URL,
      ssl: useSsl ? { rejectUnauthorized: false } : false
    });
  }
  return dbPool;
}

async function dbQuery(text, params = []) {
  const pool = getDbPool();
  if (!pool) throw new Error('DATABASE_URL is not configured.');
  return pool.query(text, params);
}

async function initDatabasePersistence() {
  if (!DB_ENABLED) {
    console.log('DB persistence disabled (DATABASE_URL not set or pg not installed).');
    return;
  }

  await dbQuery(`
    CREATE TABLE IF NOT EXISTS users (
      email TEXT PRIMARY KEY,
      password_hash TEXT NOT NULL,
      created_at TEXT,
      password_updated_at TEXT,
      auto_provisioned BOOLEAN DEFAULT FALSE
    )
  `);

  await dbQuery(`
    CREATE TABLE IF NOT EXISTS sessions (
      token TEXT PRIMARY KEY,
      email TEXT NOT NULL,
      created_at TEXT
    )
  `);

  await dbQuery(`
    CREATE TABLE IF NOT EXISTS role_overrides (
      email TEXT PRIMARY KEY,
      role TEXT NOT NULL,
      updated_at TEXT
    )
  `);

  // Hydrate runtime JSON files from DB so existing synchronous app code keeps working.
  const dbUsers = (await dbQuery('SELECT email, password_hash, created_at, password_updated_at, auto_provisioned FROM users')).rows;
  const dbSessions = (await dbQuery('SELECT token, email, created_at FROM sessions')).rows;
  const dbRoleOverrides = (await dbQuery('SELECT email, role FROM role_overrides')).rows;

  if (dbUsers.length) {
    const payload = dbUsers.map((r) => ({
      email: normalizeEmail(r.email),
      passwordHash: String(r.password_hash || ''),
      createdAt: String(r.created_at || ''),
      passwordUpdatedAt: String(r.password_updated_at || ''),
      autoProvisioned: !!r.auto_provisioned
    }));
    saveJsonFile(USERS_FILE, payload);
  }

  if (dbSessions.length) {
    const payload = dbSessions.map((r) => ({
      token: String(r.token || ''),
      email: normalizeEmail(r.email),
      createdAt: String(r.created_at || '')
    }));
    saveJsonFile(SESSIONS_FILE, payload);
  }

  if (dbRoleOverrides.length) {
    const payload = {};
    dbRoleOverrides.forEach((r) => {
      const email = normalizeEmail(r.email);
      const role = String(r.role || '').trim().toLowerCase();
      if (!email || !role) return;
      payload[email] = role;
    });
    saveJsonFile(ROLE_OVERRIDES_FILE, payload);
  }

  console.log('DB persistence initialized. Users:', dbUsers.length, '| Sessions:', dbSessions.length, '| Role overrides:', dbRoleOverrides.length);
}

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
  return '';
}

function getGooglePublishedKey(urlText) {
  const input = String(urlText || '').trim();
  const m = input.match(/\/spreadsheets\/d\/e\/([a-zA-Z0-9-_]+)/);
  return (m && m[1]) ? m[1] : '';
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
  const pubKey = getGooglePublishedKey(original);
  const gid = getGoogleGid(original);
  if (sheetId) {
    const exportUrl = 'https://docs.google.com/spreadsheets/d/' + sheetId + '/export?format=csv&gid=' + encodeURIComponent(gid);
    if (!candidates.includes(exportUrl)) candidates.push(exportUrl);

    const gvizUrl = 'https://docs.google.com/spreadsheets/d/' + sheetId + '/gviz/tq?tqx=out:csv&gid=' + encodeURIComponent(gid);
    if (!candidates.includes(gvizUrl)) candidates.push(gvizUrl);
  }

  if (pubKey) {
    const publishedCsvUrl = 'https://docs.google.com/spreadsheets/d/e/' + pubKey + '/pub?output=csv&gid=' + encodeURIComponent(gid);
    if (!candidates.includes(publishedCsvUrl)) candidates.push(publishedCsvUrl);
  }

  return candidates;
}

async function fetchCsvWithFallback(rawUrl) {
  const candidates = buildGoogleCsvCandidates(rawUrl);
  let lastStatus = 0;
  let lastStatusText = 'Unknown error';
  let lastUrl = String(rawUrl || '');
  let htmlResponseSeen = false;

  const looksLikeHtml = (text) => {
    const sample = String(text || '').trim().slice(0, 600).toLowerCase();
    return sample.startsWith('<!doctype html') ||
      sample.startsWith('<html') ||
      sample.includes('<head>') ||
      sample.includes('<meta ');
  };

  for (const candidate of candidates) {
    lastUrl = candidate;
    const response = await fetch(candidate);
    if (response.ok) {
      const text = await response.text();
      const contentType = String(response.headers.get('content-type') || '').toLowerCase();
      if (contentType.includes('text/html') || looksLikeHtml(text)) {
        htmlResponseSeen = true;
        continue;
      }
      return { ok: true, csvText: text, urlUsed: candidate };
    }
    lastStatus = response.status;
    lastStatusText = response.statusText;
  }

  const guidance = htmlResponseSeen
    ? 'Google returned an HTML page instead of CSV. Use a published CSV URL for a specific tab (gid).'
    : 'Google returned HTTP ' + lastStatus + ' for all link formats. Verify the sheet/tab is published to web as CSV and accessible to anyone with the link.';

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

function loadJsonFile(filePath, fallback = []) {
  ensureDataDir();
  if (!fs.existsSync(filePath)) {
    fs.writeFileSync(filePath, JSON.stringify(fallback, null, 2));
    return fallback;
  }
  try {
    return JSON.parse(fs.readFileSync(filePath, 'utf8') || JSON.stringify(fallback));
  } catch (e) {
    return fallback;
  }
}

function saveJsonFile(filePath, data) {
  ensureDataDir();
  fs.writeFileSync(filePath, JSON.stringify(data, null, 2));
}

function loadReportsConfig() {
  return loadJsonFile(REPORTS_CONFIG_FILE, {
    disciplineTabNames: ['discipline_records', 'disciplinary_forms'],
    evaluationTabNames: ['cadet_evaluations']
  });
}

function saveReportsConfig(config) {
  const defaults = {
    disciplineTabNames: ['discipline_records', 'disciplinary_forms'],
    evaluationTabNames: ['cadet_evaluations']
  };
  const incoming = Object.assign({}, defaults, config || {});
  incoming.disciplineTabNames = Array.isArray(incoming.disciplineTabNames)
    ? Array.from(new Set(incoming.disciplineTabNames.map(sanitizeName).filter(Boolean)))
    : defaults.disciplineTabNames;
  incoming.evaluationTabNames = Array.isArray(incoming.evaluationTabNames)
    ? Array.from(new Set(incoming.evaluationTabNames.map(sanitizeName).filter(Boolean)))
    : defaults.evaluationTabNames;
  saveJsonFile(REPORTS_CONFIG_FILE, incoming);
  return incoming;
}

function loadReportApprovals() {
  return loadJsonFile(REPORT_APPROVALS_FILE, []);
}

function saveReportApprovals(items) {
  saveJsonFile(REPORT_APPROVALS_FILE, Array.isArray(items) ? items : []);
}

function loadInternalMessages() {
  return loadJsonFile(INTERNAL_MESSAGES_FILE, []);
}

function saveInternalMessages(items) {
  saveJsonFile(INTERNAL_MESSAGES_FILE, Array.isArray(items) ? items : []);
}

function readTabRecords(tabName) {
  const safe = sanitizeName(tabName);
  const filePath = path.join(DATA_DIR, safe + '.json');
  if (!fs.existsSync(filePath)) return [];
  try {
    const parsed = JSON.parse(fs.readFileSync(filePath, 'utf8') || '[]');
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function listImportedTabNames() {
  ensureDataDir();
  return fs.readdirSync(DATA_DIR)
    .filter(f => f.endsWith('.json'))
    .filter(f => !INTERNAL_DATA_FILES.has(f))
    .map(f => f.replace(/\.json$/, ''));
}

function getFieldByAliases(row, aliases) {
  const keys = Object.keys(row || {});
  for (const alias of aliases) {
    const wanted = normalizeKey(alias);
    const found = keys.find(k => normalizeKey(k) === wanted);
    if (!found) continue;
    const value = String(row[found] || '').trim();
    if (value) return value;
  }
  return '';
}

function stableHash(value) {
  return crypto.createHash('sha1').update(String(value || '')).digest('hex');
}

function buildReportFingerprint(row, fallbackType) {
  const officer = getFieldByAliases(row, ['officer_name', 'name', 'rp_name', 'cadet_name', 'character_name']);
  const date = getFieldByAliases(row, ['date', 'incident_date', 'created_date', 'timestamp', 'submitted_at']);
  const title = getFieldByAliases(row, ['title', 'reason', 'subject', 'violation', 'report_type']);
  const details = getFieldByAliases(row, ['notes', 'description', 'summary', 'comments', 'details']);
  const typeHint = getFieldByAliases(row, ['type', 'category', 'status']) || fallbackType;
  return [officer, date, title, typeHint, details.slice(0, 180)].join('|');
}

function classifyReportType(tabName, reportsConfig) {
  const safe = sanitizeName(tabName);
  const discipline = (reportsConfig.disciplineTabNames || []).map(sanitizeName);
  const evaluation = (reportsConfig.evaluationTabNames || []).map(sanitizeName);
  if (discipline.includes(safe)) return 'discipline';
  if (evaluation.includes(safe)) return 'evaluation';
  if (safe.includes('disciplin') || safe.includes('violation')) return 'discipline';
  if (safe.includes('eval') || safe.includes('evaluation')) return 'evaluation';
  if (safe === 'officer_notes' || safe === 'internal_messages') return 'message';
  if (safe.includes('message') || safe.includes('note')) return 'message';
  return 'other';
}

function buildReportItems() {
  const reportsConfig = loadReportsConfig();
  const configuredTabs = Array.from(new Set(
    []
      .concat(reportsConfig.disciplineTabNames || [])
      .concat(reportsConfig.evaluationTabNames || [])
      .concat(['officer_notes', 'internal_messages'])
      .map(sanitizeName)
      .filter(Boolean)
  ));
  const discoveredTabs = listImportedTabNames().filter((tabName) => {
    const kind = classifyReportType(tabName, reportsConfig);
    return kind === 'discipline' || kind === 'evaluation' || kind === 'message';
  });
  const activeTabs = Array.from(new Set(configuredTabs.concat(discoveredTabs)));
  const approvals = loadReportApprovals();
  const approvalById = new Map(approvals.map(a => [String(a && a.id || ''), a]));

  const items = [];
  activeTabs.forEach((tabName) => {
    const rows = readTabRecords(tabName);
    const reportType = classifyReportType(tabName, reportsConfig);

    rows.forEach((row, idx) => {
      const fingerprint = buildReportFingerprint(row, reportType) || (tabName + '|' + idx);
      const id = stableHash(tabName + '|' + fingerprint);
      const approval = approvalById.get(id) || null;

      items.push({
        id,
        index: idx,
        sourceTab: tabName,
        type: reportType,
        subject: getFieldByAliases(row, ['title', 'subject', 'reason', 'violation', 'report_type']) || (reportType + ' report'),
        officerName: getFieldByAliases(row, ['officer_name', 'name', 'rp_name', 'cadet_name', 'character_name']),
        reportDate: getFieldByAliases(row, ['date', 'incident_date', 'created_date', 'timestamp', 'submitted_at']),
        detail: getFieldByAliases(row, ['notes', 'description', 'summary', 'comments', 'details']),
        approvalStatus: approval ? approval.status : 'pending',
        approvedBy: approval ? approval.approvedBy : '',
        approvedByEmail: approval ? approval.approvedByEmail : '',
        approvedAt: approval ? approval.approvedAt : '',
        rawRow: row
      });
    });
  });

  items.sort((a, b) => {
    const ad = String(a.reportDate || '');
    const bd = String(b.reportDate || '');
    if (ad && bd) return bd.localeCompare(ad);
    return a.sourceTab.localeCompare(b.sourceTab) || a.index - b.index;
  });

  return { items, reportsConfig, configuredTabs: activeTabs };
}

function normalizeEmail(email) {
  return String(email || '').trim().toLowerCase();
}

function formatUserDisplayName(user) {
  const rank = String((user && user.rank) || '').trim();
  const name = String((user && user.characterName) || '').trim();
  if (rank && name) return rank + ' ' + name;
  return name || rank || 'Officer';
}

function isPrivilegedRole(roleText) {
  const role = String(roleText || '').trim().toLowerCase();
  if (!role) return false;
  if (role.includes('admin')) return true;
  if (role.includes('chief')) return true;
  if (role.includes('commander')) return true;
  if (role.includes('supervisor')) return true;
  return false;
}

function canPromoteRole(roleText) {
  const role = String(roleText || '').trim().toLowerCase();
  if (!role) return false;
  if (role.includes('admin')) return true;
  if (role.includes('chief')) return true;
  if (role.includes('commander')) return true;
  return false;
}

function isPortalOwnerAuth(auth) {
  const email = normalizeEmail(auth && auth.email);
  return !!email && PORTAL_OWNER_EMAILS.includes(email);
}

function hasLeadershipAccess(auth) {
  const role = String((auth && auth.role) || '');
  return canPromoteRole(role) || isPortalOwnerAuth(auth);
}

function hasAdminAccess(auth) {
  const role = String((auth && auth.role) || '');
  return isPrivilegedRole(role) || isPortalOwnerAuth(auth);
}

function ensureHardsetRosterTab(tabs) {
  const out = Array.isArray(tabs) ? tabs.slice() : [];
  const rosterUrl = String(process.env.DEFAULT_ROSTER_URL || '').trim() || HARDSET_DEFAULT_ROSTER_URL;
  const hasRoster = out.some((t) => sanitizeName(t && t.name) === 'roster');
  if (!hasRoster && rosterUrl) out.unshift({ name: 'roster', url: rosterUrl });
  return out;
}

function normalizeAccessRole(roleText) {
  const role = String(roleText || '').trim().toLowerCase();
  if (role === 'admin') return 'admin';
  if (role === 'chief') return 'chief';
  if (role === 'commander') return 'commander';
  if (role === 'command') return 'command';
  return '';
}

function loadRoleOverrides() {
  const raw = loadJsonFile(ROLE_OVERRIDES_FILE, {});
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return {};
  const out = {};
  Object.keys(raw).forEach((k) => {
    const email = normalizeEmail(k);
    const role = String(raw[k] || '').trim().toLowerCase();
    if (!email || !role) return;
    out[email] = role;
  });
  return out;
}

function saveRoleOverrides(overrides) {
  const safe = {};
  Object.keys(overrides || {}).forEach((k) => {
    const email = normalizeEmail(k);
    const role = String(overrides[k] || '').trim().toLowerCase();
    if (!email || !role) return;
    safe[email] = role;
  });
  saveJsonFile(ROLE_OVERRIDES_FILE, safe);
  return safe;
}

async function dbUpsertUserRecord(user) {
  if (!DB_ENABLED || !user) return;
  const email = normalizeEmail(user.email);
  if (!email) return;
  await dbQuery(
    `INSERT INTO users (email, password_hash, created_at, password_updated_at, auto_provisioned)
     VALUES ($1, $2, $3, $4, $5)
     ON CONFLICT (email)
     DO UPDATE SET password_hash = EXCLUDED.password_hash,
                   created_at = COALESCE(users.created_at, EXCLUDED.created_at),
                   password_updated_at = EXCLUDED.password_updated_at,
                   auto_provisioned = EXCLUDED.auto_provisioned`,
    [
      email,
      String(user.passwordHash || ''),
      String(user.createdAt || ''),
      String(user.passwordUpdatedAt || ''),
      !!user.autoProvisioned
    ]
  );
}

async function dbInsertSessionRecord(token, email, createdAt) {
  if (!DB_ENABLED) return;
  await dbQuery(
    `INSERT INTO sessions (token, email, created_at)
     VALUES ($1, $2, $3)
     ON CONFLICT (token)
     DO UPDATE SET email = EXCLUDED.email, created_at = EXCLUDED.created_at`,
    [String(token || ''), normalizeEmail(email), String(createdAt || '')]
  );
}

async function dbDeleteSessionToken(token) {
  if (!DB_ENABLED) return;
  await dbQuery('DELETE FROM sessions WHERE token = $1', [String(token || '')]);
}

async function dbUpsertRoleOverride(email, role) {
  if (!DB_ENABLED) return;
  await dbQuery(
    `INSERT INTO role_overrides (email, role, updated_at)
     VALUES ($1, $2, $3)
     ON CONFLICT (email)
     DO UPDATE SET role = EXCLUDED.role, updated_at = EXCLUDED.updated_at`,
    [normalizeEmail(email), String(role || '').trim().toLowerCase(), new Date().toISOString()]
  );
}

async function dbDeleteRoleOverride(email) {
  if (!DB_ENABLED) return;
  await dbQuery('DELETE FROM role_overrides WHERE email = $1', [normalizeEmail(email)]);
}

function getEffectiveRole(email, baseRole) {
  const normalizedEmail = normalizeEmail(email);
  if (normalizedEmail && PORTAL_DEFAULT_ADMIN_EMAILS.includes(normalizedEmail)) {
    return 'admin';
  }
  const overrides = loadRoleOverrides();
  if (normalizedEmail && overrides[normalizedEmail]) return overrides[normalizedEmail];
  return String(baseRole || 'command').trim().toLowerCase() || 'command';
}

function hashPassword(password) {
  return crypto.createHash('sha256').update(String(password || '')).digest('hex');
}

function base64UrlEncode(input) {
  return Buffer.from(String(input || ''), 'utf8')
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
}

function base64UrlDecode(input) {
  const normalized = String(input || '').replace(/-/g, '+').replace(/_/g, '/');
  const padLen = normalized.length % 4;
  const padded = normalized + (padLen ? '='.repeat(4 - padLen) : '');
  return Buffer.from(padded, 'base64').toString('utf8');
}

function signTokenPayload(payloadText) {
  return crypto
    .createHmac('sha256', AUTH_SECRET)
    .update(String(payloadText || ''))
    .digest('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
}

function generateToken(email) {
  const normalizedEmail = normalizeEmail(email);
  if (!normalizedEmail) {
    return crypto.randomBytes(24).toString('hex');
  }
  const now = Math.floor(Date.now() / 1000);
  const payload = {
    v: 1,
    email: normalizedEmail,
    iat: now,
    exp: now + AUTH_TOKEN_TTL_SECONDS
  };
  const encoded = base64UrlEncode(JSON.stringify(payload));
  const sig = signTokenPayload(encoded);
  return 'v1.' + encoded + '.' + sig;
}

function parseSignedToken(token) {
  const text = String(token || '').trim();
  const m = text.match(/^v1\.([A-Za-z0-9_-]+)\.([A-Za-z0-9_-]+)$/);
  if (!m) return null;
  const encoded = m[1];
  const sig = m[2];
  if (signTokenPayload(encoded) !== sig) return null;
  try {
    const payload = JSON.parse(base64UrlDecode(encoded) || '{}');
    const email = normalizeEmail(payload && payload.email);
    const exp = Number(payload && payload.exp || 0);
    const now = Math.floor(Date.now() / 1000);
    if (!email || !exp || exp < now) return null;
    return { email, exp };
  } catch (e) {
    return null;
  }
}

function getCommandUsersRecords() {
  const filePath = path.join(DATA_DIR, 'command_users.json');
  if (!fs.existsSync(filePath)) return [];
  try {
    const data = JSON.parse(fs.readFileSync(filePath, 'utf8') || '[]');
    return Array.isArray(data) ? data : [];
  } catch (e) {
    return [];
  }
}

function pickField(obj, aliases) {
  const keys = Object.keys(obj || {});
  for (const alias of aliases) {
    const wanted = normalizeKey(alias);
    const found = keys.find(k => normalizeKey(k) === wanted);
    if (found) {
      const val = String(obj[found] || '').trim();
      if (val) return val;
    }
  }
  return '';
}

async function ensureCommandUsersLoaded() {
  let records = getCommandUsersRecords();
  if (records.length) return records;

  const config = loadSheetsConfig();
  const tabs = Array.isArray(config.tabs) ? config.tabs : [];
  const commandTab = tabs.find(t => sanitizeName(t && t.name) === 'command_users' && String(t && t.url || '').trim());

  const extractSheetId = (urlText) => {
    const m = String(urlText || '').match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return (m && m[1]) ? m[1] : '';
  };

  const deriveCommandUsersUrls = () => {
    const sourceUrls = [];
    tabs.forEach((t) => {
      const u = String(t && t.url || '').trim();
      if (u) sourceUrls.push(u);
    });
    const envRoster = String(process.env.DEFAULT_ROSTER_URL || '').trim();
    if (envRoster) sourceUrls.push(envRoster);

    const out = [];
    sourceUrls.forEach((u) => {
      const sheetId = extractSheetId(u);
      if (!sheetId) return;
      out.push('https://docs.google.com/spreadsheets/d/' + sheetId + '/gviz/tq?tqx=out:csv&sheet=Command_Users');
      out.push('https://docs.google.com/spreadsheets/d/' + sheetId + '/gviz/tq?tqx=out:csv&sheet=command_users');
      out.push('https://docs.google.com/spreadsheets/d/' + sheetId + '/gviz/tq?tqx=out:csv&sheet=Command Users');
    });

    // unique
    return Array.from(new Set(out));
  };

  const candidateUrls = commandTab
    ? [String(commandTab.url).trim()]
    : deriveCommandUsersUrls();

  for (const candidate of candidateUrls) {
    try {
      await importSheetsTabs([{ name: 'command_users', url: candidate }]);
      records = getCommandUsersRecords();
      if (records.length) return records;
    } catch (e) {
      // Try next candidate.
    }
  }

  return records;
}

function findCommandUserByEmailFromRecords(email, records) {
  const normalized = normalizeEmail(email);
  if (!normalized) return null;

  const isValidEmail = (value) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
  const localPart = (value) => String(value || '').trim().toLowerCase().split('@')[0] || '';
  const requestedLocal = localPart(normalized);

  const row = records.find((r) => {
    const rowEmailRaw = pickField(r, [
      'email',
      'email_address',
      'mail',
      'discord_email',
      'google_email',
      'google email',
      'googleemail'
    ]);
    const rowEmail = normalizeEmail(rowEmailRaw);
    if (!rowEmail) return false;

    // Primary: strict email match.
    if (rowEmail === normalized) return true;

    // Fallback: if row email is malformed/truncated, match by local-part.
    if (!isValidEmail(rowEmailRaw) && requestedLocal && localPart(rowEmailRaw) === requestedLocal) {
      return true;
    }

    return false;
  });
  if (!row) return null;

  return {
    email: normalized,
    characterName: pickField(row, [
      'character_name',
      'name_of_character',
      'name of character',
      'name_of_charac',
      'name of charac',
      'rp_name',
      'name',
      'officer_name'
    ]) || 'Officer',
    rank: pickField(row, ['rank', 'officer_rank']) || 'Unknown',
    role: pickField(row, ['role', 'access_role', 'permissions']) || 'command'
  };
}

async function findCommandUserByEmail(email) {
  const records = await ensureCommandUsersLoaded();
  return findCommandUserByEmailFromRecords(email, records);
}

function getAuthFromRequest(req) {
  const authHeader = String(req.headers.authorization || '');
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7).trim() : '';
  if (!token) return null;

  const sessions = loadJsonFile(SESSIONS_FILE, []);
  const session = sessions.find(s => s && s.token === token);
  const signed = parseSignedToken(token);
  const authEmail = normalizeEmail((session && session.email) || (signed && signed.email));
  if (!authEmail) return null;

  const commandProfile = findCommandUserByEmailFromRecords(authEmail, getCommandUsersRecords());
  if (!commandProfile) return null;

  return {
    token,
    email: authEmail,
    characterName: commandProfile.characterName,
    rank: commandProfile.rank,
    role: getEffectiveRole(authEmail, commandProfile.role),
    baseRole: String(commandProfile.role || 'command').trim().toLowerCase() || 'command'
  };
}

function requireAuth(req, res, next) {
  const auth = getAuthFromRequest(req);
  if (!auth) return res.status(401).json({ error: 'Unauthorized' });
  req.auth = auth;
  next();
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

function getRosterRecordId(record) {
  const candidates = [
    record && record.ID,
    record && record.id,
    record && record.Officer_ID,
    record && record.officer_id
  ];
  for (const value of candidates) {
    const text = String(value || '').trim();
    if (text) return text;
  }
  return '';
}

function loadFtoRecords() {
  return loadJsonFile(FTO_FILE, []);
}

function saveFtoRecords(records) {
  const safe = Array.isArray(records) ? records : [];
  saveJsonFile(FTO_FILE, safe);
  return safe;
}

function loadSheetsConfig() {
  const envTabs = getDefaultTabsFromEnv();

  ensureDataDir();
  if (!fs.existsSync(SHEETS_CONFIG_FILE)) {
    const tabs = ensureHardsetRosterTab(envTabs);
    return { tabs, autoSyncOnLoad: tabs.length > 0, lastSync: null };
  }
  try {
    const parsed = JSON.parse(fs.readFileSync(SHEETS_CONFIG_FILE, 'utf8') || '{}');
    const fileTabs = Array.isArray(parsed.tabs) ? parsed.tabs : [];
    const tabs = ensureHardsetRosterTab(fileTabs.length ? fileTabs : envTabs);
    return {
      tabs,
      autoSyncOnLoad: typeof parsed.autoSyncOnLoad === 'boolean' ? parsed.autoSyncOnLoad : tabs.length > 0,
      lastSync: parsed.lastSync || null
    };
  } catch (e) {
    const tabs = ensureHardsetRosterTab(envTabs);
    return { tabs, autoSyncOnLoad: tabs.length > 0, lastSync: null };
  }
}

function getDefaultTabsFromEnv() {
  const tabs = [];
  const rosterUrl = String(process.env.DEFAULT_ROSTER_URL || '').trim();
  if (rosterUrl) {
    tabs.push({ name: 'roster', url: rosterUrl });
  }

  const commandUsersUrl = String(process.env.DEFAULT_COMMAND_USERS_URL || '').trim();
  if (commandUsersUrl) {
    tabs.push({ name: 'command_users', url: commandUsersUrl });
  }

  const rawExtra = String(process.env.DEFAULT_SHEETS_TABS || '').trim();
  if (rawExtra) {
    try {
      const parsed = JSON.parse(rawExtra);
      if (Array.isArray(parsed)) {
        parsed.forEach((t) => {
          const name = sanitizeName(t && t.name);
          const url = String(t && t.url || '').trim();
          if (name && url) tabs.push({ name, url });
        });
      }
    } catch (e) {
      // Ignore invalid env JSON.
    }
  }

  // Keep unique by tab name, first wins.
  const seen = new Set();
  return tabs.filter((t) => {
    if (seen.has(t.name)) return false;
    seen.add(t.name);
    return true;
  });
}

function saveSheetsConfig(config) {
  ensureDataDir();
  const safeTabs = Array.isArray(config && config.tabs)
    ? config.tabs
      .map(t => ({ name: sanitizeName(t && t.name), url: String(t && t.url || '').trim() }))
      .filter(t => t.name && t.url)
    : [];

  const payload = {
    tabs: ensureHardsetRosterTab(safeTabs),
    autoSyncOnLoad: !!(config && config.autoSyncOnLoad),
    lastSync: (config && config.lastSync) || null
  };

  fs.writeFileSync(SHEETS_CONFIG_FILE, JSON.stringify(payload, null, 2));
  return payload;
}

async function importSheetsTabs(tabs) {
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

  const okCount = result.filter(x => x.ok).length;
  const failedCount = result.length - okCount;
  return {
    ok: failedCount === 0,
    result,
    summary: { total: result.length, ok: okCount, failed: failedCount }
  };
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
    const notes = pick('notes', 'note', 'comments', 'remarks', 'officer_notes');

    // Recover from imports where a text name landed in Callsign.
    if (!name && callsign && /[a-z]/i.test(callsign) && !/\d/.test(callsign)) {
      name = callsign;
      callsign = '';
    }

    const hasIdentity = [id, name, callsign].some(isMeaningful);
    if (!hasIdentity) return null;

    const importedFields = {};
    const originalKeys = Object.keys(r || {});
    originalKeys.forEach((key) => {
      importedFields[key] = cleanField(r[key]);
    });

    const columnsKT = {};
    for (let i = 10; i <= 19; i++) {
      const fallbackName = 'Column ' + String.fromCharCode(65 + i);
      const keyAtIndex = originalKeys[i] || fallbackName;
      columnsKT[keyAtIndex] = cleanField(r[keyAtIndex]);
    }

    const columnsNT = {};
    for (let i = 13; i <= 19; i++) {
      const fallbackName = 'Column ' + String.fromCharCode(65 + i);
      const keyAtIndex = originalKeys[i] || fallbackName;
      columnsNT[keyAtIndex] = cleanField(r[keyAtIndex]);
    }

    return {
      ID: id || ('IMP-' + Date.now() + '-' + idx),
      Name: name,
      Callsign: callsign,
      Rank: rank,
      Division: division,
      Notes: notes,
      ImportedFields: importedFields,
      ColumnsKT: columnsKT,
      ColumnsNT: columnsNT
    };
  }).filter(Boolean);
}

app.post('/api/auth/create-account', async (req, res) => {
  try {
    const email = normalizeEmail(req.body && req.body.email);
    const password = String(req.body && req.body.password || '');

    if (!email || !password) {
      return res.status(400).json({ error: 'Email and password are required.' });
    }
    if (password.length < 6) {
      return res.status(400).json({ error: 'Password must be at least 6 characters.' });
    }

    const commandProfile = await findCommandUserByEmail(email);
    if (!commandProfile) {
      return res.status(403).json({ error: 'Email not found in Command_Users tab. Verify headers (Google Email, Name of Character/Name of Charac, Rank) and re-sync sheets.' });
    }

    const users = loadJsonFile(USERS_FILE, []);
    if (users.some(u => normalizeEmail(u.email) === email)) {
      return res.status(409).json({ error: 'Account already exists. Please log in.' });
    }

    const createdAt = new Date().toISOString();
    const createdUser = {
      email,
      passwordHash: hashPassword(password),
      createdAt
    };
    users.push(createdUser);
    saveJsonFile(USERS_FILE, users);
    await dbUpsertUserRecord(createdUser);

    const sessions = loadJsonFile(SESSIONS_FILE, []);
    const token = generateToken(email);
    const sessionCreatedAt = new Date().toISOString();
    sessions.push({ token, email, createdAt: sessionCreatedAt });
    saveJsonFile(SESSIONS_FILE, sessions);
    await dbInsertSessionRecord(token, email, sessionCreatedAt);

    return res.status(201).json({
      ok: true,
      token,
      user: commandProfile,
      message: 'Welcome ' + commandProfile.rank + ' ' + commandProfile.characterName + '.'
    });
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/auth/ensure-command-users', async (req, res) => {
  try {
    const beforeCount = getCommandUsersRecords().length;
    const records = await ensureCommandUsersLoaded();
    const afterCount = Array.isArray(records) ? records.length : 0;
    return res.json({
      ok: afterCount > 0,
      beforeCount,
      afterCount,
      message: afterCount > 0 ? 'Command_Users is available.' : 'Command_Users could not be loaded yet.'
    });
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/auth/login', async (req, res) => {
  try {
    const email = normalizeEmail(req.body && req.body.email);
    const password = String(req.body && req.body.password || '');

    if (!email || !password) {
      return res.status(400).json({ error: 'Email and password are required.' });
    }

    const users = loadJsonFile(USERS_FILE, []);
    const commandProfile = await findCommandUserByEmail(email);
    if (!commandProfile) {
      return res.status(403).json({ error: 'Email is no longer authorized in Command_Users tab.' });
    }

    let user = users.find(u => normalizeEmail(u.email) === email);
    if (!user) {
      // Auto-provision account on login for authorized Command_Users entries.
      user = {
        email,
        passwordHash: hashPassword(password),
        createdAt: new Date().toISOString(),
        autoProvisioned: true
      };
      users.push(user);
      saveJsonFile(USERS_FILE, users);
      await dbUpsertUserRecord(user);
    } else if (String(user.passwordHash || '') !== hashPassword(password)) {
      return res.status(401).json({ error: 'Invalid email or password.' });
    }

    const sessions = loadJsonFile(SESSIONS_FILE, []);
    const token = generateToken(email);
    const sessionCreatedAt = new Date().toISOString();
    sessions.push({ token, email, createdAt: sessionCreatedAt });
    saveJsonFile(SESSIONS_FILE, sessions);
    await dbInsertSessionRecord(token, email, sessionCreatedAt);

    return res.json({
      ok: true,
      token,
      user: commandProfile,
      message: 'Welcome back ' + commandProfile.rank + ' ' + commandProfile.characterName + '.'
    });
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/auth/link-command-users', async (req, res) => {
  try {
    const url = String(req.body && req.body.url || '').trim();
    if (!url) return res.status(400).json({ error: 'Command_Users URL is required.' });

    const current = loadSheetsConfig();
    const tabs = Array.isArray(current.tabs) ? current.tabs.slice() : [];
    const existingIdx = tabs.findIndex(t => sanitizeName(t && t.name) === 'command_users');
    if (existingIdx >= 0) {
      tabs[existingIdx] = { name: 'command_users', url };
    } else {
      tabs.push({ name: 'command_users', url });
    }

    const updated = saveSheetsConfig({
      tabs,
      autoSyncOnLoad: current.autoSyncOnLoad,
      lastSync: current.lastSync
    });

    const imported = await importSheetsTabs([{ name: 'command_users', url }]);
    return res.json({ ok: imported.ok, linked: true, config: updated, import: imported });
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

app.get('/api/auth/me', (req, res) => {
  const auth = getAuthFromRequest(req);
  if (!auth) return res.status(401).json({ error: 'Unauthorized' });
  res.json({ ok: true, user: auth });
});

app.post('/api/auth/logout', async (req, res) => {
  try {
    const authHeader = String(req.headers.authorization || '');
    const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7).trim() : '';
    if (!token) return res.json({ ok: true });

    const sessions = loadJsonFile(SESSIONS_FILE, []);
    const filtered = sessions.filter(s => s && s.token !== token);
    saveJsonFile(SESSIONS_FILE, filtered);
    await dbDeleteSessionToken(token);
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/auth/change-password', requireAuth, (req, res) => {
  try {
    const oldPassword = String(req.body && req.body.oldPassword || '');
    const newPassword = String(req.body && req.body.newPassword || '');
    const confirmPassword = String(req.body && req.body.confirmPassword || '');

    if (!oldPassword || !newPassword || !confirmPassword) {
      return res.status(400).json({ error: 'Old password, new password, and confirm password are required.' });
    }
    if (newPassword !== confirmPassword) {
      return res.status(400).json({ error: 'New password and confirm password do not match.' });
    }
    if (newPassword.length < 6) {
      return res.status(400).json({ error: 'New password must be at least 6 characters.' });
    }

    const users = loadJsonFile(USERS_FILE, []);
    const idx = users.findIndex(u => normalizeEmail(u && u.email) === normalizeEmail(req.auth.email));
    if (idx < 0) return res.status(404).json({ error: 'User account not found.' });

    const currentHash = hashPassword(oldPassword);
    if (String(users[idx].passwordHash || '') !== currentHash) {
      return res.status(401).json({ error: 'Old password is incorrect.' });
    }

    users[idx].passwordHash = hashPassword(newPassword);
    users[idx].passwordUpdatedAt = new Date().toISOString();
    saveJsonFile(USERS_FILE, users);

    return res.json({ ok: true, message: 'Password updated successfully.' });
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/auth/admin-reset-password', requireAuth, (req, res) => {
  try {
    if (!hasAdminAccess(req.auth)) {
      return res.status(403).json({ error: 'Admin reset requires elevated role.' });
    }

    const targetEmail = normalizeEmail(req.body && req.body.email);
    const newPassword = String(req.body && req.body.newPassword || '');
    if (!targetEmail || !newPassword) {
      return res.status(400).json({ error: 'Target email and new password are required.' });
    }
    if (newPassword.length < 6) {
      return res.status(400).json({ error: 'New password must be at least 6 characters.' });
    }

    const users = loadJsonFile(USERS_FILE, []);
    const idx = users.findIndex(u => normalizeEmail(u && u.email) === targetEmail);
    if (idx < 0) {
      return res.status(404).json({ error: 'Target user account not found.' });
    }

    users[idx].passwordHash = hashPassword(newPassword);
    users[idx].passwordUpdatedAt = new Date().toISOString();
    users[idx].passwordResetBy = req.auth.email;
    saveJsonFile(USERS_FILE, users);

    return res.json({ ok: true, message: 'Password reset for ' + targetEmail + '.' });
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

app.get('/api/admin/users', requireAuth, (req, res) => {
  try {
    if (!hasAdminAccess(req.auth)) {
      return res.status(403).json({ error: 'Admin access required.' });
    }

    const users = loadJsonFile(USERS_FILE, []);
    const commandUsers = getCommandUsersRecords();
    const accountEmails = new Set(users.map((u) => normalizeEmail(u && u.email)).filter(Boolean));
    const overrides = loadRoleOverrides();

    const mapped = commandUsers.map((row) => {
      const email = normalizeEmail(pickField(row, [
        'email', 'email_address', 'mail', 'google_email', 'google email', 'googleemail'
      ]));
      if (!email) return null;
      const characterName = pickField(row, ['character_name', 'name_of_character', 'name of character', 'name_of_charac', 'name of charac', 'rp_name', 'name']);
      const rank = pickField(row, ['rank', 'officer_rank']);
      const baseRole = String(pickField(row, ['role', 'access_role', 'permissions']) || 'command').trim().toLowerCase() || 'command';
      const role = getEffectiveRole(email, baseRole);
      return {
        email,
        displayName: (rank && characterName) ? (rank + ' ' + characterName) : (characterName || email),
        hasAccount: accountEmails.has(email),
        baseRole,
        overrideRole: overrides[email] || '',
        role,
        isAdmin: isPrivilegedRole(role)
      };
    }).filter(Boolean);

    mapped.sort((a, b) => a.displayName.localeCompare(b.displayName));
    res.json({ ok: true, users: mapped });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/admin/set-admin', requireAuth, async (req, res) => {
  try {
    if (!hasAdminAccess(req.auth)) {
      return res.status(403).json({ error: 'Admin access required.' });
    }

    const email = normalizeEmail(req.body && req.body.email);
    const isAdmin = !!(req.body && req.body.isAdmin);
    if (!email) return res.status(400).json({ error: 'Target email is required.' });

    const commandProfile = findCommandUserByEmailFromRecords(email, getCommandUsersRecords());
    if (!commandProfile) {
      return res.status(404).json({ error: 'Target user is not in Command_Users.' });
    }

    const overrides = loadRoleOverrides();
    if (isAdmin) overrides[email] = 'admin';
    else delete overrides[email];
    saveRoleOverrides(overrides);
    if (isAdmin) await dbUpsertRoleOverride(email, 'admin');
    else await dbDeleteRoleOverride(email);

    const role = getEffectiveRole(email, commandProfile.role);
    res.json({ ok: true, email, role, isAdmin: isPrivilegedRole(role) });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/admin/set-role', requireAuth, async (req, res) => {
  try {
    if (!hasLeadershipAccess(req.auth)) {
      return res.status(403).json({ error: 'Role update requires chief, commander, or admin access.' });
    }

    const email = normalizeEmail(req.body && req.body.email);
    const role = normalizeAccessRole(req.body && req.body.role);
    if (!email) return res.status(400).json({ error: 'Target email is required.' });
    if (!role) return res.status(400).json({ error: 'Valid role is required (command, commander, chief, admin).' });

    const commandProfile = findCommandUserByEmailFromRecords(email, getCommandUsersRecords());
    if (!commandProfile) {
      return res.status(404).json({ error: 'Target user is not in Command_Users.' });
    }

    const overrides = loadRoleOverrides();
    overrides[email] = role;
    saveRoleOverrides(overrides);
    await dbUpsertRoleOverride(email, role);

    return res.json({ ok: true, email, role, effectiveRole: getEffectiveRole(email, commandProfile.role) });
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

app.get('/api/messages/users', requireAuth, (req, res) => {
  try {
    const users = loadJsonFile(USERS_FILE, []);
    const commandUsers = getCommandUsersRecords();

    const fromCommand = commandUsers.map((row) => {
      const email = normalizeEmail(pickField(row, [
        'email', 'email_address', 'mail', 'google_email', 'google email', 'googleemail'
      ]));
      if (!email) return null;
      const characterName = pickField(row, ['character_name', 'name_of_character', 'name of character', 'name_of_charac', 'name of charac', 'rp_name', 'name']);
      const rank = pickField(row, ['rank', 'officer_rank']);
      return {
        email,
        displayName: (rank && characterName) ? (rank + ' ' + characterName) : (characterName || email)
      };
    }).filter(Boolean);

    const fromUsers = users.map((u) => {
      const email = normalizeEmail(u && u.email);
      if (!email) return null;
      const profile = findCommandUserByEmailFromRecords(email, commandUsers);
      const name = profile ? formatUserDisplayName(profile) : email;
      return { email, displayName: name };
    }).filter(Boolean);

    const merged = Array.from(new Map(fromCommand.concat(fromUsers).map(x => [x.email, x])).values())
      .filter(x => x.email !== normalizeEmail(req.auth.email))
      .sort((a, b) => a.displayName.localeCompare(b.displayName));

    res.json({ ok: true, users: merged });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.get('/api/messages/unread-count', requireAuth, (req, res) => {
  try {
    const email = normalizeEmail(req.auth.email);
    const all = loadInternalMessages();
    const unread = all.filter(m => normalizeEmail(m && m.toEmail) === email && !m.readAt).length;
    res.json({ ok: true, unread });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.get('/api/messages/inbox', requireAuth, (req, res) => {
  try {
    const email = normalizeEmail(req.auth.email);
    const unreadOnly = String(req.query.unreadOnly || '').trim().toLowerCase() === 'true';
    const limit = Math.max(1, Math.min(250, Number(req.query.limit || 100)));
    let rows = loadInternalMessages().filter(m => normalizeEmail(m && m.toEmail) === email);
    if (unreadOnly) rows = rows.filter(m => !m.readAt);
    rows.sort((a, b) => String(b.createdAt || '').localeCompare(String(a.createdAt || '')));
    res.json({ ok: true, items: rows.slice(0, limit) });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.get('/api/messages/sent', requireAuth, (req, res) => {
  try {
    const email = normalizeEmail(req.auth.email);
    const limit = Math.max(1, Math.min(250, Number(req.query.limit || 100)));
    const rows = loadInternalMessages()
      .filter(m => normalizeEmail(m && m.fromEmail) === email)
      .sort((a, b) => String(b.createdAt || '').localeCompare(String(a.createdAt || '')))
      .slice(0, limit);
    res.json({ ok: true, items: rows });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/messages/send', requireAuth, (req, res) => {
  try {
    const rawRecipients = Array.isArray(req.body && req.body.toEmails)
      ? req.body.toEmails
      : [req.body && req.body.toEmail];
    const recipients = Array.from(new Set(rawRecipients.map(normalizeEmail).filter(Boolean)));
    const subject = String(req.body && req.body.subject || '').trim();
    const body = String(req.body && req.body.body || '').trim();

    if (!recipients.length) return res.status(400).json({ error: 'At least one recipient email is required.' });
    if (!subject && !body) return res.status(400).json({ error: 'Message subject or body is required.' });

    const fromEmail = normalizeEmail(req.auth.email);
    const directory = getCommandUsersRecords();
    const senderProfile = findCommandUserByEmailFromRecords(fromEmail, directory) || req.auth;
    const createdAt = new Date().toISOString();
    const createdItems = [];
    const invalidRecipients = [];

    recipients.forEach((toEmail) => {
      if (toEmail === fromEmail) {
        invalidRecipients.push({ email: toEmail, reason: 'Cannot send to yourself.' });
        return;
      }
      const recipientProfile = findCommandUserByEmailFromRecords(toEmail, directory);
      if (!recipientProfile) {
        invalidRecipients.push({ email: toEmail, reason: 'Not in Command_Users.' });
        return;
      }
      createdItems.push({
        id: 'msg_' + Date.now() + '_' + crypto.randomBytes(5).toString('hex'),
        fromEmail,
        fromName: formatUserDisplayName(senderProfile),
        toEmail,
        toName: formatUserDisplayName(recipientProfile),
        subject: subject || '(No Subject)',
        body,
        createdAt,
        readAt: ''
      });
    });

    if (!createdItems.length) {
      return res.status(400).json({
        error: 'No valid recipients found.',
        invalidRecipients
      });
    }

    const all = loadInternalMessages();
    createdItems.forEach((item) => all.push(item));
    saveInternalMessages(all);
    res.status(201).json({
      ok: true,
      count: createdItems.length,
      items: createdItems,
      invalidRecipients
    });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/messages/:id/read', requireAuth, (req, res) => {
  try {
    const id = String(req.params.id || '').trim();
    const markRead = !!(req.body && req.body.read);
    const email = normalizeEmail(req.auth.email);
    const all = loadInternalMessages();
    const idx = all.findIndex(m => String(m && m.id || '') === id && normalizeEmail(m && m.toEmail) === email);
    if (idx < 0) return res.status(404).json({ error: 'Message not found in your inbox.' });

    all[idx].readAt = markRead ? new Date().toISOString() : '';
    saveInternalMessages(all);
    res.json({ ok: true, item: all[idx] });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

// API: list roster
app.get('/api/roster', (req,res)=>{
  const data = loadJson();
  res.json(data);
});

app.get('/api/fto', requireAuth, (req, res) => {
  try {
    if (!hasLeadershipAccess(req.auth)) {
      return res.status(403).json({ error: 'FTO access requires chief, commander, or admin role.' });
    }
    const items = loadFtoRecords().sort((a, b) => String(a.name || '').localeCompare(String(b.name || '')));
    return res.json({ ok: true, items });
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/fto', requireAuth, (req, res) => {
  try {
    if (!hasLeadershipAccess(req.auth)) {
      return res.status(403).json({ error: 'FTO updates require chief, commander, or admin role.' });
    }

    const officerId = String(req.body && req.body.officerId || '').trim();
    if (!officerId) return res.status(400).json({ error: 'Officer ID is required.' });

    const roster = loadJson();
    const officer = roster.find((x) => getRosterRecordId(x) === officerId);
    if (!officer) return res.status(404).json({ error: 'Officer not found in roster.' });

    const list = loadFtoRecords();
    if (list.some((x) => String(x && x.officerId || '').trim() === officerId)) {
      return res.status(409).json({ error: 'Officer is already in FTO list.' });
    }

    const item = {
      officerId,
      name: String(officer.Name || officer.name || '').trim(),
      callsign: String(officer.Callsign || officer.callsign || '').trim(),
      rank: String(officer.Rank || officer.rank || '').trim(),
      division: String(officer.Division || officer.division || '').trim(),
      addedAt: new Date().toISOString(),
      addedBy: formatUserDisplayName(req.auth),
      addedByEmail: normalizeEmail(req.auth.email)
    };
    list.push(item);
    saveFtoRecords(list);

    // Mirror FTO status onto roster so officer profile can display a badge.
    const rosterIdx = roster.findIndex((x) => getRosterRecordId(x) === officerId);
    if (rosterIdx >= 0) {
      roster[rosterIdx].IsFTO = 'Yes';
      roster[rosterIdx].FTOAddedAt = item.addedAt;
      roster[rosterIdx].FTOAddedBy = item.addedBy;
      saveJson(roster);
    }

    return res.status(201).json({ ok: true, item, message: 'Officer added to FTO list.' });
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

app.delete('/api/fto/:officerId', requireAuth, (req, res) => {
  try {
    if (!hasLeadershipAccess(req.auth)) {
      return res.status(403).json({ error: 'FTO updates require chief, commander, or admin role.' });
    }

    const officerId = String(req.params.officerId || '').trim();
    if (!officerId) return res.status(400).json({ error: 'Officer ID is required.' });

    const list = loadFtoRecords();
    const before = list.length;
    const after = list.filter((x) => String(x && x.officerId || '').trim() !== officerId);
    if (after.length === before) return res.status(404).json({ error: 'Officer is not in FTO list.' });
    saveFtoRecords(after);

    // Clear mirrored roster badge fields on remove.
    const roster = loadJson();
    const rosterIdx = roster.findIndex((x) => getRosterRecordId(x) === officerId);
    if (rosterIdx >= 0) {
      roster[rosterIdx].IsFTO = '';
      roster[rosterIdx].FTOAddedAt = '';
      roster[rosterIdx].FTOAddedBy = '';
      saveJson(roster);
    }

    return res.json({ ok: true, deleted: before - after.length, message: 'Officer removed from FTO list.' });
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

// API: add roster item
app.post('/api/roster', requireAuth, (req,res)=>{
  const data = loadJson();
  const item = req.body || {};
  // Ensure an ID
  item.ID = getRosterRecordId(item) || String(Date.now());
  data.push(item);
  saveJson(data);
  res.status(201).json(item);
});

// API: update by ID
app.put('/api/roster/:id', requireAuth, (req,res)=>{
  const id = req.params.id;
  const data = loadJson();
  const idx = data.findIndex(x=>getRosterRecordId(x) === String(id));
  if (idx === -1) return res.status(404).json({error:'Not found'});
  data[idx] = Object.assign({}, data[idx], req.body);
  if (!getRosterRecordId(data[idx])) data[idx].ID = String(id);
  saveJson(data);
  res.json(data[idx]);
});

// API: promotion update by ID (chief/commander/admin only)
app.post('/api/roster/:id/promote', requireAuth, (req, res) => {
  try {
    if (!hasLeadershipAccess(req.auth)) {
      return res.status(403).json({ error: 'Promotion access requires chief, commander, or admin role.' });
    }

    const id = String(req.params.id || '').trim();
    const rank = String(req.body && req.body.rank || '').trim();
    const division = String(req.body && req.body.division || '').trim();
    const reason = String(req.body && req.body.reason || '').trim();

    if (!id) return res.status(400).json({ error: 'Officer ID is required.' });
    if (!rank && !division) return res.status(400).json({ error: 'Provide rank and/or division to update.' });

    const data = loadJson();
    const idx = data.findIndex(x => getRosterRecordId(x) === id);
    if (idx === -1) return res.status(404).json({ error: 'Officer not found.' });

    const beforeRank = String(data[idx].Rank || data[idx].rank || '').trim();
    const beforeDivision = String(data[idx].Division || data[idx].division || '').trim();
    if (rank) data[idx].Rank = rank;
    if (division) data[idx].Division = division;

    const now = new Date().toISOString();
    data[idx].LastPromotedAt = now;
    data[idx].LastPromotedBy = formatUserDisplayName(req.auth);
    data[idx].LastPromotedByEmail = normalizeEmail(req.auth.email);
    data[idx].LastPromotionReason = reason;

    const history = Array.isArray(data[idx].PromotionHistory) ? data[idx].PromotionHistory : [];
    history.push({
      at: now,
      by: formatUserDisplayName(req.auth),
      byEmail: normalizeEmail(req.auth.email),
      from: { rank: beforeRank, division: beforeDivision },
      to: { rank: String(data[idx].Rank || '').trim(), division: String(data[idx].Division || '').trim() },
      reason
    });
    data[idx].PromotionHistory = history.slice(-50);

    saveJson(data);
    return res.json({ ok: true, item: data[idx], message: 'Promotion updated successfully.' });
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

// API: delete by ID
app.delete('/api/roster/:id', requireAuth, (req,res)=>{
  const id = req.params.id;
  let data = loadJson();
  const before = data.length;
  data = data.filter(x=>getRosterRecordId(x) !== String(id));
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

    const imported = await importSheetsTabs(tabs);
    const config = loadSheetsConfig();
    config.lastSync = {
      at: new Date().toISOString(),
      ok: imported.ok,
      summary: imported.summary,
      result: imported.result
    };
    saveSheetsConfig(config);

    return res.json(imported);
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

app.get('/api/sheets/config', (req, res) => {
  const config = loadSheetsConfig();
  res.json(config);
});

app.put('/api/sheets/config', (req, res) => {
  try {
    const current = loadSheetsConfig();
    const incoming = req.body || {};
    const updated = saveSheetsConfig({
      tabs: Array.isArray(incoming.tabs) ? incoming.tabs : current.tabs,
      autoSyncOnLoad: typeof incoming.autoSyncOnLoad === 'boolean' ? incoming.autoSyncOnLoad : current.autoSyncOnLoad,
      lastSync: current.lastSync
    });
    res.json(updated);
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/sheets/sync', async (req, res) => {
  try {
    const config = loadSheetsConfig();
    const tabs = Array.isArray(req.body && req.body.tabs) && req.body.tabs.length
      ? req.body.tabs
      : config.tabs;

    if (!tabs.length) {
      return res.status(400).json({ error: 'No saved tabs to sync. Save a roster link first.' });
    }

    const imported = await importSheetsTabs(tabs);
    config.lastSync = {
      at: new Date().toISOString(),
      ok: imported.ok,
      summary: imported.summary,
      result: imported.result
    };
    saveSheetsConfig(config);
    return res.json(imported);
  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
});

app.get('/api/sheets/tabs', (req, res) => {
  ensureDataDir();
  const files = fs.readdirSync(DATA_DIR)
    .filter(f => f.endsWith('.json'))
    .filter(f => !INTERNAL_DATA_FILES.has(f))
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

app.get('/api/reports/config', requireAuth, (req, res) => {
  const config = loadReportsConfig();
  res.json(config);
});

app.put('/api/reports/config', requireAuth, (req, res) => {
  try {
    const updated = saveReportsConfig(req.body || {});
    res.json(updated);
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/reports/link-disciplinary-source', requireAuth, async (req, res) => {
  try {
    const url = String(req.body && req.body.url || '').trim();
    const tabName = sanitizeName((req.body && req.body.tabName) || 'disciplinary_forms');
    if (!url) return res.status(400).json({ error: 'Disciplinary source URL is required.' });

    const sheetsConfig = loadSheetsConfig();
    const tabs = Array.isArray(sheetsConfig.tabs) ? sheetsConfig.tabs.slice() : [];
    const existingIdx = tabs.findIndex(t => sanitizeName(t && t.name) === tabName);
    if (existingIdx >= 0) tabs[existingIdx] = { name: tabName, url };
    else tabs.push({ name: tabName, url });

    saveSheetsConfig({
      tabs,
      autoSyncOnLoad: sheetsConfig.autoSyncOnLoad,
      lastSync: sheetsConfig.lastSync
    });

    const reportsConfig = loadReportsConfig();
    const updatedReports = saveReportsConfig({
      disciplineTabNames: Array.from(new Set([].concat(reportsConfig.disciplineTabNames || [], [tabName]))),
      evaluationTabNames: reportsConfig.evaluationTabNames || ['cadet_evaluations']
    });

    const imported = await importSheetsTabs([{ name: tabName, url }]);
    res.json({ ok: imported.ok, tabName, import: imported, reportsConfig: updatedReports });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/reports/link-evaluation-source', requireAuth, async (req, res) => {
  try {
    const url = String(req.body && req.body.url || '').trim();
    const tabName = sanitizeName((req.body && req.body.tabName) || 'cadet_evaluations');
    if (!url) return res.status(400).json({ error: 'Cadet evaluations source URL is required.' });

    const sheetsConfig = loadSheetsConfig();
    const tabs = Array.isArray(sheetsConfig.tabs) ? sheetsConfig.tabs.slice() : [];
    const existingIdx = tabs.findIndex(t => sanitizeName(t && t.name) === tabName);
    if (existingIdx >= 0) tabs[existingIdx] = { name: tabName, url };
    else tabs.push({ name: tabName, url });

    saveSheetsConfig({
      tabs,
      autoSyncOnLoad: sheetsConfig.autoSyncOnLoad,
      lastSync: sheetsConfig.lastSync
    });

    const reportsConfig = loadReportsConfig();
    const updatedReports = saveReportsConfig({
      disciplineTabNames: reportsConfig.disciplineTabNames || ['discipline_records', 'disciplinary_forms'],
      evaluationTabNames: Array.from(new Set([].concat(reportsConfig.evaluationTabNames || [], [tabName])))
    });

    const imported = await importSheetsTabs([{ name: tabName, url }]);
    res.json({ ok: imported.ok, tabName, import: imported, reportsConfig: updatedReports });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.get('/api/reports/items', requireAuth, (req, res) => {
  try {
    const reportType = sanitizeName(req.query.type || 'all');
    const approvalStatus = sanitizeName(req.query.status || 'all');
    const built = buildReportItems();
    let items = built.items;

    if (reportType !== 'all') {
      items = items.filter(i => sanitizeName(i.type) === reportType);
    }
    if (approvalStatus !== 'all') {
      items = items.filter(i => sanitizeName(i.approvalStatus) === approvalStatus);
    }

    res.json({
      ok: true,
      items,
      configuredTabs: built.configuredTabs,
      reportsConfig: built.reportsConfig
    });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.get('/api/reports/summary', requireAuth, (req, res) => {
  try {
    const built = buildReportItems();
    const summarize = (type) => {
      const rows = built.items.filter(i => i.type === type);
      return {
        total: rows.length,
        pending: rows.filter(r => r.approvalStatus === 'pending').length,
        approved: rows.filter(r => r.approvalStatus === 'approved').length,
        denied: rows.filter(r => r.approvalStatus === 'denied').length
      };
    };

    const summary = {
      discipline: summarize('discipline'),
      evaluation: summarize('evaluation'),
      message: summarize('message'),
      other: summarize('other')
    };

    res.json({
      ok: true,
      summary,
      configuredTabs: built.configuredTabs,
      reportsConfig: built.reportsConfig
    });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.post('/api/reports/:id/approval', requireAuth, (req, res) => {
  try {
    const id = String(req.params.id || '').trim();
    const status = sanitizeName(req.body && req.body.status);
    if (!id) return res.status(400).json({ error: 'Report id is required.' });
    if (!['approved', 'denied', 'pending'].includes(status)) {
      return res.status(400).json({ error: 'Status must be approved, denied, or pending.' });
    }

    const built = buildReportItems();
    const item = built.items.find(x => x.id === id);
    if (!item) return res.status(404).json({ error: 'Report item not found.' });
    if (!['discipline', 'evaluation'].includes(item.type)) {
      return res.status(400).json({ error: 'Only discipline and evaluation reports require command approval.' });
    }

    const approvals = loadReportApprovals();
    const idx = approvals.findIndex(a => String(a && a.id || '') === id);
    const actorName = formatUserDisplayName(req.auth);
    const approvedAt = new Date().toISOString();
    const next = {
      id,
      type: item.type,
      sourceTab: item.sourceTab,
      status,
      approvedBy: actorName,
      approvedByEmail: req.auth.email,
      approvedAt,
      updatedAt: approvedAt
    };

    if (idx >= 0) {
      approvals[idx] = Object.assign({}, approvals[idx], next);
    } else {
      approvals.push(Object.assign({}, next, { createdAt: approvedAt }));
    }
    saveReportApprovals(approvals);

    res.json({ ok: true, approval: next });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.get('/api/health', (req, res) => {
  res.json({
    ok: true,
    service: 'FWPD Command Portal',
    time: new Date().toISOString(),
    bootedAt: BOOTED_AT,
    uptimeSeconds: Math.floor(process.uptime())
  });
});

async function startServer() {
  try {
    await initDatabasePersistence();
  } catch (err) {
    console.error('DB initialization failed. Continuing with file-based persistence only.', err.message || String(err));
  }

  app.listen(PORT, () => {
    console.log('Server running on http://localhost:' + PORT);
  });
}

startServer();
