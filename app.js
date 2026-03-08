// Simple local API calls to /api/roster endpoints

const AUTO_SYNC_SESSION_KEY = 'fwpd_auto_sync_done';
const LOCAL_SYNC_TABS_KEY = 'fwpd_sync_tabs_v1';
const AUTH_TOKEN_KEY = 'fwpd_auth_token';
const APP_BUILD = '20260308t';

let currentUser = null;

function formatUserDisplayName(user) {
  const rank = String((user && user.rank) || '').trim();
  const name = String((user && user.characterName) || '').trim();
  if (rank && name) return rank + ' ' + name;
  return name || rank || 'Officer';
}

function getAuthToken() {
  return localStorage.getItem(AUTH_TOKEN_KEY) || '';
}

function setAuthToken(token) {
  if (token) localStorage.setItem(AUTH_TOKEN_KEY, token);
  else localStorage.removeItem(AUTH_TOKEN_KEY);
}

function authHeaders(extra = {}) {
  const token = getAuthToken();
  const headers = Object.assign({}, extra);
  if (token) headers.Authorization = 'Bearer ' + token;
  return headers;
}

async function authFetch(url, options = {}) {
  const headers = authHeaders(options.headers || {});
  return fetch(url, Object.assign({}, options, { headers }));
}

function isLoggedIn() {
  return !!currentUser;
}

function showAuthBanner() {
  const existing = document.getElementById('authBanner');
  if (existing) existing.remove();

  const title = document.querySelector('.title');
  if (!title) return;

  const banner = document.createElement('div');
  banner.id = 'authBanner';
  banner.style.fontSize = '12px';
  banner.style.letterSpacing = '0';
  banner.style.marginTop = '6px';
  banner.textContent = currentUser
    ? ('Logged in as ' + formatUserDisplayName(currentUser))
    : 'Not logged in';

  title.appendChild(banner);
}

function setAuthLockedLayout(locked) {
  const sidebar = document.querySelector('.sidebar');
  if (sidebar) sidebar.style.display = locked ? 'none' : 'block';
}

function renderLoginScreen(statusText = '') {
  setAuthLockedLayout(true);
  document.getElementById('content').innerHTML = `
    <div style="max-width:680px;margin:20px auto;border:1px solid rgba(255,255,255,.25);padding:18px;background:rgba(0,0,0,.15)">
      <h2>Command Login</h2>
      <p>Only users listed in <b>Command_Users</b> can create accounts and log in.</p>
      <div style="font-size:12px;opacity:.85;margin-bottom:8px;">Build: ${APP_BUILD}</div>

      <div style="display:flex;gap:10px;flex-wrap:wrap;margin:10px 0 16px 0;">
        <button id="showLoginPane">Login</button>
        <button id="showCreatePane">Create Account</button>
      </div>

      <div style="border:1px solid rgba(255,255,255,.2);padding:10px;margin-bottom:12px;">
        <b>First-time setup</b><br>
        <span style="font-size:13px;opacity:.9">If account creation says email not found, link your Command_Users tab once.</span><br>
        <input id="commandUsersUrl" type="text" placeholder="Paste Command_Users Google Sheet/CSV link" style="margin-top:8px;width:100%;max-width:540px"><br><br>
        <button id="linkCommandUsersBtn">Link Command_Users Tab</button>
      </div>

      <div id="loginPane" style="display:block">
        <h3>Login</h3>
        <label>Email</label><br>
        <input id="loginEmail" type="email" style="width:100%;max-width:360px"><br>
        <label>Password</label><br>
        <input id="loginPassword" type="password" style="width:100%;max-width:360px"><br><br>
        <button id="loginBtn">Login</button>
      </div>

      <div id="createPane" style="display:none">
        <h3>Create Account</h3>
        <label>Email</label><br>
        <input id="createEmail" type="email" style="width:100%;max-width:360px"><br>
        <label>Password</label><br>
        <input id="createPassword" type="password" style="width:100%;max-width:360px"><br><br>
        <label>Verify Password</label><br>
        <input id="createPasswordVerify" type="password" style="width:100%;max-width:360px"><br><br>
        <button id="createAccountBtn">Create Account</button>
      </div>

      <pre id="accountStatus" style="margin-top:14px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">${statusText || 'Please login to continue.'}</pre>
    </div>
  `;

  const loginPane = document.getElementById('loginPane');
  const createPane = document.getElementById('createPane');
  document.getElementById('showLoginPane').addEventListener('click', () => {
    loginPane.style.display = 'block';
    createPane.style.display = 'none';
  });
  document.getElementById('showCreatePane').addEventListener('click', () => {
    loginPane.style.display = 'none';
    createPane.style.display = 'block';
  });
  document.getElementById('linkCommandUsersBtn').addEventListener('click', linkCommandUsersTab);
  document.getElementById('createAccountBtn').addEventListener('click', createAccount);
  document.getElementById('loginBtn').addEventListener('click', loginAccount);
}

async function linkCommandUsersTab() {
  const url = String((document.getElementById('commandUsersUrl') || {}).value || '').trim();
  const status = document.getElementById('accountStatus');
  if (!url) {
    if (status) status.textContent = 'Please paste a Command_Users tab link first.';
    return;
  }

  try {
    const response = await fetch('/api/auth/link-command-users', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ url })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to link Command_Users tab');

    const rows = (((data || {}).import || {}).result || []).find(x => x.name === 'command_users');
    const rowCount = rows && typeof rows.rows === 'number' ? rows.rows : 0;
    if (status) status.textContent = 'Command_Users linked successfully (' + rowCount + ' rows). You can now create your account.';
  } catch (err) {
    if (status) status.textContent = 'Link failed: ' + err.message;
  }
}

async function refreshAuthSession() {
  const token = getAuthToken();
  if (!token) {
    currentUser = null;
    showAuthBanner();
    renderLoginScreen();
    return;
  }

  try {
    const response = await authFetch('/api/auth/me');
    const data = await response.json();
    if (!response.ok) {
      setAuthToken('');
      currentUser = null;
    } else {
      currentUser = data.user || null;
    }
  } catch (e) {
    currentUser = null;
  }
  showAuthBanner();
  if (!currentUser) {
    renderLoginScreen();
    return;
  }

  // On initial authenticated load, land on dashboard.
  const content = document.getElementById('content');
  if (content && /Command Login/i.test(String(content.textContent || ''))) {
    loadPage('dashboard');
  }
}

function getLocalSyncTabs() {
  try {
    const raw = localStorage.getItem(LOCAL_SYNC_TABS_KEY);
    const parsed = JSON.parse(raw || '[]');
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(t => t && t.name && t.url);
  } catch (e) {
    return [];
  }
}

function setLocalSyncTabs(tabs) {
  try {
    const safe = Array.isArray(tabs) ? tabs.filter(t => t && t.name && t.url) : [];
    localStorage.setItem(LOCAL_SYNC_TABS_KEY, JSON.stringify(safe));
  } catch (e) {
    // Ignore localStorage errors (private mode/quota).
  }
}

function loadPage(page){
if(!isLoggedIn()){
renderLoginScreen();
return;
}

setAuthLockedLayout(false);

/* DASHBOARD */

if(page === "dashboard"){

document.getElementById("content").innerHTML = `
<h2>Command Dashboard</h2>

<p>Welcome to the FWPD Command Portal.</p>

<p>Use the sidebar to navigate the system.</p>

<div id="welcomeMessage" style="margin-top:8px;color:#d8f3ff"></div>

<div style="margin-top:20px">
<b>Alerts</b>
<pre id="dashboardAlerts" style="margin-top:8px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Loading alerts...</pre>
</div>

<div style="margin-top:30px">

<b>Department Structure</b>

<ul>
<li>Divisions: Adam, Nora</li>
<li>Subdivisions assigned through command qualification</li>
</ul>

</div>

<div style="margin-top:24px">
<b>Google Sync Status</b>
<pre id="syncStatusBox" style="margin-top:8px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Loading sync status...</pre>
</div>
`;

if (currentUser) {
  const welcome = document.getElementById('welcomeMessage');
  if (welcome) {
    welcome.textContent = 'Welcome ' + formatUserDisplayName(currentUser) + '.';
  }
}

loadDashboardAlerts();
loadSyncStatus();
autoSyncOnLoad();

}


/* REPORTS */

if(page === "reports"){

document.getElementById("content").innerHTML = `
<h2>Reports</h2>
<p>Command review center for discipline, cadet evaluations, and internal messages.</p>

<div style="margin-top:12px;display:flex;gap:10px;flex-wrap:wrap;">
  <button onclick="loadPage('sheettabs')">Open Sheet Tabs</button>
</div>

<pre id="reportsSummary" style="margin-top:14px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Loading reports summary...</pre>
`;

loadReportsSummary();

}


/* ACCOUNT */

if(page === "account"){

document.getElementById("content").innerHTML = `
<h2>My Account</h2>
<p>Personal command profile information.</p>
<table>
  <thead>
    <tr><th>Field</th><th>Value</th></tr>
  </thead>
  <tbody>
    <tr><td>Email</td><td>${(currentUser && currentUser.email) || '-'}</td></tr>
    <tr><td>Character Name</td><td>${(currentUser && currentUser.characterName) || '-'}</td></tr>
    <tr><td>Rank</td><td>${(currentUser && currentUser.rank) || '-'}</td></tr>
    <tr><td>Role</td><td>${(currentUser && currentUser.role) || '-'}</td></tr>
  </tbody>
</table>
<div style="margin-top:12px;">
  <button id="logoutBtn">Logout</button>
</div>
<pre id="accountStatus" style="margin-top:14px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Logged in.</pre>
`;

document.getElementById('logoutBtn').addEventListener('click', logoutAccount);

}

/* OFFICER ROSTER */

if(page === "roster"){

document.getElementById("content").innerHTML = `
<h2>Officer Roster</h2>
<p id="rosterStatus">Loading roster...</p>
<div style="margin:8px 0">
  <button id="addOfficer">Add Officer</button>
  <button id="refreshRoster">Refresh</button>
  <button id="syncSheets">Sync Now (Manual)</button>
</div>
<table id="rosterTable">
<thead>
<tr>
<th>ID</th>
<th>Name</th>
<th>Callsign</th>
<th>Rank</th>
<th>Division</th>
<th>Actions</th>
</tr>
</thead>
<tbody></tbody>
</table>
<pre id="rosterDebug" style="margin-top:8px;color:#800;white-space:pre-wrap"></pre>
`;

loadRoster();

document.getElementById('refreshRoster').addEventListener('click', loadRoster);
if (isLoggedIn()) {
  document.getElementById('addOfficer').addEventListener('click', () => showOfficerForm());
} else {
  document.getElementById('addOfficer').disabled = true;
  document.getElementById('addOfficer').title = 'Login required';
}
document.getElementById('syncSheets').addEventListener('click', syncGoogleSheets);

}


if(page === "sheettabs"){

document.getElementById("content").innerHTML = `
<h2>Sheet Tabs</h2>
<p>Imported tabs from Google Sheets will appear here.</p>
<div id="sheetTabsList">Loading tabs...</div>
<div id="sheetTabData" style="margin-top:16px"></div>
`;

loadSheetTabs();

}

}


/* ROSTER LOADER */


async function loadRoster(){
  const statusEl = document.getElementById('rosterStatus');
  statusEl.textContent = 'Loading roster...';

  try {
    const response = await fetch('/api/roster');
    const data = await response.json();
    const table = document.querySelector('#rosterTable tbody');
    table.innerHTML = '';
    let count = 0;
    const pick = (obj, keys) => {
      for (const key of keys) {
        const val = String((obj && obj[key]) || '').trim();
        if (val) return val;
      }
      return '';
    };

    (data || []).forEach((item) => {
      const id = pick(item, ['ID', 'id', 'Officer_ID', 'officer_id']);
      let name = pick(item, ['Name', 'name', 'RP_Name', 'rp_name', 'Officer_Name', 'officer_name']);
      let callsign = pick(item, ['Callsign', 'callsign', 'Call_Sign', 'call_sign']);
      const rank = pick(item, ['Rank', 'rank']);
      const division = pick(item, ['Division', 'division', 'Unit', 'unit']);

      if (!name && callsign && /[a-z]/i.test(callsign) && !/\d/.test(callsign)) {
        name = callsign;
        callsign = '';
      }

      const hasIdentity = !!(id || name || callsign);
      if (!hasIdentity) return;

      const displayName = name || '(No Name)';
      const safeIdAttr = String(id || '').replace(/"/g, '&quot;');
      const safeNameAttr = String(name || '').replace(/"/g, '&quot;');
      const safeCallsignAttr = String(callsign || '').replace(/"/g, '&quot;');

      const tr = document.createElement('tr');
      const actionButtons = isLoggedIn()
        ? `<button data-id="${safeIdAttr}" data-name="${safeNameAttr}" data-callsign="${safeCallsignAttr}" onclick="openOfficerProfileFromRow(this)">Profile</button>
            <button onclick="editOfficer('${id}')">Edit</button>
            <button onclick="deleteOfficer('${id}')">Delete</button>`
        : `<button data-id="${safeIdAttr}" data-name="${safeNameAttr}" data-callsign="${safeCallsignAttr}" onclick="openOfficerProfileFromRow(this)">Profile</button>`;
      tr.innerHTML = `
        <td>
          <button
            class="profile-link"
            data-id="${safeIdAttr}"
            data-name="${safeNameAttr}"
            data-callsign="${safeCallsignAttr}"
            onclick="openOfficerProfileFromRow(this)">
            ${id || '(No ID)'}
          </button>
        </td>
        <td>${displayName}</td>
        <td>${callsign}</td>
        <td>${rank}</td>
        <td>${division}</td>
        <td class="actions-cell">
          <div class="row-actions">
            ${actionButtons}
          </div>
        </td>
      `;
      table.appendChild(tr);
      count++;
    });
    statusEl.textContent = 'Roster loaded (' + count + ' officers)';
  } catch (err) {
    console.error('Error loading roster:', err);
    statusEl.textContent = 'Error: ' + err.message;
  }
}

async function openOfficerProfileFromRow(el) {
  const id = String((el && el.dataset && el.dataset.id) || '').trim();
  const name = String((el && el.dataset && el.dataset.name) || '').trim();
  const callsign = String((el && el.dataset && el.dataset.callsign) || '').trim();

  return openOfficerProfile({ id, name, callsign });
}

async function openOfficerProfile(criteria) {
  try {
    const response = await fetch('/api/roster');
    const data = await response.json();
    if (!response.ok) throw new Error('Failed to load roster data');

    const pick = (obj, keys) => {
      for (const key of keys) {
        const val = String((obj && obj[key]) || '').trim();
        if (val) return val;
      }
      return '';
    };

    const id = String((criteria && criteria.id) || '').trim();
    const name = String((criteria && criteria.name) || '').trim();
    const callsign = String((criteria && criteria.callsign) || '').trim();

    let officer = null;
    if (id) {
      officer = (data || []).find(x => {
        const xId = pick(x, ['ID', 'id', 'Officer_ID', 'officer_id']);
        return String(xId).trim() === id;
      });
    }
    if (!officer && name) {
      officer = (data || []).find(x => {
        const xName = pick(x, ['Name', 'name', 'RP_Name', 'rp_name', 'Officer_Name', 'officer_name']);
        return String(xName).trim().toLowerCase() === name.toLowerCase();
      });
    }
    if (!officer && callsign) {
      officer = (data || []).find(x => {
        const xCallsign = pick(x, ['Callsign', 'callsign', 'Call_Sign', 'call_sign']);
        return String(xCallsign).trim().toLowerCase() === callsign.toLowerCase();
      });
    }

    if (!officer) {
      alert('Officer profile not found. Try refreshing the roster.');
      return;
    }

    const enrichedOfficer = await enrichOfficerProfileData(officer);
    renderOfficerProfile(enrichedOfficer);
  } catch (err) {
    alert('Failed to open profile: ' + err.message);
  }
}

function pickOfficerField(obj, keys) {
  for (const key of keys) {
    const val = String((obj && obj[key]) || '').trim();
    if (val) return val;
  }
  return '';
}

function computeColumnsCTFromRecord(record) {
  const output = {};
  const keys = Object.keys(record || {});

  // Columns C through T are zero-based indexes 2..19.
  for (let i = 2; i <= 19; i++) {
    const key = keys[i] || ('Column ' + String.fromCharCode(65 + i));
    output[key] = String((record && record[key]) || '').trim();
  }

  return output;
}

function matchRawRecordToOfficer(raw, officer) {
  const rawId = pickOfficerField(raw, ['ID', 'id', 'Officer_ID', 'officer_id']);
  const rawName = pickOfficerField(raw, ['Name', 'name', 'RP_Name', 'rp_name', 'Officer_Name', 'officer_name']);
  const rawCallsign = pickOfficerField(raw, ['Callsign', 'callsign', 'Call_Sign', 'call_sign']);

  const offId = pickOfficerField(officer, ['ID', 'id', 'Officer_ID', 'officer_id']);
  const offName = pickOfficerField(officer, ['Name', 'name', 'RP_Name', 'rp_name', 'Officer_Name', 'officer_name']).toLowerCase();
  const offCallsign = pickOfficerField(officer, ['Callsign', 'callsign', 'Call_Sign', 'call_sign']).toLowerCase();

  if (offId && rawId && String(rawId).trim() === offId) return true;
  if (offName && rawName && String(rawName).trim().toLowerCase() === offName) return true;
  if (offCallsign && rawCallsign && String(rawCallsign).trim().toLowerCase() === offCallsign) return true;
  return false;
}

async function enrichOfficerProfileData(officer) {
  const hasImported = officer && officer.ImportedFields && Object.keys(officer.ImportedFields).length > 0;
  const hasCT = officer && officer.ColumnsCT && Object.keys(officer.ColumnsCT).length > 0;
  if (hasImported && hasCT) return officer;

  try {
    const response = await fetch('/api/sheets/tab/roster');
    const rows = await response.json();
    if (!response.ok || !Array.isArray(rows)) return officer;

    const rawMatch = rows.find(r => matchRawRecordToOfficer(r, officer));
    if (!rawMatch) return officer;

    const merged = Object.assign({}, officer);
    if (!hasImported) merged.ImportedFields = rawMatch;
    if (!hasCT) merged.ColumnsCT = computeColumnsCTFromRecord(rawMatch);
    return merged;
  } catch (e) {
    return officer;
  }
}

function renderOfficerProfile(officer) {
  const profileId = pickOfficerField(officer, ['ID', 'id', 'Officer_ID', 'officer_id']);
  const profileName = pickOfficerField(officer, ['Name', 'name', 'RP_Name', 'rp_name', 'Officer_Name', 'officer_name']);

  const imported = (officer && officer.ImportedFields && typeof officer.ImportedFields === 'object')
    ? officer.ImportedFields
    : ((officer && officer.ColumnsCT && typeof officer.ColumnsCT === 'object') ? officer.ColumnsCT : {});
  const importedRows = Object.keys(imported)
    .filter((k) => String(imported[k] || '').trim() !== '')
    .map((k) => [k, imported[k] || '-']);

  const rowsToTable = (rows) => {
    if (!rows.length) return '<p>No data available.</p>';
    let html = '<table><thead><tr><th>Field</th><th>Value</th></tr></thead><tbody>';
    rows.forEach((r) => {
      html += '<tr><td>' + r[0] + '</td><td>' + r[1] + '</td></tr>';
    });
    html += '</tbody></table>';
    return html;
  };

  document.getElementById('content').innerHTML = `
    <h2>Officer Profile</h2>
    <div style="margin:8px 0 16px 0;">
      <button onclick="loadPage('roster')">Back to Roster</button>
    </div>

    <p><b>ID:</b> ${profileId || '-'}</p>
    <p><b>Name:</b> ${profileName || '-'}</p>

    <h3 style="margin-top:18px;">All Imported Officer Data</h3>
    ${rowsToTable(importedRows)}
  `;
}

function showOfficerForm(id = null, data = {}) {
  const formHtml = `
    <div id="officerForm" style="position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:white;color:#111;padding:20px;border:1px solid #ccc;z-index:1000;box-shadow:0 0 10px rgba(0,0,0,0.3);min-width:260px;">
      <h3>${id ? 'Edit' : 'Add'} Officer</h3>
      <label>ID: <input id="formID" value="${data.ID || ''}"></label><br>
      <label>Name: <input id="formName" value="${data.Name || ''}"></label><br>
      <label>Callsign: <input id="formCallsign" value="${data.Callsign || ''}"></label><br>
      <label>Rank: <input id="formRank" value="${data.Rank || ''}"></label><br>
      <label>Division: <input id="formDivision" value="${data.Division || ''}"></label><br>
      <button id="saveOfficer">Save</button>
      <button onclick="closeForm()">Cancel</button>
    </div>
  `;
  document.body.insertAdjacentHTML('beforeend', formHtml);
  document.getElementById('saveOfficer').addEventListener('click', () => saveOfficer(id));
}

function closeForm() {
  const form = document.getElementById('officerForm');
  if (form) form.remove();
}

async function saveOfficer(id) {
  if (!isLoggedIn()) {
    alert('Login required for roster edits.');
    return;
  }
  const data = {
    ID: document.getElementById('formID').value,
    Name: document.getElementById('formName').value,
    Callsign: document.getElementById('formCallsign').value,
    Rank: document.getElementById('formRank').value,
    Division: document.getElementById('formDivision').value
  };
  try {
    const method = id ? 'PUT' : 'POST';
    const url = id ? `/api/roster/${id}` : '/api/roster';
    const response = await authFetch(url, {
      method: method,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    if (!response.ok) throw new Error('Save failed');
    closeForm();
    loadRoster();
  } catch (err) {
    alert('Error saving: ' + err.message);
  }
}

async function editOfficer(id) {
  if (!isLoggedIn()) {
    alert('Login required for roster edits.');
    return;
  }
  try {
    const response = await fetch(`/api/roster?id=${id}`);
    const data = await response.json();
    const item = data.find(x => x.ID === id);
    if (item) {
      showOfficerForm(id, item);
    } else {
      alert('Officer not found');
    }
  } catch (err) {
    alert('Error editing: ' + err.message);
  }
}

async function deleteOfficer(id) {
  if (!isLoggedIn()) {
    alert('Login required for roster edits.');
    return;
  }
  if (!confirm('Delete this officer?')) return;
  try {
    const response = await authFetch(`/api/roster/${id}`, { method: 'DELETE' });
    if (!response.ok) throw new Error('Delete failed');
    loadRoster();
  } catch (err) {
    alert('Error deleting: ' + err.message);
  }
}

async function createAccount() {
  const email = String(document.getElementById('createEmail').value || '').trim();
  const password = String(document.getElementById('createPassword').value || '');
  const passwordVerify = String(document.getElementById('createPasswordVerify').value || '');
  const status = document.getElementById('accountStatus');

  if (password !== passwordVerify) {
    if (status) status.textContent = 'Create account failed: Password and Verify Password do not match.';
    return;
  }

  const submitCreate = async () => {
    return fetch('/api/auth/create-account', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password })
    });
  };

  try {
    let response = await submitCreate();
    let data = await response.json();

    // If email lookup fails, try auto-linking command_users from setup URL and retry once.
    if (!response.ok && /email not found in command_users/i.test(String(data.error || ''))) {
      const setupUrlEl = document.getElementById('commandUsersUrl');
      const setupUrl = String((setupUrlEl && setupUrlEl.value) || '').trim();
      if (setupUrl) {
        const linkResp = await fetch('/api/auth/link-command-users', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ url: setupUrl })
        });
        if (linkResp.ok) {
          response = await submitCreate();
          data = await response.json();
        }
      }
    }

    if (!response.ok) throw new Error(data.error || 'Create account failed');

    setAuthToken(data.token || '');
    await refreshAuthSession();
    loadPage('dashboard');
    const welcome = 'Account created. Welcome ' + formatUserDisplayName(currentUser) + '.';
    if (status) status.textContent = welcome;
    alert(welcome);
  } catch (err) {
    if (status) status.textContent = 'Create account failed: ' + err.message;
  }
}

async function loginAccount() {
  const email = String(document.getElementById('loginEmail').value || '').trim();
  const password = String(document.getElementById('loginPassword').value || '');
  const status = document.getElementById('accountStatus');

  try {
    const response = await fetch('/api/auth/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Login failed');

    setAuthToken(data.token || '');
    await refreshAuthSession();
    loadPage('dashboard');
    const welcome = 'Login successful. Welcome ' + formatUserDisplayName(currentUser) + '.';
    if (status) status.textContent = welcome;
    alert(welcome);
  } catch (err) {
    if (status) status.textContent = 'Login failed: ' + err.message;
  }
}

function findBestField(row, aliases) {
  const keys = Object.keys(row || {});
  const normalized = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  for (const alias of aliases) {
    const wanted = normalized(alias);
    const key = keys.find(k => normalized(k) === wanted);
    if (key) {
      const val = String(row[key] || '').trim();
      if (val) return val;
    }
  }
  return '';
}

function mapAlertRow(source, row) {
  const title = findBestField(row, ['title', 'subject', 'reason', 'type', 'status']) || (source + ' item');
  const person = findBestField(row, ['rp_name', 'name', 'officer_name', 'cadet_name', 'character_name']);
  const detail = findBestField(row, ['notes', 'message', 'comments', 'description', 'summary']);
  const date = findBestField(row, ['date', 'created_date', 'created', 'last_updated', 'timestamp']);

  let line = '[' + source + '] ' + title;
  if (person) line += ' - ' + person;
  if (date) line += ' (' + date + ')';
  if (detail) line += '\n  ' + detail;
  return line;
}

async function loadDashboardAlerts() {
  const box = document.getElementById('dashboardAlerts');
  if (!box) return;

  try {
    const tabsRes = await fetch('/api/sheets/tabs');
    const tabs = await tabsRes.json();
    if (!tabsRes.ok || !Array.isArray(tabs)) throw new Error('Failed to load tab list');

    const sources = [
      { key: 'discipline_records', label: 'Discipline' },
      { key: 'cadet_evaluations', label: 'Cadet Evaluations' },
      { key: 'officer_notes', label: 'Internal Messages' },
      { key: 'internal_messages', label: 'Internal Messages' }
    ];

    const available = sources.filter(s => tabs.includes(s.key));
    if (!available.length) {
      box.textContent = 'No alert tabs imported yet.';
      return;
    }

    const lines = [];
    for (const src of available) {
      const res = await fetch('/api/sheets/tab/' + encodeURIComponent(src.key));
      const rows = await res.json();
      if (!res.ok || !Array.isArray(rows) || !rows.length) continue;

      const alertRows = rows.slice(0, 5);
      if (!alertRows.length) continue;

      alertRows.forEach(r => lines.push(mapAlertRow(src.label, r)));
    }

    box.textContent = lines.length ? lines.join('\n\n') : ('No alert records found. Tabs checked: ' + available.map(x => x.key).join(', '));
  } catch (err) {
    box.textContent = 'Alerts unavailable: ' + err.message;
  }
}

async function loadReportsSummary() {
  const box = document.getElementById('reportsSummary');
  if (!box) return;

  try {
    const tabsRes = await fetch('/api/sheets/tabs');
    const tabs = await tabsRes.json();
    if (!tabsRes.ok || !Array.isArray(tabs)) throw new Error('Failed to load sheet tabs');

    const sources = [
      { key: 'discipline_records', label: 'Discipline Reports' },
      { key: 'cadet_evaluations', label: 'Cadet Evaluations' },
      { key: 'officer_notes', label: 'Internal Messages' },
      { key: 'internal_messages', label: 'Internal Messages (Alt Tab)' }
    ];

    const lines = [];
    for (const src of sources) {
      if (!tabs.includes(src.key)) {
        lines.push(src.label + ': not imported');
        continue;
      }

      const res = await fetch('/api/sheets/tab/' + encodeURIComponent(src.key));
      const rows = await res.json();
      if (!res.ok || !Array.isArray(rows)) {
        lines.push(src.label + ': failed to load');
      } else {
        lines.push(src.label + ': ' + rows.length + ' records');
      }
    }

    lines.push('');
    lines.push('Step 2 next: approval workflow with approver + timestamp on discipline/eval actions.');
    box.textContent = lines.join('\n');
  } catch (err) {
    box.textContent = 'Reports summary unavailable: ' + err.message;
  }
}

async function logoutAccount() {
  try {
    await authFetch('/api/auth/logout', { method: 'POST' });
  } catch (e) {
    // Ignore logout errors.
  }
  setAuthToken('');
  currentUser = null;
  showAuthBanner();
  renderLoginScreen('Logged out. Please login to continue.');
}

async function syncGoogleSheets() {
  const isValidGoogleSheetLink = (input) => {
    const value = String(input || '').trim();
    if (!value) return false;
    try {
      const u = new URL(value);
      const host = String(u.hostname || '').toLowerCase();
      const path = String(u.pathname || '').toLowerCase();
      if (!host.includes('docs.google.com')) return false;
      return path.includes('/spreadsheets/');
    } catch (e) {
      return false;
    }
  };

  const rosterURL = prompt('Paste Roster Google Sheets link (tab link or published CSV):');
  if (!rosterURL) return;

  if (!isValidGoogleSheetLink(rosterURL)) {
    alert('Invalid roster link. Please paste a Google Sheets link from docs.google.com/spreadsheets/...');
    return;
  }

  const otherTabsInput = prompt(
    'Optional: add other tabs as comma-separated Name|CSV_URL entries.\nExample:\ndivisions|https://...csv,discipline|https://...csv'
  ) || '';

  const tabs = [{ name: 'roster', url: rosterURL.trim() }];

  if (otherTabsInput.trim()) {
    const invalidLinks = [];
    otherTabsInput.split(',').forEach(pair => {
      const parts = pair.split('|');
      if (parts.length >= 2) {
        const name = parts[0].trim();
        const url = parts.slice(1).join('|').trim();
        if (!name || !url) return;
        if (!isValidGoogleSheetLink(url)) {
          invalidLinks.push(name || 'unnamed');
          return;
        }
        tabs.push({ name, url });
      }
    });

    if (invalidLinks.length) {
      alert('Invalid Google Sheets link for tab(s): ' + invalidLinks.join(', ') + '.\nUse docs.google.com/spreadsheets/... links.');
      return;
    }
  }

  try {
    setLocalSyncTabs(tabs);

    await fetch('/api/sheets/config', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tabs, autoSyncOnLoad: true })
    });

    const response = await fetch('/api/sheets/import', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tabs })
    });

    const result = await response.json();
    if (!response.ok) throw new Error(result.error || 'Import failed');

    const rosterResult = (result.result || []).find(x => x.name === 'roster');
    const tabSummary = (result.result || []).map(x => {
      if (!x.ok) return `${x.name}: FAILED (${x.error || 'error'})`;
      if (typeof x.rows === 'number') return `${x.name}: OK (${x.rows} rows)`;
      return `${x.name}: OK`;
    }).join(', ');

    let message = 'Import completed. ' + tabSummary;
    if (rosterResult && rosterResult.ok) {
      const parsed = Number(rosterResult.parsedRows || 0);
      const headers = Array.isArray(rosterResult.headers) ? rosterResult.headers.join(', ') : '';
      message += `\nRoster parser: ${parsed} source rows.`;
      if (headers) message += `\nDetected headers: ${headers}`;
    }
    if (rosterResult && rosterResult.ok && Number(rosterResult.rows || 0) === 0) {
      message += '\n\nRoster imported 0 rows. Confirm the roster tab has officer data and a Name/RP_Name or Callsign column.';
    }

    alert(message);
    loadRoster();
    loadSyncStatus();
  } catch (err) {
    alert('Google sync failed: ' + err.message + '\n\nUse a published CSV link or standard sheet link with correct tab gid.');
  }
}

function formatSyncSummary(result) {
  const okRows = (result || []).filter(x => x.ok).map(x => {
    const rows = typeof x.rows === 'number' ? ` (${x.rows} rows)` : '';
    return `${x.name}: OK${rows}`;
  });
  const failRows = (result || []).filter(x => !x.ok).map(x => `${x.name}: FAILED (${x.error || 'error'})`);
  return okRows.concat(failRows).join('\n');
}

async function loadSyncStatus() {
  const box = document.getElementById('syncStatusBox');
  if (!box) return;

  try {
    const response = await fetch('/api/sheets/config');
    const config = await response.json();
    if (!response.ok) throw new Error(config.error || 'Failed to load sync status');

    const lines = [];
    lines.push('Auto-sync on load: ' + (config.autoSyncOnLoad ? 'Enabled' : 'Disabled'));
    lines.push('Saved tabs: ' + ((config.tabs || []).map(t => t.name).join(', ') || 'none'));
    if (config.lastSync && config.lastSync.at) {
      lines.push('Last sync: ' + new Date(config.lastSync.at).toLocaleString());
      if (config.lastSync.summary) {
        lines.push('Summary: ' + config.lastSync.summary.ok + ' OK / ' + config.lastSync.summary.failed + ' failed');
      }
      if (Array.isArray(config.lastSync.result) && config.lastSync.result.length) {
        lines.push('');
        lines.push(formatSyncSummary(config.lastSync.result));
      }
    } else {
      lines.push('Last sync: none yet');
    }

    box.textContent = lines.join('\n');
  } catch (err) {
    box.textContent = 'Sync status unavailable: ' + err.message;
  }
}

async function autoSyncOnLoad() {
  if (sessionStorage.getItem(AUTO_SYNC_SESSION_KEY) === '1') return;

  try {
    const configResponse = await fetch('/api/sheets/config');
    const config = await configResponse.json();
    if (!configResponse.ok) return;

    let tabsToSync = Array.isArray(config.tabs) ? config.tabs : [];
    let autoEnabled = !!config.autoSyncOnLoad;

    // Render free deployments can reset local files; restore from browser cache when available.
    if (!tabsToSync.length) {
      const cachedTabs = getLocalSyncTabs();
      if (cachedTabs.length) {
        tabsToSync = cachedTabs;
        autoEnabled = true;
        await fetch('/api/sheets/config', {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tabs: tabsToSync, autoSyncOnLoad: true })
        });
      }
    }

    if (!autoEnabled || !tabsToSync.length) {
      await loadSyncStatus();
      return;
    }

    sessionStorage.setItem(AUTO_SYNC_SESSION_KEY, '1');
    const syncResponse = await fetch('/api/sheets/sync', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tabs: tabsToSync })
    });
    const syncResult = await syncResponse.json();
    if (!syncResponse.ok) throw new Error(syncResult.error || 'Auto-sync failed');
    await loadSyncStatus();
  } catch (err) {
    const box = document.getElementById('syncStatusBox');
    if (box) box.textContent = 'Auto-sync failed: ' + err.message;
  }
}

async function loadSheetTabs() {
  const listEl = document.getElementById('sheetTabsList');
  const dataEl = document.getElementById('sheetTabData');

  try {
    const response = await fetch('/api/sheets/tabs');
    const tabs = await response.json();
    if (!Array.isArray(tabs) || !tabs.length) {
      listEl.textContent = 'No imported tabs yet. Use Officer Roster > Sync Google Sheets.';
      return;
    }

    listEl.innerHTML = '';
    tabs.forEach(name => {
      const btn = document.createElement('button');
      btn.textContent = 'View ' + name;
      btn.style.marginRight = '8px';
      btn.style.marginBottom = '8px';
      btn.addEventListener('click', () => loadSheetTabData(name));
      listEl.appendChild(btn);
    });

    dataEl.textContent = 'Select a tab to preview.';
  } catch (err) {
    listEl.textContent = 'Failed to load tabs: ' + err.message;
  }
}

async function loadSheetTabData(name) {
  const dataEl = document.getElementById('sheetTabData');
  try {
    const response = await fetch('/api/sheets/tab/' + encodeURIComponent(name));
    const rows = await response.json();
    if (!response.ok) throw new Error(rows.error || 'Failed to load tab');

    if (!Array.isArray(rows) || !rows.length) {
      dataEl.textContent = 'No rows found in ' + name;
      return;
    }

    const cols = Object.keys(rows[0]);
    let html = '<h3>' + name + '</h3><table><thead><tr>';
    cols.forEach(c => { html += '<th>' + c + '</th>'; });
    html += '</tr></thead><tbody>';

    rows.slice(0, 100).forEach(r => {
      html += '<tr>';
      cols.forEach(c => { html += '<td>' + (r[c] || '') + '</td>'; });
      html += '</tr>';
    });

    html += '</tbody></table><p>Showing up to first 100 rows.</p>';
    dataEl.innerHTML = html;
  } catch (err) {
    dataEl.textContent = 'Failed to load tab: ' + err.message;
  }
}

if (!getAuthToken()) {
  currentUser = null;
  showAuthBanner();
  renderLoginScreen();
}

refreshAuthSession();