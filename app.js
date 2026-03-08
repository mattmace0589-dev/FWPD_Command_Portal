// Simple local API calls to /api/roster endpoints

const AUTO_SYNC_SESSION_KEY = 'fwpd_auto_sync_done';
const LOCAL_SYNC_TABS_KEY = 'fwpd_sync_tabs_v1';

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

/* DASHBOARD */

if(page === "dashboard"){

document.getElementById("content").innerHTML = `
<h2>Command Dashboard</h2>

<p>Welcome to the FWPD Command Portal.</p>

<p>Use the sidebar to navigate the system.</p>

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

loadSyncStatus();
autoSyncOnLoad();

}


/* DIVISIONS */

if(page === "divisions"){

document.getElementById("content").innerHTML = `
<h2>Department Divisions</h2>

<ul>
<li>Adam Division</li>
<li>Nora Division</li>
</ul>

<p style="margin-top:20px;">
Subdivisions such as CIU, MBU, K9, Detectives, Air-1, Air-2 and Drone Surveillance
are assigned through officer qualification and command approval.
</p>
`;

}


/* DISCIPLINE */

if(page === "discipline"){

document.getElementById("content").innerHTML = `
<h2>Discipline Records</h2>

<p>Command staff can review disciplinary records here.</p>
`;

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
document.getElementById('addOfficer').addEventListener('click', () => showOfficerForm());
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
            <button data-id="${safeIdAttr}" data-name="${safeNameAttr}" data-callsign="${safeCallsignAttr}" onclick="openOfficerProfileFromRow(this)">Profile</button>
            <button onclick="editOfficer('${id}')">Edit</button>
            <button onclick="deleteOfficer('${id}')">Delete</button>
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

    renderOfficerProfile(officer);
  } catch (err) {
    alert('Failed to open profile: ' + err.message);
  }
}

function renderOfficerProfile(officer) {
  const coreRows = [
    ['ID', officer.ID || ''],
    ['Name', officer.Name || ''],
    ['Callsign', officer.Callsign || ''],
    ['Rank', officer.Rank || ''],
    ['Division', officer.Division || '']
  ];

  const imported = (officer && officer.ImportedFields && typeof officer.ImportedFields === 'object')
    ? officer.ImportedFields
    : {};
  const importedRows = Object.keys(imported)
    .filter((k) => String(imported[k] || '').trim() !== '')
    .map((k) => [k, imported[k]]);

  const columnsNT = (officer && officer.ColumnsNT && typeof officer.ColumnsNT === 'object')
    ? officer.ColumnsNT
    : {};
  const ntRows = Object.keys(columnsNT).map((k) => [k, columnsNT[k] || '-']);

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

    <h3>Core Information</h3>
    ${rowsToTable(coreRows)}

    <h3 style="margin-top:18px;">Columns N Through T</h3>
    ${rowsToTable(ntRows)}

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
    const response = await fetch(url, {
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
  if (!confirm('Delete this officer?')) return;
  try {
    const response = await fetch(`/api/roster/${id}`, { method: 'DELETE' });
    if (!response.ok) throw new Error('Delete failed');
    loadRoster();
  } catch (err) {
    alert('Error deleting: ' + err.message);
  }
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

loadPage('dashboard');