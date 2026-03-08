// Simple local API calls to /api/roster endpoints

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
`;

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
  <button id="syncSheets">Sync Google Sheets</button>
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

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${id}</td>
        <td>${name}</td>
        <td>${callsign}</td>
        <td>${rank}</td>
        <td>${division}</td>
        <td class="actions-cell">
          <div class="row-actions">
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
  const rosterURL = prompt('Paste Roster CSV URL (published Google Sheets CSV link):');
  if (!rosterURL) return;

  const otherTabsInput = prompt(
    'Optional: add other tabs as comma-separated Name|CSV_URL entries.\nExample:\ndivisions|https://...csv,discipline|https://...csv'
  ) || '';

  const tabs = [{ name: 'roster', url: rosterURL.trim() }];

  if (otherTabsInput.trim()) {
    otherTabsInput.split(',').forEach(pair => {
      const parts = pair.split('|');
      if (parts.length >= 2) {
        const name = parts[0].trim();
        const url = parts.slice(1).join('|').trim();
        if (name && url) tabs.push({ name, url });
      }
    });
  }

  try {
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
    if (rosterResult && rosterResult.ok && Number(rosterResult.rows || 0) === 0) {
      message += '\n\nRoster imported 0 rows. Confirm the roster tab has officer data and a Name/RP_Name or Callsign column.';
    }

    alert(message);
    loadRoster();
  } catch (err) {
    alert('Google sync failed: ' + err.message);
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