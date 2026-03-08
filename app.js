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
    (data || []).forEach((item, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${item.ID || ''}</td>
        <td>${item.Name || ''}</td>
        <td>${item.Callsign || ''}</td>
        <td>${item.Rank || ''}</td>
        <td>${item.Division || ''}</td>
        <td>
          <button onclick="editOfficer('${item.ID}')">Edit</button>
          <button onclick="deleteOfficer('${item.ID}')">Delete</button>
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
    <div id="officerForm" style="position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:white;padding:20px;border:1px solid #ccc;z-index:1000;box-shadow:0 0 10px rgba(0,0,0,0.3);">
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