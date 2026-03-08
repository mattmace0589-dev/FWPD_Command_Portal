FWPD Command Portal — Quick Start

This is a local portal that stores roster data in a simple JSON file (zero cost, no cloud needed).

Quick Setup (5 minutes)
-----------------------
1. Download and install Node.js from https://nodejs.org/
2. Restart your terminal (PowerShell / CMD).
3. Open a terminal in this project folder and run:
```bash
npm install
npm start
```
4. Open http://localhost:3000 in your browser.
5. Click **Officer Roster** → Click **Add Officer** to add/edit/delete from the portal.

Data Storage
------------
All roster data is stored locally in `data/roster.json` — no cloud, no fees, no Firebase paywall.

API Endpoints (for reference)
-----------------------------
The server exposes these endpoints:
- `GET /api/roster` — list all officers
- `POST /api/roster` — add new officer (JSON body)
- `PUT /api/roster/:id` — update officer
- `DELETE /api/roster/:id` — delete officer

Example add (using curl or Postman):
```bash
curl -X POST http://localhost:3000/api/roster \
  -H "Content-Type: application/json" \
  -d '{"ID":"1","Name":"John Doe","Callsign":"Alpha-1","Rank":"Officer","Division":"Adam"}'
```

Import from Google Sheets
--------------------------
If you have roster data in Google Sheets:
1. File → Download → CSV.
2. Save as `roster.csv` in the project folder.
3. Restart the server (it will auto-import on first run).
4. Data is now available in the portal.

That's it! No Firebase, no paywall, just local.

Local editable server (add/edit/delete roster)
---------------------------------------------
If you want to edit the roster from the application, you can run the included Node server which exposes a simple REST API and serves the portal.

1) Install Node.js (if not installed) and from the project folder run:
```bash
npm install
npm start
```

2) The server runs at `http://localhost:3000` by default. Open that URL (or use Live Server) and go to Officer Roster.

3) Click the **Use Local Server** button to load data from the local API (`/api/roster`).

4) Data is persisted in `data/roster.json`. You can edit it directly or use the API endpoints:
- `GET /api/roster` — list
- `POST /api/roster` — add (JSON body)
- `PUT /api/roster/:id` — update
- `DELETE /api/roster/:id` — delete

Notes:
- On first run, if `roster.csv` exists in the project root the server will import it into `data/roster.json`.
- This local server is for development only — do not expose it publicly without adding authentication.