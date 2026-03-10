FWPD Command Portal

Command workflow portal for roster management, reports review/approval, account controls, and internal messaging.

Quick Start
-----------
1. Install Node.js from https://nodejs.org/
2. Open this project in a terminal.
3. Run:

```bash
npm install
npm start
```

4. Open `http://localhost:3000`

Core Features
-------------
- Command login + account creation (backed by `command_users` import)
- Officer roster CRUD (with per-officer notes)
- Reports queue with filters, details, and approval actions
- Internal messaging (inbox/sent/unread count)
- Google Sheets tab linking and sync status

Data Files
----------
- `data/roster.json`
- `data/sheets-config.json`
- `data/reports-config.json`
- `data/report_approvals.json`
- `data/users.json`
- `data/sessions.json`
- `data/internal_mailbox.json`

API Highlights
--------------
- `GET /api/health` health/uptime check
- `GET /api/roster` list officers
- `POST /api/roster` add officer
- `PUT /api/roster/:id` update officer
- `DELETE /api/roster/:id` delete officer
- `GET /api/reports/items` report list
- `POST /api/reports/:id/approval` set report status
- `GET /api/messages/inbox` inbox
- `POST /api/messages/send` send message

Beta Readiness Checklist
------------------------
1. Confirm `GET /api/health` returns `ok: true`.
2. Log in with a command account and verify sidebar/pages load.
3. Test roster CRUD on one officer, including notes save.
4. Open Reports, filter by officer, open details, approve/deny one report.
5. Send one internal message and confirm inbox/unread updates.
6. Verify Google sync status and configured tabs appear as expected.
7. Verify mobile layout (sidebar, tables, profile buttons).

Deployment Note
---------------
If you use multiple local clones, always deploy from the same path your Git client tracks (to avoid pushing stale files).