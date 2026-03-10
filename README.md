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
- DB-backed persistence for users/sessions/admin-role overrides/FTO assignments (when `DATABASE_URL` is set)
- Optional auto-launch links for roster/training/discipline tabs on dashboard login

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

Environment Variables
---------------------
- `DATABASE_URL` Postgres connection string to enable DB persistence layer.
- `DEFAULT_ROSTER_URL` optional default roster sheet URL.
- `DEFAULT_COMMAND_USERS_URL` optional default `command_users` URL.
- `DEFAULT_TRAINING_URL` optional default training tab URL.
- `DEFAULT_DISCIPLINE_URL` optional default discipline tab URL.
- `DEFAULT_TRAINING_TAB_NAME` optional tab name for training URL (default `training_records`).
- `DEFAULT_DISCIPLINE_TAB_NAME` optional tab name for discipline URL (default `discipline_records`).

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