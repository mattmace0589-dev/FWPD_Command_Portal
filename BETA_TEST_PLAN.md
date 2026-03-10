# FWPD Command Portal Beta Test Plan

## Scope
Validate key command workflows before opening broader beta access.

## Test Accounts
- One command-level account
- One secondary command account (for messaging)

## Smoke Tests (Must Pass)
1. `GET /api/health` returns `ok: true`.
2. Login succeeds and dashboard loads.
3. Officer roster loads without console/API errors.
4. Reports page loads summary and table rows.
5. Messages page loads recipients/inbox.

## Functional Tests
1. Roster:
- Add officer
- Edit officer fields
- Save officer notes in profile
- Delete test officer

2. Reports:
- Filter by type/status/officer
- Open report details panel
- Approve, deny, and reset one report

3. Messaging:
- Send message to second user
- Confirm unread count changes
- Mark read/unread

4. Account:
- Change own password
- If privileged, reset another user password

5. Google Sync:
- Verify configured tabs in sync status
- Run manual sync and confirm status refresh

## UI/UX Checks
1. Header/sidebar colors consistent with brand palette.
2. Center patch background is visible but not obstructive.
3. NTXRP logo appears below patch and remains legible.
4. Roster action buttons display full labels on desktop/mobile.

## Exit Criteria
- All smoke tests pass
- No data-loss behavior observed
- No blocking auth/session issues
- No major layout break on desktop/mobile
