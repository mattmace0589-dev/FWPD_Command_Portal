# FWPD Command Portal Beta Training Checklist

## Pre-Session (2 min)
- Ensure everyone uses latest build by hard refreshing (`Ctrl+F5`).
- Confirm trainer account can see `Admin` and `FTO` tabs.
- Verify command users can access login/create account page.

## 1) Login and Account Access (2 min)
- Click `Sync Command_Users` on login screen.
- Create account for a valid `Command_Users` email.
- Verify successful login and top banner identity.

Pass criteria:
- No manual URL paste needed.
- Valid command emails can create accounts.

## 2) Role Visibility and Permissions (2 min)
- Admin account sees `Admin` and `FTO` tabs.
- Chief/Commander sees `FTO` and profile promotion controls.
- Standard command account cannot access restricted admin actions.

Pass criteria:
- Sidebar tabs match permissions.

## 3) Admin Delegation (2 min)
- Go to `Admin`.
- Grant admin to one user (including users with `Account = No`).
- Revoke admin and grant again to confirm both directions.

Pass criteria:
- Admin grant/revoke succeeds without needing target to already have account.

## 4) FTO Workflow (3 min)
- Open `FTO` tab.
- Add one officer from dropdown.
- Open officer profile and confirm `FTO: ACTIVE TRAINER`.
- Remove same officer and confirm badge clears.

Pass criteria:
- Add/remove works and profile updates immediately.

## 5) Promotion + Status Editing (2 min)
- Open an officer profile.
- Apply promotion update (rank/division).
- Edit officer fields for `Status` and `Activity Status`.

Pass criteria:
- Changes persist and are visible in roster/profile views.

## 6) Reports + Messaging Spot Check (2 min)
- Open `Reports` and verify disciplinary/cadet data loads.
- Run one filter and verify expected results.
- Send a test message and verify unread indicator updates.

Pass criteria:
- Reports and internal messaging are operational.

## Rapid Issue Capture Template
- User email:
- Role shown in account:
- Feature/page:
- Exact action taken:
- Exact error text:
- Time observed:

## Go/No-Go
Go if:
- Login/account creation succeeds for valid users.
- Role-based visibility and permissions behave correctly.
- FTO, promotions, reports, and messaging pass checks.

No-Go if:
- Any critical auth/permission path fails.
- Core workflows (FTO/promotion/reports) fail for command leadership.
