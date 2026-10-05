# WORKSHOP Automation V1.4.1 — Incremental Communication Update

This release builds on the last-ship V1.3.5 backend. The existing scheduler, queue, certificate generation, PDF, and verification architecture is preserved.

## New
- Apple-style sheet-native Control Center with a step-by-step workshop creation flow.
- Online/offline branching.
- Smart meeting URL parsing for Google Meet/Zoom/Teams where information is objectively available. Missing meeting fields are hidden from participant emails.
- Workshop WhatsApp/group invitation URL. It appears in the Google Form confirmation message and Registration Confirmation email only.
- Registration email uses the supplied Finance Club PSTU Apple-white design.
- Reminder email uses the supplied Join Meeting design for online workshops and switches to venue details for offline workshops.
- Attendance Required email when Attendance is ON and the attendance window opens.
- Certificate email uses the supplied Apple-white certificate design.
- Announcement is now MANUAL ONLY. Use Custom Email.
- Self Test does not require a real Google Form submission or a production participant row. It uses a temporary test participant snapshot, supports multiple recipient emails, configurable selected stages, and configurable delays (immediately / 30 seconds / 1 / 2 / 5 minutes).
- Custom Email is a separate operator sheet. It supports All registered participants, Selected participants by email list, or Specific email addresses, plus presets, personalization, optional CTA, queue-safe delivery, recipient-count confirmation, preview draft, and audit records.
- Global verification base URL can be configured through `PUBLIC_VERIFY_URL`.
- Social footer supports Facebook, LinkedIn, Instagram, Website, and X.

## Install / upgrade
1. Replace the changed Apps Script project files with `WORKSHOP_Automation.gs` and `appsscript.json`. Keep your existing `Verify.html` unchanged.
2. Save the Apps Script project.
3. Deploy > Manage deployments > edit the existing Web app > Version = New version.
4. Deploy. Keep the existing `/exec` deployment URL.
5. Run `WORKSHOP Automation > Run Setup / Repair` once.
6. Open the Control Center.

## Preview Draft
The Custom Email preview button creates one Gmail draft for review. This uses the Gmail Compose permission; the draft is not sent automatically.

## Self Test
The main Control Center keeps the communication tools hidden. Use the separate `Self Test` sheet or the WORKSHOP Automation menu.

1. Select a workshop.
2. Enter one or more recipient emails, one per line.
3. Edit the temporary participant profile if needed.
4. Select Registration, Reminder, Attendance, and/or Certificate stages.
5. Choose the delay between selected stages.
6. Tick RUN SELF TEST. The first selected stage runs immediately; later selected stages use the configured delay.

Self Test emails use the production visual appearance. The Test Runs sheet and System Logs identify them as test runs.

## Custom Email
Use the separate `Custom Email` sheet. The exact recipient count is shown before sending.

1. Select a workshop.
2. Choose a preset or Custom.
3. Choose All registered participants, Selected participants, or Specific email addresses.
4. Fill Subject, Heading, and Message.
5. Optionally add a CTA.
6. Tick SEND CUSTOM EMAIL only after reviewing the recipient count.

Custom Email is manual communication. It is not part of the automatic registration/reminder/attendance/certificate lifecycle.

## Group invitation
The group link is stored per workshop. For new workshops, enter the invite URL in the Create Workshop area. It is shown in the Google Form confirmation message and Registration Confirmation email. It is not repeated in the Join Workshop/Reminder email.

## Meeting information
For Online workshops, enter the meeting URL. The system attempts to derive the platform and meeting ID when the URL contains objective information. Passcode and host remain manual optional values. Empty fields are omitted from participant emails.

For Offline workshops, meeting fields are hidden and venue information is used.

## Backward compatibility
V1.4 uses header-based schema migration. Existing workshop rows are preserved and new fields are appended by header name. Legacy scheduled Announcement settings are no longer processed automatically; use Custom Email for manual announcements/updates.

## Vercel
The `VERCEL UI` folder is unchanged from V1.3.5 in this release. Its backend continues to call the Apps Script `/exec?page=verify-api` endpoint. Update `api/backend-config.js` only when connecting Vercel to a completely different Apps Script project/account.

## Package policy
This is an update-only release. Unchanged V1.3.5 files are intentionally not duplicated in the package.

## Self Test quota note
Self Test sends directly so the test can happen immediately and at one-minute intervals. Each test email consumes normal Apps Script email quota; use a small number of test recipients.


## V1.4.1 incremental communication update

This release starts from the original V1.4.0 baseline and intentionally preserves the production schema and existing workflow. It adds dedicated `Self Test` and `Custom Email` sheets, keeps the old Control Center communication rows hidden for compatibility, adds configurable Self Test stages/recipients/delay, routes Custom Email through the existing central queue, fixes the malformed Reminder paragraph, uses direct Vercel certificate verification paths for newly generated certificates, removes the duplicate certificate attachment card, and adds render/send validation for malformed HTML and unresolved placeholders.

V1.4.1 does not add jsDelivr and does not migrate or rewrite the production data sheets.

## V1.4.1 safety baseline

V1.4.1 is built directly from the original V1.4.0 baseline. Later V1.4.6/V1.4.7/V1.4.8/V1.4.9 experimental lines are not used. Production sheet schemas and existing certificate/verification architecture are preserved. jsDelivr is intentionally not included.
