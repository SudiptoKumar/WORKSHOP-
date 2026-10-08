# WORKSHOP Automation V2

Complete workshop management system for Google Workspace — built by **Sudipto Kumar**.

Handles the full workshop lifecycle: **creation → registration → reminders → attendance → certificates → verification**.

## 📁 Project Structure

```
workshop-automation-v2/
├── README.md                    ← You are here
├── Workshop_Automation.gs       ← Main Apps Script (all logic)
├── Verify.html                  ← Public certificate verification page
└── email/
    ├── templates.js             ← Email templates (JavaScript module)
    ├── registration.html        ← Registration email (standalone)
    ├── reminder.html            ← Reminder email (standalone)
    ├── certificate.html         ← Certificate email (standalone)
    └── custom.html              ← Custom broadcast email (standalone)
```

## 🚀 Setup Instructions

### Prerequisites
- Google account
- Google Sheet (for data storage)
- Google Slides (for certificate generation)
- Google Forms (auto-created for registrations)

### Step 1: Create the Apps Script Project
1. Go to [script.google.com](https://script.google.com)
2. Click **New Project**
3. Delete the default `Code.gs` content
4. Copy the entire contents of `Workshop_Automation.gs` into the editor
5. Rename the file to `Workshop_Automation` (optional)

### Step 2: Add the Verification Page
1. In the Apps Script editor, click **+** next to Files → **HTML**
2. Name it `Verify`
3. Copy the entire contents of `Verify.html` into it

### Step 3: Enable Required Services
1. In Apps Script editor, click **Services** (+ icon in left sidebar)
2. Add **Google Slides API** (Advanced Service)
3. Add **Drive API** if needed for file operations

### Step 4: Create the Google Sheet
1. Create a new Google Sheet
2. Run the `setupSystem` function from the Apps Script editor
3. This creates all required sheets:
   - `Control Center` — Workshop creation dashboard
   - `Self Test` — Testing interface
   - `Custom Email` — Broadcast email tool
   - `Workshops` — Workshop database
   - `Participants` — Registration database
   - `Certificates` — Certificate records
   - `Test Runs` — Test execution logs
   - `Settings` — System configuration

### Step 5: Configure Settings
1. Open the `Control Center` sheet
2. Fill in your organization details:
   - Organization name
   - Logo URL
   - Social media links (Facebook, Instagram, LinkedIn)
   - Support email
   - Verification page URL

### Step 6: Deploy as Web App (for verification page)
1. Click **Deploy** → **New Deployment**
2. Select **Web App**
3. Set **Execute as:** Me
4. Set **Who has access:** Anyone
5. Copy the Web App URL — this is your certificate verification link

### Step 7: Install Triggers
Run `reinstallScheduler` from the function dropdown to set up the 1-minute scheduler trigger that processes:
- Scheduled reminders
- Scheduled certificates
- Test run steps

## 🔄 How It Works

### System Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                    CONTROL CENTER (Sheet)                     │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌──────────────┐  │
│  │ Workshop │  │ Reminder │  │Certificate│  │   CREATE     │  │
│  │ Details  │  │ Schedule │  │ Schedule  │  │  WORKSHOP    │  │
│  └──────────┘  └──────────┘  └──────────┘  └──────────────┘  │
└──────────────────────────┬──────────────────────────────────┘
                           │ Creates
                           ▼
┌─────────────────────────────────────────────────────────────┐
│                      WORKSHOPS Sheet                         │
│  Event ID │ Name │ Date │ Meeting Link │ Reminder │ Cert    │
└──────────────────────────┬──────────────────────────────────┘
                           │ Auto-creates
                           ▼
┌─────────────────────────────────────────────────────────────┐
│                   Google Form (Registration)                 │
└──────────────────────────┬──────────────────────────────────┘
                           │ Participant submits
                           ▼
┌─────────────────────────────────────────────────────────────┐
│                    PARTICIPANTS Sheet                        │
└──────────────────────────┬──────────────────────────────────┘
                           │
              ┌────────────┼────────────┐
              ▼            ▼            ▼
     ┌────────────┐ ┌────────────┐ ┌────────────┐
     │Registration│ │  Reminder  │ │Certificate │
     │   Email    │ │   Email    │ │   Email    │
     │(immediate) │ │(scheduled) │ │(scheduled) │
     └────────────┘ └────────────┘ └────────────┘
                                          │
                                          ▼
                                 ┌────────────────┐
                                 │ Google Slides  │
                                 │ → PDF Export   │
                                 │ → Drive Upload │
                                 └────────────────┘
                                          │
                                          ▼
                                 ┌────────────────┐
                                 │ Verify.html    │
                                 │ (Public Page)  │
                                 └────────────────┘
```

### Workflow Steps

1. **Create Workshop** — Fill in Control Center → Tick CREATE WORKSHOP
   - System creates workshop record
   - Auto-generates Google Form for registration
   - Schedules reminder and certificate emails

2. **Registration** — Participant fills Google Form
   - Immediate confirmation email sent
   - Participant added to Participants sheet
   - If reminder is OFF, registration email includes Join Workshop button

3. **Reminder** — Sent at scheduled time (before workshop)
   - Contains Join Workshop button with meeting link
   - Meeting details and joining checklist

4. **Certificate** — Generated and sent at scheduled time
   - Google Slides template → personalized → PDF export
   - PDF attached to email
   - Verification URL included
   - QR code for quick verification

5. **Verification** — Anyone can verify at the public URL
   - `Verify.html` web app
   - Enter certificate ID → shows certificate details

## 🏆 Certificate System

### How Certificates Work

```
Certificate Master (Google Slides)
         │
         │ 1. Copy template
         ▼
┌─────────────────────┐
│ Personalized Slide  │
│ - Participant name  │
│ - Workshop name     │
│ - Date              │
│ - Certificate ID    │
│ - Signatory         │
└─────────────────────┘
         │
         │ 2. Export as PDF
         ▼
┌─────────────────────┐
│ PDF in Google Drive │
└─────────────────────┘
         │
         │ 3. Attach to email
         ▼
┌─────────────────────┐
│ Certificate Email   │
│ - PDF attached      │
│ - Verify button     │
│ - QR code           │
└─────────────────────┘
```

### Changing the Certificate Design

1. Open your **Certificate Master** Google Slides presentation
2. Edit the design (background, fonts, layout, images)
3. Keep these **placeholder texts** (they get replaced automatically):
   - `{{participant_name}}` — Participant's full name
   - `{{workshop_name}}` — Workshop title
   - `{{event_date}}` — Workshop date
   - `{{certificate_type}}` — e.g., "Certificate of Participation"
   - `{{certificate_id}}` — Unique certificate ID
   - `{{signatory_name}}` — e.g., "Abdullah Muhsin"
   - `{{signatory_title}}` — e.g., "Secretary"

4. To use a different master for a workshop:
   - In Control Center → Certificate Master dropdown
   - Select your new master
   - Or set as default in Settings

### Updating Certificate Settings

| Setting | Where | Description |
|---------|-------|-------------|
| Certificate Type | Control Center B31 | e.g., "Certificate of Participation" |
| Signatory | Control Center B32 | Name and title |
| Master Template | Control Center F32 | Which Slides template to use |
| Send Certificate | Control Center H27 | ☑ = send, ☐ = don't send |
| Certificate Date/Time | Control Center C27/F27 | When to generate & send |

### Certificate Timing
- Can be **before, during, or after** the workshop
- No restrictions — set any date/time
- Scheduler checks every 1 minute for precise delivery

## 📧 Email Templates

All templates support **dark mode** (auto-adapts to reader's device theme).

| File | Description |
|------|-------------|
| `email/templates.js` | JavaScript module with all 4 templates + `renderEmailTemplate()` function |
| `email/registration.html` | Standalone registration email |
| `email/reminder.html` | Standalone reminder email |
| `email/certificate.html` | Standalone certificate email |
| `email/custom.html` | Standalone custom broadcast email |

### Using Templates in JavaScript

```javascript
import { renderEmailTemplate } from './email/templates.js';

const html = renderEmailTemplate('REGISTRATION', {
  participant_name: 'John Doe',
  workshop_name: 'My Workshop',
  event_date: '08 Oct 2026',
  // ... see placeholders below
});
```

### Placeholders

**Common:** `{{organization_name}}`, `{{logo_url}}`, `{{facebook_url}}`, `{{instagram_url}}`, `{{linkedin_url}}`, `{{support_email}}`, `{{current_year}}`

**Registration/Reminder:** `{{participant_name}}`, `{{workshop_name}}`, `{{event_date}}`, `{{event_time}}`, `{{timezone}}`, `{{platform}}`, `{{meeting_link}}`, `{{meeting_id}}`

**Certificate:** `{{participant_name}}`, `{{workshop_name}}`, `{{certificate_type}}`, `{{certificate_id}}`, `{{event_date}}`, `{{season}}`, `{{verification_url}}`

## 🧪 Self Test

Test the full flow without affecting production data:

1. Go to the **Self Test** sheet
2. Select a workshop
3. Enter test recipient email
4. Enable steps to test (Registration / Reminder / Certificate / Attendance)
5. Set timing (Immediate or specific date/time)
6. Tick **RUN SELF TEST**
7. Check the test email inbox

Test runs are logged in the **Test Runs** sheet.

## 📬 Custom Email

Send broadcast emails to workshop participants:

1. Go to **Custom Email** sheet
2. Pick a workshop
3. Choose recipients (All / Attended / Not Attended)
4. Write subject and message
5. Preview → Send

## ⚙️ Key Functions

| Function | Description |
|----------|-------------|
| `setupSystem` | Initial sheet setup |
| `createWorkshop` | Creates workshop from Control Center |
| `startTestRun` | Starts a self-test (called from sheet) |
| `processScheduler` | Processes due reminders/certificates (runs every minute) |
| `reinstallScheduler` | Reinstalls the scheduler trigger |
| `rebuildAllSheetsUI` | Rebuilds Control Center/Self Test/Custom Email UI |
| `verifyCertificate` | Verifies a certificate ID |
| `renderEmailTemplate_` | Renders email HTML with data |

## 🔧 Troubleshooting

| Issue | Solution |
|-------|----------|
| Emails arrive late | Scheduler runs every 1 min; check triggers in Apps Script |
| "Workshop not found" | Ensure workshop exists in Workshops sheet |
| Certificate not generating | Check Slides API is enabled; verify master template ID |
| Dark mode not working | Ensure reader's email client supports `prefers-color-scheme` |
| Form not created | Check Drive permissions; re-run `createWorkshop` |

## 📄 License

MIT License — feel free to use and modify.

## 👥 Credits

Built by **Sudipto Kumar**
