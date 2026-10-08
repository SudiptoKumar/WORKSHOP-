# WORKSHOP Automation V2.1.0

Complete workshop management system for Google Workspace — built by **Finance Club PSTU**.

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

### Step 2: Add the Verification Page
1. In the Apps Script editor, click **+** next to Files → **HTML**
2. Name it `Verify`
3. Copy the entire contents of `Verify.html` into it

### Step 3: Enable Required Services
1. In Apps Script editor, click **Services** (+ icon in left sidebar)
2. Add **Google Slides API** (Advanced Service)

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
Run `reinstallScheduler` from the function dropdown to set up the 1-minute scheduler trigger.

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
└──────────────────────────┬──────────────────────────────────┘
                           │ Auto-creates Google Form
                           ▼
┌─────────────────────────────────────────────────────────────┐
│              Participant fills Registration Form             │
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
                                 │ → PDF → Drive  │
                                 └────────────────┘
```

### Drive Folder Structure (per workshop)
```
{EventID} - {Workshop Name}/
├── Registration/       ← Google Form + responses
├── Attendance/         ← Attendance form (if enabled)
└── Certificates/       ← Generated PDFs
    └── Test Runs/      ← Test certificates
```

## 🏆 Certificate System

### Changing the Certificate Design
1. Open your **Certificate Master** Google Slides presentation
2. Edit the design (background, fonts, layout, images)
3. Keep these **placeholders** (auto-replaced):
   - `{{participant_name}}`
   - `{{workshop_name}}`
   - `{{event_date}}`
   - `{{certificate_type}}`
   - `{{certificate_id}}`
   - `{{signatory_name}}`
   - `{{signatory_title}}`

### Certificate Settings (Control Center)
| Setting | Description |
|---------|-------------|
| Certificate type | e.g., "Certificate of Participation" |
| Signatory | Name and title |
| Master Template | Which Slides template to use |
| Send Certificate | ☑ = send, ☐ = don't send |
| Date/Time | Can be before, during, or after workshop |

## 📧 Email Templates

All templates support **dark mode** (auto-adapts to reader's device).

| File | Description |
|------|-------------|
| `email/templates.js` | JS module with all templates + `renderEmailTemplate()` |
| `email/*.html` | Standalone HTML versions |

### Placeholders
**Common:** `{{organization_name}}`, `{{logo_url}}`, `{{facebook_url}}`, `{{instagram_url}}`, `{{linkedin_url}}`, `{{support_email}}`, `{{current_year}}`

**Registration/Reminder:** `{{participant_name}}`, `{{workshop_name}}`, `{{event_date}}`, `{{event_time}}`, `{{timezone}}`, `{{platform}}`, `{{meeting_link}}`, `{{meeting_id}}`

**Certificate:** `{{participant_name}}`, `{{workshop_name}}`, `{{certificate_type}}`, `{{certificate_id}}`, `{{event_date}}`, `{{season}}`, `{{verification_url}}`

## 🧪 Self Test

Fixed timing — no configuration needed:
- **0 min:** Registration email (immediate)
- **+1 min:** Reminder email
- **+3 min:** Certificate email
- **Attendance** (optional): When ON, certificate sends 1 min after form fill

## 📬 Custom Email

Send broadcasts: pick workshop → choose recipients → write message → preview → send.

## ⚙️ Key Functions

| Function | Description |
|----------|-------------|
| `setupSystem` | Initial sheet setup |
| `createWorkshop` | Creates workshop from Control Center |
| `reinstallScheduler` | Installs 1-minute scheduler trigger |
| `rebuildAllSheetsUI` | Rebuilds sheet interfaces |
| `verifyCertificate` | Verifies certificate ID |

## 📄 License

MIT License.

## 👥 Credits

Built by **Finance Club PSTU** — Learn · Grow · Connect · Create Impact

---
*V2.1.0 — Cleaned codebase, simplified Self Test, dark mode emails, precise 1-minute scheduling.*
