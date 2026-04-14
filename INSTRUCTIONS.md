# INFRA SalesPlay Tracker — User Instructions

**App URL:** https://orange-smoke-04675d40f.7.azurestaticapps.net

> Sign in with your Progress SSO credentials (same as Outlook / Teams).

---

## For Sales Reps

### Getting Started

1. Open the app URL on your phone or laptop.
2. Click **Sign In** and use your Progress SSO login (e.g. `langer@progress.com`).
3. You'll land on your **Dashboard** showing your assigned play(s), target counts, and your current score.

### Logging an Engagement

1. From the dashboard, click a **play card** to open your account list.
2. Find the account you spoke with and click it.
3. **Step 1 — Interaction:**
   - Select how you reached out: Call, Email, Text, or In-Person Meeting.
   - Select the **outcome**: Interested or Not Interested.
4. **Step 2 — Details:**
   - If Interested: select a reason and a next step (required).
   - If Not Interested: select a reason (required).
   - Add notes and timing if applicable.
5. Tap **Submit ✓**.
6. You'll see a confirmation screen with your points earned and a confetti celebration if you're on a streak!
7. You're automatically returned to your account list after a few seconds.

### Dashboard Summary

Your play card shows:
| Metric | Meaning |
|--------|---------|
| **Targets** | Accounts assigned to you for this play |
| **Engaged** | Accounts you've logged any outcome for |
| **Interested** | Accounts whose latest outcome is Interested |
| **Progress %** | Engaged ÷ Targets |

### Scoring

| Action | Points |
|--------|--------|
| Log Not Interested | 15 pts |
| Log Interested | 30 pts |
| Interested + Next Step booked | 45 pts |

**Levels:**
- 🌱 Rookie — 0 pts
- 🎯 Pitcher — 50 pts
- ⭐ Pro — 150 pts
- 🥇 Champion — 300 pts
- 🏆 Legend — 500 pts

> Top 3 reps win a special prize!

### Tips

- You can log the same account multiple times. The **latest** entry is what counts for your dashboard status.
- Use the filter tabs (All / Not Started / Interested / Not Interested) to focus your outreach.
- Bookmark the app URL on your phone's home screen for quick access.

---

## For Admins

Admin access is available at:
**https://orange-smoke-04675d40f.7.azurestaticapps.net/admin.html**

> Admin access requires an owner or manager role. Contact Phil Langer to request access.

---

### Creating a New Play

1. Go to **admin.html** → **Plays** tab.
2. Click **+ New Play** to open the play creation wizard.

**Step 1 — Basics:**
- Enter a play name, goal (Qualify In or Qualify Out), ICP, target persona, and elevator pitch.
- Click **Continue to Outcomes →**

**Step 2 — Outcomes:**
- Configure the outcomes, reasons, and next steps reps will see.
- Click **Continue to Questions →** or **💾 Save Config Only** to save without uploading accounts yet.

**Step 3 — Questions:**
- Enable/disable standard boilerplate fields reps see on the log form.
- Add custom questions if needed.
- Click **Continue to Accounts →**

**Step 4 — Accounts:**
- Upload a CSV or Excel file of target accounts.
- Each row should include: `account_id`, `account_name`, `rep_sso_login` (rep's login email), `rep_name`.
- Choose **Launch Play** (new) or **Add Targets** / **Replace Targets** (existing play).
- Click **Upload & Launch** to go live.

---

### Editing an Existing Play

From the **Plays** tab:
- Click **✏️ Edit Play** to update the play config (name, goal, ICP, outcomes) without touching accounts.
- Click **➕ Add Targets** to upload additional accounts to an existing play without losing current assignments.

---

### Account Upload Format

Use the template file `account-upload-template.xlsx` in the repo.

Required columns:
| Column | Description |
|--------|-------------|
| `account_id` | Unique account ID (e.g. Salesforce Account ID) |
| `account_name` | Account display name |
| `rep_sso_login` | Rep's Progress login email (e.g. `klesel@progress.com`) |
| `rep_name` | Rep's display name |

Optional columns: `arr`, `region`, `segment`, `target_persona`, `icp_notes`

> **Email matching:** The system automatically handles email alias mismatches (e.g. `klesel@` vs `micah.klesel@`). Upload with either format.

---

### Managing Rep Identity (Owner Only)

The **Identity** tab in admin.html lets you upload `rep-identity.json` to SharePoint, which enables full alias-based email matching for all 93 reps.

1. Go to **admin.html** → **Identity** tab.
2. Click **Choose File** → select `rep-identity.json` from `OneDrive/Copilot Apps/SalesPlay Tracker/Data Dump/`.
3. Click **Upload to SharePoint**.

This only needs to be done once. After upload, reps are matched automatically regardless of which email format was used in the account upload.

---

### Viewing Results

| Tab | What it shows |
|-----|--------------|
| **Leaderboard** | All reps ranked by total points, with engaged and interested counts |
| **Accounts** | All assigned accounts with rep, play, and status |
| **Plays** | All plays with target counts and actions |
| **Activity Log** | Every rep login and engagement submission |
| **Health** | System file status (SharePoint data files) |
| **Coverage** | Rep coverage across plays |

---

### Backup & Reset (Danger Zone)

In the **Plays** tab, scroll to the bottom to find the Danger Zone:

- **Backup + Reset:** Downloads a full backup ZIP, then clears all engagement data. Type `RESET` to confirm.
- **Undo Last Reset:** Restores data from the last backup. Type `RESTORE` to confirm.

> ⚠️ These actions cannot be undone beyond the most recent backup.

---

### Exporting Data

Click **⬇️ Export All** (top right of admin.html) to download a full Excel export containing:
- All plays and their configuration
- All account assignments
- All engagement logs per play
- Leaderboard summary

---

## Troubleshooting

| Issue | Solution |
|-------|---------|
| App shows "Not Authorized" | Your email isn't in the admin list. Contact Phil Langer. |
| Rep can't see their accounts | Check that `rep_sso_login` in the upload matches their login email. Upload the rep-identity.json via the Identity tab to enable alias matching. |
| Dashboard counts not updating | Hard refresh the page (Ctrl+Shift+R / Cmd+Shift+R). |
| "Loading…" spinner hangs | Check your network connection. SharePoint calls may time out on slow connections. |
| Submission fails or times out | Try again. If it persists, check the Activity Log in admin to see if the entry was saved. |

---

## Key URLs

| Page | URL |
|------|-----|
| Rep App | https://orange-smoke-04675d40f.7.azurestaticapps.net |
| Admin Console | https://orange-smoke-04675d40f.7.azurestaticapps.net/admin.html |
| Play Wizard | https://orange-smoke-04675d40f.7.azurestaticapps.net/ingest.html |
| Security & Privacy | https://orange-smoke-04675d40f.7.azurestaticapps.net/security.html |

---

*INFRA SalesPlay Tracker — Progress Software Proprietary & Confidential. Progress employees only.*
