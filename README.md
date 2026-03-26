# Chef SaaS Motion Tracker

A mobile-first web app for tracking Chef SaaS sales motions. Reps log their account status from any phone browser — no app install required. Managers see a live admin dashboard with filters and exports.

---

## Quick Start

**You need:** A Progress Microsoft 365 account with SharePoint access and permission to register an Azure AD app (or ask IT to do Step 1 for you).

---

## Step 1: Register Azure AD App (~10 minutes)

1. Go to [portal.azure.com](https://portal.azure.com) and sign in with your Progress credentials.
2. Search **"App registrations"** in the top search bar → click **New registration**.
3. Fill in:
   - **Name:** `Chef SaaS Motion Tracker`
   - **Supported account types:** _Accounts in this organizational directory only_
   - **Redirect URI:** Select **Single-page application (SPA)** → enter the URL where you'll host the app (e.g. `https://progresssoftware.sharepoint.com/sites/SalesOps/SiteAssets/chef-saas-tracker/index.html` or `http://localhost` for testing)
4. Click **Register**.
5. On the **Overview** page, copy:
   - **Application (client) ID** → paste into `config.js` as `clientId`
   - **Directory (tenant) ID** → paste into `config.js` as `tenantId`
6. Left menu → **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** → search `Sites.ReadWrite.All` → **Add permissions**.
7. Click **Grant admin consent for [your org]** → **Yes**.
   > If you don't see that button, ask your IT/Azure admin to grant consent.

---

## Step 2: Fill in config.js

Open `config.js` in any text editor. Replace the placeholder values:

```javascript
const CONFIG = {
  tenantId: "YOUR_TENANT_ID_HERE",   // ← paste from Step 1
  clientId: "YOUR_CLIENT_ID_HERE",   // ← paste from Step 1
  sharepointSiteUrl: "YOUR_SHAREPOINT_SITE_URL_HERE",  // ← e.g. https://progresssoftware.sharepoint.com/sites/SalesOps
  ...
};
```

The `sharepointSiteUrl` is the root URL of your SharePoint site — just the site, not a library or page. Example: `https://progresssoftware.sharepoint.com/sites/SalesOps`

---

## Step 3: Create SharePoint Lists (one time)

1. Upload all app files to a folder in your SharePoint site (e.g. Site Assets → chef-saas-tracker).
2. Open **setup.html** in your browser.
3. Sign in with your Progress account.
4. Click **"Create All Lists & Seed Data"** — wait for all ✅ confirmations.

This creates 5 SharePoint lists and seeds the initial Chef SaaS play configuration:
- `sales_plays` — the plays reps can select
- `form_fields` — the questions for each play
- `field_options` — the dropdown choices
- `target_accounts` — your account list
- `motion_log` — where submissions are stored

It also adds 20 sample accounts so you can test the app right away.

---

## Step 4: Add Your Target Accounts

**Option A — Manual entry:**
1. Go to your SharePoint site → Site contents → `target_accounts` list.
2. Click **New** → fill in each row.

**Option B — Bulk import:**
1. Open the `target_accounts` list in SharePoint.
2. Click **Edit in grid view** and paste from Excel.

Required columns: `account_id` (unique), `account_name`, `rep_name`, `csm_name`, `region`, `segment`, `active_flag` (Yes/No).

---

## Step 5: Share with Reps

1. Share the SharePoint folder (or Azure Static Web App) with your rep/CSM security group.
2. Send reps the direct link to **index.html**.
3. Reps sign in with their Progress credentials — no IT setup needed on their end.

---

## Hosting Options

**Option A — SharePoint Document Library (simplest):**
- Upload all files to a SharePoint document library folder.
- Reps click the link to `index.html` — it opens directly in their browser.
- Add the full SharePoint file URL as a redirect URI in your Azure AD app registration.

**Option B — Azure Static Web Apps (cleaner URL):**
1. Create a Static Web App in the Azure portal.
2. Point it at a GitHub repo containing these files.
3. Add the Static Web App URL as a redirect URI in the Azure AD app registration.
4. Reps get a clean URL like `https://your-app.azurestaticapps.net`.

---

## Adding a New Sales Play (No Rebuild Needed)

The app is fully data-driven. To add a new play:

1. Go to SharePoint → `sales_plays` list → **New item**:
   - `play_id`: a unique slug (e.g. `chef-desktop`)
   - `play_name`: display name reps will see
   - `active_flag`: Yes
   - `display_order`: a number controlling sort order

2. Go to `form_fields` list → add one row per question:
   - `play_id`: must match the play_id above
   - `field_name`: internal name (no spaces, e.g. `status`)
   - `field_label`: what reps see (e.g. `Status`)
   - `field_type`: `dropdown`, `text`, `text_area`, or `date`
   - `is_required`: Yes/No
   - `display_order`: sort order
   - `active_flag`: Yes

3. Go to `field_options` list → add one row per dropdown choice:
   - `play_id`, `field_name` must match
   - `option_value`: the text reps see and that gets saved
   - `display_order`, `active_flag`

The new play appears automatically in the app. No code changes.

---

## Conditional Field Logic

In `form_fields`, the `show_when_field` and `show_when_value` columns control which fields are shown based on another field's value. Special token values for `show_when_value`:

| Token | Shows when `show_when_field` value is… |
|---|---|
| `PITCHED_OR_WON` | Any "Pitched - …" status, or "Closed Won" |
| `BLOCKER_STATUSES` | "Pitched - No Interest", "Pitched - Exploring", or "Closed Lost" |
| `NOT_NOT_PITCHED` | Anything except "Not Pitched" (and not blank) |
| `NEXT_STEP_STATUSES` | "Pitched - Exploring" or "Pitched - Active Deal" |
| _(any other value)_ | Exact string match against the controlling field |

---

## Files Reference

| File | Purpose |
|---|---|
| `index.html` | Rep-facing mobile form (share this link with reps) |
| `admin.html` | Manager dashboard — filters, table, exports |
| `setup.html` | One-time setup — creates SharePoint lists |
| `config.js` | **Edit this** — tenant ID, client ID, site URL |
| `msal-auth.js` | Azure AD sign-in (PKCE flow via MSAL.js) |
| `graph-api.js` | All Microsoft Graph / SharePoint API calls |
| `form-engine.js` | Dynamic form renderer |
| `export.js` | CSV and Excel export |
| `styles.css` | All styles |

---

## Troubleshooting

**"Need admin approval" error on sign-in:**
IT needs to grant admin consent for the `Sites.ReadWrite.All` permission in your Azure AD app registration.

**What IT does (30 seconds of work):**
> Entra ID → App registrations → Chef SaaS Motion Tracker → API permissions → **Grant admin consent for Progress Software**

**Email/Slack template to send IT:**

> Hi — I've registered an internal Azure AD app called **Chef SaaS Motion Tracker** (Client ID: `48db2acc-9dfb-4337-9abb-302c9dfb88fc`) for a sales tracking tool I'm building for the team. It uses SharePoint lists to store data and needs the `Sites.ReadWrite.All` delegated permission approved.
>
> Could someone with admin rights grant consent here:
> **Entra ID → App registrations → Chef SaaS Motion Tracker → API permissions → Grant admin consent for Progress Software**
>
> A few things worth knowing if there are any concerns:
> - **Delegated permission only** — the app acts as the logged-in user and can never do more than that user is already allowed to do in SharePoint. It cannot act autonomously or access data when no one is signed in.
> - **No data leaves Microsoft** — all data is stored in Progress SharePoint lists. Nothing goes to a third-party server.
> - **No new software or licenses** — this is a web page that runs in the browser using our existing M365 infrastructure (Azure AD + SharePoint).
> - **Reps log in with their normal Progress credentials** — no new accounts or passwords.
> - **Scope is limited to one SharePoint site** — in practice it only writes to the lists I set up on our Sales SharePoint site, not the entire org.

**Lists not found / setup errors:**
Re-run `setup.html`. Errors for lists that already exist are harmless (marked ⚠️). Look for ❌ errors with specific messages.

**Can't sign in / redirect error:**
The redirect URI in your Azure AD app registration must exactly match the URL you're opening the app from (including the path and trailing slash if any). Add the exact URL as a redirect URI of type SPA.

**Accounts not showing up:**
Check that the `target_accounts` list has items with `active_flag = Yes (true)`.

**Submitted data not appearing in admin:**
Open `admin.html`, click Apply Filters (with no filters set). If nothing appears, check that the `motion_log` list exists in SharePoint and that the account used for the rep form has write access.
