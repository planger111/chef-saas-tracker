# INFRA SalesPlay Tracker

A mobile-first web app for Progress INFRA Sales. Reps log account engagement outcomes from any phone — no app install required. Leadership sees a live dashboard with leaderboard and conversion data.

---

## Current Setup (already done)

| Item | Value |
|------|-------|
| Azure AD App | `INFRA SalesPlay Tracker` (Client ID: `48db2acc-9dfb-4337-9abb-302c9dfb88fc`) |
| SharePoint site | `https://progresssoftware.sharepoint.com/sites/INFRASalesLeadership` |
| Graph permission | `Sites.ReadWrite.All` (admin consent granted) |
| GitHub repo | `planger111/chef-saas-tracker` (private) |
| Demo (public) | https://planger111.github.io/chef-saas-demo/ |

---

## Hosting: Azure Static Web Apps (IT Request)

This is the right long-term hosting solution. The GitHub Actions workflow is already set up in `.github/workflows/azure-static-web-apps.yml` — IT just needs to create the Azure resource and paste one token into GitHub.

### What IT needs to do (~10 minutes total)

#### Step 1 — Create the Static Web App (5 min)

1. Sign in to [portal.azure.com](https://portal.azure.com)
2. Search **"Static Web Apps"** → **Create**
3. Fill in:
   - **Subscription:** Progress Software (whichever is appropriate)
   - **Resource Group:** create new or use existing Sales/INFRA group
   - **Name:** `INFRASalesPlayTracker`
   - **Plan type:** Free
   - **Region:** East US 2 (or closest)
   - **Deployment source:** GitHub
4. Click **Sign in with GitHub** → authorize → select:
   - **Organization:** `planger111`
   - **Repository:** `chef-saas-tracker`
   - **Branch:** `main`
5. **Build details:**
   - Build preset: **Custom**
   - App location: `/`
   - Output location: _(leave blank)_
6. Click **Review + Create** → **Create**
7. Copy the URL shown (e.g. `https://infrasalesplaytracker.azurestaticapps.net`)

> GitHub Actions will automatically deploy on every push to `main`. No manual deploys needed.

#### Step 2 — Add redirect URI to Azure AD App (2 min)

1. In Azure portal → **Azure Active Directory** → **App registrations**
2. Find **INFRA SalesPlay Tracker** (Client ID: `48db2acc-9dfb-4337-9abb-302c9dfb88fc`)
3. **Authentication** → under **Single-page application** → **Add URI**
4. Add the Static Web App URL: `https://infrasalesplaytracker.azurestaticapps.net`
5. Click **Save**

---

## After IT sets up hosting — Phil's steps (~5 min)

1. Go to `https://infrasalesplaytracker.azurestaticapps.net/setup.html`
2. Sign in with your Progress account
3. Click **Initialize Data Files** — creates the `ChefSaaS/` folder in SharePoint
4. Upload real account list to `ChefSaaS/assignments.json`
5. Share `https://infrasalesplaytracker.azurestaticapps.net/app.html` with reps

---

## IT Email Template

> Hi — I'm running a pilot sales tracking tool for the INFRA Sales team and need help hosting it on Azure Static Web Apps.
>
> **What I need:**
> 1. Create an Azure Static Web App resource connected to my private GitHub repo (`planger111/chef-saas-tracker`, branch `main`)
> 2. Add the resulting URL as a redirect URI on our existing Azure AD app registration (Client ID: `48db2acc-9dfb-4337-9abb-302c9dfb88fc`)
>
> The GitHub Actions deployment workflow is already in the repo — the app will deploy automatically once the resource is created. No build step needed, it's plain HTML/JS.
>
> The app uses our existing M365 SSO (Progress credentials) and stores data in SharePoint. No new permissions or licenses are needed — admin consent for `Sites.ReadWrite.All` was already granted.

---

## Files Reference

| File | Purpose |
|------|---------|
| `index.html` | Pilot landing page (links to demo + real app) |
| `app.html` | Real rep app — Azure AD SSO, reads from SharePoint |
| `admin.html` | Leadership dashboard |
| `setup.html` | One-time data initialization (admin only) |
| `DEMO-index.html` | Self-contained demo — no login, localStorage only |
| `DEMO-admin.html` | Demo leadership view |
| `config.js` | Azure AD + SharePoint config (already filled in) |
| `msal-auth.js` | Azure AD sign-in (popup flow) |
| `msal-browser.js` | MSAL library (local copy, CDN was broken) |
| `graph-api.js` | All Graph API / SharePoint file operations |

---

## Data Architecture

Engagement data lives in SharePoint Documents → `ChefSaaS/` folder:

| File | Contents |
|------|---------|
| `assignments.json` | Rep → account assignments (upload to assign accounts) |
| `response-options.json` | Dropdown options for the engagement form |
| `chef-saas_engagements.csv` | All rep engagements (open in Excel for reporting) |

One CSV per play. Admin downloads and opens in Excel — no special tooling needed.


---

## Quick Start

**You need:** A Progress Microsoft 365 account with SharePoint access and permission to register an Azure AD app (or ask IT to do Step 1 for you).

---

## Step 1: Register Azure AD App (~10 minutes)

1. Go to [portal.azure.com](https://portal.azure.com) and sign in with your Progress credentials.
2. Search **"App registrations"** in the top search bar → click **New registration**.
3. Fill in:
   - **Name:** `INFRA SalesPlay Tracker`
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

This creates 5 SharePoint lists and seeds the initial INFRA SalesPlay configuration:
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
> Entra ID → App registrations → INFRA SalesPlay Tracker → API permissions → **Grant admin consent for Progress Software**

**Email/Slack template to send IT:**

> Hi — I've registered an internal Azure AD app called **INFRA SalesPlay Tracker** (Client ID: `48db2acc-9dfb-4337-9abb-302c9dfb88fc`) for a sales tracking tool I'm building for the team. It uses SharePoint lists to store data and needs the `Sites.ReadWrite.All` delegated permission approved.
>
> Could someone with admin rights grant consent here:
> **Entra ID → App registrations → INFRA SalesPlay Tracker → API permissions → Grant admin consent for Progress Software**
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
