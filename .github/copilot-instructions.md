# Copilot Instructions — INFRA SalesPlay Tracker

## What This App Is
A mobile-first, gamified sales play tracker for the Progress INFRA Sales team. Reps log in with Microsoft (Azure AD) credentials, see assigned accounts, log engagement outcomes, earn points, and compete on leaderboards.

**Static web app** — plain HTML/JavaScript, no frameworks, no build tools, no npm. Runs entirely in the browser. All data lives in Progress SharePoint (JSON/CSV files).

**Live URLs:**
- Rep app: `https://orange-smoke-04675d40f.7.azurestaticapps.net/app.html`
- Admin: `https://orange-smoke-04675d40f.7.azurestaticapps.net/admin.html`
- Reports: `https://orange-smoke-04675d40f.7.azurestaticapps.net/reports.html`

---

## Key Files

| File | What it does |
|---|---|
| `app.html` | Rep-facing app — dashboard, play cards, engagement form, leaderboard |
| `admin.html` | Admin HQ — activity log, leaderboard, plays, accounts, corrections, coverage, health, broadcast, identity, config, snapshots |
| `reports.html` | Per-play reporting — 5 views, CSV export |
| `graph-api.js` | Core engine — ALL SharePoint reads/writes, scoring, leaderboard, engagement/activity logging |
| `msal-auth.js` | Authentication — MSAL login with mobile redirect flow, token management, eTag concurrency helpers |
| `config.js` | App constants — clientId, tenantId, SharePoint site URL, scopes |
| `form-engine.js` | Dynamic engagement form builder |
| `styles.css` | All visual styles |
| `ingest.html` | Play wizard — create/edit plays, upload accounts |
| `intake.html` | Play brief intake form for reps to submit play requests |
| `email-lookup.json` | Authoritative rep identity map (93 reps) |
| `staticwebapp.config.json` | Azure SWA routing + no-cache headers |

---

## Critical Rules — Read Before Changing Anything

### 1. Cache busting — ALWAYS increment the version number
All `<script>` tags use `?v=N` query strings (e.g., `graph-api.js?v=3`). **When you modify any JS file, increment its version number in every HTML file that loads it.** Otherwise browsers serve the old cached version and the fix never reaches users.

### 2. Never break eTag concurrency
`msal-auth.js` has `getFileWithETag()` and `putFileWithETag()`. These protect against lost engagement entries when multiple reps submit simultaneously. Any code that writes to SharePoint CSV files MUST use these helpers — never do a plain PUT.

### 3. SharePoint paths use `SP_ROOT`
Never hardcode SharePoint paths. The `SP_ROOT` constant in `config.js` points to the correct folder (`Chef SaaS Tracker/ChefSaaS`). The dev environment uses a separate folder (`ChefSaaS-Dev`). Always use `SP_ROOT`.

### 4. Mobile authentication
On mobile (iOS Safari, Teams, webviews), popup auth is blocked. `msal-auth.js` auto-detects mobile via `_useRedirectFlow()` and uses redirect-based login. Do not add popup-only auth flows.

### 5. No server-side code
This is a static web app. No Node.js, no serverless functions, no backend. All logic must be client-side JavaScript reading/writing SharePoint via Microsoft Graph API.

### 6. Rep identity is complex
Reps sign in under many email variants. Never match reps by a single email field. Use `_canonicalEmail()` in `graph-api.js` which handles aliases. See the multi-tier identity resolution system in the code comments.

---

## SharePoint Data Model

All data in `Chef SaaS Tracker/ChefSaaS/` on `https://progresssoftware.sharepoint.com/sites/INFRASalesApps`:

| File | Contents |
|---|---|
| `plays.json` | Play definitions |
| `assignments.json` | Account-to-rep assignments |
| `{play_id}_engagements.csv` | Per-play engagement logs |
| `activity_log.csv` | Login + engagement events |
| `email-lookup.json` | Rep identity map |
| `rep-identity.json` | Manual alias overrides |

---

## Deployment
Push to `main` → GitHub Actions → Azure Static Web Apps → live in ~1 minute. No manual steps.

The `dev/v2-upgrade` branch deploys to a separate preview environment with an orange "DEV ENVIRONMENT" banner.

---

## How to Fix Issues
1. Read the issue carefully — identify which file and function is responsible
2. Check `graph-api.js` first for anything data/SharePoint related
3. Check `msal-auth.js` for auth/login issues
4. Make the targeted fix — this codebase has no tests, so be precise and conservative
5. Increment the `?v=N` version on any JS file you change
6. Do not refactor unrelated code
