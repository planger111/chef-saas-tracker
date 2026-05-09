# TEAMS-DEV-README — INFRA SalesPlay Tracker Teams Dev App

> ⚠️ **DEV ONLY** — This is a development/sideload version for testing the Teams experience.
> Do not share with reps. Do not promote to production without review.

---

## 1. What This Is

A Microsoft Teams personal tab version of the INFRA SalesPlay Tracker **rep experience**.

- Rep-only. No admin controls. No setup. No exports. No backup/reset.
- Reuses the existing shared JS stack (MSAL auth, Graph API, form engine).
- Isolated: `teams-dev.html` is a standalone file. Deleting it removes Teams support entirely.
- The existing `app.html` and `admin.html` are completely untouched.

Architecture:
```
Shared JS (msal-auth.js, graph-api.js, form-engine.js, config.js)
  ├── Web Rep View:   app.html          ← unchanged
  ├── Web Admin View: admin.html        ← unchanged
  └── Teams Rep View: teams-dev.html   ← NEW (this file)
```

---

## 2. Files Added

| File | What it is |
|---|---|
| `teams-dev.html` | Teams rep tab — adapted copy of `app.html` |
| `teams-package-dev/manifest.json` | Teams app manifest v1.17 (personal tab) |
| `teams-package-dev/color.png` | 192×192 blue placeholder icon |
| `teams-package-dev/outline.png` | 32×32 white placeholder icon |
| `TEAMS-DEV-README.md` | This file |

---

## 3. Files Changed

| File | Change | Scope |
|---|---|---|
| `staticwebapp.config.json` | Added one route entry at top of `routes[]` for `/teams-dev.html` | Only affects `/teams-dev.html` — all other routes unchanged |

**The route override:**
- Removes `X-Frame-Options: DENY` for this file only (Teams iframes the tab)
- Sets a Teams-specific CSP with `frame-ancestors` for Teams domains
- Allows `https://res.cdn.office.net` in `script-src` for the Teams JS SDK
- All other routes are unaffected; `globalHeaders` unchanged

**`app.html`, `admin.html`, `setup.html`, `msal-auth.js`, `graph-api.js`, `config.js` — NOT touched.**

---

## 4. What Was Changed in teams-dev.html vs app.html

Four tiny patches, all additive and Teams-specific:

1. **Title**: `INFRA SalesPlay Tracker` → `[DEV] Sales Play Tracker — Teams`
2. **DEV banner text**: Updated to say `TEAMS DEV` so it's obvious what environment you're in
3. **Before scripts**: Added `window.APP_HOST = 'teams'` flag + Teams JS SDK CDN
4. **Bookmark banner**: Added `if (window.APP_HOST === 'teams') return;` early exit (bookmark prompt makes no sense in Teams)
5. **After app load**: Added `microsoftTeams.app.initialize()` call so Teams knows the tab finished loading

---

## 5. How to Run / Test Locally

**In a browser (no Teams):**
```
https://orange-smoke-04675d40f.7.azurestaticapps.net/teams-dev.html
```
It runs just like `app.html`. Auth redirect flow works the same way. The orange "TEAMS DEV" banner confirms you're on the right file.

**In Teams Desktop / Teams Web:**
Sideload the package (see section 9). Teams will iframe `teams-dev.html` from the hosted URL.

---

## 6. How to Host

The app is already hosted on Azure Static Web Apps. Deployment is automatic:

1. Push commits to the `dev/v2-upgrade` branch → triggers a **preview deploy** (new URL per PR)
2. Merge to `main` → triggers the **production deploy** to `https://orange-smoke-04675d40f.7.azurestaticapps.net/`

The Teams manifest points to the **production URL**. For sideload testing this weekend:
- If using `main`/production deploy → manifest is ready as-is
- If using a preview deploy URL → update `contentUrl` and `websiteUrl` in `manifest.json` AND add that preview URL to Azure AD redirect URIs

---

## 7. Exact URL the Teams Tab Expects

```
https://orange-smoke-04675d40f.7.azurestaticapps.net/teams-dev.html
```

This is set in `teams-package-dev/manifest.json` → `staticTabs[0].contentUrl`.

---

## 8. Azure AD Redirect URI

**No new redirect URI is needed for the production URL.**

The existing registered redirect URI is:
```
https://orange-smoke-04675d40f.7.azurestaticapps.net
```

MSAL uses `window.location.origin` as the redirect URI, which resolves to the origin (not the full path). Since `teams-dev.html` is on the same origin as `app.html`, the same registration covers it.

**Exception — preview deploy URLs:**
If you're testing from a preview deploy URL like `https://orange-smoke-04675d40f-5.eastus2.7.azurestaticapps.net`, you MUST add that full origin to the Azure AD redirect URIs:

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Find **Chef SaaS Motion Tracker** (Client ID: `48db2acc-9dfb-4337-9abb-302c9dfb88fc`)
3. Authentication → Redirect URIs → Add URI:
   ```
   https://orange-smoke-04675d40f-5.eastus2.7.azurestaticapps.net
   ```
4. Save

**Teams-specific notes:**
- No Teams SSO / `teamsSSO` is configured — the app uses standard MSAL redirect auth
- Teams will redirect to the Azure AD login page when the tab first loads (normal behavior)
- After first login, MSAL's sessionStorage cache means re-auth is usually silent
- If reps see a blank page or auth loop in Teams: check the redirect URI is registered

---

## 9. How to Zip the Teams Package

From the repo root:
```bash
cd teams-package-dev
zip ../SalesPlayTracker-DEV.zip manifest.json color.png outline.png
cd ..
```

Or from any directory:
```bash
zip -j SalesPlayTracker-DEV.zip teams-package-dev/manifest.json teams-package-dev/color.png teams-package-dev/outline.png
```

The zip file must contain the files at the **root** of the archive (not in a subfolder). The `-j` flag strips the directory prefix.

Verify the zip:
```bash
unzip -l SalesPlayTracker-DEV.zip
# Should show: manifest.json, color.png, outline.png (no subdirectory prefix)
```

---

## 10. How to Sideload in Teams

**Teams Desktop or Teams Web:**

1. Open Microsoft Teams
2. Go to **Apps** (left sidebar)
3. Click **Manage your apps** (bottom left of Apps panel)
4. Click **Upload an app** → **Upload a custom app**
5. Select `SalesPlayTracker-DEV.zip`
6. Click **Add** in the preview dialog

The app will appear as "SalesPlay DEV" in your personal apps sidebar.

**If "Upload a custom app" is not visible:**
Your tenant may have restricted custom app uploads. To enable:
- Teams Admin Center → Teams apps → Setup policies → Global → Allow uploading custom apps: ON
- Or ask your Teams admin to enable it for your account

---

## 11. Testing Checklist

### Existing app — confirm unchanged
- [ ] `app.html` loads in browser — same behavior as before
- [ ] `admin.html` loads in browser — same behavior as before
- [ ] `setup.html` loads in browser — same behavior as before
- [ ] Sign in / sign out works in browser rep app
- [ ] Engagement submission works in browser rep app

### teams-dev.html — browser smoke test (before sideloading)
- [ ] `https://orange-smoke-04675d40f.7.azurestaticapps.net/teams-dev.html` loads in browser
- [ ] Orange "TEAMS DEV" banner is visible
- [ ] Page title shows `[DEV] Sales Play Tracker — Teams` (in browser tab)
- [ ] Sign in with Microsoft works (redirect flow)
- [ ] Assigned accounts load
- [ ] Plays load
- [ ] Engagement form opens and submits
- [ ] Leaderboard loads
- [ ] No admin controls visible (no admin tab, no corrections, no coverage, no broadcast)
- [ ] No setup / export / backup / reset controls visible
- [ ] Bookmark banner does NOT appear (suppressed in Teams mode)

### Teams desktop / web
- [ ] Sideload package uploads successfully
- [ ] App appears in personal apps sidebar
- [ ] Tab opens and loads `teams-dev.html`
- [ ] Sign in prompt appears (expected on first open)
- [ ] Authentication completes (redirect back to tab)
- [ ] Plays and accounts load
- [ ] Form submission works and logs correctly
- [ ] DEV data folder (`ChefSaaS-Dev`) confirmed — not hitting production data

### Teams mobile (stretch)
- [ ] App visible in Teams mobile after sideload on desktop
- [ ] Tab loads (may take longer due to redirect auth on mobile Teams)
- [ ] Authentication completes
- [ ] Basic functionality works

---

## 12. Known Risks / Issues

| Risk | Severity | Notes |
|---|---|---|
| MSAL redirect auth in Teams | Low | Redirect flow already works universally in this app; tested on mobile Safari. Teams Chromium should handle it fine. |
| `X-Frame-Options: DENY` from globalHeaders | Low | Teams uses modern Chromium; `frame-ancestors` CSP takes precedence over `X-Frame-Options` per spec. Works in Teams Desktop, Teams Web, modern browsers. |
| Teams mobile auth | Medium | Redirect flow in Teams mobile WebView is untested. May show auth prompt every session if sessionStorage is isolated per tab. |
| Preview deploy URL not in Azure AD | Medium | If sideloading from a preview URL, auth will fail until that origin is added to Azure AD redirect URIs. See section 8. |
| Teams SDK CDN availability | Very Low | SDK loads from `https://res.cdn.office.net`. If unavailable, `microsoftTeams` is undefined — the `if (window.microsoftTeams)` guard handles this gracefully. |
| Icon quality | Info | Icons are solid-color placeholders. Fine for sideload/testing. Replace before any org-wide deploy. |
| Custom app upload permission | Info | Requires org to allow custom app uploads. Teams admin setting. |

---

## 13. Rollback Instructions

To completely remove Teams support:

```bash
# From repo root
rm teams-dev.html
rm -rf teams-package-dev/
rm TEAMS-DEV-README.md
rm SalesPlayTracker-DEV.zip  # if you created the zip
```

Then revert the one line in `staticwebapp.config.json`:

```bash
git diff staticwebapp.config.json   # review the change
git checkout staticwebapp.config.json  # revert it
```

Or manually: open `staticwebapp.config.json`, remove the `/teams-dev.html` route entry (the first entry in the `routes` array).

**Impact on existing app:** Zero. `app.html`, `admin.html`, and all shared JS are unchanged.

---

## 14. Before Promoting to Production

If you decide to make this a real production Teams app (not just dev):

1. **Replace icons** — create proper 192×192 and 32×32 PNG icons with the Progress INFRA branding
2. **Remove DEV markers** — rename to `teams.html`, change app name to "Sales Play Tracker" (remove DEV)
3. **New manifest version** — bump to `"version": "1.0.0"` and assign a new stable app ID
4. **Register in Teams Admin Center** — upload as an org-wide app instead of sideload
5. **Test auth on Teams mobile** — validate redirect flow on iOS/Android Teams app
6. **Teams SSO** — optionally implement proper Teams SSO (`microsoftTeams.authentication.getAuthToken`) for seamless sign-in; this is a more involved auth upgrade
7. **Separate app registration** — optionally create a dedicated Azure AD app for the Teams app (currently shares with web app)
8. **Custom domain** — consider adding a custom domain to avoid Azure URL changes (see open items in WHAT-WE-BUILT.md)

---

*Created: 2026-05-08 — Phil Langer + GitHub Copilot*
*Branch: dev/v2-upgrade*
