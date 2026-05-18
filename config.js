// INFRA SalesPlay Tracker — Configuration
// Fill in these 3 values after registering your Azure AD app
// See README.md for step-by-step instructions

const APP_VERSION = '2.1.0';

const CONFIG = (() => {
  const host = window.location.hostname || '';

  // Production: exact match on the main Azure SWA domain only.
  // Dev preview URLs look like orange-smoke-04675d40f-5.eastus2.7.azurestaticapps.net
  // and must NOT be treated as prod — they need the dev data folder and banner.
  const isProd = host === 'orange-smoke-04675d40f.7.azurestaticapps.net';

  // Dev: preview environments, localhost, or any non-prod host
  const isDev = !isProd;

  // SharePoint data folder — dev uses isolated folder to protect production data
  const dataFolder = isProd ? 'ChefSaaS' : 'ChefSaaS-Dev';

  return {
    // From Azure AD App Registration → Overview
    tenantId: "db266a67-cbe0-4d26-ae1a-d0581fe03535",
    clientId: "48db2acc-9dfb-4337-9abb-302c9dfb88fc",

    // Your SharePoint site URL
    sharepointSiteUrl: "https://progresssoftware.sharepoint.com/sites/INFRASalesApps",

    // Environment flags
    isDev,
    isProd,
    dataFolder,

    // Role-based access — shared across admin.html, ingest.html, reports.html
    // Roles: 'owner' (full access), 'manager' (manage plays + reps), 'product' (play content only, no account upload)
    // To grant access: add the user's Progress SSO email here with their role.
    ADMIN_ROLES: {
      'langer@progress.com':           'owner',
      'philip.langer@progress.com':    'owner',
      'planger@progress.com':          'owner',
      'phil.langer@progress.com':      'owner',
      'kathleen.faria@progress.com':   'manager',
      'faria@progress.com':            'manager',
      'kfaria@progress.com':           'manager',
      'k.faria@progress.com':          'manager',
      'kathy.faria@progress.com':      'manager',
      'scheaney@progress.com':         'manager',
      'sara.scheaney@progress.com':    'manager',
    },

    // Leave these as-is
    graphBaseUrl: "https://graph.microsoft.com/v1.0",
    scopes: [
      "https://graph.microsoft.com/Sites.ReadWrite.All",
      "https://graph.microsoft.com/User.Read",
      // User.ReadBasic.All (for manager directReports) requires admin consent —
      // requested incrementally in getDirectReports() instead of at login.
    ],
  };
})();

// Dev environment banner — auto-injected on non-production hosts
if (CONFIG.isDev) {
  document.addEventListener('DOMContentLoaded', () => {
    // Update the existing static banner (present in every page's HTML) rather
    // than prepending a second one.  Fall back to creating one if absent.
    let banner = document.getElementById('dev-env-banner');
    if (!banner) {
      banner = document.createElement('div');
      banner.id = 'dev-env-banner';
      banner.style.cssText = 'position:sticky;top:0;left:0;right:0;z-index:99999;border-bottom:2px solid #b34400;';
      document.body.prepend(banner);
    }
    banner.style.background = '#e65c00';
    banner.style.color = '#fff';
    banner.style.textAlign = 'center';
    banner.style.padding = '6px 12px';
    banner.style.fontSize = '12px';
    banner.style.fontWeight = '700';
    banner.textContent = `⚠️ DEV  v${APP_VERSION}  —  Data: ${CONFIG.dataFolder}  —  Changes go to test data only`;
  });
}
