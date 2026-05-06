// INFRA SalesPlay Tracker — Configuration
// Fill in these 3 values after registering your Azure AD app
// See README.md for step-by-step instructions

const CONFIG = (() => {
  const host = window.location.hostname || '';

  // Production: the main Azure SWA domain
  const isProd = host.includes('orange-smoke-04675d40f');

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

    // Leave these as-is
    graphBaseUrl: "https://graph.microsoft.com/v1.0",
    scopes: [
      "https://graph.microsoft.com/Sites.ReadWrite.All",
      "https://graph.microsoft.com/User.Read",
    ],
  };
})();

// Dev environment banner — auto-injected on non-production hosts
if (CONFIG.isDev) {
  document.addEventListener('DOMContentLoaded', () => {
    const banner = document.createElement('div');
    banner.id = 'dev-env-banner';
    banner.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:99999;background:#ff6600;color:#fff;text-align:center;padding:4px 8px;font:bold 12px/1.4 system-ui;pointer-events:none;';
    banner.textContent = `⚠️ DEV ENVIRONMENT — Data folder: ${CONFIG.dataFolder}`;
    document.body.prepend(banner);
    document.body.style.paddingTop = '28px';
  });
}
