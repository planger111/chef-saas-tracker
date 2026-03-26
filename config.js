// Chef SaaS Motion Tracker — Configuration
// Fill in these 3 values after registering your Azure AD app
// See README.md for step-by-step instructions

const CONFIG = {
  // From Azure AD App Registration → Overview
  tenantId: "db266a67-cbe0-4d26-ae1a-d0581fe03535",
  clientId: "48db2acc-9dfb-4337-9abb-302c9dfb88fc",

  // Your SharePoint site URL (e.g. https://progresssoftware.sharepoint.com/sites/SalesOps)
  sharepointSiteUrl: "https://progresssoftware.sharepoint.com/sites/INFRASalesLeadership",

  // Leave these as-is
  graphBaseUrl: "https://graph.microsoft.com/v1.0",
  scopes: ["https://graph.microsoft.com/Sites.ReadWrite.All"],
};
