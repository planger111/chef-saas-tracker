// Chef SaaS Motion Tracker — MSAL Authentication (popup flow)

let _msalInstance = null;
let _account = null;

function _getMsalConfig() {
  return {
    auth: {
      clientId: CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: window.location.origin,
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: true,
    },
  };
}

async function initAuth() {
  if (!_msalInstance) {
    _msalInstance = new msal.PublicClientApplication(_getMsalConfig());
  }

  // Handle any pending redirect (in case redirect flow was used previously)
  try {
    const response = await _msalInstance.handleRedirectPromise();
    if (response && response.account) {
      _account = response.account;
    }
  } catch (err) {
    console.warn("MSAL redirect error (safe to ignore on popup flow):", err);
  }

  // Check cache for existing account
  if (!_account) {
    const accounts = _msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      _account = accounts[0];
    }
  }

  return _account;
}

async function signIn() {
  if (!_msalInstance) {
    _msalInstance = new msal.PublicClientApplication(_getMsalConfig());
    await _msalInstance.handleRedirectPromise();
  }
  // Use popup — no redirect URI issues, works on any page
  const response = await _msalInstance.loginPopup({ scopes: CONFIG.scopes });
  if (response && response.account) {
    _account = response.account;
  }
  return _account;
}

async function signOut() {
  await _msalInstance.logoutPopup({ account: _account });
  _account = null;
}

async function getAccessToken() {
  if (!_account) throw new Error("No signed-in account");

  const request = { scopes: CONFIG.scopes, account: _account };

  try {
    const result = await _msalInstance.acquireTokenSilent(request);
    return result.accessToken;
  } catch (err) {
    if (err instanceof msal.InteractionRequiredAuthError) {
      // Fall back to popup for token refresh
      const result = await _msalInstance.acquireTokenPopup(request);
      return result.accessToken;
    }
    throw err;
  }
}

function getCurrentUser() {
  if (!_account) return null;
  return {
    name: _account.name || _account.username,
    email: _account.username,
    username: _account.username,
  };
}
