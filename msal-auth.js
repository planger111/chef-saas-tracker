// Chef SaaS Motion Tracker — MSAL Authentication (PKCE flow)

let _msalInstance = null;
let _account = null;

function _getMsalConfig() {
  // Build redirectUri from current page URL (strips query string and hash)
  // This allows the same app to work from localhost AND GitHub Pages
  const redirectUri = window.location.origin + window.location.pathname;
  return {
    auth: {
      clientId: CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: redirectUri,
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false,
    },
  };
}

async function initAuth() {
  _msalInstance = new msal.PublicClientApplication(_getMsalConfig());

  // Handle redirect promise (called after loginRedirect returns)
  try {
    const response = await _msalInstance.handleRedirectPromise();
    if (response && response.account) {
      _account = response.account;
    }
  } catch (err) {
    console.error("MSAL redirect error:", err);
  }

  // If still no account, check the cache
  if (!_account) {
    const accounts = _msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      _account = accounts[0];
    }
  }

  return _account;
}

function signIn() {
  _msalInstance.loginRedirect({ scopes: CONFIG.scopes });
}

function signOut() {
  _msalInstance.logoutRedirect({ account: _account });
}

async function getAccessToken() {
  if (!_account) throw new Error("No signed-in account");

  const request = { scopes: CONFIG.scopes, account: _account };

  try {
    const result = await _msalInstance.acquireTokenSilent(request);
    return result.accessToken;
  } catch (err) {
    if (err instanceof msal.InteractionRequiredAuthError) {
      _msalInstance.acquireTokenRedirect(request);
    }
    throw err;
  }
}

function getCurrentUser() {
  if (!_account) return null;
  return {
    name: _account.name || _account.username,
    username: _account.username,
  };
}
