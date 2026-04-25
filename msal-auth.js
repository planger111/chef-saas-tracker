// INFRA SalesPlay Tracker — MSAL Authentication (popup with redirect fallback)

let _msalInstance = null;
let _account = null;

// Detect contexts where popups are unreliable — use redirect flow instead
function _useRedirectFlow() {
  try {
    const ua = navigator.userAgent;
    // Mobile browsers — popups are blocked or unreliable on phones/tablets
    if (/iPhone|iPad|iPod|Android|Mobile|Tablet/i.test(ua)) return true;
    // Running inside an iframe
    if (window !== window.parent) return true;
    // Teams in-app browser
    if (/Teams/i.test(ua)) return true;
    // Embedded webviews (Facebook, Instagram, etc.)
    if (/wv|WebView|FBAN|FBAV|Instagram|Line\//i.test(ua)) return true;
    // Already inside a popup
    if (window.opener && window.opener !== window) return true;
  } catch (_) { /* cross-origin frame check may throw — treat as embedded */ return true; }
  return false;
}

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

  const loginRequest = { scopes: CONFIG.scopes };

  if (_useRedirectFlow()) {
    // Redirect flow for Teams / embedded browsers / webviews
    await _msalInstance.loginRedirect(loginRequest);
    return null; // page will reload on redirect return
  }

  // Popup flow for normal browsers
  try {
    const response = await _msalInstance.loginPopup(loginRequest);
    if (response && response.account) {
      _account = response.account;
    }
    return _account;
  } catch (popupErr) {
    // If popup fails at runtime (blocked by browser), fall back to redirect
    if (popupErr.errorCode === "block_nested_popups" ||
        popupErr.errorCode === "popup_window_error") {
      await _msalInstance.loginRedirect(loginRequest);
      return null;
    }
    throw popupErr;
  }
}

async function signOut() {
  if (_useRedirectFlow()) {
    await _msalInstance.logoutRedirect({ account: _account });
  } else {
    await _msalInstance.logoutPopup({ account: _account });
  }
  _account = null;
}

async function getAccessToken() {
  if (!_account) throw new Error("No signed-in account");

  const request = { scopes: CONFIG.scopes, account: _account };

  try {
    const result = await _msalInstance.acquireTokenSilent(request);
    return result.accessToken;
  } catch (err) {
    // Any error that requires user interaction or iframe timeout on mobile
    const needsInteraction = err instanceof msal.InteractionRequiredAuthError;
    const iframeFailed = err.errorCode === "monitor_window_timeout" ||
                         err.errorCode === "token_renewal_error";

    if (needsInteraction || iframeFailed) {
      if (_useRedirectFlow()) {
        await _msalInstance.acquireTokenRedirect(request);
        return null; // page will reload
      }
      try {
        const result = await _msalInstance.acquireTokenPopup(request);
        return result.accessToken;
      } catch (popupErr) {
        if (popupErr.errorCode === "block_nested_popups" ||
            popupErr.errorCode === "popup_window_error") {
          await _msalInstance.acquireTokenRedirect(request);
          return null;
        }
        throw popupErr;
      }
    }
    throw err;
  }
}

function getCurrentUser() {
  if (!_account) return null;
  const claims = _account.idTokenClaims || {};
  // Collect every email-like identifier available from the token
  const allEmails = [
    _account.username,
    claims.preferred_username,
    claims.email,
    claims.upn,
    claims.unique_name,
  ].filter(Boolean).map(e => e.toLowerCase());
  const uniqueEmails = [...new Set(allEmails)];
  const primaryEmail = uniqueEmails[0] || '';
  return {
    name:     _account.name || primaryEmail,
    email:    primaryEmail,
    username: _account.username || primaryEmail,
    oid:      claims.oid || _account.localAccountId || '',
    emails:   uniqueEmails,   // all candidates for role lookup
  };
}
