let msalApp = null;
let currentAccount = null;

function getMsalConfig() {
  return {
    auth: {
      clientId: APP_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${APP_CONFIG.tenantId}`,
      redirectUri: APP_CONFIG.redirectUri,
      postLogoutRedirectUri: APP_CONFIG.redirectUri
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false
    },
    system: {
      allowNativeBroker: false
    }
  };
}

async function initAuth() {
  if (!window.msal) throw new Error("Microsoft authentication library did not load.");
  msalApp = new msal.PublicClientApplication(getMsalConfig());

  try {
    const redirectResult = await msalApp.handleRedirectPromise();
    if (redirectResult?.account) {
      currentAccount = redirectResult.account;
      msalApp.setActiveAccount(currentAccount);
    }
  } catch (error) {
    console.error("Redirect handling error", error);
  }

  const accounts = msalApp.getAllAccounts();
  if (!currentAccount && accounts.length) {
    currentAccount = accounts[0];
    msalApp.setActiveAccount(currentAccount);
  }

  return currentAccount;
}

async function signIn() {
  if (!msalApp) await initAuth();

  const result = await msalApp.loginPopup({
    scopes: APP_CONFIG.scopes,
    prompt: "select_account"
  });

  currentAccount = result.account;
  msalApp.setActiveAccount(currentAccount);
  return currentAccount;
}

async function signOut() {
  if (!msalApp) return;
  const account = currentAccount || msalApp.getActiveAccount();
  currentAccount = null;
  await msalApp.logoutPopup({
    account,
    mainWindowRedirectUri: APP_CONFIG.redirectUri
  });
}

async function getAccessToken() {
  if (!msalApp) await initAuth();

  currentAccount = currentAccount || msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
  if (!currentAccount) {
    throw new Error("Please sign in with your Microsoft 365 account first.");
  }

  msalApp.setActiveAccount(currentAccount);

  try {
    const result = await msalApp.acquireTokenSilent({
      account: currentAccount,
      scopes: APP_CONFIG.scopes
    });
    return result.accessToken;
  } catch (silentError) {
    console.warn("Silent token failed; opening Microsoft sign-in.", silentError);
    const result = await msalApp.acquireTokenPopup({
      account: currentAccount,
      scopes: APP_CONFIG.scopes
    });
    return result.accessToken;
  }
}

function getAccount() {
  return currentAccount || msalApp?.getActiveAccount() || msalApp?.getAllAccounts()?.[0] || null;
}
