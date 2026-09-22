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
  if (!window.msal) {
    throw new Error("Microsoft authentication library did not load.");
  }

  msalApp = new msal.PublicClientApplication(getMsalConfig());

  try {
    // Important for redirect authentication.
    const redirectResult = await msalApp.handleRedirectPromise();

    if (redirectResult?.account) {
      currentAccount = redirectResult.account;
      msalApp.setActiveAccount(currentAccount);
    }
  } catch (error) {
    console.error("Redirect handling error:", error);
    throw error;
  }

  const accounts = msalApp.getAllAccounts();

  if (!currentAccount && accounts.length > 0) {
    currentAccount = accounts[0];
    msalApp.setActiveAccount(currentAccount);
  }

  return currentAccount;
}


/* ---------------------------------------------------------
   MICROSOFT LOGIN
   Uses full-page redirect instead of popup.
--------------------------------------------------------- */

async function signIn() {
  if (!msalApp) {
    await initAuth();
  }

  await msalApp.loginRedirect({
    scopes: APP_CONFIG.scopes,
    prompt: "select_account"
  });
}


/* ---------------------------------------------------------
   SIGN OUT
--------------------------------------------------------- */

async function signOut() {
  if (!msalApp) {
    await initAuth();
  }

  const account =
    currentAccount ||
    msalApp.getActiveAccount() ||
    msalApp.getAllAccounts()[0];

  currentAccount = null;

  await msalApp.logoutRedirect({
    account,
    postLogoutRedirectUri: APP_CONFIG.redirectUri
  });
}


/* ---------------------------------------------------------
   GET GRAPH ACCESS TOKEN
   First try silent authentication.
   If Microsoft requires interaction, use redirect instead
   of opening a popup.
--------------------------------------------------------- */

async function getAccessToken() {
  if (!msalApp) {
    await initAuth();
  }

  currentAccount =
    currentAccount ||
    msalApp.getActiveAccount() ||
    msalApp.getAllAccounts()[0];

  if (!currentAccount) {
    await msalApp.loginRedirect({
      scopes: APP_CONFIG.scopes,
      prompt: "select_account"
    });

    throw new Error("Redirecting to Microsoft sign-in...");
  }

  msalApp.setActiveAccount(currentAccount);

  try {
    const result = await msalApp.acquireTokenSilent({
      account: currentAccount,
      scopes: APP_CONFIG.scopes
    });

    return result.accessToken;

  } catch (silentError) {

    console.warn(
      "Silent token acquisition failed. Redirecting to Microsoft.",
      silentError
    );

    await msalApp.acquireTokenRedirect({
      account: currentAccount,
      scopes: APP_CONFIG.scopes
    });

    throw new Error("Redirecting to Microsoft authentication...");
  }
}


function getAccount() {
  return (
    currentAccount ||
    msalApp?.getActiveAccount() ||
    msalApp?.getAllAccounts()?.[0] ||
    null
  );
}
