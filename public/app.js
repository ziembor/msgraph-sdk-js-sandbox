const GRAPH_SCOPES = ["Mail.Read", "Calendars.Read", "Calendars.Read.Shared"];
const statusEl = document.getElementById("status");
const outputEl = document.getElementById("output");

const appConfig = window.__APP_CONFIG__ || {};

const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: appConfig.clientId || "",
    authority: `https://login.microsoftonline.com/${appConfig.tenantId || "common"}`,
    redirectUri: appConfig.redirectUri || window.location.origin,
  },
});

const debugState = {
  config: null,
  account: null,
  token: null,
  lastGraphRequest: null,
  lastError: null,
};

const debugConfigEl = document.getElementById("debugConfig");
const debugAccountEl = document.getElementById("debugAccount");
const debugTokenEl = document.getElementById("debugToken");
const debugGraphEl = document.getElementById("debugGraph");
const debugErrorEl = document.getElementById("debugError");

function renderDebugPanel() {
  debugConfigEl.textContent = debugState.config ? JSON.stringify(debugState.config, null, 2) : "(none yet)";
  debugAccountEl.textContent = debugState.account ? JSON.stringify(debugState.account, null, 2) : "(signed out)";
  debugTokenEl.textContent = debugState.token ? JSON.stringify(debugState.token, null, 2) : "(none yet)";
  debugGraphEl.textContent = debugState.lastGraphRequest
    ? JSON.stringify(debugState.lastGraphRequest, null, 2)
    : "(none yet)";
  debugErrorEl.textContent = debugState.lastError ? JSON.stringify(debugState.lastError, null, 2) : "(none)";
}

function decodeJwtPayload(token) {
  // Graph access tokens are officially opaque; this decoding is best-effort, for learning only.
  try {
    const payload = token.split(".")[1];
    return JSON.parse(atob(payload.replace(/-/g, "+").replace(/_/g, "/")));
  } catch {
    return null;
  }
}

function updateAccountDebug(account) {
  debugState.account = account
    ? {
        username: account.username,
        homeAccountId: account.homeAccountId,
        tenantId: account.tenantId,
        environment: account.environment,
        idTokenClaims: account.idTokenClaims,
      }
    : null;
  renderDebugPanel();
}

function updateTokenDebug({ flow, tokenResult }) {
  debugState.token = {
    flow,
    fromCache: tokenResult.fromCache,
    tokenType: tokenResult.tokenType,
    expiresOn: tokenResult.expiresOn,
    requestedScopes: GRAPH_SCOPES,
    grantedScopes: tokenResult.scopes,
    accessTokenClaims: decodeJwtPayload(tokenResult.accessToken),
    rawAccessToken: tokenResult.accessToken,
  };
  renderDebugPanel();
}

function updateGraphDebug({ method, url, status, headers, errorBody }) {
  debugState.lastGraphRequest = {
    timestamp: new Date().toISOString(),
    method,
    url,
    status,
    headers,
    errorBody,
  };
  renderDebugPanel();
}

function recordError(action, error) {
  debugState.lastError = {
    action,
    name: error?.name,
    message: error?.message,
    errorCode: error?.errorCode,
    correlationId: error?.correlationId,
  };
  console.error(`${action} failed.`, error);
  renderDebugPanel();
}

debugState.config = { ...msalInstance.getConfiguration().auth, scopes: GRAPH_SCOPES };
renderDebugPanel();

function setOutput(value) {
  outputEl.textContent = value;
}

function setStatus(value) {
  statusEl.textContent = value;
}

function getCurrentAccount() {
  return msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
}

async function initialize() {
  await msalInstance.initialize();
  const response = await msalInstance.handleRedirectPromise();
  if (response?.account) {
    msalInstance.setActiveAccount(response.account);
  }
  const account = getCurrentAccount();
  if (account) {
    msalInstance.setActiveAccount(account);
    setStatus(`Signed in as ${account.username}`);
  }
  updateAccountDebug(account);
}

async function signIn() {
  const loginResponse = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
  msalInstance.setActiveAccount(loginResponse.account);
  setStatus(`Signed in as ${loginResponse.account.username}`);
  updateAccountDebug(loginResponse.account);
}

async function signOut() {
  const account = getCurrentAccount();
  await msalInstance.logoutPopup({ account });
  setStatus("Signed out");
  setOutput("");
  updateAccountDebug(null);
  debugState.token = null;
  renderDebugPanel();
}

async function acquireAccessToken() {
  const account = getCurrentAccount();
  if (!account) {
    throw new Error("No authenticated account found. Please sign in to continue.");
  }

  try {
    const tokenResult = await msalInstance.acquireTokenSilent({
      account,
      scopes: GRAPH_SCOPES,
    });
    updateTokenDebug({ flow: "silent", tokenResult });
    return tokenResult.accessToken;
  } catch (error) {
    const needsInteraction =
      error instanceof msal.InteractionRequiredAuthError ||
      ["interaction_required", "consent_required", "login_required"].includes(
        error?.errorCode
      );
    if (!needsInteraction) {
      throw error;
    }

    const tokenResult = await msalInstance.acquireTokenPopup({
      scopes: GRAPH_SCOPES,
    });
    updateTokenDebug({ flow: "popup (interaction required)", tokenResult });
    return tokenResult.accessToken;
  }
}

async function callGraph(path) {
  const method = "GET";
  const url = `https://graph.microsoft.com/v1.0${path}`;
  const token = await acquireAccessToken();
  const response = await fetch(url, {
    method,
    headers: {
      Authorization: "Bearer ".concat(token),
    },
  });

  const headers = [...response.headers.entries()];

  if (!response.ok) {
    const errText = await response.text();
    updateGraphDebug({ method, url, status: response.status, headers, errorBody: errText });
    throw new Error(`Graph request failed (${response.status}): ${errText}`);
  }
  updateGraphDebug({ method, url, status: response.status, headers, errorBody: null });
  return response.json();
}

async function loadMail() {
  const result = await callGraph(
    "/me/messages?$top=10&$select=subject,from,receivedDateTime"
  );
  setOutput(JSON.stringify(result, null, 2));
}

async function loadCalendar() {
  const result = await callGraph(
    "/me/events?$top=10&$select=subject,start,end,organizer"
  );
  setOutput(JSON.stringify(result, null, 2));
}

async function loadSharedCalendar() {
  const address = document.getElementById("sharedMailbox").value.trim();
  if (!address) {
    setOutput("Enter a shared mailbox address (e.g. shared.mailbox@contoso.com).");
    return;
  }
  const result = await callGraph(
    `/users/${encodeURIComponent(address)}/events?$top=10&$select=subject,start,end,organizer`
  );
  setOutput(JSON.stringify(result, null, 2));
}

function wireUpButtons() {
  document.getElementById("signIn").addEventListener("click", async () => {
    try {
      await signIn();
    } catch (error) {
      recordError("Sign-in", error);
      setOutput("Sign-in failed. Check browser console for details.");
    }
  });

  document.getElementById("signOut").addEventListener("click", async () => {
    try {
      await signOut();
    } catch (error) {
      recordError("Sign-out", error);
      setOutput("Sign-out failed. Check browser console for details.");
    }
  });

  document.getElementById("loadMail").addEventListener("click", async () => {
    try {
      await loadMail();
    } catch (error) {
      recordError("Loading mail", error);
      setOutput("Loading mail failed. Check browser console for details.");
    }
  });

  document.getElementById("loadCalendar").addEventListener("click", async () => {
    try {
      await loadCalendar();
    } catch (error) {
      recordError("Loading calendar", error);
      setOutput("Loading calendar failed. Check browser console for details.");
    }
  });

  document.getElementById("openSharedCalendar").addEventListener("click", async () => {
    try {
      await loadSharedCalendar();
    } catch (error) {
      recordError("Loading shared calendar", error);
      setOutput("Loading shared calendar failed. Check browser console for details.");
    }
  });
}

initialize()
  .then(() => wireUpButtons())
  .catch((error) => {
    recordError("App initialization", error);
    setOutput("App initialization failed. Check browser console for details.");
  });
