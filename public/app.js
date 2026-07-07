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
}

async function signIn() {
  const loginResponse = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
  msalInstance.setActiveAccount(loginResponse.account);
  setStatus(`Signed in as ${loginResponse.account.username}`);
}

async function signOut() {
  const account = getCurrentAccount();
  await msalInstance.logoutPopup({ account });
  setStatus("Signed out");
  setOutput("");
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
    return tokenResult.accessToken;
  }
}

async function callGraph(path) {
  const token = await acquireAccessToken();
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: {
      Authorization: "Bearer ".concat(token),
    },
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Graph request failed (${response.status}): ${errText}`);
  }
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
      console.error("Sign-in failed.", error);
      setOutput("Sign-in failed. Check browser console for details.");
    }
  });

  document.getElementById("signOut").addEventListener("click", async () => {
    try {
      await signOut();
    } catch (error) {
      console.error("Sign-out failed.", error);
      setOutput("Sign-out failed. Check browser console for details.");
    }
  });

  document.getElementById("loadMail").addEventListener("click", async () => {
    try {
      await loadMail();
    } catch (error) {
      console.error("Loading mail failed.", error);
      setOutput("Loading mail failed. Check browser console for details.");
    }
  });

  document.getElementById("loadCalendar").addEventListener("click", async () => {
    try {
      await loadCalendar();
    } catch (error) {
      console.error("Loading calendar failed.", error);
      setOutput("Loading calendar failed. Check browser console for details.");
    }
  });

  document.getElementById("openSharedCalendar").addEventListener("click", async () => {
    try {
      await loadSharedCalendar();
    } catch (error) {
      console.error("Loading shared calendar failed.", error);
      setOutput("Loading shared calendar failed. Check browser console for details.");
    }
  });
}

initialize()
  .then(() => wireUpButtons())
  .catch((error) => {
    console.error("App initialization failed.", error);
    setOutput("App initialization failed. Check browser console for details.");
  });
