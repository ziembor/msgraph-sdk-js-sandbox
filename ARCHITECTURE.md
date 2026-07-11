# Architecture

This is a deliberately minimal single-page application (SPA) demonstrating the **OAuth 2.0 authorization code flow (with PKCE, via MSAL.js)** and **Microsoft Graph API calls** for mail and calendar scopes — including opening a shared mailbox calendar. There is no framework, no bundler, and no npm dependency: the entire app is one HTML page, one JavaScript file, and a tiny static file server.

## Project layout

```
msgraph-sdk-js-sandbox/
├── package.json                  # name + one script: "start": "node server.js" (no dependencies)
├── server.js                     # zero-dependency Node.js static file server (port 3000)
├── ARCHITECTURE.md               # this file
├── README.md                     # setup and deployment docs
├── .github/workflows/deploy.yml  # CI: generates config.js, deploys public/ to Cloudflare Pages
└── public/                       # everything served to the browser
    ├── index.html                # SPA markup + inline <style> (no external CSS)
    ├── app.js                    # all MSAL + Graph + UI logic
    ├── config.example.js         # committed template for local config
    └── config.js                 # git-ignored; real clientId/tenantId/redirectUri
```

## Runtime pieces

### Static server (`server.js`)

A plain `node:http` server with no dependencies:

- Indexes every file under `public/` at startup into a path→file map (`indexPublicFiles`).
- Normalizes and validates each request path (`normalizeRequestPath`): rejects null bytes, backslash tricks, and any `..` segment — path-traversal safe by construction.
- Serves known files with the right `Content-Type`; extensionless paths fall back to `index.html` (SPA-style routing); everything else is 404/403.
- Listens on `PORT` env var or `3000` (which matches the local `redirectUri`).

### Page (`public/index.html`)

Flat, framework-free markup:

1. Status line (`#status`) — signed-in/out state.
2. Toolbar — `#signIn`, `#signOut`, `#loadMail`, `#loadCalendar` buttons, a `#sharedMailbox` email input, and `#openSharedCalendar`.
3. Output panel — a single `<pre id="output">` where Graph JSON results are printed.
4. **Debug Panel** — five collapsible `<details>` sections (`#debugConfig`, `#debugAccount`, `#debugToken`, `#debugGraph`, `#debugError`) that expose the OAuth/Graph internals live (see below).
5. Scripts, loaded in order at the end of `<body>`:
   - `msal-browser@2.39.0` from the unpkg CDN (the only external dependency),
   - `./config.js` (sets `window.__APP_CONFIG__`),
   - `./app.js`.

### Configuration (`config.js` / `config.example.js`)

`config.js` is git-ignored and sets `window.__APP_CONFIG__ = { clientId, tenantId, redirectUri }`. Locally you copy `config.example.js` and fill in your Entra app registration values; in CI the deploy workflow generates it from GitHub repository Variables. Because it's plain client-side JS, these are public identifiers, not secrets.

## Application logic (`public/app.js`)

All logic lives in this one file. Key constants:

- `GRAPH_SCOPES = ["Mail.Read", "Calendars.Read", "Calendars.Read.Shared"]` — the delegated Graph scopes this demo is about.
- `msalInstance = new msal.PublicClientApplication(...)` — configured from `window.__APP_CONFIG__` with authority `https://login.microsoftonline.com/{tenantId}` (falls back to `common`) and `redirectUri` defaulting to `window.location.origin`.

### Auth flow

```
page load
  └─ initialize()
       ├─ msalInstance.initialize()
       ├─ handleRedirectPromise()          # completes a redirect flow if one is in flight
       ├─ set active account (if any cached) → status "Signed in as <user>"
       └─ updateAccountDebug(account)      # debug panel reflects session state
  └─ wireUpButtons()                       # attaches all click handlers

Sign in  → signIn():  msalInstance.loginPopup({ scopes: GRAPH_SCOPES })
                      → setActiveAccount → status + debug panel update
Sign out → signOut(): msalInstance.logoutPopup({ account })
                      → clears status, output, account + token debug state
```

Token acquisition (`acquireAccessToken()`) implements the standard MSAL silent-first pattern:

1. Require a signed-in account (throw a friendly error otherwise).
2. Try `acquireTokenSilent({ account, scopes })` — served from MSAL's cache or via a hidden refresh.
3. If (and only if) the error is interaction-required (`InteractionRequiredAuthError`, or error codes `interaction_required` / `consent_required` / `login_required`), fall back to `acquireTokenPopup({ scopes })`. Any other error is re-thrown.
4. Either path records which flow ran ("silent" vs "popup (interaction required)") in the debug panel before returning the raw access token.

### Graph calls

`callGraph(path)` is the single Graph gateway: acquire a token, then plain `fetch` to `https://graph.microsoft.com/v1.0{path}` with an `Authorization: Bearer` header (no Graph SDK is used). Non-OK responses throw with the status and raw error body. Callers:

| Action | Endpoint |
|---|---|
| Load Mail | `GET /me/messages?$top=10&$select=subject,from,receivedDateTime` |
| Load Calendar | `GET /me/events?$top=10&$select=subject,start,end,organizer` |
| Open Shared Mailbox Calendar | `GET /users/{address}/events?$top=10&$select=...` — requires `Calendars.Read.Shared` and actual access to that mailbox |

Results are pretty-printed as JSON into `#output`.

### Error handling

Every button handler and the app-init chain follows one pattern:

```js
try { await action(); }
catch (error) {
  recordError("<Action>", error);   // console.error + structured entry in the debug panel
  setOutput("<Action> failed. Check browser console for details.");
}
```

`recordError` captures `action`, `name`, `message`, and MSAL's `errorCode` / `correlationId` into the debug panel, so auth failures are diagnosable in the UI, not just the console.

### Debug panel

Since the app exists to *demonstrate* the OAuth2/Graph flow, a `debugState` object (`config`, `account`, `token`, `lastGraphRequest`, `lastError`) is rendered into the panel by `renderDebugPanel()` and updated at every hook point:

| Section | Populated by | Shows |
|---|---|---|
| App Config | immediately after MSAL construction | resolved MSAL auth config (clientId, authority, redirectUri) + requested scopes |
| Account / Session | `initialize()`, `signIn()`, `signOut()` | username, homeAccountId, tenantId, environment, ID-token claims |
| Last Token Acquisition | `acquireAccessToken()` | flow used (silent vs popup), `fromCache`, expiry, **requested vs granted scopes**, decoded access-token claims (best-effort via `decodeJwtPayload` — Graph tokens are officially opaque), raw token |
| Last Graph Request | `callGraph()` | method, URL, status, CORS-exposed response headers, error body on failure |
| Last Error | `recordError()` | action, error name/message, MSAL errorCode + correlationId |

The raw token and decoded claims are shown unredacted by design — this is a local teaching sandbox, not a production app.

## Deployment (`.github/workflows/deploy.yml`)

On push to `main` (or manual dispatch):

1. **Generate `public/config.js`** from repository *Variables* (`AZURE_CLIENT_ID`, `AZURE_TENANT_ID`, `REDIRECT_URI`) using Node's `JSON.stringify` for safe escaping. Variables, not Secrets, because these values ship in client-side JS anyway.
2. **Deploy `public/`** to Cloudflare Pages via `wrangler-action` (`pages deploy public --project-name=zbmsgraph-sdk-js-sandbox`), authenticated with `CLOUDFLARE_API_TOKEN` / `CLOUDFLARE_ACCOUNT_ID` secrets.

There is no build step — the deployed artifact is the `public/` folder as-is plus the generated `config.js`.

## Design principles

- **Zero dependencies, zero build** — everything is readable in place; the only external code is the MSAL CDN script.
- **Transparency over abstraction** — raw `fetch` instead of the Graph SDK, and a debug panel that exposes config, tokens, scopes, requests, and errors, because the point of the app is to make the OAuth2/Graph flow visible.
- **Secrets stay out of git** — `config.js` is git-ignored locally and generated in CI; the values themselves are public client identifiers.
