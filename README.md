# msgraph-sdk-js-sandbox

Minimal JavaScript single-page app (SPA) for Azure Web App hosting that signs in with an Entra ID app and calls Microsoft Graph for Exchange Online data using delegated scopes:

- `Mail.Read`
- `Calendars.Read`
- `Calendars.Read.Shared`

## Project structure

- `/public/index.html` - SPA UI
- `/public/app.js` - MSAL + Graph logic
- `/public/config.example.js` - template app configuration (committed)
- `/public/config.js` - your local app configuration (client/tenant/redirect); git-ignored
- `/server.js` - lightweight static file server for App Service

## Configure Entra ID app

1. Register an app in Entra ID.
2. Add SPA redirect URI (for local dev: `http://localhost:3000`).
3. Add delegated Graph permissions:
   - `Mail.Read`
   - `Calendars.Read`
   - `Calendars.Read.Shared`
4. Grant admin consent if your tenant requires it.
5. Copy `/public/config.example.js` to `/public/config.js` and set your values:
   - `clientId`
   - `tenantId`
   - `redirectUri`

   `config.js` is git-ignored so your real client/tenant IDs stay local.

## Run locally

```bash
node server.js
```

Open `http://localhost:3000`.

## Deploy to Azure Web App

Deploy this repository to an Azure Web App (Node runtime). App Service uses `server.js` and `process.env.PORT` automatically when started with:

```bash
node server.js
```