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
- `/.github/workflows/deploy.yml` - CI: build config.js + deploy to Cloudflare Pages

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

## Deploy to Cloudflare Pages (GitHub Actions)

`.github/workflows/deploy.yml` runs on every push to `main` (and via manual dispatch). It
generates `public/config.js` from repository Variables and deploys the static `public/` folder to
Cloudflare Pages. Because `config.js` is git-ignored, the workflow rebuilds it each run — nothing
sensitive lives in the repo.

**Repository Variables** (Settings → Secrets and variables → Actions → *Variables*) — these are
public identifiers that ship in client-side JS, so they are Variables, not Secrets:

| Variable | Value |
| --- | --- |
| `AZURE_CLIENT_ID` | Entra app (client) ID |
| `AZURE_TENANT_ID` | Entra directory (tenant) ID |
| `REDIRECT_URI` | Production URL, e.g. `https://zbmsgraph-sdk-js-sandbox.pages.dev` |

**Repository Secrets** (same page → *Secrets*):

| Secret | Value |
| --- | --- |
| `CLOUDFLARE_API_TOKEN` | API token with the **Cloudflare Pages: Edit** permission (see below) |
| `CLOUDFLARE_ACCOUNT_ID` | Your Cloudflare account ID (see below) |

### Create the Cloudflare credentials

**`CLOUDFLARE_API_TOKEN`** — use a scoped Custom Token, never the Global API Key:

1. Cloudflare dashboard → **My Profile → API Tokens → Create Token → Create Custom Token**.
2. **Permissions:** `Account` → `Cloudflare Pages` → `Edit`. This single permission is all
   `pages deploy` needs — no Workers, Zone, or Account Settings scopes.
3. **Account Resources:** Include → *your account only* (not "All accounts").
4. Optionally set a **TTL / expiration** so the token gets rotated. Leave IP filtering off —
   GitHub-hosted runners have dynamic IPs.
5. Create it, copy the token value once, and paste it into the `CLOUDFLARE_API_TOKEN` secret.

> Note: Cloudflare's Pages permission is account-wide — this token can deploy to *any* Pages
> project in the account (there is no per-project scoping). For a hard blast-radius limit, use a
> dedicated Cloudflare account for this project.

**`CLOUDFLARE_ACCOUNT_ID`** — find it any of these ways, then paste into the secret:

- Dashboard → **Workers & Pages** → right sidebar **Account details → Account ID** (copy button).
- Or read the hex segment in the dashboard URL: `https://dash.cloudflare.com/<account-id>/...`.
- Or run `npx wrangler whoami` (this also verifies the API token works before you push).

**One-time setup**

1. Create the Pages project (name must match the workflow's `--project-name`):
   ```bash
   npx wrangler pages project create zbmsgraph-sdk-js-sandbox --production-branch=main
   ```
   (Or create it once in the Cloudflare dashboard: Workers & Pages → Create → Pages → Direct upload.)
2. In the Entra app registration, add the production URL (the `REDIRECT_URI` value) as a
   **Single-page application** redirect URI, alongside `http://localhost:3000`.

Push to `main` and the workflow deploys to `https://zbmsgraph-sdk-js-sandbox.pages.dev`.