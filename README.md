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

You need two values for the repository **Secrets**: a scoped API token and your account ID.

#### `CLOUDFLARE_API_TOKEN`

Use a scoped **Custom Token**, never the Global API Key — the Global Key has full account access
and can't be narrowed.

1. Open [dash.cloudflare.com/profile/api-tokens](https://dash.cloudflare.com/profile/api-tokens)
   (or dashboard → top-right avatar → **My Profile → API Tokens**).
2. Click **Create Token**, then next to **Create Custom Token** click **Get started**.
3. **Token name:** something you'll recognise later, e.g. `pages-deploy-msgraph-sandbox`.
4. **Permissions:** add exactly one row — `Account` → `Cloudflare Pages` → `Edit`. This is all
   `wrangler pages deploy` needs; do **not** add Workers, Zone, or Account Settings scopes.
5. **Account Resources:** `Include` → *your account only* (not "All accounts").
6. **Client IP Address Filtering:** leave empty — GitHub-hosted runners have dynamic IPs.
7. **TTL:** optionally set a start/expiry date so the token is rotated on a schedule.
8. Click **Continue to summary → Create Token**. Copy the token value **now** — Cloudflare shows
   it only once.
9. In GitHub: repo → **Settings → Secrets and variables → Actions → Secrets → New repository
   secret**, name it `CLOUDFLARE_API_TOKEN`, and paste the value.

Verify the token works before you rely on it (a healthy token returns `"status": "active"`):

```bash
curl -s https://api.cloudflare.com/client/v4/user/tokens/verify \
  -H "Authorization: Bearer <your-token>"
```

> Note: Cloudflare's Pages permission is account-wide — this token can deploy to *any* Pages
> project in the account (there is no per-project scoping). For a hard blast-radius limit, use a
> dedicated Cloudflare account for this project.

#### `CLOUDFLARE_ACCOUNT_ID`

This is a public identifier for your account. Find it any of these ways, then add it as the
`CLOUDFLARE_ACCOUNT_ID` repository secret (same GitHub page as above):

- Dashboard → **Workers & Pages** → right sidebar **Account details → Account ID** → copy button.
- Read the hex segment in any dashboard URL: `https://dash.cloudflare.com/<account-id>/...`.
- Run `npx wrangler whoami` — it prints the account name and ID, and confirms the API token works.
  On Windows PowerShell, supply the token first:

  ```powershell
  $env:CLOUDFLARE_API_TOKEN = "<your-token>"; npx wrangler whoami
  ```

**One-time setup**

1. Create the Pages project (name must match the workflow's `--project-name`):
   ```bash
   npx wrangler pages project create zbmsgraph-sdk-js-sandbox --production-branch=main
   ```
   (Or create it once in the Cloudflare dashboard: Workers & Pages → Create → Pages → Direct upload.)
2. In the Entra app registration, add the production URL (the `REDIRECT_URI` value) as a
   **Single-page application** redirect URI, alongside `http://localhost:3000`.

Push to `main` and the workflow deploys to `https://zbmsgraph-sdk-js-sandbox.pages.dev`.