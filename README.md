# Xero P&L Excel Add-in

A proper Excel add-in — appears as a button in the Excel ribbon, opens a sidebar, and pulls live P&L data from Xero directly into your spreadsheet. No Python, no terminal, no Streamlit. Just Excel.

---

## How it works

1. You click **Refresh P&L** in the Excel ribbon
2. A sidebar opens inside Excel
3. You click **Connect to Xero** — a Xero login popup appears
4. After logging in, the add-in reads your spreadsheet's month columns (row 5) and account rows (column A)
5. It fetches the P&L from Xero for each month
6. It writes the figures directly into the yellow cells

---

## What you need to set up (one-time)

### A. GitHub account (free)
Create one at **github.com** if you don't have one.

### B. Netlify account (free)
Sign up at **netlify.com** — use your GitHub account to log in.

### C. Xero developer app
Go to **developer.xero.com** → My Apps → **New App**:
- App name: anything (e.g. "Excel P&L Add-in")
- OAuth 2.0 grant type: **Authorization Code**
- Redirect URI: `https://YOUR-APP.netlify.app/auth-dialog.html`
  *(you'll fill in your real Netlify URL in step 2 below)*

Copy the **Client ID** and **Client Secret**.

---

## Deployment steps

### Step 1 — Upload to GitHub

1. Go to **github.com** → click **+** → **New repository**
2. Name it `xero-excel-addin`, keep it Public
3. Click **uploading an existing file** and drag in all the files from this folder
4. Click **Commit changes**

### Step 2 — Deploy to Netlify

1. Go to **netlify.com** → **Add new site** → **Import an existing project**
2. Choose **GitHub** → select your `xero-excel-addin` repository
3. Leave all settings as defaults → click **Deploy site**
4. Wait ~1 minute. Netlify gives you a URL like `https://amazing-name-123.netlify.app`
5. **Copy this URL** — you need it in the next two steps.

### Step 3 — Add your Xero secrets to Netlify

1. In Netlify dashboard → your site → **Site configuration** → **Environment variables**
2. Add these two variables:
   ```
   XERO_CLIENT_ID     = your-client-id-from-xero
   XERO_CLIENT_SECRET = your-client-secret-from-xero
   ```
3. Click **Save**
4. Go to **Deploys** → **Trigger deploy** → **Deploy site** (so it picks up the new variables)

### Step 4 — Update the Xero redirect URI

1. Go to **developer.xero.com** → My Apps → your app → **Configuration**
2. Under **Redirect URIs**, add: `https://YOUR-APP.netlify.app/auth-dialog.html`
   (replace with your actual Netlify URL)
3. Save

### Step 5 — Update the manifest and app URL

Two places need your Netlify URL:

**In `manifest.xml`:** Replace every `YOUR-NETLIFY-APP` with your actual app name
(e.g. `amazing-name-123`). There are ~8 occurrences — use Find & Replace.

**In `taskpane.js`:** Change line 1:
```js
const APP_URL = "https://YOUR-NETLIFY-APP.netlify.app";
```
to your real URL.

Then commit these changes to GitHub. Netlify will automatically redeploy.

### Step 6 — Sideload the manifest into Excel

"Sideloading" is how you install a custom add-in without going through Microsoft's app store.

**On Windows:**
1. Open Excel
2. Go to **Insert** tab → **Add-ins** → **My Add-ins** → **Manage My Add-ins**
3. Click **Upload My Add-in** → browse to `manifest.xml` → click **Upload**
4. A "Xero" group and "Refresh P&L" button should appear in the **Home** ribbon

**On Mac:**
1. Open Excel
2. Go to **Tools** menu → **Excel Add-ins**
3. Click **Add** or browse for the manifest file
4. (Or use: Insert → Add-ins → My Add-ins → Upload)

**For Excel on the Web (browser):**
1. Open Excel Online
2. Insert → Add-ins → My Add-ins → Upload My Add-in → upload `manifest.xml`

---

## Using the add-in

1. Open your P&L Excel template (the `Xero_PnL_Template.xlsx` from earlier)
2. Click **Refresh P&L** in the Excel ribbon → sidebar opens
3. Enter your Xero **Client ID** and **Client Secret** → click **Save credentials**
4. Click **Connect to Xero** → log in in the popup
5. Click **Fetch from Xero & Update Sheet**
6. Wait for the log to show ✅ Done — cells update automatically

---

## Troubleshooting

| Problem | Fix |
|---|---|
| "Not connected" after login | Check the redirect URI in Xero matches your Netlify URL exactly |
| Token exchange error | Check XERO_CLIENT_ID and XERO_CLIENT_SECRET are set in Netlify environment variables |
| Cells stay yellow | Account name in column A doesn't match Xero exactly — check spelling |
| Add-in doesn't appear in ribbon | Re-upload the manifest.xml, or clear Office cache |
| "Could not open login dialog" | Make sure popup blockers aren't blocking the Xero window |

---

## File guide

| File | What it does |
|---|---|
| `manifest.xml` | Registers the add-in with Excel — tells Excel what to show and where to find it |
| `taskpane.html` | The sidebar UI — the panel that appears inside Excel |
| `taskpane.js` | All the logic: OAuth, reading Excel, calling Xero API, writing cells |
| `auth-dialog.html` | The OAuth popup — opens in a separate window for Xero login |
| `netlify/functions/token.js` | Serverless function: handles the Xero token exchange (avoids CORS) |
| `netlify.toml` | Netlify build config |
| `assets/` | Icons for the ribbon button |
