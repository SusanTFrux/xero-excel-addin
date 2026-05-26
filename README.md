# Xero P&L Refresher — No-Install Versions

Two versions — pick whichever suits you best.

---

## Option A: Web App (recommended — truly no install for users)

Users open a URL in their browser. Nothing to download or install.

### Files
- `app.py` — the Streamlit web app
- `requirements.txt` — Python packages needed

### Deploy to Streamlit Community Cloud (free)

1. **Create a free GitHub account** at github.com
2. **Create a new repository** called `xero-pnl-refresher`
3. **Upload these files** to the repository:
   - `app.py`
   - `requirements.txt`
4. **Go to streamlit.io/cloud** and sign in with your GitHub account
5. Click **New app** → select your repository → set main file to `app.py` → Deploy
6. **Add your secrets**: in the Streamlit Cloud dashboard → your app → Settings → Secrets, paste:
   ```
   XERO_CLIENT_ID = "your-client-id"
   XERO_CLIENT_SECRET = "your-client-secret"
   ```
7. **Update your Xero app redirect URI**: go to developer.xero.com → your app →
   Configuration → add your Streamlit app URL as a redirect URI
   (e.g. `https://your-app-name.streamlit.app`)

Your app will be live at `https://your-app-name.streamlit.app`.

### Test locally first (optional)
```bash
pip install -r requirements.txt
streamlit run app.py
```
Add your credentials to `.streamlit/secrets.toml` before running locally.

---

## Option B: Desktop App (standalone .exe / .app — no Python needed to run)

A window-based desktop app that runs without Python installed.
You need Python once to build it — the resulting .exe runs forever without Python.

### Files
- `app_desktop.py` — the desktop app source code

### Build on Windows (creates a .exe)
```
pip install pyinstaller openpyxl python-dateutil
pyinstaller --onefile --windowed --name "Xero PnL Refresher" app_desktop.py
```
Find the finished file at: `dist\Xero PnL Refresher.exe`

### Build on Mac (creates a .app)
```
pip3 install pyinstaller openpyxl python-dateutil
pyinstaller --onefile --windowed --name "Xero PnL Refresher" app_desktop.py
```
Find the finished file at: `dist/Xero PnL Refresher`

### Notes
- The first time Windows users run the .exe, Windows may show a blue
  "Windows protected your PC" warning — click "More info" → "Run anyway"
  This is normal for unsigned apps
- Credentials are saved to your home folder (~/.xero_pnl_config.json)
- Xero tokens are saved to ~/.xero_pnl_token.json — keep this private
- The Xero app redirect URI must stay as http://localhost:8765/callback

---

## Which option should I choose?

| | Web app (A) | Desktop app (B) |
|---|---|---|
| Install needed for users | None | None (after build) |
| Works on any device | Yes | Windows/Mac only |
| Internet required | Always | Only for Xero |
| Data stored | In your browser session | On your computer |
| Share with others | Send them a URL | Send them the .exe |
| Hosting cost | Free (Streamlit Cloud) | None |
