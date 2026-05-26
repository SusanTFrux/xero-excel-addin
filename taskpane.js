/*
  taskpane.js
  ===========
  All the logic for the Excel add-in sidebar.

  What this file does:
    1. Waits for Office to be ready (Office.onReady)
    2. Manages saving/loading Xero credentials in the browser's localStorage
    3. Handles the Xero OAuth login flow using Office's Dialog API
    4. Reads the spreadsheet to find month columns and account rows
    5. Calls the Xero P&L API for each month
    6. Writes the results back into the correct Excel cells

  The token exchange (swapping the one-time auth code for an access token)
  is handled by a Netlify serverless function (netlify/functions/token.js)
  because Xero's token endpoint blocks direct browser requests for security reasons.
  Everything else — reading Excel, calling the Xero data API — happens right here.
*/

"use strict";

// ── Replace with your actual Netlify app URL ────────────────────────────────
// e.g. "https://xero-pnl.netlify.app"
const APP_URL = "https://YOUR-NETLIFY-APP.netlify.app";

// Xero OAuth endpoints
const XERO_AUTH_URL = "https://login.xero.com/identity/connect/authorize";
const XERO_API_BASE = "https://api.xero.com/api.xro/2.0";
const XERO_TENANTS  = "https://api.xero.com/connections";
const SCOPES = "accounting.reports.profitandloss.read accounting.settings.read offline_access";

// Keys used for localStorage — where we save credentials and tokens
const KEY_CLIENT_ID     = "xero_client_id";
const KEY_CLIENT_SECRET = "xero_client_secret";
const KEY_TOKEN         = "xero_token";
const KEY_TENANT_ID     = "xero_tenant_id";
const KEY_TENANT_NAME   = "xero_tenant_name";

// Used to pass the PKCE code verifier from buildAuthUrl() to the OAuth callback
let _pkceVerifier = null;
let _oauthState   = null;
let _authDialog   = null;  // reference to the Office Dialog window

// ── Month name → number lookup table ────────────────────────────────────────
const MONTH_MAP = {
  jan:1, feb:2, mar:3, apr:4, may:5, jun:6,
  jul:7, aug:8, sep:9, oct:10, nov:11, dec:12,
  january:1, february:2, march:3, april:4, june:6,
  july:7, august:8, september:9, october:10, november:11, december:12
};


// ══════════════════════════════════════════════════════════════════════════
// STARTUP — waits for Office.js to be ready before doing anything
// ══════════════════════════════════════════════════════════════════════════

Office.onReady(function(info) {
  /*
    Office.onReady fires once the Office.js library has finished loading
    and Excel (or Word, etc.) is ready to accept our API calls.
    We must not call any Office API before this fires.
  */
  if (info.host === Office.HostType.Excel) {
    // Restore saved credentials into the input fields
    const savedId     = localStorage.getItem(KEY_CLIENT_ID)     || "";
    const savedSecret = localStorage.getItem(KEY_CLIENT_SECRET) || "";
    document.getElementById("clientId").value     = savedId;
    document.getElementById("clientSecret").value = savedSecret;

    // Restore connection state from a previous session
    updateConnectionUI();

    log("Office.js ready. Add-in loaded in Excel.");
  }
});


// ══════════════════════════════════════════════════════════════════════════
// CREDENTIALS — save/load from localStorage
// ══════════════════════════════════════════════════════════════════════════

function saveCreds() {
  /*
    localStorage is like a small text database in the browser.
    Data saved here persists between sessions on the same machine.
    We never send credentials to our server — they stay in the browser.
  */
  const cid    = document.getElementById("clientId").value.trim();
  const secret = document.getElementById("clientSecret").value.trim();

  if (!cid || !secret) {
    alert("Please enter both Client ID and Client Secret.");
    return;
  }

  localStorage.setItem(KEY_CLIENT_ID,     cid);
  localStorage.setItem(KEY_CLIENT_SECRET, secret);
  log("✅ Credentials saved.");
}

function getClientId()     { return localStorage.getItem(KEY_CLIENT_ID)     || ""; }
function getClientSecret() { return localStorage.getItem(KEY_CLIENT_SECRET) || ""; }


// ══════════════════════════════════════════════════════════════════════════
// TOKEN MANAGEMENT — load, check expiry
// ══════════════════════════════════════════════════════════════════════════

function getToken() {
  /*
    The access token is the "visitor badge" Xero gives us after login.
    We save it to localStorage so we don't have to log in every time.
    Tokens expire after 30 minutes — we check that here.
  */
  try {
    const raw = localStorage.getItem(KEY_TOKEN);
    if (!raw) return null;
    const tok = JSON.parse(raw);
    // Check if the token is still valid (with 60-second buffer)
    const expiresAt = (tok.obtained_at || 0) + (tok.expires_in || 1800) - 60;
    return Date.now() / 1000 < expiresAt ? tok : null;
  } catch(e) {
    return null;
  }
}

function saveToken(tok) {
  tok.obtained_at = Date.now() / 1000;  // record when we got it (Unix timestamp)
  localStorage.setItem(KEY_TOKEN, JSON.stringify(tok));
}

function isConnected() {
  return !!getToken() && !!localStorage.getItem(KEY_TENANT_ID);
}

function disconnect() {
  // Clear all saved auth data
  [KEY_TOKEN, KEY_TENANT_ID, KEY_TENANT_NAME].forEach(k => localStorage.removeItem(k));
  updateConnectionUI();
  log("Disconnected from Xero.");
}


// ══════════════════════════════════════════════════════════════════════════
// UI STATE — update the sidebar to reflect current connection
// ══════════════════════════════════════════════════════════════════════════

function updateConnectionUI() {
  const connected  = isConnected();
  const orgName    = localStorage.getItem(KEY_TENANT_NAME) || "";
  const statusEl   = document.getElementById("connStatus");
  const connectBtn = document.getElementById("connectBtn");
  const disconnBtn = document.getElementById("disconnectBtn");
  const refreshBtn = document.getElementById("refreshBtn");

  if (connected) {
    statusEl.textContent  = `● Connected: ${orgName}`;
    statusEl.className    = "status status-connected";
    connectBtn.style.display  = "none";
    disconnBtn.style.display  = "block";
    refreshBtn.disabled       = false;
  } else {
    statusEl.textContent  = "● Not connected";
    statusEl.className    = "status status-disconnected";
    connectBtn.style.display  = "block";
    disconnBtn.style.display  = "none";
    refreshBtn.disabled       = true;
  }
}

function setProgress(pct, label) {
  const wrap = document.getElementById("progressWrap");
  const bar  = document.getElementById("progressBar");
  const lbl  = document.getElementById("progressLabel");
  wrap.style.display = "block";
  bar.style.width    = pct + "%";
  lbl.textContent    = label;
  if (pct >= 100) {
    setTimeout(() => { wrap.style.display = "none"; }, 2000);
  }
}

function log(msg, type = "") {
  /*
    Appends a line to the log panel at the bottom of the sidebar.
    type can be "ok", "warn", or "err" to colour-code the message.
  */
  const area = document.getElementById("logArea");
  const line = document.createElement("div");
  line.className = "log-line" + (type ? " log-" + type : "");
  const time = new Date().toLocaleTimeString("en-NZ", { hour12: false });
  line.textContent = `[${time}]  ${msg}`;
  area.appendChild(line);
  area.scrollTop = area.scrollHeight;  // auto-scroll to the newest line
}


// ══════════════════════════════════════════════════════════════════════════
// OAUTH — connect to Xero using the Office Dialog API
// ══════════════════════════════════════════════════════════════════════════

function generateCodeVerifier() {
  /*
    Creates a random 64-character string used in the PKCE security flow.
    PKCE prevents someone intercepting the auth code from using it — they
    would also need this secret verifier, which never leaves our device.
  */
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~";
  let result = "";
  const arr = new Uint8Array(64);
  crypto.getRandomValues(arr);  // cryptographically secure random numbers
  arr.forEach(n => result += chars[n % chars.length]);
  return result;
}

async function generateCodeChallenge(verifier) {
  /*
    Creates a SHA-256 hash of the verifier, then base64url-encodes it.
    This "challenge" is sent to Xero in the auth URL.
    Xero stores it. When we later send the verifier, Xero hashes it
    and confirms it matches — proving we made the original request.
  */
  const data      = new TextEncoder().encode(verifier);
  const hashBuf   = await crypto.subtle.digest("SHA-256", data);
  const hashArr   = Array.from(new Uint8Array(hashBuf));
  const hashStr   = String.fromCharCode(...hashArr);
  const b64       = btoa(hashStr);
  // Convert from base64 to base64url (replace + → -, / → _, remove =)
  return b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
}

async function connectXero() {
  /*
    Starts the Xero OAuth login flow:
    1. Generate PKCE verifier + challenge
    2. Build the Xero auth URL
    3. Open it in an Office Dialog popup
    4. Wait for the dialog to send us back the auth code
    5. Exchange the code for a token (via our Netlify function)
    6. Fetch the list of Xero organisations and store the first one
  */
  if (!getClientId() || !getClientSecret()) {
    // Make sure they've entered credentials first
    saveCreds();
    if (!getClientId()) return;
  }

  // Generate PKCE pair
  _pkceVerifier = generateCodeVerifier();
  _oauthState   = Math.random().toString(36).slice(2);  // random state for CSRF protection
  const challenge = await generateCodeChallenge(_pkceVerifier);

  // Build the Xero login URL
  const params = new URLSearchParams({
    response_type:         "code",
    client_id:             getClientId(),
    redirect_uri:          `${APP_URL}/auth-dialog.html`,
    scope:                 SCOPES,
    state:                 _oauthState,
    code_challenge:        challenge,
    code_challenge_method: "S256",
  });
  const authUrl = `${XERO_AUTH_URL}?${params}`;

  log("Opening Xero login…");

  /*
    Office.context.ui.displayDialogAsync opens a popup window for OAuth.
    This is how Office add-ins handle external logins — the popup is a
    real browser window, not the same sandboxed context as the task pane.

    When auth-dialog.html finishes, it calls Office.context.ui.messageParent()
    to send a message back to this task pane, which we handle below.
  */
  Office.context.ui.displayDialogAsync(
    authUrl,
    { height: 60, width: 40, displayInIframe: false },
    function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        log("❌ Could not open login dialog: " + asyncResult.error.message, "err");
        return;
      }

      _authDialog = asyncResult.value;

      // Listen for messages from the auth-dialog.html popup
      _authDialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        onDialogMessage
      );

      // Listen for the dialog being closed by the user
      _authDialog.addEventHandler(
        Office.EventType.DialogEventReceived,
        function(args) {
          if (args.error === 12006) {
            log("Login dialog closed by user.", "warn");
          }
        }
      );
    }
  );
}

async function onDialogMessage(args) {
  /*
    Called when auth-dialog.html sends us a message via messageParent().
    The message is a JSON string containing either:
      { type: "code", code: "...", state: "..." }   — auth success
      { type: "error", message: "..." }              — something went wrong
  */
  _authDialog.close();

  let msg;
  try {
    msg = JSON.parse(args.message);
  } catch(e) {
    log("❌ Unexpected message from auth dialog.", "err");
    return;
  }

  if (msg.type === "error") {
    log("❌ Login failed: " + msg.message, "err");
    return;
  }

  if (msg.type !== "code") {
    log("❌ Unexpected message type: " + msg.type, "err");
    return;
  }

  // Verify the state matches what we sent — CSRF protection
  if (msg.state !== _oauthState) {
    log("❌ Security check failed (state mismatch). Please try again.", "err");
    return;
  }

  log("Auth code received. Exchanging for access token…");

  // Exchange the one-time code for an access token
  // This goes via our Netlify Function because Xero's token endpoint
  // blocks direct browser-to-Xero requests (CORS restriction)
  try {
    const resp = await fetch(`${APP_URL}/.netlify/functions/token`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        code:         msg.code,
        verifier:     _pkceVerifier,
        redirect_uri: `${APP_URL}/auth-dialog.html`,
      })
    });

    if (!resp.ok) {
      const err = await resp.text();
      throw new Error(`Token exchange failed (${resp.status}): ${err}`);
    }

    const token = await resp.json();
    saveToken(token);
    log("✅ Token received and saved.");

    // Now find out which Xero organisation(s) this token can access
    await loadTenants(token.access_token);
    updateConnectionUI();

  } catch(e) {
    log("❌ " + e.message, "err");
  }
}

async function loadTenants(accessToken) {
  /*
    Asks Xero which organisations (called "tenants") this token can access.
    Most users will have just one — we use the first one automatically.
    If they have multiple, we could add a selector in a future version.
  */
  const resp = await fetch(XERO_TENANTS, {
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Accept":        "application/json"
    }
  });

  if (!resp.ok) throw new Error("Could not fetch Xero organisations.");

  const tenants = await resp.json();
  if (!tenants || tenants.length === 0) {
    throw new Error("No Xero organisations found for this account.");
  }

  // Use the first organisation
  localStorage.setItem(KEY_TENANT_ID,   tenants[0].tenantId);
  localStorage.setItem(KEY_TENANT_NAME, tenants[0].tenantName);
  log(`✅ Connected to: ${tenants[0].tenantName}`, "ok");
}


// ══════════════════════════════════════════════════════════════════════════
// EXCEL — read the spreadsheet to find months and accounts
// ══════════════════════════════════════════════════════════════════════════

function parseMonthHeader(value) {
  /*
    Converts a cell value like "May 2026" into { year: 2026, month: 5 }.
    Returns null if the cell doesn't look like a month header.
    Handles Excel date objects and plain text strings.
  */
  if (!value) return null;

  // Excel sometimes returns dates as JavaScript Date objects
  if (value instanceof Date) {
    return { year: value.getFullYear(), month: value.getMonth() + 1 };
  }

  if (typeof value !== "string") return null;

  const parts = value.trim().split(/\s+/);  // split on whitespace
  if (parts.length !== 2) return null;

  const monthNum = MONTH_MAP[parts[0].slice(0, 3).toLowerCase()];
  const year     = parseInt(parts[1], 10);

  if (monthNum && year >= 2000 && year <= 2100) {
    return { year, month: monthNum };
  }
  return null;
}

async function readSpreadsheetConfig() {
  /*
    Uses the Excel JavaScript API to read the active worksheet and find:
      - periods:  which columns have month headers in row 5
      - accounts: which rows have account names in column A (non-bold only)

    Returns { periods, accounts } where:
      periods  = [{ col: 2, colLetter: "B", year: 2026, month: 5 }, ...]
      accounts = [{ row: 8, name: "Sales" }, { row: 9, name: "Advertising" }, ...]
  */
  return await Excel.run(async context => {
    /*
      Excel.run() gives us a context object. We use it to queue up read
      operations, then call context.sync() to actually execute them all
      at once (batching is more efficient than one request per cell).
    */
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // ── Read row 5 (month headers) ────────────────────────────────────────
    // We read a wide range — up to column Z (column 26)
    const headerRange = sheet.getRange("A5:Z5");
    headerRange.load("values");          // queue a read of the cell values
    await context.sync();                // execute all queued reads

    const headerValues = headerRange.values[0];  // array of values from row 5
    const periods = [];

    // Start from index 1 (column B) — column A (index 0) is "Account"
    for (let i = 1; i < headerValues.length; i++) {
      const parsed = parseMonthHeader(headerValues[i]);
      if (parsed) {
        // Convert column index (0-based) to Excel column letter
        const colLetter = String.fromCharCode(65 + i);  // 0 → A, 1 → B, etc.
        periods.push({ col: i + 1, colLetter, ...parsed });
      }
    }

    if (periods.length === 0) {
      throw new Error(
        "No month headers found in row 5.\n\n" +
        "Expected format: 'May 2026', 'Apr 2026' etc."
      );
    }

    // ── Read column A (account names) ─────────────────────────────────────
    // Read from row 6 down to row 60 (adjust if you have more rows)
    const accountRange = sheet.getRange("A6:A60");
    accountRange.load("values, format/font/bold");  // read values AND bold formatting
    await context.sync();

    const accounts = [];
    const rowValues = accountRange.values;      // array of arrays: [[val], [val], ...]
    const boldFlags = accountRange.format.font.bold;  // array: [true/false, ...]

    for (let i = 0; i < rowValues.length; i++) {
      const val  = rowValues[i][0];
      const bold = boldFlags[i][0];
      if (!val) continue;          // skip empty cells
      if (bold) continue;          // skip bold cells (section headers and totals)
      accounts.push({ row: i + 6, name: String(val).trim() });
    }

    return { periods, accounts };
  });
}


// ══════════════════════════════════════════════════════════════════════════
// XERO DATA — fetch P&L for a single month
// ══════════════════════════════════════════════════════════════════════════

function formatDate(year, month, day) {
  /*
    Formats a date as "YYYY-MM-DD" — the format Xero's API expects.
    padStart(2, "0") adds a leading zero if the number is single-digit
    (e.g. month 5 → "05").
  */
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function daysInMonth(year, month) {
  /*
    Returns the number of days in a given month.
    new Date(year, month, 0) gives the last day of the previous month,
    which is the same as the last day of month-1 in 0-indexed terms.
    Example: daysInMonth(2026, 2) → 28 (February 2026)
  */
  return new Date(year, month, 0).getDate();
}

async function fetchPnL(year, month) {
  /*
    Calls the Xero P&L report API for one month.
    Returns a map of { "account name (lowercase)": amount }.

    Xero's P&L response is a nested structure we have to walk through
    to extract the individual account line items.
  */
  const token    = getToken();
  const tenantId = localStorage.getItem(KEY_TENANT_ID);

  const fromDate = formatDate(year, month, 1);
  const toDate   = formatDate(year, month, daysInMonth(year, month));

  const url = `${XERO_API_BASE}/Reports/ProfitAndLoss?` +
    `fromDate=${fromDate}&toDate=${toDate}&standardLayout=true&paymentsOnly=false`;

  const resp = await fetch(url, {
    headers: {
      "Authorization":  `Bearer ${token.access_token}`,
      "Xero-Tenant-Id": tenantId,
      "Accept":         "application/json"
    }
  });

  if (!resp.ok) {
    throw new Error(`Xero API error ${resp.status} for ${fromDate}`);
  }

  const data    = await resp.json();
  const results = {};

  // Walk the nested rows structure to find individual account amounts
  function walk(rows) {
    for (const row of rows) {
      const rt = row.RowType || "";
      if (rt === "Section" || rt === "SummaryRow") {
        // Sections contain child rows — recurse into them
        walk(row.Rows || []);
      } else if (rt === "Row") {
        const cells = row.Cells || [];
        if (cells.length >= 2) {
          const name = (cells[0].Value || "").trim();
          const val  = (cells[1].Value || "0").replace(/,/g, "");  // remove commas
          if (name) {
            results[name.toLowerCase()] = parseFloat(val) || 0;
          }
        }
      }
    }
  }

  walk((data.Reports?.[0]?.Rows) || []);
  return results;
}


// ══════════════════════════════════════════════════════════════════════════
// MAIN REFRESH — ties everything together
// ══════════════════════════════════════════════════════════════════════════

async function runRefresh() {
  /*
    The main function that runs when the user clicks "Fetch from Xero":
    1. Reads the spreadsheet to find months and accounts
    2. Fetches Xero P&L data for each month
    3. Writes the amounts back into the correct cells
  */

  if (!isConnected()) {
    log("❌ Not connected to Xero. Please connect first.", "err");
    return;
  }

  // Disable the button while running so it can't be clicked twice
  const btn = document.getElementById("refreshBtn");
  btn.disabled = true;

  try {
    // ── Step 1: Read the spreadsheet ─────────────────────────────────────
    log("Reading spreadsheet…");
    setProgress(5, "Reading spreadsheet…");
    const { periods, accounts } = await readSpreadsheetConfig();
    log(`Found ${periods.length} month columns and ${accounts.length} account rows.`);

    // ── Step 2: Fetch data from Xero for each month ───────────────────────
    const allData = {};  // { "2026-5": { "sales": 85000, ... }, ... }

    for (let i = 0; i < periods.length; i++) {
      const { year, month } = periods[i];
      const label = `${new Date(year, month - 1).toLocaleString("en-NZ", { month: "long" })} ${year}`;
      const pct   = 10 + Math.round((i / periods.length) * 70);

      setProgress(pct, `Fetching ${label}…`);
      log(`Fetching ${label}…`);

      try {
        allData[`${year}-${month}`] = await fetchPnL(year, month);
        log(`  → ${Object.keys(allData[`${year}-${month}`]).length} accounts found.`);
      } catch(e) {
        log(`  ⚠️ ${label}: ${e.message}`, "warn");
        allData[`${year}-${month}`] = {};
      }
    }

    // ── Step 3: Write data back to Excel ─────────────────────────────────
    setProgress(85, "Writing to spreadsheet…");
    log("Writing data to spreadsheet…");

    let written  = 0;
    const notFound = [];

    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Update the "Refreshed from Xero" line in cell A3
      const org = localStorage.getItem(KEY_TENANT_NAME) || "Xero";
      const ts  = new Date().toLocaleString("en-NZ");
      sheet.getRange("A2").values = [[org]];
      sheet.getRange("A3").values = [[`Refreshed from Xero: ${ts}`]];

      for (const { row, name } of accounts) {
        const key = name.toLowerCase();

        for (const { col, colLetter, year, month } of periods) {
          const periodData = allData[`${year}-${month}`] || {};
          const cellAddr   = `${colLetter}${row}`;
          const cell       = sheet.getRange(cellAddr);

          // Try exact match first, then partial match
          let amount = null;
          if (key in periodData) {
            amount = periodData[key];
          } else {
            const matches = Object.entries(periodData)
              .filter(([k]) => k.includes(key) || key.includes(k));
            if (matches.length === 1) amount = matches[0][1];
          }

          if (amount !== null) {
            // Write the number and format it as currency
            cell.values        = [[amount]];
            cell.numberFormat  = [['#,##0;(#,##0);"-"']];
            // Clear the yellow background — data has been populated
            cell.format.fill.color = "#FFFFFF";
            written++;
          } else {
            // Leave cell yellow — account not found in Xero
            cell.values = [[null]];
            cell.format.fill.color = "#FFFF00";
            if (!notFound.includes(name)) notFound.push(name);
          }
        }
      }

      // Execute all the writes in one batch (much faster than one at a time)
      await context.sync();
    });

    // ── Summary ────────────────────────────────────────────────────────────
    setProgress(100, "Done!");
    log(`✅ Done — ${written} cells updated.`, "ok");

    if (notFound.length > 0) {
      log(`⚠️ ${notFound.length} account(s) not matched — cells left yellow:`, "warn");
      notFound.forEach(n => log(`   • ${n}`, "warn"));
      log("   Check spelling matches your Xero chart of accounts exactly.", "warn");
    }

  } catch(e) {
    log("❌ Error: " + e.message, "err");
    setProgress(0, "");
    console.error(e);
  } finally {
    btn.disabled = false;
  }
}
