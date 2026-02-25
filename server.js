require("dotenv").config();
const express = require("express");
const axios = require("axios");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = process.env.PORT || 3000;

// ─────────────────────────────────────────────
// Microsoft Graph — Token Management
// ─────────────────────────────────────────────
// We use the OAuth 2.0 client credentials flow to get a bearer token.
// Tokens are cached and refreshed automatically when they expire.

let tokenCache = {
  accessToken: null,
  expiresAt: 0,
};

async function getAccessToken() {
  // Return cached token if still valid (with 60s buffer)
  if (tokenCache.accessToken && Date.now() < tokenCache.expiresAt - 60000) {
    return tokenCache.accessToken;
  }

  const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;

  if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
    throw new Error("Missing Azure credentials. Check AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET in .env");
  }

  const tokenUrl = `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`;

  const params = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: AZURE_CLIENT_ID,
    client_secret: AZURE_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
  });

  const response = await axios.post(tokenUrl, params.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });

  tokenCache.accessToken = response.data.access_token;
  // expires_in is in seconds
  tokenCache.expiresAt = Date.now() + response.data.expires_in * 1000;

  console.log("✓ Microsoft Graph token acquired");
  return tokenCache.accessToken;
}

// ─────────────────────────────────────────────
// Microsoft Graph API client
// ─────────────────────────────────────────────

async function graphGet(url, params = {}) {
  const token = await getAccessToken();
  return axios.get(url, {
    baseURL: "https://graph.microsoft.com/v1.0",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    params,
    timeout: 15000,
  });
}

// ─────────────────────────────────────────────
// Serve static frontend files
// ─────────────────────────────────────────────

app.use(express.static(path.join(__dirname, "public")));

// ─────────────────────────────────────────────
// Helper: Fetch profile photo and convert to base64
// Graph API returns photos as binary — we embed them
// as base64 data URIs so they render in email clients.
// ─────────────────────────────────────────────

async function getPhotoBase64(userId) {
  try {
    const token = await getAccessToken();
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${userId}/photo/$value`,
      {
        headers: { Authorization: `Bearer ${token}` },
        responseType: "arraybuffer",
        timeout: 8000,
      }
    );
    const contentType = response.headers["content-type"] || "image/jpeg";
    const base64 = Buffer.from(response.data).toString("base64");
    return `data:${contentType};base64,${base64}`;
  } catch (err) {
    // 404 means no photo on file — not an error worth logging loudly
    if (err.response?.status !== 404) {
      console.warn(`Could not fetch photo for user ${userId}:`, err.message);
    }
    return null;
  }
}

// ─────────────────────────────────────────────
// Helper: Map a raw Graph API user object to
// only the fields we care about.
// ─────────────────────────────────────────────

function mapEmployee(raw) {
  return {
    id: raw.id,
    firstName: raw.givenName || "",
    lastName: raw.surname || "",
    jobTitle: raw.jobTitle || "",
    phoneNumber: raw.businessPhones?.[0] || raw.mobilePhone || "",
    emailAddress: raw.mail || raw.userPrincipalName || "",
    department: raw.department || "",
    // accountEnabled: false means the account is disabled/terminated
    active: raw.accountEnabled !== false,
  };
}

// ─────────────────────────────────────────────
// GET /api/employees
// Returns all enabled users from Entra ID.
// Paginates through all pages using @odata.nextLink.
// ─────────────────────────────────────────────

app.get("/api/employees", async (req, res) => {
  try {
    // Fields we request from Graph — only ask for what we need
    const $select = "id,givenName,surname,jobTitle,department,mail,userPrincipalName,mobilePhone,businessPhones,accountEnabled";

    let allUsers = [];
    let nextLink = `https://graph.microsoft.com/v1.0/users?$select=${$select}&$top=999&$filter=accountEnabled eq true`;

    // Paginate through all pages using @odata.nextLink
    while (nextLink) {
      const token = await getAccessToken();
      const response = await axios.get(nextLink, {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          // Required for $filter queries
          ConsistencyLevel: "eventual",
        },
        params: nextLink.includes("?") ? {} : {
          $select,
          $top: 999,
          $filter: "accountEnabled eq true",
        },
        timeout: 15000,
      });

      const page = response.data?.value || [];
      allUsers = allUsers.concat(page);
      console.log(`Fetched page — ${page.length} users (total so far: ${allUsers.length})`);

      nextLink = response.data?.["@odata.nextLink"] || null;
    }

    console.log(`Total users fetched from Entra ID: ${allUsers.length}`);

    // Write raw debug file
    const debugPath = path.join(__dirname, "entra_debug.json");
    fs.writeFileSync(debugPath, JSON.stringify(allUsers, null, 2));
    console.log(`Raw response written to ${debugPath}`);

    const employees = allUsers
      .map(mapEmployee)
      .filter((e) => e.firstName || e.lastName)
      .sort((a, b) => `${a.lastName}${a.firstName}`.localeCompare(`${b.lastName}${b.firstName}`));

    console.log(`Employees loaded: ${employees.length}`);

    res.json({ success: true, employees });
  } catch (err) {
    console.error("Error fetching users from Entra ID:", err.message);

    if (err.message.includes("Missing Azure credentials")) {
      return res.status(500).json({ success: false, error: err.message });
    }

    if (err.response?.status === 401) {
      return res.status(401).json({
        success: false,
        error: "Entra ID authentication failed. Check your AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET in .env",
      });
    }

    if (err.response?.status === 403) {
      return res.status(403).json({
        success: false,
        error: "Insufficient permissions. Make sure User.Read.All is granted with admin consent in Azure Portal.",
      });
    }

    res.status(500).json({
      success: false,
      error: "Failed to fetch users from Entra ID.",
      detail: err.message,
    });
  }
});

// ─────────────────────────────────────────────
// GET /api/employees/:id
// Returns full detail for a single user.
// ─────────────────────────────────────────────

app.get("/api/employees/:id", async (req, res) => {
  try {
    const $select = "id,givenName,surname,jobTitle,department,mail,userPrincipalName,mobilePhone,businessPhones,accountEnabled";
    const response = await graphGet(`https://graph.microsoft.com/v1.0/users/${req.params.id}`, { $select });
    const employee = mapEmployee(response.data);
    res.json({ success: true, employee });
  } catch (err) {
    console.error(`Error fetching user ${req.params.id}:`, err.message);

    if (err.response?.status === 404) {
      return res.status(404).json({ success: false, error: "User not found in Entra ID." });
    }

    res.status(500).json({
      success: false,
      error: "Failed to fetch user detail.",
      detail: err.message,
    });
  }
});

// ─────────────────────────────────────────────
// GET /signature/:id
// Fetches user data + photo, renders the signature
// HTML template, and returns a full preview page.
// ─────────────────────────────────────────────

app.get("/signature/:id", async (req, res) => {
  try {
    // 1. Fetch user from Graph
    const $select = "id,givenName,surname,jobTitle,department,mail,userPrincipalName,mobilePhone,businessPhones,accountEnabled";
    const response = await graphGet(`https://graph.microsoft.com/v1.0/users/${req.params.id}`, { $select });
    const employee = mapEmployee(response.data);

    // 2. Fetch and base64-encode profile photo
    const photoBase64 = await getPhotoBase64(req.params.id);

    // 3. Load the signature HTML template
    const templatePath = path.join(__dirname, "templates", "signature.html");
    let template = fs.readFileSync(templatePath, "utf8");

    // 4. Build display values
    const fullName = `${employee.firstName} ${employee.lastName}`.trim();
    const phoneDisplay = employee.phoneNumber ? formatPhoneNumber(employee.phoneNumber) : "";

    // 5. Replace all placeholders
    const replacements = {
      "{{FULL_NAME}}":       fullName || "—",
      "{{FIRST_NAME}}":      employee.firstName || "",
      "{{LAST_NAME}}":       employee.lastName || "",
      "{{JOB_TITLE}}":       employee.jobTitle || "",
      "{{DEPARTMENT}}":      employee.department || "",
      "{{EMAIL_ADDRESS}}":   employee.emailAddress || "",
      "{{PHONE_NUMBER}}":    phoneDisplay,
      "{{PHONE_RAW}}":       employee.phoneNumber || "",
      "{{PHOTO_SRC}}":       photoBase64 || "",
      "{{COMPANY_NAME}}":    process.env.COMPANY_NAME || "",
      "{{COMPANY_WEBSITE}}": process.env.COMPANY_WEBSITE || "#",
      "{{ACCENT_COLOR}}":    process.env.SIGNATURE_ACCENT_COLOR || "#500000",
      "{{YEAR}}":            new Date().getFullYear().toString(),
    };

    for (const [placeholder, value] of Object.entries(replacements)) {
      template = template.split(placeholder).join(value);
    }

    // 6. Wrap in preview page
    res.send(buildSignaturePage(template, fullName, employee.emailAddress));
  } catch (err) {
    console.error(`Error generating signature for ${req.params.id}:`, err.message);

    if (err.code === "ENOENT") {
      return res.status(500).send("<h3>Error: signature.html template not found in /templates</h3>");
    }
    if (err.response?.status === 404) {
      return res.status(404).send("<h3>User not found in Entra ID.</h3>");
    }

    res.status(500).send(`<h3>Error generating signature: ${err.message}</h3>`);
  }
});

// ─────────────────────────────────────────────
// Helper: Format phone number for display
// ─────────────────────────────────────────────

function formatPhoneNumber(raw) {
  const digits = raw.replace(/\D/g, "");
  if (digits.length === 10) {
    return `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;
  }
  if (digits.length === 11 && digits[0] === "1") {
    return `+1 (${digits.slice(1, 4)}) ${digits.slice(4, 7)}-${digits.slice(7)}`;
  }
  return raw;
}

// ─────────────────────────────────────────────
// Helper: Build the full signature preview page
// ─────────────────────────────────────────────

function buildSignaturePage(signatureHtml, fullName, email) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Signature — ${fullName}</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&display=swap');
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'DM Sans', sans-serif; background: #f0f2f5; min-height: 100vh; display: flex; flex-direction: column; align-items: center; padding: 40px 20px; color: #1a1a2e; }
    .page-header { width: 100%; max-width: 760px; margin-bottom: 24px; }
    .page-header h1 { font-size: 22px; font-weight: 600; color: #1a1a2e; }
    .page-header p { font-size: 14px; color: #666; margin-top: 4px; }
    .signature-card { background: #fff; border-radius: 12px; padding: 40px; width: 100%; max-width: 760px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); }
    .actions { display: flex; gap: 12px; margin-top: 28px; flex-wrap: wrap; }
    .btn { display: inline-flex; align-items: center; gap: 8px; padding: 10px 20px; border-radius: 8px; font-family: 'DM Sans', sans-serif; font-size: 14px; font-weight: 500; cursor: pointer; border: none; transition: all 0.15s ease; text-decoration: none; }
    .btn-primary { background: #500000; color: white; }
    .btn-primary:hover { background: #3a0000; }
    .btn-success { background: #16a34a; color: white; }
    .btn-secondary { background: #f0f2f5; color: #1a1a2e; }
    .btn-secondary:hover { background: #e2e5ea; }
    .divider { height: 1px; background: #eee; margin: 28px 0; }
    .label { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; color: #999; margin-bottom: 16px; }
    .instructions { margin-top: 28px; padding: 16px 20px; background: #f8f9ff; border-left: 3px solid #500000; border-radius: 0 8px 8px 0; font-size: 13px; color: #444; line-height: 1.9; }
    .instructions strong { color: #1a1a2e; }
    .instructions ol { padding-left: 18px; margin-top: 6px; }
    .instructions ol li { margin-bottom: 4px; }
    .instructions .section { margin-bottom: 16px; }
    .instructions .section:last-child { margin-bottom: 0; }
    .toast { position: fixed; bottom: 30px; right: 30px; background: #1a1a2e; color: white; padding: 12px 20px; border-radius: 8px; font-size: 14px; opacity: 0; transform: translateY(8px); transition: all 0.25s ease; pointer-events: none; }
    .toast.show { opacity: 1; transform: translateY(0); }
    #sig-preview { user-select: all; cursor: default; }
  </style>
</head>
<body>
  <div class="page-header">
    <h1>Email Signature</h1>
    <p>${fullName} &mdash; ${email}</p>
  </div>
  <div class="signature-card">
    <div class="label">Your Signature — Click to select, then copy</div>
    <div id="sig-preview">${signatureHtml}</div>
    <div class="divider"></div>
    <div class="actions">
      <button class="btn btn-primary" id="copyBtn" onclick="copySignature()">Copy to Clipboard</button>
      <button class="btn btn-secondary" onclick="window.close()">✕ Close</button>
    </div>
    <div class="instructions">
      <div class="section">
        <strong>How to add your signature in Outlook Desktop:</strong>
        <ol>
          <li>Click <strong>Copy to Clipboard</strong> above</li>
          <li>Open Outlook and go to <strong>File → Options → Mail → Signatures</strong></li>
          <li>Click <strong>New</strong> and give your signature a name</li>
          <li>Click inside the signature editor box and press <strong>Ctrl + V</strong> to paste</li>
          <li>Click <strong>OK</strong> to save</li>
        </ol>
      </div>
      <div class="section">
        <strong>How to add your signature in Outlook Web:</strong>
        <ol>
          <li>Click <strong>Copy to Clipboard</strong> above</li>
          <li>In Outlook, click the <strong>gear icon</strong> next to your profile photo in the top right</li>
          <li>Go to <strong>Account → Signatures</strong></li>
          <li>Click <strong>New signature</strong>, give it a name, then press <strong>Ctrl + V</strong> to paste</li>
          <li>Set it as the default for <strong>New messages</strong> and <strong>Replies &amp; forwards</strong></li>
          <li>Click <strong>Save</strong></li>
        </ol>
      </div>
    </div>
  </div>
  <div class="toast" id="toast">✓ Signature copied to clipboard!</div>
  <script>
    async function copySignature() {
      const preview = document.getElementById('sig-preview');
      const btn = document.getElementById('copyBtn');

      try {
        // Use the Clipboard API to copy rich content (HTML + text with images)
        // This preserves formatting and the profile photo when pasting into Outlook
        const htmlContent = preview.innerHTML;
        const textContent = preview.innerText;

        const clipboardItem = new ClipboardItem({
          'text/html': new Blob([htmlContent], { type: 'text/html' }),
          'text/plain': new Blob([textContent], { type: 'text/plain' }),
        });

        await navigator.clipboard.write([clipboardItem]);
        showToast();
        btn.textContent = '✓ Copied!';
        btn.className = 'btn btn-success';
        setTimeout(() => {
          btn.innerHTML = '📋 Copy to Clipboard';
          btn.className = 'btn btn-primary';
        }, 2500);

      } catch (err) {
        // Fallback: select the signature element so user can manually copy
        console.warn('Clipboard API not available, falling back to selection:', err);
        const range = document.createRange();
        range.selectNodeContents(preview);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);
        btn.textContent = '✓ Selected — press Ctrl+C to copy';
        btn.className = 'btn btn-success';
        setTimeout(() => {
          btn.innerHTML = '📋 Copy to Clipboard';
          btn.className = 'btn btn-primary';
        }, 3000);
      }
    }

    function showToast() {
      const t = document.getElementById('toast');
      t.classList.add('show');
      setTimeout(() => t.classList.remove('show'), 2500);
    }
  </script>
</body>
</html>`;
}

// ─────────────────────────────────────────────
// Health check
// ─────────────────────────────────────────────

app.get("/health", (req, res) => {
  res.json({
    status: "ok",
    company: process.env.COMPANY_NAME || "unknown",
    azureConfigured: !!(process.env.AZURE_TENANT_ID && process.env.AZURE_CLIENT_ID && process.env.AZURE_CLIENT_SECRET),
  });
});

// ─────────────────────────────────────────────
// Start server
// ─────────────────────────────────────────────

app.listen(PORT, () => {
  const azureReady = !!(process.env.AZURE_TENANT_ID && process.env.AZURE_CLIENT_ID && process.env.AZURE_CLIENT_SECRET);
  console.log(`\n✅  Signature Generator (Entra ID) running at http://localhost:${PORT}`);
  console.log(`    Company: ${process.env.COMPANY_NAME || "(not set — check .env)"}`);
  console.log(`    Azure credentials: ${azureReady ? "✓ configured" : "✗ MISSING — check .env"}\n`);
});
