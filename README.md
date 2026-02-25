# Outlook Signature Generator — Entra ID Edition

pulls employee data from **Microsoft Entra ID** via the Microsoft Graph API instead of Rippling.

---

## Azure Setup (do this first)

### 1. Create an App Registration in Azure Portal

1. Go to **portal.azure.com** and sign in
2. Search for **Entra ID** → **App Registrations** → **New Registration**
3. Name it something like `Signature Generator`
4. Leave **Supported account types** as "Single tenant"
5. No redirect URI needed — click **Register**

### 2. Note your credentials

On the app overview page, copy:
- **Application (client) ID** → this is your `AZURE_CLIENT_ID`
- **Directory (tenant) ID** → this is your `AZURE_TENANT_ID`

### 3. Create a Client Secret

1. In your app registration, go to **Certificates & Secrets**
2. Click **New client secret**
3. Give it a description and expiry (24 months recommended)
4. Click **Add**
5. Copy the **Value** immediately — it's only shown once

This is your `AZURE_CLIENT_SECRET`

### 4. Grant API Permissions

1. Go to **API Permissions** → **Add a permission**
2. Choose **Microsoft Graph** → **Application permissions**
3. Search for and add:
   - `User.Read.All` — read all users' profiles
   - `User.ReadBasic.All` — read profile photos
4. Click **Add permissions**
5. Click **Grant admin consent for [your org]** — this is required for application permissions

### 5. Verify permissions

The permissions page should show both permissions with a green ✓ under **Status** (Granted).

---

## Local Setup

### 1. Install dependencies

```bash
npm install
```

### 2. Configure environment

```bash
cp .env.example .env
```

Fill in your `.env`:

```env
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
COMPANY_NAME=WireStar Networks
COMPANY_WEBSITE=https://wirestar.net
PORT=3000
```

### 3. Run

```bash
npm start
```

Open **http://localhost:3000**

---

## How it works

- Uses **OAuth 2.0 client credentials flow** — the app authenticates as itself using the client secret, no user login required
- Tokens are cached and refreshed automatically when they expire
- Fetches all enabled users from Entra ID using `$filter=accountEnabled eq true`
- Paginates through all pages via `@odata.nextLink`
- Profile photos are fetched from Graph and base64-encoded so they render in email clients

---

## Graph API Fields Used

| Field | Graph property |
|---|---|
| First name | `givenName` |
| Last name | `surname` |
| Job title | `jobTitle` |
| Department | `department` |
| Email | `mail` → fallback `userPrincipalName` |
| Phone | `businessPhones[0]` → fallback `mobilePhone` |
| Active/disabled | `accountEnabled` |
| Profile photo | `/users/{id}/photo/$value` |

---

## Project Structure

```
/
├── server.js              # Express server + Graph API integration
├── .env.example           # Environment variable template
├── .gitignore
├── package.json
├── public/
│   └── index.html         # Frontend UI
└── templates/
    └── signature.html     # WireStar Outlook-compatible signature template
```

---