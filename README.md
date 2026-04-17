# Cognos Analytics Excel Add-in — Sysadmin Deployment Guide

On-premises replacement for IBM Cognos Analysis for Excel (CAFÉ).  
All data, auth tokens, and session state remain on the campus network at all times.

---

## Table of Contents

1. [Architecture Overview](#architecture-overview)
2. [Prerequisites](#prerequisites)
3. [Node.js Version Requirements](#nodejs-version-requirements)
4. [SSL Certificate Configuration](#ssl-certificate-configuration)
5. [AD / Kerberos SPN Registration](#ad--kerberos-spn-registration)
6. [Installation and Configuration](#installation-and-configuration)
7. [Switching Between Mock and Production Cognos](#switching-between-mock-and-production-cognos)
8. [Deployment — Microsoft 365 Admin Center](#deployment--microsoft-365-admin-center)
9. [Deployment — SharePoint App Catalog (Office 2016 / 2019)](#deployment--sharepoint-app-catalog-office-2016--2019)
10. [Deployment — Group Policy (Office 2016 VMs)](#deployment--group-policy-office-2016-vms)
11. [Firewall Requirements](#firewall-requirements)
12. [User Troubleshooting Reference](#user-troubleshooting-reference)
13. [Cognos 11.0.12 End-of-Support Notice](#cognos-11012-end-of-support-notice)

---

## Architecture Overview

```
Excel Task Pane (Office JS)
        |
        | HTTPS (campus network only)
        v
Add-in Host Server (Node.js, on-prem)
  - Serves static HTML/JS/CSS
  - Holds Cognos session tokens in server-side sessions
  - Proxies all Cognos API calls
        |
        | HTTPS + Kerberos / AD credentials
        v
IBM Cognos Analytics 11.0.12 (on-prem)
```

**No data, tokens, or credentials leave the campus network.**  
The Node.js server brokers authentication — the browser never sees a Cognos token.

---

## Prerequisites

- Windows Server 2016 or later (domain-joined to campus AD)
- Node.js 18 LTS or 20 LTS
- Valid SSL certificate issued by the campus internal CA
- Cognos 11.0.12 accessible from the host server
- AD service account for the host server (for Kerberos delegation, if used)

---

## Node.js Version Requirements

Use Node.js **18 LTS** or **20 LTS**. Node 16 will work but is EOL.

```powershell
# Verify version
node --version   # should be v18.x.x or v20.x.x
```

Download: https://nodejs.org (download installer on a campus machine; do not pull from internet on the server if restricted).

---

## SSL Certificate Configuration

The add-in **requires HTTPS** — Office JS will refuse to load task pane content over HTTP in production.

### Using a Campus Internal CA Certificate (production)

1. Request a certificate for the add-in host's FQDN (e.g., `cognos-addin.your-college.edu`) from your campus CA.
2. Place the `.crt` and `.key` files in the `certs/` directory:
   ```
   certs/server.crt
   certs/server.key
   ```
3. If you have a CA bundle (intermediate chain), combine it:
   ```powershell
   type intermediate.crt server.crt > combined.crt
   ```
   Then set `SSL_CERT_PATH=./certs/combined.crt` in `.env`.

4. Set the CA cert path so the Node server can verify Cognos:
   ```
   SSL_CA_CERT_PATH=./certs/campus-ca.crt
   ```

### Client Machines — Trusting the Campus CA

All campus machines joined to the domain should already trust the campus CA via Group Policy.  
If end users see certificate errors, push the campus CA cert via GPO:

```
Computer Configuration → Policies → Windows Settings → Security Settings
→ Public Key Policies → Trusted Root Certification Authorities
```

### Development / Testing — Self-Signed Cert

```bash
bash deploy/gen-dev-certs.sh
```

Trust the generated `certs/server.crt` on each test machine (double-click → Install → Local Machine → Trusted Root Certification Authorities).

---

## AD / Kerberos SPN Registration

The host server uses the user's existing Windows session to pass credentials to Cognos.  
Users enter their AD username and password once in the task pane; the server exchanges them for a Cognos session token via the Cognos REST API `PUT /api/v1/session` endpoint.

### Register an SPN for the Host Server (if using Kerberos delegation)

If you configure Kerberos constrained delegation so the server can impersonate users:

```powershell
# Run as Domain Admin
setspn -A HTTP/cognos-addin.your-college.edu DOMAIN\cognos-addin-svc
setspn -A HTTP/cognos-addin DOMAIN\cognos-addin-svc
```

Configure the service account in Active Directory Users and Computers:
- Account tab → check "Account is trusted for delegation"
- Delegation tab → "Trust this computer for delegation to specified services only" → add the Cognos SPN

Set in `.env`:
```
COGNOS_SPN=HTTP/cognos.your-college.edu@YOUR-COLLEGE.EDU
AD_DOMAIN=YOUR-COLLEGE.EDU
```

If you are **not** using Kerberos delegation (simpler setup), users enter their AD credentials in the task pane and the server passes them to Cognos's `CAMUsername`/`CAMPassword` authentication. This is the default configuration.

---

## Installation and Configuration

```powershell
# 1. Copy project to server
xcopy \\share\cognos-addin C:\apps\cognos-addin /E /I

# 2. Install dependencies (no internet required if node_modules is pre-populated)
cd C:\apps\cognos-addin
npm install

# 3. Copy and edit the environment file
copy .env.example .env
notepad .env
```

### `.env` Key Settings

| Variable | Description |
|---|---|
| `COGNOS_MODE` | `mock` or `production` |
| `COGNOS_BASE_URL` | Production Cognos URL, e.g. `https://cognos.your-college.edu` |
| `COGNOS_MOCK_URL` | Mock server URL (default `https://localhost:3001`) |
| `COGNOS_NAMESPACE` | Cognos AD namespace name (check Cognos Configuration tool) |
| `SERVER_PORT` | Port the add-in server listens on (default `3000`) |
| `SESSION_SECRET` | Long random string — change this before production |
| `SSL_CERT_PATH` | Path to SSL certificate |
| `SSL_KEY_PATH` | Path to SSL private key |
| `SSL_CA_CERT_PATH` | Path to campus CA cert (used to verify Cognos TLS) |

### Start the Server

```powershell
# Run interactively (testing)
node server.js

# Install as a Windows service (production) using node-windows or NSSM:
# NSSM: https://nssm.cc
nssm install CognosAddin "C:\Program Files\nodejs\node.exe" "C:\apps\cognos-addin\server.js"
nssm set CognosAddin AppDirectory "C:\apps\cognos-addin"
nssm start CognosAddin
```

### Run the Mock Server (for testing without Cognos)

```powershell
node mock-server\server.js
```

Or both together:
```powershell
npm run dev
```

---

## Switching Between Mock and Production Cognos

Edit **one line** in `.env`:

```
# Development / testing
COGNOS_MODE=mock

# Production
COGNOS_MODE=production
```

Restart the server after changing. No code changes required.

---

## Manifest Configuration

Before deploying, edit `manifest.xml` and replace every instance of `REPLACE_ADDIN_HOST` with the FQDN of your add-in host server (e.g., `cognos-addin.your-college.edu:3000`).

Also replace `REPLACE_WITH_NEW_GUID` with a freshly generated GUID. Generate one on an offline/campus machine using:

```powershell
[System.Guid]::NewGuid().ToString()
```

---

## Deployment — Microsoft 365 Admin Center

For Microsoft 365 users (Exchange-licensed accounts):

1. Sign in to [Microsoft 365 Admin Center](https://admin.microsoft.com) **on a campus machine**.
2. Go to **Settings → Integrated Apps → Upload custom apps**.
3. Upload `manifest.xml`.
4. Assign to the Finance department security group.
5. Users will see the add-in in Excel under **Home → Finance Reports → Cognos Reports**.

The add-in connects to your on-premises host server only — no data goes to Microsoft (the manifest merely tells Excel where to load the task pane HTML from).

---

## Deployment — SharePoint App Catalog (Office 2016 / 2019)

1. Ensure a SharePoint App Catalog site exists (Central Administration → Apps → Manage App Catalog).
2. Navigate to the App Catalog site → **Apps for Office** document library.
3. Upload `manifest.xml`.
4. In Excel, go to **Insert → My Add-ins → My Organization** — the add-in will appear.

Alternatively, manually insert via the Shared Folder method (see Group Policy below).

---

## Deployment — Group Policy (Office 2016 VMs)

For legacy Office 2016 machines on the domain, use a shared network folder:

1. Copy `manifest.xml` to a read-only network share accessible to all users, e.g.:  
   `\\campus-fileserver\OfficAddins\CognosFinance\manifest.xml`

2. Create or edit a GPO for the Finance OU:
   ```
   User Configuration → Administrative Templates → Microsoft Office 2016
   → Security Settings → Trust Center → Trusted Add-in Catalogs
   ```
   Add the UNC path: `\\campus-fileserver\OfficeAddins\CognosFinance`  
   Enable: "Allow Trusted Locations on the network"

3. Force a Group Policy update:
   ```powershell
   gpupdate /force
   ```

4. Users open Excel → **Insert → My Add-ins → Shared Folder** → Cognos Finance Reports.

---

## Firewall Requirements

| Source | Destination | Port | Protocol | Purpose |
|---|---|---|---|---|
| Client workstations | Add-in host server | 3000 (or your port) | TCP/HTTPS | Task pane loading |
| Add-in host server | Cognos server | 443 | TCP/HTTPS | Cognos REST API calls |
| Add-in host server | AD Domain Controller | 88 | TCP/UDP | Kerberos (if delegation enabled) |
| Add-in host server | AD Domain Controller | 389 / 636 | TCP | LDAP / LDAPS (if needed) |

All traffic stays within the campus network. No outbound internet access is required or should be permitted for this service.

---

## User Troubleshooting Reference

| Symptom | Likely Cause | Resolution |
|---|---|---|
| "Sign In" fails with auth error | Wrong AD credentials or wrong Cognos namespace | Verify `COGNOS_NAMESPACE` in `.env` matches the Cognos Configuration namespace name |
| Task pane shows blank / won't load | SSL cert not trusted on client machine | Install campus CA cert in Windows Trusted Root (see SSL section) |
| "Failed to load" in report browser | Add-in server unreachable | Check server is running; verify firewall allows traffic on configured port |
| Report runs but no data appears | Cognos report has required parameters not supplied | Check parameter panel; ensure Department or Date Range is populated |
| "Session expired" after idle | Cognos session timed out (default 8 hours) | Sign out and sign back in |
| Add-in doesn't appear in Excel | Manifest not deployed or Office cached an old version | In Excel: File → Options → Trust Center → Trusted Add-in Catalogs → clear cache checkbox; restart Excel |

---

## Cognos 11.0.12 End-of-Support Notice

> **IBM ended support for Cognos Analytics 11.0.x on April 30, 2023.**
>
> The institution is running a version that no longer receives IBM security patches, bug fixes, or technical support. This creates the following risks:
>
> - Unpatched security vulnerabilities in the Cognos application tier
> - Inability to obtain vendor support for issues that arise
> - Potential incompatibility with future browser/OS security changes
>
> **Recommended action:** Plan an upgrade path to Cognos Analytics 11.2.x or 12.x (currently supported), or evaluate alternative on-premises reporting platforms compatible with your ERP/SIS. This add-in is written against the standard Cognos REST API (`/api/v1/`) which is available in all supported Cognos versions — upgrading Cognos should not require changes to the add-in itself beyond updating `COGNOS_BASE_URL`.
>
> Contact your IBM account representative or a certified IBM partner for upgrade planning.

---

*Internal use only — not for distribution outside the campus network.*
