# GMOO Excel Add-in — Install & Troubleshooting Guide

This guide is for customers installing the GMOO Excel add-in against
their own globalMOO API server (on-prem, private cloud, or
`https://app.globalmoo.com`).

---

## 1. How the add-in is delivered

The add-in is an **Office Web Add-in**, not a downloaded executable.
The `install.ps1` script registers a small manifest XML file on your
machine; that manifest tells Excel to load the task pane HTML/JS live
from `https://globalmoo.github.io/gmoo-excel-plugin/` every time you
open it.

**What this means for your data:**

- `globalmoo.github.io` only serves static code — HTML, CSS, JS.
- Once the page is loaded inside Excel's WebView, the code runs on
  **your** machine, in your Excel process.
- Every API call (project data, API key, results) flows **directly**
  from your Excel to whatever API URL you configure. github.io has no
  backend, sees none of it.
- Updates roll out automatically the next time you open the task pane;
  there is no reinstall step.

---

## 2. Quick install

Open PowerShell (not Excel) and run:

```powershell
irm https://globalmoo.github.io/gmoo-excel-plugin/install.ps1 | iex
```

The script will:

1. Prompt for your API URL (e.g. `https://10.0.0.5/api/` or
   `https://app.globalmoo.com/api/`).
2. Probe TLS trust. If your server uses a self-signed or private-CA
   cert, the script offers to import it into your `CurrentUser\Root`
   store (no admin rights needed).
3. Download and register the manifest under
   `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`.
4. Tell you to open Excel and add a Connection in the task pane.

You can also pre-supply parameters non-interactively:

```powershell
$args = @{
    ApiUrl        = 'https://10.0.0.5/api/'
    CertFile      = 'C:\acme-ca.crt'   # optional
    NoInteractive = $true
}
& ([scriptblock]::Create((irm https://globalmoo.github.io/gmoo-excel-plugin/install.ps1))) @args
```

---

## 3. Server requirements

For the add-in to work against your server, two things must be true:

### 3a. The server's TLS cert must be trusted by Windows

The Excel WebView uses the Windows certificate store. A cert from a
public CA (Let's Encrypt, DigiCert, etc.) works out of the box.

For a self-signed or private-CA cert, you have three options:

1. **Recommended.** Run `install.ps1` (above). It prompts you to trust
   the cert into your user-level Trusted Root store. No admin needed,
   only affects your own Windows account.
2. **Manual.** Open `certmgr.msc` → *Trusted Root Certification
   Authorities* → *Certificates* → right-click → *All Tasks* → *Import*
   → select your `.crt`/`.cer`/`.pem` file.
3. **Org-wide via Group Policy.** Push the CA cert to user/machine
   Trusted Root stores. (Out of scope for this doc; ask your IT.)

**Important caveats:**

- The cert's **Common Name (CN) or Subject Alternative Name (SAN)** must
  match the hostname or IP you enter in the Connection. A cert valid for
  `server.local` will be rejected when you connect to `https://1.1.1.1`,
  even if the cert is trusted.
- After importing a cert, restart Excel so the WebView re-reads the
  trust store.

### 3b. The server must allow CORS from `https://globalmoo.github.io`

Because the task pane HTML is served from `globalmoo.github.io`, every
API call is a cross-origin request. Your server must respond with the
right CORS headers, or the browser will silently drop the response.

**Minimum required headers on all API responses:**

```
Access-Control-Allow-Origin:  https://globalmoo.github.io
Access-Control-Allow-Headers: Authorization, Content-Type
Access-Control-Allow-Methods: GET, POST, PUT, DELETE, OPTIONS
```

And the server must respond `2xx` to `OPTIONS` preflight requests on
every API path.

#### nginx

```nginx
location /api/ {
    add_header Access-Control-Allow-Origin  "https://globalmoo.github.io" always;
    add_header Access-Control-Allow-Headers "Authorization, Content-Type" always;
    add_header Access-Control-Allow-Methods "GET, POST, PUT, DELETE, OPTIONS" always;

    if ($request_method = OPTIONS) { return 204; }

    proxy_pass http://your-backend;
}
```

#### Apache (with `mod_headers` and `mod_rewrite`)

```apache
<Location /api/>
    Header always set Access-Control-Allow-Origin  "https://globalmoo.github.io"
    Header always set Access-Control-Allow-Headers "Authorization, Content-Type"
    Header always set Access-Control-Allow-Methods "GET, POST, PUT, DELETE, OPTIONS"

    RewriteEngine On
    RewriteCond %{REQUEST_METHOD} OPTIONS
    RewriteRule ^ - [R=204,L]
</Location>
```

#### IIS (`web.config`)

```xml
<system.webServer>
  <httpProtocol>
    <customHeaders>
      <add name="Access-Control-Allow-Origin"  value="https://globalmoo.github.io" />
      <add name="Access-Control-Allow-Headers" value="Authorization, Content-Type" />
      <add name="Access-Control-Allow-Methods" value="GET, POST, PUT, DELETE, OPTIONS" />
    </customHeaders>
  </httpProtocol>
</system.webServer>
```

> **Note.** If you ever rehost a fork of this add-in under a different
> GitHub Pages account or custom domain, replace `https://globalmoo.github.io`
> in the snippets above with whatever the task pane is actually served
> from. The CORS error message in the task pane will show you the exact
> origin to allow.

---

## 4. Troubleshooting

The task pane's "Test connection" button is the diagnostic. Each failure
mode produces a different message:

| Message | What it means | What to do |
|---|---|---|
| **Invalid API key.** | Server returned HTTP 401. | Re-check the key. Generate a new one if needed. |
| **API error (4xx/5xx): &lt;message&gt;** | Server responded but rejected the request. | Read the message; check server logs. |
| **Couldn't reach &lt;host&gt;** + cert-trust PowerShell snippet | The browser couldn't complete the TLS handshake. Usually means the cert isn't trusted by Windows, but also fires on DNS / network failures. | Copy and run the PowerShell snippet (it re-runs `install.ps1 -CertOnly`). If the cert was already trusted, check DNS / firewall / that the server is up. |
| **Server reachable, but the browser blocked the response (CORS)** | TCP and TLS worked; the server returned a response that lacked the right `Access-Control-Allow-Origin` header (or didn't handle `OPTIONS`). | Apply the snippets in section 3b. |

The "Technical details" disclosure under each error shows the raw
browser error message — useful when reporting issues.

### Other things to check

- **Excel must be the desktop app (Microsoft 365)**. The add-in won't
  load in Excel Online for self-hosted servers because Excel Online
  ignores the `Developer` registry key.
- **Restart Excel** after install or after importing a cert. The
  WebView caches state per-session.
- **Don't use `http://`** — only `https://` is supported, including for
  on-prem servers. Mixed-content rules block the rest.
- **Port-mismatched URLs are fine** — e.g. `https://10.0.0.5:8443` —
  but the cert must include that hostname/IP in its SAN.

---

## 5. Uninstalling

```powershell
Remove-ItemProperty `
  -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" `
  -Name "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
Remove-Item -Recurse "$env:LOCALAPPDATA\GlobalMOO\ExcelAddin"
```

This removes the manifest registration. Imported certificates remain in
your user trust store — remove them via `certmgr.msc` if desired.

---

## 6. Reporting problems

Open an issue at
https://github.com/globalMOO/gmoo-excel-plugin/issues with:

- The exact error message from the task pane (including "Technical
  details").
- The output of `install.ps1` if installation failed.
- Your server type (nginx / Apache / IIS / custom) and whether it sits
  behind a reverse proxy.
- Whether the cert is from a public CA, a private CA, or self-signed.
