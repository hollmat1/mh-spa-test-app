MSAL PKCE SPA (Multi-tenant & Single-tenant)

This is a minimal public single-page app demonstrating the OAuth2 Authorization Code flow with PKCE using msal-browser.

Files added:
- `index.html` — UI and script includes
- `authConfig.js` — configuration (set clientId and tenant)
- `app.js` — MSAL usage (loginPopup, acquireTokenSilent + popup fallback, logout)

Quick start:
1. Register an app in Azure AD (App registrations):
   - Platform -> Add a Redirect URI -> Single-page application -> e.g. http://localhost:8080
   - For multi-tenant: under Authentication, allow "Accounts in any organizational directory" or use "common" as the tenant in authConfig.js.
   - For single-tenant: set Supported account types to "Accounts in this organizational directory only" and set tenant in authConfig.js to your tenant ID or tenant name.
   - Expose the API / add delegated scopes if you intend to call a protected API.

2. Edit authConfig.js:
   - Replace YOUR_CLIENT_ID with the registered app client id.
   - Set tenant to common (multi-tenant) or your tenant id (single-tenant).

3. Run a simple static server from this folder. Example commands:
   - npx http-server -c-1 -p 8080
   - python -m http.server 8080

4. Open http://localhost:8080 in your browser.

Notes:
- msal-browser implements the Authorization Code flow with PKCE for SPAs; the library handles PKCE details.
- For production, use https, stronger cache/cookie settings, and follow MSAL security guidance.

Local msal-browser copy
 - To prefer a local copy of `msal-browser`, run the helper script:
    - `.
fetch-msal.ps1` (PowerShell) which will `npm install @azure/msal-browser` and copy the
       minified bundle to `libs/msal-browser.min.js`.
 - `index.html` prefers `libs/msal-browser.min.js` and will fall back to the CDN if the local file is missing.
