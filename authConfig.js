/*
  Configure MSAL here.
  - Set `clientId` to your Application (client) ID from Azure AD app registration.
  - Set `tenant` to either a tenant id (single-tenant) or 'common' (multi-tenant).
  - Set `redirectUri` to a URL registered in the app (for local testing, http://localhost:PORT)
*/

const authConfig = {
  // Replace with your Application (client) ID from Azure AD app registration.
  // Leave empty to enter it at runtime via the UI.
  clientId: "",

  // Use 'common' for multi-tenant, or a tenant GUID / 'organizations' / tenant name for single-tenant
  tenant: "common",

  // Where Azure will redirect after auth. Use your registered redirect URI.
  // Determine redirect URI at runtime so local testing (localhost) works
  // and the same build can run on GitHub Pages (including repository subpaths)
  // Example for Pages: https://hollmat1.github.io/mh-spa-test-app/
  redirectUri: (typeof window !== 'undefined' && window.location && window.location.origin)
    ? (function(){
        // preserve any path prefix (e.g. /mh-spa-test-app/). Strip index.html if present.
        try {
          let p = window.location.pathname || '/';
          p = p.replace(/\/index\.html$/i, '/');
          if (!p.endsWith('/')) p = p + '/';
          return window.location.origin + p;
        } catch (e) {
          return window.location.origin + '/';
        }
      })()
    : 'http://localhost:8080/',

  // Scopes you want to request when acquiring a token
  loginRequest: {
    scopes: ["openid", "profile", "User.Read"]
  }
  ,
  // Default protected API endpoint (Microsoft Graph - current user)
  apiEndpoint: "https://graph.microsoft.com/v1.0/me"
};

// Build msalConfig dynamically from above
authConfig.msalConfig = {
  auth: {
    clientId: authConfig.clientId,
    authority: `https://login.microsoftonline.com/${authConfig.tenant}`,
    redirectUri: authConfig.redirectUri,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};
