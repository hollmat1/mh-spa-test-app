/*
  Configure MSAL here.
  - Set `clientId` to your Application (client) ID from Azure AD app registration.
  - Set `tenant` to either a tenant id (single-tenant) or 'common' (multi-tenant).
  - Set `redirectUri` to a URL registered in the app (for local testing, http://localhost:PORT)
*/

const authConfig = {
  // Replace with your client id
  clientId: "1f416560-172d-43f6-84fe-54c51bf25035",

  // Use 'common' for multi-tenant, or a tenant GUID / 'organizations' / tenant name for single-tenant
  tenant: "common",

  // Where Azure will redirect after auth. Use your registered redirect URI.
  // For GitHub Pages, set this to: https://<your-username>.github.io/<repo-name>/
  // Replace <your-username> and <repo-name> with your GitHub username and repository name.
  // Example: "https://mattygithub.github.io/mh-spa-test-app/"
  redirectUri: "https://<your-username>.github.io/<repo-name>/",

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
