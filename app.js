// app.js - MSAL PKCE SPA example (msal-browser)

async function ensureMsalLoaded() {
  if (window.msal) return window.msal;
  const localSrc = 'libs/msal-browser.min.js';
  const cdnSrc = 'https://unpkg.com/@azure/msal-browser@2.46.1/dist/msal-browser.min.js';

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src; s.async = false; s.crossOrigin = 'anonymous';
      s.onload = () => resolve(window.msal);
      s.onerror = () => reject(new Error('Failed to load script: ' + src));
      document.head.appendChild(s);
    });
  }

  // Try local copy first
  try {
    return await loadScript(localSrc);
  } catch (eLocal) {
    // then CDN
    try {
      return await loadScript(cdnSrc);
    } catch (eCdn) {
      throw new Error('Failed to load msal-browser from local and CDN. Please add libs/msal-browser.min.js or enable network access.');
    }
  }
}

(async function main(){
  const msalApi = await ensureMsalLoaded();

  // In msal-browser v5+, use the factory to create a standard public client application (async init)
  const msalInstance = await (msalApi.createStandardPublicClientApplication
    ? msalApi.createStandardPublicClientApplication(authConfig.msalConfig)
    : Promise.resolve(new msalApi.PublicClientApplication(authConfig.msalConfig)));

  const ui = {
    status: document.getElementById('status'),
    account: document.getElementById('account'),
    result: document.getElementById('result'),
    tokenLink: document.getElementById('tokenLink'),
    decoded: document.getElementById('decoded'),
    rawToken: document.getElementById('rawToken'),
    copyToken: document.getElementById('copyToken'),
    tokenInfo: document.getElementById('tokenInfo'),
    signin: document.getElementById('signin'),
    signout: document.getElementById('signout'),
    acquire: document.getElementById('acquire'),
    callApi: document.getElementById('callApi')
  };

  let activeAccount = null;

  function updateUI() {
    if (activeAccount) {
      ui.status.textContent = `Signed in: ${activeAccount.username}`;
      ui.account.textContent = JSON.stringify(activeAccount, null, 2);
      ui.signout.disabled = false;
      ui.acquire.disabled = false;
      ui.callApi.disabled = false;
    } else {
      ui.status.textContent = 'Not signed in';
      ui.account.textContent = '(none)';
      ui.result.textContent = '(none)';
      ui.signout.disabled = true;
      ui.acquire.disabled = true;
      ui.callApi.disabled = true;
    }
  }

  // Initialize: if an account exists in cache, set it active
  function init() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      activeAccount = accounts[0];
    }
    updateUI();
  }

  async function signIn() {
    try {
      // Use redirect to avoid popup communication issues
      await msalInstance.loginRedirect(authConfig.loginRequest);
      // control will return via handleRedirectPromise
    } catch (err) {
      console.error(err);
      ui.result.textContent = 'Login error: ' + err;
    }
  }

  async function signOut() {
    try {
      const logoutRequest = { account: activeAccount };
      // prefer redirect logout for reliability
      await msalInstance.logoutRedirect(logoutRequest);
    } catch (err) {
      console.error(err);
    }
  }

  async function getToken() {
    if (!activeAccount) throw new Error('No active account');
    const silentRequest = { account: activeAccount, scopes: authConfig.loginRequest.scopes };

    try {
      const tokenResponse = await msalInstance.acquireTokenSilent(silentRequest);
      ui.result.textContent = JSON.stringify(tokenResponse, null, 2);
      if (tokenResponse && tokenResponse.accessToken) showToken(tokenResponse.accessToken, 'access_token');
      return tokenResponse.accessToken;
    } catch (err) {
      // interaction required -> redirect fallback
      if (err instanceof msalApi.InteractionRequiredAuthError || (err.errorMessage && err.errorMessage.indexOf('interaction_required') !== -1)) {
        try {
          await msalInstance.acquireTokenRedirect(silentRequest);
          // token response will be handled by handleRedirectPromise
          return null;
        } catch (rerr) {
          console.error('acquireTokenRedirect error', rerr);
          ui.result.textContent = 'Token request error: ' + rerr;
          throw rerr;
        }
      }
      console.error(err);
      ui.result.textContent = 'Token error: ' + err;
      throw err;
    }
  }

  async function callProtectedApi() {
    try {
      const token = await getToken();
      const sampleApi = authConfig.apiEndpoint || null;
      if (!sampleApi) {
        ui.result.textContent = 'No API configured. Token length: ' + (token ? token.length : 0);
        return;
      }
      const res = await fetch(sampleApi, { headers: { Authorization: `Bearer ${token}` } });
      const text = await res.text();
      ui.result.textContent = text;
    } catch (err) {
      console.error(err);
      ui.result.textContent = 'API call error: ' + err;
    }
  }

  // Token helpers: decode JWT and show link to jwt.ms
  function base64UrlDecode(str) {
    try {
      let s = str.replace(/-/g, '+').replace(/_/g, '/');
      while (s.length % 4) s += '=';
      return decodeURIComponent(escape(window.atob(s)));
    } catch (e) {
      return null;
    }
  }

  function decodeJwt(token) {
    if (!token) return null;
    const parts = token.split('.');
    if (parts.length < 2) return null;
    const header = base64UrlDecode(parts[0]);
    const payload = base64UrlDecode(parts[1]);
    try {
      return { header: header ? JSON.parse(header) : null, payload: payload ? JSON.parse(payload) : null };
    } catch (e) {
      return { header: header || null, payload: payload || null };
    }
  }

  function showToken(token, type) {
    if (!token) return;
    const param = type === 'id_token' ? 'id_token' : (type === 'access_token' ? 'access_token' : 'token');
    const href = `https://jwt.ms#${param}=`;
    if (ui.tokenLink) {
      ui.tokenLink.href = href;
      ui.tokenLink.style.display = 'inline';
      ui.tokenLink.textContent = 'Open token in jwt.ms';
      // remove previous handler if any
      ui.tokenLink.onclick = null;
      ui.tokenLink.addEventListener('click', (e) => {
        e.preventDefault();
        // open jwt.ms with the raw token so it is not percent-encoded in the fragment
        window.open(`${href}${token}`, '_blank', 'noopener');
      });
    }
    const decoded = decodeJwt(token);
    if (ui.decoded) {
      if (decoded && (decoded.header || decoded.payload)) {
        ui.decoded.textContent = `Header:\n${JSON.stringify(decoded.header, null, 2)}\n\nPayload:\n${JSON.stringify(decoded.payload, null, 2)}`;
      } else {
        ui.decoded.textContent = '(token is not a JWT or could not be decoded)';
      }
    }
    // populate raw token and diagnostics
    if (ui.rawToken) ui.rawToken.value = token;
    if (ui.tokenInfo) {
      const segs = token.split('.');
      ui.tokenInfo.textContent = `length=${token.length} chars · segments=${segs.length}`;
      // extra help for JWE (5 parts)
      if (segs.length === 5) {
        const hdr = base64UrlDecode(segs[0]);
        let hdrObj = hdr ? (() => { try { return JSON.parse(hdr) } catch(e){ return hdr } })() : '(not decodable)';
        ui.decoded.textContent = `Detected JWE (encrypted JWT).\nProtected header:\n${JSON.stringify(hdrObj, null, 2)}\n\nPayload is encrypted and cannot be decoded without the private key.`;
      }
    }
  }

  if (ui.copyToken) {
    ui.copyToken.addEventListener('click', () => {
      const t = ui.rawToken && ui.rawToken.value ? ui.rawToken.value : null;
      if (!t) return;
      navigator.clipboard.writeText(t).then(()=>{
        ui.copyToken.textContent = 'Copied';
        setTimeout(()=> ui.copyToken.textContent = 'Copy token', 1500);
      }).catch(()=>{
        // fallback
        try { ui.rawToken.select(); document.execCommand('copy'); ui.copyToken.textContent = 'Copied'; setTimeout(()=> ui.copyToken.textContent = 'Copy token',1500);} catch(e){ alert('Copy failed'); }
      });
    });
  }

  ui.signin.addEventListener('click', signIn);
  ui.signout.addEventListener('click', signOut);
  ui.acquire.addEventListener('click', () => getToken().catch(()=>{}));
  ui.callApi.addEventListener('click', () => callProtectedApi().catch(()=>{}));

  // Handle redirect responses (loginRedirect / acquireTokenRedirect / logoutRedirect)
  try {
    const redirectResponse = await msalInstance.handleRedirectPromise();
    if (redirectResponse) {
      // login or token response handled here
      if (redirectResponse.account) {
        activeAccount = redirectResponse.account;
      } else {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) activeAccount = accounts[0];
      }
      updateUI();
    } else {
      init();
    }
  } catch (err) {
    console.error('handleRedirectPromise error', err);
    ui.result.textContent = 'Redirect handling error: ' + err;
    init();
  }

})();
