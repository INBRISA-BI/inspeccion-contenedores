// Inicio de sesión Microsoft compartido (MSAL). El inicio de sesión interactivo solo ocurre en la página principal.
(function () {
  const CFG = window.APP_CONFIG;
  let app = null;
  const embedded = window.top !== window;
  async function init() {
    if (app) return app;
    app = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: 'https://login.microsoftonline.com/' + (CFG.tenant || 'organizations'), redirectUri: CFG.redirectUri },
      cache: { cacheLocation: 'localStorage' },
    });
    if (typeof app.initialize === 'function') await app.initialize();
    if (!embedded) { try { await app.handleRedirectPromise(); } catch (e) { window.__authError = e.message; } }
    return app;
  }
  function account() { const a = app && app.getAllAccounts(); return a && a.length ? a[0] : null; }
  async function token(scopes) {
    await init();
    const acc = account();
    if (!acc) return null;
    try { return (await app.acquireTokenSilent({ scopes, account: acc })).accessToken; }
    catch (e) {
      if (embedded) { const er = new Error('Falta aprobar el permiso. Pide al administrador que apruebe la aplicación.'); er.consent = true; throw er; }
      await app.acquireTokenRedirect({ scopes, account: acc }); return null;
    }
  }
  async function login(scopes) {
    await init();
    if (embedded) { window.top.postMessage({ type: 'ic-login' }, location.origin); return; }
    await app.loginRedirect({ scopes });
  }
  async function logout() { await init(); const acc = account(); await app.logoutRedirect({ account: acc || undefined }); }
  window.icAuth = { init, account, token, login, logout, embedded };
})();
