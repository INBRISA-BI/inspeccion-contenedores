// Contraseña de acceso (solo guarda el hash SHA-256). Es una barrera ligera: la protección real es el inicio de sesión Microsoft.
(function () {
  const PASS_HASH = '24bc205ab6bb2fc5d75be56834817b1856a6121ced63166c4713954e77feeacc';
  async function sha256(t) {
    const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(t));
    return [...new Uint8Array(b)].map((x) => x.toString(16).padStart(2, '0')).join('');
  }
  function unlocked() { try { return sessionStorage.getItem('ic-ok') === PASS_HASH; } catch (e) { return false; } }
  window.icGate = function (onOk) {
    if (unlocked()) { onOk(); return; }
    const o = document.createElement('div');
    o.style.cssText = 'position:fixed;inset:0;background:var(--color-bg);display:flex;align-items:center;justify-content:center;z-index:99;padding:1rem';
    o.innerHTML = '<form style="display:flex;flex-direction:column;gap:.75rem;width:min(20rem,100%)"><h1 style="margin:0;font-family:var(--font-anthropic-serif);font-weight:400;font-size:1.5rem">Calidad INBRISA</h1><label style="font-size:.875rem;color:var(--color-fg-muted)">Contraseña<input type="password" autocomplete="off" style="display:block;width:100%;box-sizing:border-box;margin-top:.25rem;padding:.5rem;font:inherit;color:var(--color-fg);background:var(--color-panel);border:1px solid var(--color-border-line);border-radius:var(--cds-radius)"></label><button class="ic-login" style="margin:0">Entrar</button><p style="margin:0;color:var(--color-bad);font-size:.875rem" hidden>Contraseña incorrecta.</p></form>';
    document.body.appendChild(o);
    const f = o.querySelector('form'), i = o.querySelector('input'), e = o.querySelector('p');
    i.focus();
    f.addEventListener('submit', async (ev) => {
      ev.preventDefault();
      if ((await sha256(i.value)) === PASS_HASH) {
        try { sessionStorage.setItem('ic-ok', PASS_HASH); } catch (x) { /* sin almacenamiento */ }
        o.remove(); onOk();
      } else { e.hidden = false; i.select(); }
    });
  };
})();
