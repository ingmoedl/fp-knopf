/* FP-Übersicht als eigene Seite (Browser, Edge-App, Handy): Anmeldung mit MSAL per Weiterleitung,
 * danach die Übersicht aus uebersicht.js. Gleiche App-Registrierung wie der FP-Knopf (einstellungen.js);
 * die Adresse dieser Seite muss dort als Umleitungs-URI (Single-Page-Anwendung) eingetragen sein. */
(async function () {
  "use strict";

  const E = window.FP_EINSTELLUNGEN || {};
  const S = E.sharepoint || {};
  const scopes = [S.host + "/AllSites.Write"];
  const app = document.getElementById("app");
  const html = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const msg = (e) => (e && (e.errorMessage || e.message)) || String(e);
  const warte = (ms) => new Promise((s) => setTimeout(s, ms));

  function hinweis(text, knopf, fn) {
    app.innerHTML = `<div class="start"><div class="logo">FP</div><h1>FP-Aufgaben</h1><p>${text}</p>${knopf ? `<button>${html(knopf)}</button>` : ""}</div>`;
    if (knopf) app.querySelector("button").onclick = fn;
  }

  function fehler(e) {
    const m = msg(e);
    if (/AADSTS65001|AADSTS90094|consent|Administrator|admin/i.test(m)) {
      hinweis("Die App „FP-Knopf“ braucht noch die einmalige Freigabe durch einen Administrator. " +
        "Danach funktioniert diese Seite ohne weitere Einrichtung.<br><small>" + html(m.slice(0, 200)) + "</small>");
    } else if (/AADSTS50011|redirect/i.test(m)) {
      hinweis("Diese Adresse ist in der App-Registrierung „FP-Knopf“ noch nicht als Umleitungs-URI eingetragen:<br><code>" +
        html(location.origin + location.pathname) + "</code>");
    } else {
      hinweis("Anmeldung fehlgeschlagen: " + html(m), "Nochmal versuchen", () => location.reload());
    }
  }

  if (!E.clientId || !window.msal) { hinweis("Einstellungen fehlen (einstellungen.js) oder MSAL wurde nicht geladen."); return; }

  let pca;
  try {
    pca = await msal.PublicClientApplication.createPublicClientApplication({
      auth: { clientId: E.clientId, authority: "https://login.microsoftonline.com/" + E.tenantId, redirectUri: location.origin + location.pathname },
      cache: { cacheLocation: "localStorage" },
    });
    const rr = await pca.handleRedirectPromise();
    if (rr && rr.account) pca.setActiveAccount(rr.account);
  } catch (e) { fehler(e); return; }

  let konto = pca.getActiveAccount() || pca.getAllAccounts()[0];
  if (!konto) {
    try { const r = await pca.ssoSilent({ scopes }); konto = r.account; pca.setActiveAccount(konto); }
    catch (e) {
      hinweis("Einmal mit dem Firmenkonto anmelden, danach öffnet sich die Übersicht direkt.", "Anmelden",
        () => pca.loginRedirect({ scopes }).catch(fehler));
      return;
    }
  }

  async function token() {
    try { return (await pca.acquireTokenSilent({ scopes, account: konto })).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) { await pca.acquireTokenRedirect({ scopes, account: konto }); return new Promise(() => { }); }
      throw e;
    }
  }

  async function sp(path, opts = {}) {
    for (let v = 1; v <= 4; v++) {
      const headers = { Authorization: "Bearer " + (await token()), Accept: "application/json;odata=nometadata" };
      if (opts.body !== undefined) headers["Content-Type"] = "application/json;odata=nometadata";
      const r = await fetch(S.host + path, { method: opts.method || "GET", headers, body: opts.body === undefined ? undefined : JSON.stringify(opts.body) });
      if (r.status === 429 || r.status === 503) { await warte((Number(r.headers.get("Retry-After")) || 3 * v) * 1000); continue; }
      if (r.ok && opts.roh) return r.arrayBuffer();                 // Dateiinhalt (Mails)
      const t = await r.text();
      let j = null; try { j = t ? JSON.parse(t) : null; } catch (_) { }
      if (!r.ok) {
        const m = (j && j["odata.error"] && j["odata.error"].message && j["odata.error"].message.value) || (j && j.error && j.error.message) || t.slice(0, 200);
        const e = new Error(r.status + " " + m); e.status = r.status; throw e;
      }
      return j;
    }
    throw new Error("SharePoint ist gerade überlastet. Gleich nochmal versuchen.");
  }

  try { await token(); } catch (e) { fehler(e); return; }

  const mail = (konto.username || "").toLowerCase();
  const b = (E.bearbeiter || []).find((x) => x.mail.toLowerCase() === mail);
  const ich = b || { kuerzel: (konto.name || mail).split(/[,\s@.]+/).filter(Boolean).slice(0, 2).map((s) => s[0]).join("").toUpperCase(), name: konto.name || mail, mail };
  window.FPUebersicht.start(app, { sp, cfg: S, ich, bearbeiter: E.bearbeiter, speicher: "seite" });
})();
