/*
  FP-Karte – Anmeldung an SharePoint
  ----------------------------------
  Quelle der Karte im Betrieb. Anmeldung mit MSAL per Weiterleitung, gleiche App-Registrierung und gleiche
  Einstellungen wie die Übersicht (einstellungen.js). Die Adresse dieser Seite muss in der App-Registrierung
  als Umleitungs-URI (Single-Page-Anwendung) eingetragen sein.

  Die Karte liest nur. sp() lässt deshalb nur GET durch und die eine Listenabfrage (POST …/GetItems), alles
  andere wird abgelehnt, bevor eine Anfrage hinausgeht. Verlangt wird dasselbe Recht wie in der Übersicht,
  damit keine neue Zustimmung nötig ist und die Anmeldung der Übersicht mitgilt.

  Ergebnis: window.FP_KARTE_QUELLE (Versprechen) mit { sp, cfg, konto: { mail, name }, maildomain }.
  Braucht die Anmeldung einen Klick oder scheitert sie, steht der Hinweis im Ladefeld der Seite und das
  Versprechen bleibt offen. Nach der Weiterleitung lädt die Seite neu.

  Für Prüfungen ohne Microsoft: window.FPKarteSharepoint.quelle({ msal, einstellungen, fetch, warte, zeige, ort, rahmen })
*/
(function () {
  'use strict';

  const msg = (e) => (e && (e.errorMessage || e.message)) || String(e);
  const offen = () => new Promise(() => { });

  /* Hinweis im Ladefeld der Seite, auf Wunsch mit einer Taste. Nur Text, kein HTML. */
  function zeigeImLadefeld(text, knopf, fn) {
    const el = document.getElementById('lade');
    if (!el) return;
    el.replaceChildren(document.createTextNode(text));
    if (knopf) {
      const b = document.createElement('button');
      b.type = 'button'; b.className = 'lade-knopf'; b.textContent = knopf; b.onclick = fn;
      el.append(b);
    }
  }

  /* Nur Lesen: GET und die Abfrage der Mappen eines Ordners. */
  function erlaubt(pfad, opts) {
    const m = String((opts && opts.method) || 'GET').toUpperCase();
    return m === 'GET' || (m === 'POST' && /\/_api\/web\/lists\(guid'[0-9a-f-]+'\)\/GetItems\?/i.test(String(pfad)));
  }

  function quelle(o) {
    o = o || {};
    const E = o.einstellungen || window.FP_EINSTELLUNGEN || {};
    const S = E.sharepoint || {};
    const M = o.msal || window.msal;
    const holen = o.fetch || ((adresse, init) => fetch(adresse, init));
    const warte = o.warte || ((ms) => new Promise((r) => setTimeout(r, ms)));
    const zeige = o.zeige || zeigeImLadefeld;
    const ort = o.ort || window.location;
    const scopes = [S.host + '/AllSites.Write'];

    function zeigeFehler(e) {
      const m = msg(e);
      if (/AADSTS65001|AADSTS90094|consent|Administrator|admin/i.test(m)) {
        zeige('Die App „FP-Knopf“ braucht noch die einmalige Freigabe durch einen Administrator. Danach öffnet sich die Karte ohne weitere Einrichtung. (' + m.slice(0, 200) + ')');
      } else if (/AADSTS50011|redirect/i.test(m)) {
        zeige('Diese Adresse ist in der App-Registrierung „FP-Knopf“ noch nicht als Umleitungs-URI eingetragen: ' + ort.origin + ort.pathname);
      } else {
        zeige('Anmeldung fehlgeschlagen: ' + m, 'Nochmal versuchen', () => ort.reload());
      }
    }

    return (async () => {
      // Im unsichtbaren Rahmen, in den MSAL diese Seite für die stille Anmeldung lädt, startet nichts. Dort zählt nur
      // die Adresse. In fremde Seiten lässt sich die Karte nicht einbetten.
      if (o.rahmen !== false && window.top !== window.self) {
        let eigen = false;
        try { eigen = window.top.location.origin === ort.origin; } catch (e) { eigen = false; }
        if (!eigen) zeige('Die Karte bitte direkt öffnen: ' + ort.origin + ort.pathname);
        return offen();
      }
      if (!E.clientId || !E.tenantId || !S.host || !S.web || !S.listId || !S.libRel || !S.wissenRel || !M) {
        zeige('Einstellungen fehlen (einstellungen.js) oder MSAL wurde nicht geladen.');
        return offen();
      }

      let pca;
      try {
        pca = await M.PublicClientApplication.createPublicClientApplication({
          auth: { clientId: E.clientId, authority: 'https://login.microsoftonline.com/' + E.tenantId, redirectUri: ort.origin + ort.pathname },
          cache: { cacheLocation: 'localStorage' },
        });
        const rr = await pca.handleRedirectPromise();
        if (rr && rr.account) pca.setActiveAccount(rr.account);
      } catch (e) { zeigeFehler(e); return offen(); }

      let konto = pca.getActiveAccount() || pca.getAllAccounts()[0];
      if (!konto) {
        try { const r = await pca.ssoSilent({ scopes }); konto = r.account; pca.setActiveAccount(konto); }
        catch (e) {
          zeige('Einmal mit dem Firmenkonto anmelden, danach öffnet sich die Karte direkt.', 'Anmelden', () => pca.loginRedirect({ scopes }).catch(zeigeFehler));
          return offen();
        }
      }

      async function token() {
        try { return (await pca.acquireTokenSilent({ scopes, account: konto })).accessToken; }
        catch (e) {
          if (M.InteractionRequiredAuthError && e instanceof M.InteractionRequiredAuthError) { await pca.acquireTokenRedirect({ scopes, account: konto }); return offen(); }
          throw e;
        }
      }

      /* Wie sp() der Übersicht: JSON hin und zurück, bei Drosselung warten und wiederholen, Fehler tragen .status. */
      async function sp(pfad, opts) {
        opts = opts || {};
        if (!erlaubt(pfad, opts)) throw new Error('Die Karte liest nur. Abgelehnt: ' + String(opts.method || 'GET') + ' ' + String(pfad).slice(0, 120));
        for (let v = 1; v <= 4; v++) {
          const headers = { Authorization: 'Bearer ' + (await token()), Accept: 'application/json;odata=nometadata' };
          if (opts.body !== undefined) headers['Content-Type'] = 'application/json;odata=nometadata';
          const r = await holen(S.host + pfad, { method: opts.method || 'GET', headers, body: opts.body === undefined ? undefined : JSON.stringify(opts.body) });
          if (r.status === 429 || r.status === 503) {
            const nach = r.headers && r.headers.get ? r.headers.get('Retry-After') : null;
            await warte((nach !== null && nach !== '' && isFinite(Number(nach)) ? Number(nach) : 3 * v) * 1000);
            continue;
          }
          const t = await r.text();
          let j = null; try { j = t ? JSON.parse(t) : null; } catch (_) { }
          if (!r.ok) {
            const m = (j && j['odata.error'] && j['odata.error'].message && j['odata.error'].message.value) || (j && j.error && j.error.message) || t.slice(0, 200);
            const e = new Error(r.status + ' ' + m); e.status = r.status; throw e;
          }
          return j;
        }
        throw new Error('SharePoint ist gerade überlastet. Gleich nochmal versuchen.');
      }

      try { await token(); } catch (e) { zeigeFehler(e); return offen(); }

      return { sp, cfg: S, konto: { mail: String(konto.username || '').toLowerCase(), name: String(konto.name || '') }, maildomain: E.maildomain };
    })();
  }

  window.FPKarteSharepoint = { quelle, erlaubt };
  if (!window.FP_KARTE_TEST) window.FP_KARTE_QUELLE = quelle();
})();
