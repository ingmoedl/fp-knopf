/*
  FP-Karte – Start
  ----------------
  Verbindet Ladeschicht, Kartenansicht und Gebäudeseite. Die Datenquelle kommt von außen als
  window.FP_KARTE_QUELLE (Versprechen mit { sp, cfg, konto, maildomain, schreiben } und für lokale Quellen dazu
  beschriftung, fussnote, heute, eigeneParteien). Mit schreiben lassen sich Aufgaben auf der Gebäudeseite
  bearbeiten (bearbeiten.js).
  Lokal liefert sie quelle-lokal.js, im Betrieb quelle-sharepoint.js mit der Anmeldung an SharePoint.

  Adresse: #g=<Nummer> öffnet die Gebäudeseite, #g=ohne die Aufgaben ohne Gebäude, #g=liste die Aufgabenliste der Auswahl.
  Mit ?ruhig läuft die Karte ohne Bewegung.
*/
(async function () {
  'use strict';

  const $ = (s) => document.querySelector(s);
  const A = window.FPKarteAuswertung, GS = window.FPGebaeudeseite, h = GS.h;
  const app = $('#app'), lade = $('#lade'), seiteEl = $('#seite'), gs = $('#gs');
  const msg = (e) => (e && (e.errorMessage || e.message)) || String(e);
  const lies = (k) => { try { return localStorage.getItem(k); } catch (e) { return null; } };
  const schreib = (k, v) => { try { localStorage.setItem(k, v); } catch (e) { } };
  const breit = () => window.innerWidth > 720;
  const OHNE = '_ohne', LISTE = '_liste';
  const istGebaeude = (key) => key !== OHNE && key !== LISTE;

  let Q;
  try { Q = await window.FP_KARTE_QUELLE; if (!Q || !Q.sp || !Q.cfg) throw new Error('Keine Datenquelle eingebunden.'); }
  catch (e) { lade.textContent = 'Die Datenquelle fehlt. ' + msg(e); return; }

  const daten = window.FPKarteDaten.erzeugen({
    sp: Q.sp, cfg: Q.cfg, konto: Q.konto, maildomain: Q.maildomain,
    gebaeudeAusText: window.FPEinordnung ? (text, register) => window.FPEinordnung.gebaeudeFinden(text, '', register) : null,
  });

  let grund;
  try { grund = await daten.grundlagen(); }
  catch (e) { lade.textContent = 'Die Karte konnte nicht geladen werden. ' + msg(e); return; }

  // Welche Angabe bei „offen bei“ das eigene Büro meint: von der Quelle, sonst aus dem Register (Feld eigene_parteien)
  const eigeneParteien = Q.eigeneParteien || (Array.isArray(grund.register.eigene_parteien) ? grund.register.eigene_parteien : []);

  const Z = {
    filter: 'alle', firmen: [], firmenOffen: false, wahl: null, heute: Q.heute || A.heuteIso(), team: [], ich: null,
    aufgaben: [], fehlerAufgaben: '', index: new Map(), indexStand: '', indexFehlt: false,
    bild: { gebaeude: new Map(), ohne: { aufgaben: 0 } }, dateien: new Map(), dateiFehler: new Map(),
    offen: null, arbeit: null, meldung: null, zurueck: null,       // Bearbeitung: offene Aufgabe, laufendes Speichern, Rückmeldung, zuletzt Erledigtes
  };
  const bearb = Q.schreiben && window.FPKarteBearbeiten
    ? window.FPKarteBearbeiten.erzeugen({ sp: Q.sp, cfg: Q.cfg, ich: () => Z.ich, heute: () => Q.heute || A.heuteIso() }) : null;

  const karte = window.FPKarteAnsicht.erzeugen($('#karte'), grund.karte, grund.gebaeude, {
    beiWahl: (nr) => oeffnen(nr),
    beiLeer: () => { if (breit()) schliessen(); },
    beiAnsicht: (a) => { $('#w-ansicht').textContent = a.draufsicht ? 'Schrägansicht' : 'Draufsicht'; },
    ohneBewegung: new URLSearchParams(location.search).has('ruhig'),      // ?ruhig: ohne Fahrten und Einblenden, für Bilder und Prüfungen
  });
  lade.remove();
  window.FP_KARTE = { karte, daten, zustand: Z };           // zum Nachsehen in der Konsole und für die Prüfseite

  const nord = grund.karte.meta && grund.karte.meta.nord;
  $('#hinweis').textContent = (nord && nord.geprueft ? '' : 'Norden ungeprüft · ') + 'Höhen schematisch';

  /* ---------- Rechnen und Zeigen ---------- */
  function filterGueltig(f) { return f === 'alle' || (f === 'meine' && Z.ich) || Z.team.some((b) => b.kuerzel === f); }

  function neuRechnen() {
    if (!filterGueltig(Z.filter)) Z.filter = 'alle';
    // Gemerkte Firmen, die es in den Aufgaben nicht mehr gibt, fallen heraus. Solange die Aufgaben fehlen, bleibt die Auswahl.
    const alleFirmen = A.firmenListe(Z.aufgaben, 'alle', Z.ich).map((f) => f.name);
    if (Z.aufgaben.length) Z.firmen = Z.firmen.filter((f) => alleFirmen.includes(f));
    Z.bild = A.lagebild(grund.gebaeude, Z.aufgaben, Z.index, Z.filter, Z.ich, Z.heute, Z.firmen);
    karte.zustand(Z.bild.gebaeude);
    $('#filter').replaceChildren(GS.filterLeiste(Z.team, Z.ich, Z.filter, A.zaehler(A.nachFirmen(Z.aufgaben, Z.firmen), Z.team, Z.ich), filterSetzen));
    firmenZeigen();
    $('#liste-zahl').textContent = String(Z.aufgaben.filter((a) => A.passt(a, Z.filter, Z.ich, Z.firmen)).length);
    abseitsZeigen();
    standZeigen();
    if (Z.wahl) seiteZeigen(true);
  }

  function filterSetzen(f) {
    if (!filterGueltig(f) || f === Z.filter) return;
    Z.filter = f; schreib('fpk-filter', f);
    neuRechnen();
  }

  /* Firmenwahl in der Leiste. Die Zahlen je Firma folgen dem Personenfilter. Offen bleibt sie über das Neuzeichnen hinweg. */
  function firmenZeigen() {
    const liste = A.firmenListe(Z.aufgaben, Z.filter, Z.ich);
    for (const f of Z.firmen) if (!liste.some((x) => x.name === f)) liste.push({ name: f, anzahl: 0 });
    const el = GS.firmenWahl(liste, Z.firmen, firmenSetzen, Z.firmenOffen);
    el.addEventListener('toggle', () => { Z.firmenOffen = el.open; });
    $('#firmen').replaceChildren(el);
  }

  function firmenSetzen(liste) {
    Z.firmen = Array.isArray(liste) ? liste.map(String) : [];
    schreib('fpk-firmen', JSON.stringify(Z.firmen));
    neuRechnen();
  }

  function standZeigen() {
    const teile = [];
    if (Q.beschriftung) teile.push(Q.beschriftung);
    if (Z.fehlerAufgaben) teile.push('Aufgaben nicht geladen');
    teile.push(Z.indexFehlt ? 'ohne Auswertung der Mails' : Z.indexStand ? 'Stand ' + A.standText(Z.indexStand) : 'wird geladen …');
    const e = $('#stand'); e.textContent = teile.join(' · '); e.classList.toggle('probe', !!Q.beschriftung || !!Z.fehlerAufgaben);
  }

  /* Gebäude mit Einträgen, die auf der Karte keinen Ort haben, und die Aufgaben ohne Gebäude */
  function abseitsZeigen() {
    const el = $('#abseits'), marken = [];
    for (const [nr, e] of Z.bild.gebaeude) {
      if (!e.relevant || karte.hat(nr)) continue;
      const ueber = e.dringend ? window.FPKarteAnsicht.ueberText(e.ueberfaellig) : '';
      marken.push(h('button', { type: 'button', class: 'schild relevant' + (e.dringend ? ' dringend' : '') + (Z.wahl === nr ? ' gewaehlt' : ''), 'data-ueber': ueber || null,
        'aria-label': 'Gebäude ' + nr + ', nicht auf der Karte verortet' + (ueber ? ', ' + ueber + ' überfällig' : ''), onclick: () => oeffnen(nr) }, h('span', { text: nr })));
    }
    if (Z.bild.ohne.aufgaben) {
      const ueber = window.FPKarteAnsicht.ueberText(Z.bild.ohne.ueberfaellig);
      marken.push(h('button', { type: 'button', class: 'schild' + (ueber ? ' dringend' : '') + (Z.wahl === OHNE ? ' gewaehlt' : ''), 'data-ueber': ueber || null,
        'aria-label': 'Ohne Gebäude, ' + Z.bild.ohne.aufgaben + ' Aufgaben' + (ueber ? ', ' + ueber + ' überfällig' : ''), onclick: () => oeffnen(OHNE) },
        h('span', { text: 'Ohne Gebäude · ' + Z.bild.ohne.aufgaben })));
    }
    el.hidden = !marken.length;
    el.replaceChildren(h('span', { class: 'titel', text: 'Nicht auf der Karte' }), ...marken);
  }

  /* ---------- Gebäudeseite ---------- */
  function seiteZeigen(behalten) {
    const key = Z.wahl; if (!key) return;
    const bearbeitung = { bearbeiten: !!bearb, offen: Z.offen, arbeit: Z.arbeit, meldung: Z.meldung };
    const rueckrufe = { beiFilter: filterSetzen, beiFirmen: firmenSetzen, beiZu: () => schliessen(), behalten,
      beiBearbeiten: bearbeitenSetzen, beiErledigt: erledigt, beiFaellig: faelligSetzen, beiNotiz: notiz, beiRueck: zurueckNehmen };
    if (key === LISTE) {
      GS.zeichnenListe(gs, Object.assign({
        liste: A.liste(Z.aufgaben, grund.gebaeude, { filter: Z.filter, ich: Z.ich, heute: Z.heute, firmen: Z.firmen }),
        filter: Z.filter, firmen: Z.firmen, ich: Z.ich, team: Z.team, zaehler: A.zaehler(A.nachFirmen(Z.aufgaben, Z.firmen), Z.team, Z.ich),
        fehlerAufgaben: Z.fehlerAufgaben, fussnote: Q.fussnote || '',
      }, bearbeitung), Object.assign({ beiGebaeude: (nr) => oeffnen(nr) }, rueckrufe));
      return;
    }
    const ohne = key === OHNE;
    const nr = ohne ? null : key;
    const datei = ohne ? null : Z.dateien.get(key) || null;
    const amGebaeude = Z.aufgaben.filter((a) => (ohne ? !a.gebaeude.length : a.gebaeude.some((g) => g.nr === nr)));
    GS.zeichnen(gs, {
      art: ohne ? 'ohne' : 'gebaeude', g: ohne ? null : grund.gebaeude.get(key),
      sicht: A.seite(nr, Z.aufgaben, datei, { filter: Z.filter, ich: Z.ich, heute: Z.heute, eigeneParteien, firmen: Z.firmen }),
      filter: Z.filter, firmen: Z.firmen, ich: Z.ich, team: Z.team, zaehler: A.zaehler(A.nachFirmen(amGebaeude, Z.firmen), Z.team, Z.ich),
      laedt: !ohne && !datei && !Z.dateiFehler.has(key), datei, fehler: Z.dateiFehler.get(key) || '', fehlerAufgaben: Z.fehlerAufgaben,
      fussnote: Q.fussnote || '', aufDerKarte: !ohne && karte.hat(key),
      bearbeiten: bearbeitung.bearbeiten, offen: bearbeitung.offen, arbeit: bearbeitung.arbeit, meldung: bearbeitung.meldung,
    }, rueckrufe);
  }

  /* ---------- Bearbeiten ---------- */
  // Die Gebäudeseite arbeitet mit Kopien. Geändert wird die Aufgabe im Zustand.
  const original = (k) => Z.aufgaben.find((x) => x.id === (k && k.id)) || null;

  function bearbeitenSetzen(id) {
    if (Z.arbeit) return;
    Z.offen = id;
    if (Z.meldung && Z.meldung.id !== id) Z.meldung = null;
    seiteZeigen(true);
  }

  /* Stand aus SharePoint in die Aufgabe übernehmen. Ist sie dort nicht mehr offen, fällt sie aus der Liste. */
  function standUebernehmen(a, s) {
    a.geaendert = s.geaendert;
    if (s.status !== 'Offen') { Z.aufgaben = Z.aufgaben.filter((x) => x !== a); if (Z.offen === a.id) Z.offen = null; }
    else a.faellig = s.faellig;
  }

  /* Eine Änderung speichern. Bei einem Fehler bleibt die Anzeige, wie sie war. Hat jemand anderes genau dieses
     Feld inzwischen geändert, wird nichts überschrieben und die Aufgabe zeigt danach dessen Stand. */
  async function aendern(a, tun, danach, text) {
    if (!bearb || !a || Z.arbeit) return;
    Z.arbeit = a.id; Z.meldung = null;
    seiteZeigen(true);
    try {
      const s = await tun();
      a.geaendert = s.geaendert;
      Z.meldung = { id: a.id, text, art: 'ok' };
      danach(s);
    } catch (e) {
      Z.meldung = { id: a.id, text: msg(e), art: 'warn' };
      if (e && e.konflikt && e.stand) { standUebernehmen(a, e.stand); Z.meldung.text += ' Die Anzeige ist jetzt aktuell.'; if (!Z.aufgaben.includes(a)) delete Z.meldung.id; }
    }
    Z.arbeit = null;
    neuRechnen();
  }

  function erledigt(k) {
    const a = original(k);
    aendern(a, () => bearb.erledigen(a), () => {
      Z.aufgaben = Z.aufgaben.filter((x) => x !== a);
      Z.offen = null; Z.zurueck = a;
      Z.meldung = { text: 'Erledigt: ' + a.titel, art: 'ok', rueck: true };
    }, '');
  }

  async function zurueckNehmen() {
    const a = Z.zurueck;
    if (!bearb || !a || Z.arbeit) return;
    Z.arbeit = a.id; Z.meldung = { text: 'Wird zurückgenommen …', art: 'ok' };
    seiteZeigen(true);
    try {
      const s = await bearb.wiederOeffnen(a);
      a.geaendert = s.geaendert; a.faellig = s.faellig;
      Z.aufgaben = Z.aufgaben.concat(a); Z.zurueck = null;
      Z.meldung = { text: 'Wieder offen: ' + a.titel, art: 'ok' };
    } catch (e) {
      Z.meldung = { text: msg(e), art: 'warn' };
      if (e && e.konflikt && e.stand && e.stand.status === 'Offen') { a.faellig = e.stand.faellig; Z.aufgaben = Z.aufgaben.concat(a); Z.zurueck = null; Z.meldung.text += ' Die Anzeige ist jetzt aktuell.'; }
    }
    Z.arbeit = null;
    neuRechnen();
  }

  function faelligSetzen(k, iso) {
    const a = original(k);
    aendern(a, () => bearb.faelligSetzen(a, iso), () => { a.faellig = iso; }, iso ? 'Fälligkeit gespeichert: ' + A.tagKurz(iso) : 'Die Aufgabe ist jetzt ohne Termin.');
  }

  function notiz(k, text) {
    const a = original(k);
    aendern(a, () => bearb.notiz(a, text), () => { }, 'Notiz angefügt.');
  }

  function raender() { karte.rand({ r: breit() && Z.wahl ? seiteEl.offsetWidth : 0, o: breit() ? 56 : 104, u: breit() ? 0 : 56 }); }

  function oeffnen(key, ausAdresse) {
    if (istGebaeude(key) && !grund.gebaeude.has(key)) return false;
    const neu = Z.wahl !== key;
    if (neu && !Z.arbeit) { Z.offen = null; Z.meldung = null; }
    Z.wahl = key;
    seiteEl.classList.add('offen'); seiteEl.setAttribute('aria-hidden', 'false'); app.classList.add('seite-offen');
    karte.waehle(istGebaeude(key) ? key : null);
    raender();
    if (istGebaeude(key) && karte.hat(key)) karte.zeige(key);
    $('#liste-auf').setAttribute('aria-pressed', String(key === LISTE));
    seiteZeigen(!neu);
    abseitsZeigen();
    if (!ausAdresse) adresseSetzen(key);
    if (!breit()) seiteEl.focus({ preventScroll: true });
    if (istGebaeude(key) && !Z.dateien.has(key)) {
      Z.dateiFehler.delete(key);
      daten.gebaeude(key).then((d) => { Z.dateien.set(key, d); }, (e) => { Z.dateiFehler.set(key, msg(e)); })
        .then(() => { if (Z.wahl === key) seiteZeigen(true); });
    }
    return true;
  }

  function schliessen(ausAdresse) {
    if (!Z.wahl) return;
    Z.wahl = null;
    seiteEl.classList.remove('offen'); seiteEl.setAttribute('aria-hidden', 'true'); app.classList.remove('seite-offen');
    karte.waehle(null);
    $('#liste-auf').setAttribute('aria-pressed', 'false');
    raender(); abseitsZeigen();
    if (!ausAdresse) adresseSetzen(null);
    if (seiteEl.contains(document.activeElement)) $('#karte').focus({ preventScroll: true });
  }

  /* ---------- Adresse ---------- */
  const ausAdresseLesen = () => { const m = /^#g=(.+)$/.exec(location.hash); if (!m) return null; const k = decodeURIComponent(m[1]); return k === 'ohne' ? OHNE : k === 'liste' ? LISTE : k; };
  function adresseSetzen(key) {
    const ziel = key ? '#g=' + encodeURIComponent(key === OHNE ? 'ohne' : key === LISTE ? 'liste' : key) : '';
    if (location.hash === ziel || (!ziel && !location.hash)) return;
    if (ziel) location.hash = ziel; else history.replaceState(null, '', location.pathname + location.search);
  }
  function ausAdresse() { const k = ausAdresseLesen(); if (k) { if (!oeffnen(k, true)) schliessen(true); } else schliessen(true); }
  window.addEventListener('hashchange', ausAdresse);

  /* ---------- Suche ---------- */
  const suche = $('#suche'), treffer = $('#treffer');
  let fund = [], an = 0;
  function sucheZeigen() {
    const q = suche.value.trim().toLowerCase();
    if (!q) { treffer.hidden = true; fund = []; return; }
    const alle = [...grund.gebaeude.values()];
    const genau = grund.nummern.get(q.replace(/^geb\.?\s*/, ''));
    const nachNr = alle.filter((g) => g.nr.toLowerCase().startsWith(q) || g.aliase.some((x) => x.toLowerCase().startsWith(q)));
    const nachName = q.length >= 2 ? alle.filter((g) => !nachNr.includes(g) && g.name.toLowerCase().includes(q)) : [];
    fund = nachNr.concat(nachName);
    if (genau && !fund.some((g) => g.nr === genau)) fund.unshift(grund.gebaeude.get(genau));
    const rang = (g) => { const e = Z.bild.gebaeude.get(g.nr); return (g.nr === genau ? 0 : 2) + (e && e.relevant ? 0 : 1); };
    fund = fund.map((g, i) => ({ g, i })).sort((x, y) => rang(x.g) - rang(y.g) || x.i - y.i).map((x) => x.g).slice(0, 8);
    an = 0;
    treffer.replaceChildren(...(fund.length ? fund.map((g, i) => h('button', { type: 'button', class: i === 0 ? 'an' : '', onclick: () => sucheWahl(g.nr) },
      h('span', { class: 'nr', text: g.nr }), h('span', { class: 'name', text: g.name }), h('span', { class: 'ort', text: karte.hat(g.nr) ? '' : 'nicht auf der Karte' })))
      : [h('div', { class: 'leer', text: 'Kein Gebäude gefunden.' })]));
    treffer.hidden = false;
  }
  function sucheWahl(nr) { suche.value = ''; treffer.hidden = true; fund = []; suche.blur(); oeffnen(nr); }
  suche.addEventListener('input', sucheZeigen);
  suche.addEventListener('focus', sucheZeigen);
  suche.addEventListener('keydown', (ev) => {
    if (ev.key === 'Enter' && fund[an]) { ev.preventDefault(); sucheWahl(fund[an].nr); }
    else if ((ev.key === 'ArrowDown' || ev.key === 'ArrowUp') && fund.length) {
      ev.preventDefault();
      an = (an + (ev.key === 'ArrowDown' ? 1 : fund.length - 1)) % fund.length;
      [...treffer.children].forEach((b, i) => b.classList.toggle('an', i === an));
    } else if (ev.key === 'Escape') { suche.value = ''; treffer.hidden = true; fund = []; ev.stopPropagation(); }
  });
  treffer.addEventListener('pointerdown', (ev) => ev.preventDefault());      // Feld behält den Fokus, der Klick kommt an
  suche.addEventListener('blur', () => { treffer.hidden = true; });

  /* ---------- Werkzeuge ---------- */
  $('#liste-auf').onclick = () => { if (Z.wahl === LISTE) schliessen(); else oeffnen(LISTE); };
  $('#w-plus').onclick = () => karte.zoom(1.6);
  $('#w-minus').onclick = () => karte.zoom(1 / 1.6);
  $('#w-ansicht').onclick = () => karte.ansichtWechseln();
  $('#w-start').onclick = () => karte.start();
  window.addEventListener('keydown', (ev) => {
    if (ev.key !== 'Escape') return;
    if (Z.firmenOffen) { Z.firmenOffen = false; firmenZeigen(); } else if (Z.wahl) schliessen();
  });
  // Klick neben die Firmenwahl schließt sie
  document.addEventListener('pointerdown', (ev) => { if (Z.firmenOffen && !$('#firmen').contains(ev.target)) { Z.firmenOffen = false; firmenZeigen(); } });
  const platzhalter = () => { suche.placeholder = breit() ? 'Gebäude suchen' : 'Gebäude'; };
  window.addEventListener('resize', () => { raender(); platzhalter(); });
  platzhalter();

  /* ---------- Daten nachladen ---------- */
  raender();
  neuRechnen();
  try { [Z.team, Z.ich] = await Promise.all([daten.team(), daten.ich()]); } catch (e) { Z.team = []; Z.ich = null; }
  const gemerkt = lies('fpk-filter'); if (gemerkt && filterGueltig(gemerkt)) Z.filter = gemerkt;
  try { const f = JSON.parse(lies('fpk-firmen') || '[]'); if (Array.isArray(f)) Z.firmen = f.filter((x) => typeof x === 'string'); } catch (e) { Z.firmen = []; }
  const [a, i] = await Promise.allSettled([daten.aufgaben(), daten.index()]);
  if (a.status === 'fulfilled') Z.aufgaben = a.value.liste; else Z.fehlerAufgaben = msg(a.reason);
  if (i.status === 'fulfilled') { Z.index = i.value.gebaeude; Z.indexStand = i.value.stand; Z.indexFehlt = i.value.fehlt; } else Z.indexFehlt = true;
  karte.messen();
  neuRechnen();
  ausAdresse();
  document.documentElement.dataset.bereit = '1';
})();
