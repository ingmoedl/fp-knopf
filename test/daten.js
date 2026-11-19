/*
  FP-Karte – Ladeschicht
  ----------------------
  Die einzige Stelle, an der die Karte Daten holt. Karte und Gebäudeseite kennen nur die fertigen Formen
  weiter unten, nie SharePoint selbst.

  Quelle ist immer SharePoint über die Funktion sp(pfad, {method, body}), genau wie in der Übersicht.
  Für den Bau ohne Anmeldung liefert quelle-lokal.js ein nachgebildetes sp, das dieselben Aufrufe aus
  lokalen Dateien beantwortet. So läuft der ganze Weg auch mit Beispieldaten durch.

  Gelesen wird:
    Bibliothek der Aufgaben   Mappen des Stammordners (je Ordner abgefragt, wegen des Schwellenwerts)
                              _Import/team.json
    Bibliothek des Wissens    karte/karte.json        Geometrie
                              karte/gebaeude.json     Gebäude mit Lage und Höhe
                              regeln/gebaeude.json    Register (Zugehörigkeiten, Namen)
                              gebaeude/_index.json    Übersicht je Gebäude
                              gebaeude/<Nr>.json      Mails, Fristen und offene Punkte je Gebäude

  FPKarteDaten.erzeugen({ sp, cfg, konto, gebaeudeAusText }) liefert:
    grundlagen()  { karte, gebaeude: Map(nr → Gebäude), register }
    team()        [{ kuerzel, name, mail, hinter, schrift }]
    ich()         { kuerzel, name, mail } oder null
    aufgaben()    { liste: [Aufgabe], doppelt: Anzahl zusammengelegter Mappen }
    index()       { stand, gebaeude: Map(nr → { mails, handlungsbedarf, offene_punkte, naechste_frist, letzte_mail }) }
    gebaeude(nr)  { nr, name, stand, mails, fristen, offene_punkte } (leer, wenn es keine Datei gibt)

  Aufgabe: { id, name, titel, rel, url, firma, ansprech, kuerzel: [..], leute: [mail], faellig: 'JJJJ-MM-TT',
             geaendert, doppel, gebaeude: [{ nr, ampel, grund, quelle: 'spalte' | 'name' }], bereich }
*/
(function (wurzel) {
  'use strict';

  const z2 = (n) => String(n).padStart(2, '0');
  const tagIso = (d) => `${d.getFullYear()}-${z2(d.getMonth() + 1)}-${z2(d.getDate())}`;
  const ausSp = (iso) => { if (!iso) return ''; const d = new Date(iso); return isNaN(d) ? '' : tagIso(d); };   // SharePoint speichert UTC
  const woerter = (s) => String(s || '').split(/[^A-Za-zÄÖÜäöüß]+/).filter(Boolean);
  const pfad = (s) => encodeURIComponent(String(s).replace(/'/g, "''"));
  const KENNUNG = /\s*\[[0-9a-f]{8}\]\s*$/i;
  const FARBE = /^#[0-9a-f]{6}$/i;

  /* Team wie in der Übersicht streng prüfen, die Farben landen später in der Darstellung. */
  function teamPruefen(liste, domain) {
    return (Array.isArray(liste) ? liste : []).filter((b) => b &&
      /^[A-Z]{2,4}$/.test(b.kuerzel) && typeof b.name === 'string' && b.name.length <= 80 &&
      /^[a-z0-9._%+-]{1,64}@[a-z0-9.-]{1,100}$/i.test(b.mail || '') && (!domain || b.mail.toLowerCase().endsWith('@' + String(domain).toLowerCase())) &&
      FARBE.test(b.hinter) && FARBE.test(b.schrift)).slice(0, 30)
      .map((b) => ({ kuerzel: b.kuerzel, name: b.name, mail: b.mail.toLowerCase(), hinter: b.hinter, schrift: b.schrift }));
  }

  /* Gebäude der Karte, ergänzt um alles, was nur im Register steht. */
  function gebaeudeListe(lagen, register) {
    const aus = new Map();
    const roh = lagen && lagen.gebaeude ? lagen.gebaeude : {};
    for (const g of Array.isArray(roh) ? roh : Object.values(roh)) {
      if (!g || g.nr == null) continue;
      const lage = g.lage && typeof g.lage === 'object' ? g.lage : { art: 'unbekannt' };
      aus.set(String(g.nr), {
        nr: String(g.nr), name: String(g.name || ''), areal: String(g.areal || ''), gruppe: String(g.gruppe || ''),
        planer: String(g.planer || ''), status: String(g.status || ''), aliase: Array.isArray(g.aliase) ? g.aliase.map(String) : [],
        hoehe: Number(g.hoehe) > 0 ? Number(g.hoehe) : 8, lage,
      });
    }
    for (const g of (register && register.gebaeude) || []) {
      if (!g || g.nr == null || aus.has(String(g.nr))) continue;
      aus.set(String(g.nr), {
        nr: String(g.nr), name: String(g.name || ''), areal: String(g.areal || ''), gruppe: String(g.gruppe || ''), planer: '', status: '',
        aliase: Array.isArray(g.aliase) ? g.aliase.map(String) : [], hoehe: 8, lage: { art: 'unbekannt' },
      });
    }
    return aus;
  }

  /* Nachschlagen einer Nummer: genaue Nummer, Alias oder Zugehörigkeit aus dem Register. */
  function nummernIndex(gebaeude, register) {
    const idx = new Map();
    for (const g of gebaeude.values()) idx.set(g.nr.toLowerCase(), g.nr);
    for (const g of gebaeude.values()) for (const a of g.aliase) if (!idx.has(a.toLowerCase())) idx.set(a.toLowerCase(), g.nr);
    const zu = (register && register.gehoert_zu) || {};
    for (const k of Object.keys(zu)) { const ziel = idx.get(String(zu[k]).toLowerCase()); if (ziel && !idx.has(k.toLowerCase())) idx.set(k.toLowerCase(), ziel); }
    return idx;
  }

  /* Spalte „Gebäude“ der Aufgabe: mehrere Nummern mit Semikolon, sonst ein Bereich. */
  function spalteLesen(wert, idx) {
    const nrn = [], bereiche = [];
    for (const t of String(wert || '').split(';').map((s) => s.trim()).filter(Boolean)) {
      const k = t.toLowerCase().replace(/^geb(?:äude|aeude)?\.?\s*/, '').replace(/\s+/g, ' ');
      const nr = idx.get(k) || idx.get(k.replace(/\s/g, '')) || idx.get(k.replace(/\.[\dx]$/, ''));
      if (nr) { if (!nrn.includes(nr)) nrn.push(nr); } else bereiche.push(t);
    }
    return { nrn, bereich: bereiche.join('; ') };
  }

  /* Nummern, mit denen der Name einer Mappe beginnt: „660 …“, „660_…“, „660_279 …“, auch nach „FP_“.
     Das Team stellt die Gebäudenummer voran. Steht dort eine Nummer aus dem Register, gilt die Zuordnung als sicher.
     Gesehen an den echten FP-Aufgaben am 19.11.2026. */
  function nummernVorn(titel, idx) {
    const m = /^\s*(?:FP[_\s]+)?(\d{2,4}[a-z]?(?:\.[\dx])?(?:\s?[_\/+&,]\s?\d{2,4}[a-z]?(?:\.[\dx])?)*)(?=$|[\s_:.,;–-])/i.exec(String(titel || ''));
    if (!m) return [];
    const aus = [];
    for (const t of m[1].split(/\s?[_\/+&,]\s?/)) {
      const k = t.toLowerCase();
      const nr = idx.get(k) || idx.get(k.replace(/\.[\dx]$/, ''));
      if (nr && !aus.includes(nr)) aus.push(nr);
    }
    return aus;
  }

  /* Eine Zeile aus SharePoint in die Form der Karte bringen. */
  function aufgabeAusZeile(x, team, idx, register, gebaeudeAusText, url) {
    const p = Array.isArray(x.FPBearbeiter) ? x.FPBearbeiter : [];
    const leute = p.map((q) => String((q && q.EMail) || '').toLowerCase()).filter(Boolean);
    const ansprech = String(x.FPAnsprechperson || '').trim();
    const w = woerter(ansprech);
    const kuerzel = team.filter((b) => w.includes(b.kuerzel) || leute.includes(b.mail)).map((b) => b.kuerzel);
    const name = String(x.FileLeafRef || '');
    const titel = name.replace(KENNUNG, '').trim();

    let gebaeude = [], bereich = '';
    const spalte = spalteLesen(x.FPGebaeude, idx);
    if (spalte.nrn.length || spalte.bereich) {
      gebaeude = spalte.nrn.map((nr) => ({ nr, ampel: 'gruen', grund: 'Steht in der Spalte „Gebäude“ der Aufgabe.', quelle: 'spalte' }));
      bereich = spalte.bereich;
    } else if (gebaeudeAusText) {
      // Solange die Spalte leer ist: Nummer aus dem Namen der Mappe
      let fund = [];
      try { fund = gebaeudeAusText(titel, register) || []; } catch (e) { fund = []; }
      const vorn = nummernVorn(titel, idx);
      for (const f of fund) {
        const nr = idx.get(String(f.nr).toLowerCase());
        if (!nr || gebaeude.some((g) => g.nr === nr)) continue;
        if (vorn.includes(nr)) gebaeude.push({ nr, ampel: 'gruen', grund: 'Die Nummer steht am Anfang des Mappennamens.', quelle: 'name' });
        else gebaeude.push({ nr, ampel: f.ampel === 'gruen' ? 'gruen' : 'gelb', grund: 'Aus dem Namen der Mappe. ' + String(f.grund || ''), quelle: 'name' });
      }
      // Reihenfolge wie im Namen: die Nummern vom Anfang zuerst
      gebaeude.sort((a, b) => (vorn.indexOf(a.nr) < 0 ? 99 : vorn.indexOf(a.nr)) - (vorn.indexOf(b.nr) < 0 ? 99 : vorn.indexOf(b.nr)));
    }
    return {
      id: x.Id, name, titel, rel: String(x.FileRef || ''), url: url(String(x.FileRef || '')),
      firma: String(x.FPFirma || '').trim(), ansprech, kuerzel, leute,
      faellig: ausSp(x.FPFaelligkeit), geaendert: String(x.Modified || ''), doppel: 1, gebaeude, bereich,
    };
  }

  /* Mappen, die es einmal ohne und einmal mit Kennung gibt, zeigt die Karte einmal. Es bleibt die mit Kennung,
     sonst die zuletzt geänderte. Entfernt wird nichts, das ist nur die Anzeige. */
  function doppelteZusammenlegen(liste) {
    const nach = new Map();
    for (const a of liste) {
      const k = a.titel.toLowerCase();
      const alt = nach.get(k);
      if (!alt) { nach.set(k, a); continue; }
      const besser = (KENNUNG.test(a.name) && !KENNUNG.test(alt.name)) || (KENNUNG.test(a.name) === KENNUNG.test(alt.name) && a.geaendert > alt.geaendert) ? a : alt;
      besser.doppel = alt.doppel + a.doppel;
      nach.set(k, besser);
    }
    return [...nach.values()];
  }

  const leereDatei = (nr, name) => ({ nr, name: name || '', stand: '', mails: [], fristen: [], offene_punkte: [], fehlt: true });
  function gebaeudeDatei(d, nr, name) {
    if (!d || typeof d !== 'object') return leereDatei(nr, name);
    const liste = (x) => (Array.isArray(x) ? x.filter((e) => e && typeof e === 'object') : []);
    return { nr, name: String(d.name || name || ''), stand: String(d.stand || ''), mails: liste(d.mails), fristen: liste(d.fristen), offene_punkte: liste(d.offene_punkte), fehlt: false };
  }

  function erzeugen(o) {
    const sp = o.sp, cfg = o.cfg;
    const gebaeudeAusText = o.gebaeudeAusText || null;
    const merke = new Map();
    const einmal = (k, fn) => { if (!merke.has(k)) merke.set(k, fn().catch((e) => { merke.delete(k); throw e; })); return merke.get(k); };

    /* Datei aus einer Bibliothek. Fehlt sie, kommt null zurück. */
    async function datei(bibliothek, name) {
      try { return await sp(`${cfg.web}/_api/web/GetFileByServerRelativePath(decodedurl='${pfad(bibliothek + '/' + name)}')/$value`); }
      catch (e) { if (e && e.status === 404) return null; throw e; }
    }

    /* Mappen des Stammordners, nicht rekursiv, per CAML über GetItems (wie in der Übersicht). Eine Abfrage über die
       ganze Bibliothek scheitert am Schwellenwert von 5.000 Einträgen. GetItems liefert Personen nur als Id. */
    const CAML = "<View><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq></Where></Query><RowLimit>5000</RowLimit></View>";
    const SPALTEN = 'Id,FileLeafRef,FileRef,FPFirma,FPAnsprechperson,FPFaelligkeit,FPStatus,FPErledigtAm,Modified,FPBearbeiterId';
    const nutzer = (id) => einmal('nutzer:' + id, () => sp(`${cfg.web}/_api/web/getuserbyid(${Number(id)})?$select=Title,Email`)
      .then((u) => ({ Title: (u && u.Title) || '', EMail: (u && u.Email) || '' }), () => ({ Title: '', EMail: '' })));
    async function mappen() {
      const frage = (spalten) => sp(`${cfg.web}/_api/web/lists(guid'${cfg.listId}')/GetItems?$select=${spalten}`,
        { method: 'POST', body: { query: { ViewXml: CAML, FolderServerRelativeUrl: cfg.libRel } } });
      let r;
      try { r = await frage(SPALTEN + ',FPGebaeude'); }
      catch (e) { if (e && e.status === 400) r = await frage(SPALTEN); else throw e; }      // Spalte „Gebäude“ fehlt noch
      const liste = (r && r.value) || [];
      for (const x of liste) x.FPBearbeiter = await Promise.all((Array.isArray(x.FPBearbeiterId) ? x.FPBearbeiterId : []).map(nutzer));
      return liste;
    }

    const grundlagen = () => einmal('grundlagen', async () => {
      const [karte, lagen, register] = await Promise.all([
        datei(cfg.wissenRel, 'karte/karte.json'), datei(cfg.wissenRel, 'karte/gebaeude.json'), datei(cfg.wissenRel, 'regeln/gebaeude.json'),
      ]);
      if (!karte) throw new Error('Die Kartengeometrie fehlt (karte/karte.json).');
      const reg = register || { gebaeude: [], gehoert_zu: {} };
      const gebaeude = gebaeudeListe(lagen, reg);
      return { karte, gebaeude, register: reg, nummern: nummernIndex(gebaeude, reg) };
    });

    const team = () => einmal('team', async () => teamPruefen(await datei(cfg.libRel, '_Import/team.json'), o.maildomain));

    const ich = () => einmal('ich', async () => {
      const k = o.konto || {};
      const mail = String(k.mail || '').toLowerCase();
      if (!mail) return null;
      const b = (await team()).find((x) => x.mail === mail);
      return b ? { kuerzel: b.kuerzel, name: b.name, mail } : { kuerzel: '', name: String(k.name || mail), mail };
    });

    async function aufgaben() {
      const [zeilen, t, g] = await Promise.all([mappen(), team(), grundlagen()]);
      const url = (rel) => (cfg.host && rel ? cfg.host + rel.split('/').map(encodeURIComponent).join('/') : '');
      const offen = zeilen.filter((x) => x && x.FPStatus === 'Offen' && String(x.FileRef || '').slice(0, String(x.FileRef || '').lastIndexOf('/')) === cfg.libRel);
      const alle = offen.map((x) => aufgabeAusZeile(x, t, g.nummern, g.register, gebaeudeAusText, url));
      const liste = doppelteZusammenlegen(alle);
      return { liste, doppelt: alle.length - liste.length };
    }

    const index = () => einmal('index', async () => {
      const d = await datei(cfg.wissenRel, 'gebaeude/_index.json');
      const aus = new Map();
      for (const e of (d && Array.isArray(d.gebaeude) ? d.gebaeude : [])) {
        if (!e || e.nr == null) continue;
        aus.set(String(e.nr), {
          mails: Number(e.mails) || 0, handlungsbedarf: Number(e.handlungsbedarf) || 0, offene_punkte: Number(e.offene_punkte) || 0,
          naechste_frist: String(e.naechste_frist || ''), letzte_mail: String(e.letzte_mail || ''),
        });
      }
      return { stand: String((d && d.stand) || ''), gebaeude: aus, fehlt: !d };
    });

    const gebaeude = (nr) => einmal('gebaeude:' + nr, async () => {
      const [grund, i] = await Promise.all([grundlagen(), index().catch(() => null)]);
      const g = grund.gebaeude.get(String(nr)), name = g ? g.name : '';
      // Der Index nennt jedes Gebäude, zu dem es eine Datei gibt. Steht es dort nicht, entfällt die Abfrage.
      if (i && !i.fehlt && !i.gebaeude.has(String(nr))) return gebaeudeDatei(null, String(nr), name);
      return gebaeudeDatei(await datei(cfg.wissenRel, 'gebaeude/' + String(nr) + '.json'), String(nr), name);
    });

    return { grundlagen, team, ich, aufgaben, index, gebaeude, vergessen: () => merke.clear() };
  }

  wurzel.FPKarteDaten = { erzeugen, teamPruefen, gebaeudeListe, nummernIndex, spalteLesen, nummernVorn, aufgabeAusZeile, doppelteZusammenlegen, gebaeudeDatei };
})(typeof window !== 'undefined' ? window : this);
