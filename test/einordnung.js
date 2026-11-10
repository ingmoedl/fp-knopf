/*
  FP-Knopf – Einordnung nach festen Regeln (ohne Claude)
  ------------------------------------------------------
  Liefert sofort beim Öffnen einer Mail: Gebäude, Partei, Ablageordner, Prüffrist der Planplattform.
  Jede Angabe trägt eine Ampel (gruen, gelb, rot) und einen Grund in einem Satz.
  Reine Funktionen ohne Zugriff auf Outlook oder SharePoint, damit sie sich im Browser testen lassen
  (werkzeuge/test_einordnung.html).

  Register: regeln/gebaeude.json und regeln/parteien.json, siehe werkzeuge/register_bauen.py.
*/
(function (wurzel) {
  'use strict';

  const NUMMER = '\\d{3}[a-z]?';
  // Nummer mit möglichem Teil (123.1, 123.x), danach keine weitere Ziffer. So wird aus einer Jahreszahl keine Gebäudenummer.
  const NUM = NUMMER + '(?:\\.[\\dx])?(?!\\d)';
  // "Geb. 123", "Geb.123", "GEB 123", "Gebäude 123", auch Aufzählungen wie "Geb. 123 und 456"
  const MIT_MARKE = new RegExp('\\bGeb(?:äude|aeude)?\\.?\\s*(' + NUM + '(?:\\s*(?:,|/|\\+|&|und|u\\.|bis|-)\\s*' + NUM + ')*)', 'gi');
  // Projektkürzel vor der Nummer wie "FP_123_" oder "FPA 123". Der Unterstrich davor zählt nicht als Wortgrenze, deshalb die eigene Klasse.
  const MIT_PROJEKT = new RegExp('(?:^|[^A-Za-z0-9])FPA?[ _-]{1,3}(' + NUMMER + ')(?!\\d)', 'gi');
  // Antwort- und Weiterleitungskürzel am Anfang, auch gestapelt und mit Datum dazwischen
  const VORSPANN = /^\s*(?:(\d{6})\s+|((?:AW|WG|RE|FW|FWD|Antwort|Zugesagt|Abgelehnt|Mit Vorbehalt)\s*:)\s*|\[(?:External|Extern|EXT)\]\s*)/i;

  function nummerNormal(roh, reg) {
    let n = String(roh).toLowerCase().replace(/\s/g, '');
    const zu = reg.gehoert_zu || {};
    if (zu[n]) return zu[n];
    const ohneTeil = n.replace(/\.[\dx]$/, '');           // 123.3 -> 123
    if (zu[ohneTeil]) return zu[ohneTeil];
    return ohneTeil;
  }

  function registerIndex(reg) {
    if (reg._index) return reg._index;
    const idx = {};
    for (const g of reg.gebaeude || []) {
      idx[String(g.nr).toLowerCase()] = g;
      for (const a of g.aliase || []) if (/^\d{3}[a-z]?$/i.test(a)) idx[String(a).toLowerCase()] = g;
    }
    Object.defineProperty(reg, '_index', { value: idx, enumerable: false });
    return idx;
  }

  // Zerlegt "123 und 456", "123.1-3", "123-130" in einzelne Nummern. Bereiche werden nicht aufgefüllt.
  function nummernAus(text) {
    return (String(text).match(new RegExp(NUMMER + '(?:\\.[\\dx])?', 'gi')) || []);
  }

  /*
    Sucht Gebäudenummern.
    betreff, text: Zeichenketten. Rückgabe: Liste { nr, name, ampel, grund }, stärkster Beleg zuerst.
      gruen  ausdrücklich genannt ("Geb. 123", "FP_123_") oder Nummer im Betreff und zusätzlich im Text bestätigt
      gelb   nur blanke Nummer im Betreff, oder nur im Text genannt
      rot    gibt es hier nicht. Ohne Fund ist die Liste leer und der Aufrufer zeigt "kein Gebäude erkannt".
  */
  function gebaeudeFinden(betreff, text, reg) {
    const idx = registerIndex(reg);
    const fund = {};                                      // nr -> { betreffMarke, betreffBlank, textMarke }
    const merke = (roh, feld) => {
      const nr = nummerNormal(roh, reg);
      if (!idx[nr]) { (fund._unbekannt = fund._unbekannt || {})[nr] = feld; return; }
      (fund[nr] = fund[nr] || {})[feld] = true;
    };
    const b = String(betreff || ''), t = String(text || '');

    for (const [quelle, feld] of [[b, 'betreffMarke'], [t, 'textMarke']]) {
      let m;
      MIT_MARKE.lastIndex = 0;
      while ((m = MIT_MARKE.exec(quelle))) for (const n of nummernAus(m[1])) merke(n, feld);
      MIT_PROJEKT.lastIndex = 0;
      while ((m = MIT_PROJEKT.exec(quelle))) merke(m[1], feld);
    }

    // Blanke Nummern nur im Betreff und nur, wenn sie im Register stehen. Datumsangaben, Maße und Kennungen fallen heraus.
    const ohneVorspann = betreffKern(b).kern;
    const blank = new RegExp('(^|[\\s_(/+&,;:-])(' + NUMMER + '(?:\\.[\\dx])?)(?=$|[\\s_)/+&,;:.-])', 'gi');
    let m;
    while ((m = blank.exec(ohneVorspann))) {
      // Teil einer Kennung oder eines Datums wie "2026-123" oder "12.11.123": Ziffer, Trennzeichen, Nummer
      if (m[1] && /[-.\/]/.test(m[1]) && /\d/.test(ohneVorspann.charAt(m.index - 1))) continue;
      const davor = ohneVorspann.slice(Math.max(0, m.index - 4), m.index + m[1].length);
      if (/(?:DN|KV|Az\.?|Nr\.?|Tel\.?|Flur)\s*$/i.test(davor)) continue;      // Nennweiten, Kennungen, Flurnummern
      const danach = ohneVorspann.slice(m.index + m[0].length, m.index + m[0].length + 6);
      if (/^\s*(?:m[²2³3]?|mm|cm|kV|V|kW|A|€|EUR|Stk|%)\b/i.test(danach)) continue; // 400 V, 200 mm
      if (/^[.\/-]\d/.test(danach) && !/\.[\dx]$/.test(m[2])) continue;          // 12.11., 2026-11
      const nr = nummerNormal(m[2], reg);
      if (idx[nr]) (fund[nr] = fund[nr] || {}).betreffBlank = true;
    }

    // Dokumentnummern wie "ABC-123-4-…" tragen das Gebäude an zweiter Stelle
    const DOK = /(?:^|[^A-Za-z0-9])[A-Z]{2,4}-(\d{3}[a-z]?)-\d/g;
    for (const quelle of [b, t]) {
      DOK.lastIndex = 0;
      while ((m = DOK.exec(quelle))) { const nr = nummerNormal(m[1], reg); if (idx[nr]) (fund[nr] = fund[nr] || {}).dokument = true; }
    }
    // Eindeutige Gebäudenamen im Betreff (Liste "namen" im Register, bewusst kurz gehalten)
    for (const g of reg.gebaeude || []) {
      for (const name of g.namen || []) {
        const re = new RegExp('(^|[^A-Za-zÄÖÜäöüß])' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i');
        if (re.test(ohneVorspann)) { const k = String(g.nr).toLowerCase(); (fund[k] = fund[k] || {}).name = name; }
      }
    }

    const aus = [];
    for (const nr of Object.keys(fund)) {
      if (nr === '_unbekannt') continue;
      const f = fund[nr], g = idx[nr];
      let ampel, grund;
      if (f.betreffMarke) { ampel = 'gruen'; grund = 'Im Betreff ausdrücklich genannt.'; }
      else if (f.betreffBlank && f.textMarke) { ampel = 'gruen'; grund = 'Nummer im Betreff, im Text bestätigt.'; }
      else if (f.betreffBlank) { ampel = 'gelb'; grund = 'Nummer steht im Betreff ohne „Geb.“.'; }
      else if (f.textMarke) { ampel = 'gelb'; grund = 'Nur im Text genannt, nicht im Betreff.'; }
      else if (f.dokument) { ampel = 'gelb'; grund = 'Aus der Dokumentnummer abgeleitet.'; }
      else { ampel = 'gelb'; grund = 'Name „' + f.name + '“ im Betreff, keine Nummer.'; }
      aus.push({ nr: g.nr, name: g.name, ampel, grund });
    }
    const rang = { gruen: 0, gelb: 1 };
    aus.sort((x, y) => rang[x.ampel] - rang[y.ampel] || String(x.nr).localeCompare(String(y.nr)));
    aus.unbekannt = Object.keys(fund._unbekannt || {});    // genannte Nummern, die nicht im Register stehen
    return aus;
  }

  // Bereich für Themen ohne Gebäude. Die Stichworte stehen im Register (bereichsregeln: [{ muster, name }]),
  // nicht im Code. Ohne Treffer gibt es keinen Bereich.
  function bereichFinden(betreff, reg) {
    const b = String(betreff || '');
    for (const r of (reg && reg.bereichsregeln) || []) {
      let re;
      try { re = new RegExp(r.muster, r.schalter === undefined ? 'i' : r.schalter); } catch (e) { continue; }
      if (re.test(b)) return { name: r.name, ampel: 'gelb', grund: 'Stichwort im Betreff.' };
    }
    return null;
  }

  // Partei aus der Absenderadresse. Unbekannte Domain ergibt rot, nie eine geratene Partei.
  function parteiFinden(adresse, reg) {
    const dom = String(adresse || '').toLowerCase().split('@')[1] || '';
    if (!dom) return { ordner_kom: '', ampel: 'rot', grund: 'Keine Absenderadresse.' };
    for (const p of reg.parteien || []) {
      for (const d of p.domains || []) {
        if (dom === d || dom.endsWith('.' + d)) {
          return { ordner_kom: p.ordner_kom, ordner_ext: p.ordner_ext || '', rolle: p.rolle || '', ampel: 'gruen', grund: 'Absender ' + d + ' ist bekannt.' };
        }
      }
    }
    return { ordner_kom: '', ordner_ext: '', rolle: '', ampel: 'rot', grund: 'Absender ' + dom + ' ist keiner Partei zugeordnet.' };
  }

  // Benachrichtigung der Planplattform mit Prüffrist. Domain und Muster stehen im Register (pruefplattform).
  function pruefFrist(adresse, betreff, reg) {
    const pp = reg && reg.pruefplattform;
    if (!pp || !pp.domain) return null;
    const dom = String(adresse || '').toLowerCase().split('@')[1] || '';
    if (!(dom === pp.domain || dom.endsWith('.' + pp.domain))) return null;
    let m = null;
    try { m = new RegExp(pp.fristMuster, 'i').exec(String(betreff || '')); } catch (e) { m = null; }
    const frist = m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : '';
    return { art: frist ? 'pruefung' : 'eingang', frist, name: pp.name || 'Planplattform' };
  }

  // Trennt den Vorspann aus Datumspräfixen und Antwortkürzeln vom eigentlichen Betreff.
  function betreffKern(betreff) {
    let rest = String(betreff || ''), art = '', m;
    while ((m = VORSPANN.exec(rest))) {
      if (m[2] && !art) {
        const k = m[2].replace(/\s*:$/, '').toUpperCase();
        if (k === 'AW' || k === 'RE' || k === 'ANTWORT') art = 'AW';
        else if (k === 'WG' || k === 'FW' || k === 'FWD') art = 'WG';
      }
      rest = rest.slice(m[0].length);
    }
    return { art, kern: rest.replace(/\s+/g, ' ').replace(/[\s\-–|_]+$/, '').trim() };
  }

  const jjmmtt = d => String(d.getFullYear()).slice(-2) + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);

  /*
    Sauberer Betreff beim Senden: genau ein Datum vorn, alte Präfixe weg, Gebäude ergänzt, wenn es fehlt.
    art: 'AW' für Antwort, 'WG' für Weiterleitung, '' für neue Mail. Der Betreff wird nicht gekürzt.
  */
  function betreffSauber(betreff, datum, art, gebaeudeNr) {
    const k = betreffKern(betreff);
    let kern = k.kern;
    if (gebaeudeNr && !new RegExp('(^|\\D)' + String(gebaeudeNr).replace(/[.]/g, '\\.') + '(\\D|$)', 'i').test(kern)) kern = 'Geb. ' + gebaeudeNr + ' ' + kern;
    return jjmmtt(datum) + ' ' + (art ? art + ': ' : '') + kern;
  }

  // Dateiname für die Ablage: Datum vorn, unzulässige Zeichen ersetzt, auf maxLaenge gekürzt. Endung kommt vom Aufrufer.
  function dateiname(betreff, datum, maxLaenge) {
    const k = betreffKern(betreff);
    let n = jjmmtt(datum) + ' ' + (k.art ? k.art + ' ' : '') + k.kern;
    n = n.replace(/["*:<>?\/\\|#%~&{}]/g, '_').replace(/\s+/g, ' ').replace(/[. ]+$/, '');
    const max = maxLaenge || 100;
    return n.length > max ? n.slice(0, max).replace(/[. _]+$/, '') : n;
  }

  const norm = s => String(s).toLowerCase().replace(/[^a-z0-9äöüß]+/g, ' ').trim();
  // Titel einer Aufgabenmappe ohne Kennung "[1a2b3c4d]" und ohne führende Gebäudenummern
  function aufgabeTitel(mappe) {
    return norm(String(mappe).replace(/\s*\[[0-9a-f]{8}\]\s*$/i, '')
      .replace(/^\s*(?:(?:FP|NST|Geb\.?)[ _]*)?(?:\d{3}[a-z]?(?:\.[\dx])?[\s_+&,.-]*)+/i, ''));
  }

  /*
    Sucht die Aufgabenmappe, zu der ein Betreff gehört.
    gruen  Betreff und Titel sind gleich
    gelb   einer enthält den anderen (mindestens 12 Zeichen), etwa bei abgeschnittenen Betreffs
  */
  function aufgabeFinden(betreff, mappen) {
    const kern = norm(betreffKern(betreff).kern.replace(/^(?:FPA?\s*[:|_-]\s*)/i, ''));
    const kernOhneNr = aufgabeTitel(betreffKern(betreff).kern);
    let best = null;
    for (const mappe of mappen || []) {
      const titel = aufgabeTitel(mappe);
      if (titel.length < 4) continue;
      let treffer = null;
      if (titel === kern || titel === kernOhneNr) treffer = { mappe, ampel: 'gruen', grund: 'Betreff entspricht der Aufgabe.', laenge: titel.length + 1000 };
      else if (titel.length >= 12 && (kern.includes(titel) || (kernOhneNr.length >= 12 && titel.includes(kernOhneNr)))) {
        treffer = { mappe, ampel: 'gelb', grund: 'Betreff ähnelt der Aufgabe.', laenge: titel.length };
      }
      if (treffer && (!best || treffer.laenge > best.laenge)) best = treffer;
    }
    if (best) delete best.laenge;
    return best;
  }

  /*
    Gesamte Einordnung einer Mail.
    mail: { von, betreff, text, eingang (Date), anhaenge: [name] }
    reg:  { gebaeude, gehoert_zu, parteien }
    mappen: Namen der offenen Aufgabenmappen (optional)
  */
  function einordnen(mail, reg, mappen) {
    const gebaeude = gebaeudeFinden(mail.betreff, mail.text, reg);
    const aufgabe = aufgabeFinden(mail.betreff, mappen);
    if (aufgabe && !gebaeude.length) {
      // Kein Gebäude im Betreff, aber die Aufgabe trägt eines im Namen
      for (const g of gebaeudeFinden(aufgabe.mappe, '', reg)) gebaeude.push({ nr: g.nr, name: g.name, ampel: 'gelb', grund: 'Über die Aufgabe „' + aufgabe.mappe + '“ erkannt.' });
    }
    const bereich = gebaeude.length ? null : bereichFinden(mail.betreff, reg);
    const partei = parteiFinden(mail.von, reg);
    // Absender aus dem eigenen Haus (gleiche Domain wie das angemeldete Postfach): Bei einer Weiterleitung
    // sagt der Absender nichts über die Partei, die Ablage ist dann nur ein Anhalt.
    const vonDomain = String(mail.von || '').toLowerCase().split('@')[1] || '';
    const intern = !!mail.eigeneDomain && vonDomain === String(mail.eigeneDomain).toLowerCase();
    const pv = pruefFrist(mail.von, mail.betreff, reg);
    const ext = (reg.ablage && reg.ablage.ext) || 'der Ablage für Unterlagen';
    const vorschlaege = [];

    if (pv && pv.art === 'pruefung') {
      vorschlaege.push({ art: 'pruefaufgabe', frist: pv.frist, name: pv.name, ampel: 'gruen', grund: pv.name + ' nennt die Prüffrist im Betreff.' });
    }
    if (aufgabe) vorschlaege.push({ art: 'aufgabe_zuordnen', mappe: aufgabe.mappe, ampel: aufgabe.ampel, grund: aufgabe.grund });
    if (partei.ordner_kom) {
      vorschlaege.push(intern
        ? { art: 'ablegen', ziel_kom: partei.ordner_kom, ampel: 'gelb', grund: 'Absender ist das eigene Haus. Bei einer Weiterleitung kann die Partei im Verlauf die richtige sein.' }
        : { art: 'ablegen', ziel_kom: partei.ordner_kom, ampel: 'gruen', grund: 'Ablage folgt dem Absender.' });
    } else {
      vorschlaege.push({ art: 'ablegen', ziel_kom: '', ampel: 'rot', grund: 'Kein Ablageordner bekannt, bitte wählen.' });
    }
    if ((mail.anhaenge || []).length) {
      vorschlaege.push(partei.ordner_ext
        ? { art: 'anhang_ablegen', ziele: [partei.ordner_ext], ampel: 'gelb', grund: 'Ordner der Partei in ' + ext + ', Unterordner nicht geprüft.' }
        : { art: 'anhang_ablegen', ziele: [], ampel: 'rot', grund: 'Kein passender Ordner in ' + ext + ' gefunden.' });
    }
    return {
      durch: 'regeln',
      gebaeude: gebaeude.map(g => g),
      unbekannteNummern: gebaeude.unbekannt || [],
      bereich,
      partei,
      pruefplattform: pv,
      vorschlaege,
      // Was Regeln nicht können, sagt der Reiter offen: keine Aussage zu Antwort, Zuständigkeit oder Dringlichkeit.
      offen: ['ob eine Antwort nötig ist', 'wer zuständig ist', 'Thema und Zusammenfassung'],
    };
  }

  const api = { gebaeudeFinden, bereichFinden, parteiFinden, pruefFrist, betreffKern, betreffSauber, dateiname, aufgabeFinden, einordnen };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else wurzel.FPEinordnung = api;
})(typeof window !== 'undefined' ? window : this);
