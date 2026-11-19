/*
  FP-Karte – Gebäudeseite
  -----------------------
  Zeigt zu einem Gebäude, was jetzt zählt: Aufgaben, Termine und Fristen, offene Punkte, darunter zugeklappt
  den Verlauf der Mails. Die Reihenfolge kommt aus auswertung.js.

  Alle Texte aus Aufgaben und Mails sind fremder Inhalt. Sie werden nur als Text eingesetzt, nie als HTML.

  FPGebaeudeseite.zeichnen(root, modell, { beiFilter(f), beiFirmen(liste), beiZu(), behalten,
                                           beiBearbeiten(id | null), beiErledigt(a), beiFaellig(a, iso), beiNotiz(a, text), beiRueck() })
    modell: { art: 'gebaeude' | 'ohne', g, sicht (aus FPKarteAuswertung.seite), filter, firmen, ich, team, zaehler,
              laedt, datei, fehler, fehlerAufgaben, fussnote, aufDerKarte,
              bearbeiten (darf geschrieben werden), offen (Nummer der Aufgabe in Bearbeitung), arbeit (Nummer der
              Aufgabe, die gerade gespeichert wird), meldung { text, art: 'ok' | 'warn', id, rueck } }
  FPGebaeudeseite.filterLeiste(team, ich, filter, zaehler, beiFilter) baut die Tasten für den Personenfilter.
  FPGebaeudeseite.firmenWahl(liste, gewaehlt, beiWahl, offen) baut die Firmenwahl für die Leiste: eine Taste,
    darunter je Firma eine Zeile zum Anhaken. liste kommt aus FPKarteAuswertung.firmenListe.
*/
(function (wurzel) {
  'use strict';

  const A = () => wurzel.FPKarteAuswertung;

  function h(tag, eig) {
    const e = document.createElement(tag);
    for (const k of Object.keys(eig || {})) {
      const v = eig[k];
      if (v == null || v === false) continue;
      if (k === 'class') e.className = v;
      else if (k === 'text') e.textContent = v;
      else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), v);
      else e.setAttribute(k, v === true ? '' : String(v));
    }
    for (let i = 2; i < arguments.length; i++) {
      for (const c of [].concat(arguments[i])) if (c != null && c !== false) e.append(c.nodeType ? c : document.createTextNode(String(c)));
    }
    return e;
  }
  const kurz = (s, max) => { s = String(s || ''); return s.length > max ? s.slice(0, max - 1).replace(/\s+\S*$/, '') + '…' : s; };
  const mehrzahl = (n, eins, viele) => `${n} ${n === 1 ? eins : viele}`;

  // Arten der Vorschläge aus der Bewertung, in Worten
  const ARTEN = {
    antworten: 'Antworten', entscheidung: 'Entscheidung einholen', weiterleiten: 'Weiterleiten', zuweisen: 'Zuweisen',
    aufgabe_neu: 'Neue Aufgabe', aufgabe_zuordnen: 'Zur Aufgabe legen', ablegen: 'Ablegen', anhang_ablegen: 'Anhang ablegen',
    pruefaufgabe: 'Prüfaufgabe', kenntnis: 'Zur Kenntnis',
  };

  function filterLeiste(team, ich, filter, zaehler, beiFilter) {
    const z = zaehler || {};
    const taste = (wert, text, farbe, name) => {
      const b = h('button', { type: 'button', 'aria-pressed': String(filter === wert), 'aria-label': name || null, onclick: () => beiFilter(wert) });
      if (farbe) { const p = h('span', { class: 'punkt' }); p.style.background = farbe; b.append(p); }
      b.append(h('span', { text }));
      if (z[wert] !== undefined) b.append(h('span', { class: 'zahl', text: String(z[wert]) }));
      return b;
    };
    const el = h('div', { class: 'filter', role: 'group', 'aria-label': 'Aufgaben nach Person' });
    el.append(taste('alle', 'Alle'));
    if (ich && (ich.kuerzel || ich.mail)) el.append(taste('meine', 'Meine'));
    if (team.length) el.append(h('span', { class: 'trenner' }));
    for (const b of team) el.append(taste(b.kuerzel, b.kuerzel, b.hinter, 'Aufgaben von ' + b.name));
    return el;
  }

  const firmaName = (name) => name || 'Ohne Firma';
  // Nimmt eine Firma in die Auswahl oder heraus. Die Reihenfolge folgt der Liste, nicht der Klickfolge.
  function firmaUmschalten(gewaehlt, name, reihenfolge) {
    const neu = gewaehlt.includes(name) ? gewaehlt.filter((x) => x !== name) : gewaehlt.concat(name);
    return reihenfolge.filter((x) => neu.includes(x)).concat(neu.filter((x) => !reihenfolge.includes(x)));
  }

  function firmenWahl(liste, gewaehlt, beiWahl, offen) {
    gewaehlt = gewaehlt || [];
    const namen = liste.map((f) => f.name);
    const el = h('details', { class: 'firmen' + (gewaehlt.length ? ' aktiv' : ''), open: !!offen });
    el.append(h('summary', { 'aria-label': 'Aufgaben nach Firma' },
      h('span', { text: 'Firmen' }), gewaehlt.length ? h('span', { class: 'zahl', text: String(gewaehlt.length) }) : null));
    const feld = h('div', { class: 'firmen-feld', role: 'group', 'aria-label': 'Firmen wählen' });
    feld.append(h('button', { type: 'button', class: 'alle', disabled: !gewaehlt.length, onclick: () => beiWahl([]) }, 'Alle Firmen'));
    for (const f of liste) {
      const haken = h('input', { type: 'checkbox', checked: gewaehlt.includes(f.name), onchange: () => beiWahl(firmaUmschalten(gewaehlt, f.name, namen)) });
      feld.append(h('label', null, haken, h('span', { class: 'name', text: firmaName(f.name) }), h('span', { class: 'zahl', text: String(f.anzahl) })));
    }
    if (!liste.length) feld.append(h('div', { class: 'leer', text: 'Keine Aufgabe mit Firma.' }));
    el.append(feld);
    return el;
  }

  /* Firmen am Gebäude als Marken. Ein Klick nimmt die Firma in die Auswahl der ganzen Karte oder heraus. */
  function firmenMarken(liste, gewaehlt, beiWahl) {
    gewaehlt = gewaehlt || [];
    const namen = liste.map((f) => f.name);
    const el = h('div', { class: 'gs-firmen', role: 'group', 'aria-label': 'Firmen an diesem Gebäude' });
    for (const f of liste) {
      el.append(h('button', { type: 'button', 'aria-pressed': String(gewaehlt.includes(f.name)), onclick: () => beiWahl(firmaUmschalten(gewaehlt, f.name, namen)) },
        h('span', { text: firmaName(f.name) }), h('i', { text: String(f.anzahl) })));
    }
    return el;
  }

  function kuerzelMarken(a, team) {
    const el = h('span', { class: 'gs-wer' });
    for (const k of a.kuerzel) {
      const b = team.find((x) => x.kuerzel === k);
      const m = h('span', { class: 'kz', text: k });
      if (b) { m.style.background = b.hinter; m.style.color = b.schrift; }
      el.append(m);
    }
    if (!a.kuerzel.length && a.ansprech) el.append(h('span', { class: 'kz fremd', text: kurz(a.ansprech, 22) }));
    return el;
  }

  function aufgabeZeile(a, m) {
    const Aw = A();
    const klasse = 'gs-zeile' + (a.ueberfaellig ? ' ueber' : a.heute ? ' heute' : '');
    const wann = a.faellig
      ? h('span', { class: 'gs-wann' }, h('b', { text: Aw.tagKurz(a.faellig) }), h('i', { text: Aw.abstand(a.tage) }))
      : h('span', { class: 'gs-wann ohne' }, h('i', { text: 'ohne Termin' }));
    const neben = h('span', { class: 'gs-neben' });
    const teile = [];
    if (a.firma) teile.push(a.firma);
    if (a.bereich) teile.push(a.bereich);
    if (a.doppel > 1) teile.push(a.doppel + ' Mappen mit diesem Namen');
    neben.append(teile.join(' · '));
    if (a.zuordnung && a.zuordnung.ampel !== 'gruen') {
      if (teile.length) neben.append(' · ');
      neben.append(h('span', { class: 'unsicher', title: a.zuordnung.grund, text: 'Zuordnung zum Gebäude unsicher' }));
    }
    const was = h('span', { class: 'gs-was' }, h('span', { class: 'gs-titel', text: a.titel }), neben.textContent ? neben : null);
    const eig = { class: klasse };
    if (/^https:\/\//.test(a.url)) { eig.href = a.url; eig.target = '_blank'; eig.rel = 'noopener'; }
    return h(eig.href ? 'a' : 'div', eig, wann, was, kuerzelMarken(a, m.team));
  }

  /* Rückmeldung nach einer Änderung. Der Text kann den Namen einer Aufgabe enthalten und wird nur als Text eingesetzt. */
  function meldungEl(mel, o, oben) {
    const el = h('div', { class: 'gs-meldung ' + (mel.art === 'warn' ? 'warn' : 'ok') + (oben ? ' oben' : ''), role: 'status' }, h('span', { text: String(mel.text || '') }));
    if (mel.rueck) el.append(h('button', { type: 'button', onclick: () => o.beiRueck && o.beiRueck() }, 'Rückgängig'));
    return el;
  }

  /* Bearbeitung einer Aufgabe: erledigen, Fälligkeit, Notiz. Jede Taste meldet nur, gespeichert wird außerhalb. */
  function bearbeitung(a, m, o) {
    const zu = m.arbeit === a.id;
    const datum = h('input', { type: 'date', class: 'faellig', value: a.faellig || '', disabled: zu, 'aria-label': 'Fälligkeit' });
    const notiz = h('input', { type: 'text', class: 'notiz', maxlength: 300, placeholder: 'Notiz für den Verlauf', disabled: zu, 'aria-label': 'Notiz für den Verlauf' });
    const sende = () => { if (notiz.value.trim() && o.beiNotiz) o.beiNotiz(a, notiz.value); };
    notiz.addEventListener('keydown', (ev) => { if (ev.key === 'Enter') { ev.preventDefault(); sende(); } });
    const taste = (klasse, text, fn, aus) => h('button', { type: 'button', class: klasse, disabled: zu || aus, onclick: fn }, text);
    const el = h('div', { class: 'gs-bearb' },
      h('div', { class: 'reihe' },
        taste('erledigt', 'Erledigt', () => o.beiErledigt && o.beiErledigt(a)),
        h('span', { class: 'lab', text: 'Fällig' }), datum,
        taste('setzen', 'Setzen', () => o.beiFaellig && o.beiFaellig(a, datum.value)),
        taste('ohne', 'Ohne Termin', () => o.beiFaellig && o.beiFaellig(a, ''), !a.faellig)),
      h('div', { class: 'reihe' }, notiz, taste('anfuegen', 'Notiz anfügen', sende)));
    if (zu) el.append(h('div', { class: 'gs-meldung', role: 'status' }, h('span', { text: 'Wird gespeichert …' })));
    else if (m.meldung && m.meldung.id === a.id) el.append(meldungEl(m.meldung, o, false));
    return el;
  }

  /* Eine Aufgabe mit der Taste „Bearbeiten“. Ohne Schreibrecht bleibt es bei der Zeile. */
  function aufgabeBlock(a, m, o) {
    const zeile = aufgabeZeile(a, m);
    if (!m.bearbeiten) return zeile;
    const offen = m.offen === a.id;
    const block = h('div', { class: 'gs-aufgabe' + (offen ? ' offen' : ''), 'data-id': a.id }, zeile,
      h('button', { type: 'button', class: 'gs-stift', 'aria-expanded': String(offen), 'aria-label': (offen ? 'Bearbeitung schließen: ' : 'Aufgabe bearbeiten: ') + a.titel,
        onclick: () => o.beiBearbeiten && o.beiBearbeiten(offen ? null : a.id) }, offen ? 'Schließen' : 'Bearbeiten'));
    if (offen) block.append(bearbeitung(a, m, o));
    return block;
  }

  function terminZeile(f) {
    const Aw = A();
    return h('div', { class: 'gs-zeile zwei' + (f.ueberfaellig ? ' ueber' : f.tage === 0 ? ' heute' : '') },
      h('span', { class: 'gs-wann' }, h('b', { text: Aw.tagKurz(f.datum) }), h('i', { text: Aw.abstand(f.tage) })),
      h('span', { class: 'gs-was' }, h('span', { class: 'gs-titel', text: f.text }),
        h('span', { class: 'gs-neben', text: (f.art === 'frist' ? 'Frist' : 'Termin') + (f.betreff ? ' · aus „' + kurz(f.betreff, 60) + '“' : '') })));
  }

  function abschnitt(id, titel, anzahl) {
    const el = h('section', { class: 'gs-abschnitt', id });
    el.append(h('h2', null, h('span', { text: titel }), anzahl !== '' && anzahl != null ? h('span', { class: 'anz', text: String(anzahl) }) : null));
    for (let i = 3; i < arguments.length; i++) for (const c of [].concat(arguments[i])) if (c) el.append(c);
    return el;
  }

  function zeichnen(root, m, o) {
    o = o || {};
    const Aw = A();
    const s = m.sicht, z = s.zahlen, g = m.g || {};
    const istOhne = m.art === 'ohne';
    const oben = root.scrollTop;

    /* Kopf */
    const chips = [];
    if (!istOhne) {
      if (g.areal) chips.push(h('span', { class: 'chip', text: 'Areal ' + g.areal }));
      const lage = g.lage || {};
      if (lage.art === 'pin') chips.push(h('span', { class: 'chip warn', text: 'Lage nur als Platzhalter' + (lage.bezug ? ': ' + lage.bezug : '') }));
      else if (lage.art !== 'polygon' || !m.aufDerKarte) chips.push(h('span', { class: 'chip warn', text: 'Nicht auf der Karte verortet' }));
      else if (lage.sicher === false) chips.push(h('span', { class: 'chip warn', text: 'Lage unsicher' }));
    }
    const kopf = h('header', { class: 'gs-kopf' },
      h('button', { type: 'button', class: 'gs-zu', 'aria-label': 'Gebäudeseite schließen, zurück zur Karte', onclick: () => o.beiZu && o.beiZu() },
        h('span', { class: 'nur-schmal', text: '‹ Karte' }), h('span', { class: 'nur-breit', text: 'Schließen' })),
      h('div', { class: 'gs-nr' + (istOhne ? ' wort' : ''), text: istOhne ? 'Ohne Gebäude' : g.nr }),
      h('div', { class: 'gs-name', text: istOhne ? 'Aufgaben, die keinem Gebäude zugeordnet sind' : g.name }),
      chips.length ? h('div', { class: 'gs-chips' }, chips) : null);

    const innen = h('div', { class: 'gs-innen' });
    innen.append(filterLeiste(m.team, m.ich, m.filter, m.zaehler, (f) => o.beiFilter && o.beiFilter(f)));
    const firmenGewaehlt = m.firmen || [];
    if (s.firmen && s.firmen.length) innen.append(firmenMarken(s.firmen, firmenGewaehlt, (neu) => o.beiFirmen && o.beiFirmen(neu)));

    /* Zahlen oben: das Lagebild in drei Feldern, zugleich Sprungmarken */
    const sprung = (id) => () => { const e = root.querySelector('#' + id); if (e) e.scrollIntoView({ block: 'start', behavior: 'smooth' }); };
    const feld = (id, zahl, wort, zusatz, rot) => h('button', { type: 'button', class: 'gs-zahl', onclick: sprung(id) },
      h('b', { text: String(zahl) }), h('span', { text: wort }), h('i', { class: rot ? 'rot' : '', text: zusatz || '' }));
    const zahlen = h('nav', { class: 'gs-zahlen' + (istOhne ? ' eins' : ''), 'aria-label': 'Überblick' });
    zahlen.append(feld('gs-aufgaben', z.aufgaben, z.aufgaben === 1 ? 'Aufgabe' : 'Aufgaben',
      z.ueberfaellig ? z.ueberfaellig + ' überfällig' : z.heute ? z.heute + ' heute fällig' : '', z.ueberfaellig > 0));
    if (!istOhne) {
      zahlen.append(feld('gs-termine', m.laedt ? '…' : z.termine, z.termine === 1 ? 'Termin' : 'Termine',
        z.fristUeber ? mehrzahl(z.fristUeber, 'Frist', 'Fristen') + ' überfällig' : z.naechster ? 'nächster ' + Aw.tagKurz(z.naechster) : '', z.fristUeber > 0));
      zahlen.append(feld('gs-punkte', m.laedt ? '…' : z.punkte, z.punkte === 1 ? 'offener Punkt' : 'offene Punkte',
        z.punkteEigen ? z.punkteEigen + ' bei ' + z.eigenName : ''));
    }
    innen.append(zahlen);

    /* Aufgaben */
    const gefiltert = (m.filter && m.filter !== 'alle') || firmenGewaehlt.length > 0;
    const anzahl = gefiltert && s.aufgabenAlle !== s.aufgaben.length ? `${s.aufgaben.length} von ${s.aufgabenAlle}` : s.aufgaben.length || '';
    let aufgabenInhalt;
    if (m.fehlerAufgaben) aufgabenInhalt = h('div', { class: 'gs-leer warn', text: 'Die Aufgaben konnten nicht geladen werden: ' + m.fehlerAufgaben });
    else if (s.aufgaben.length) aufgabenInhalt = h('div', { class: 'gs-liste' }, s.aufgaben.map((a) => aufgabeBlock(a, m, o)));
    else aufgabenInhalt = h('div', { class: 'gs-leer', text: s.aufgabenAlle ? `Keine offene Aufgabe für diese Auswahl. Am Gebäude sind insgesamt ${s.aufgabenAlle} offen.` : 'Keine offene Aufgabe.' });
    // Meldung über der Liste, wenn sie zu keiner offenen Bearbeitung gehört (zum Beispiel nach „Erledigt“)
    const inZeile = m.meldung && m.meldung.id != null && m.offen === m.meldung.id && s.aufgaben.some((a) => a.id === m.meldung.id);
    const melOben = m.bearbeiten && m.meldung && !inZeile ? meldungEl(m.meldung, o, true) : null;
    innen.append(abschnitt('gs-aufgaben', 'Aufgaben', anzahl, melOben, aufgabenInhalt));

    if (!istOhne) {
      const warte = m.laedt ? h('div', { class: 'gs-leer', text: 'Wird geladen …' }) : null;
      const fehlt = !m.laedt && (m.fehler || (m.datei && m.datei.fehlt))
        ? h('div', { class: 'gs-leer' + (m.fehler ? ' warn' : ''), text: m.fehler ? 'Die Auswertung der Mails konnte nicht geladen werden: ' + m.fehler : 'Zu diesem Gebäude liegt keine Auswertung der Mails vor.' }) : null;

      /* Termine und Fristen */
      const termine = [];
      if (warte || fehlt) termine.push(warte || fehlt);
      else {
        termine.push(s.termine.length ? h('div', { class: 'gs-liste' }, s.termine.map(terminZeile)) : h('div', { class: 'gs-leer', text: 'Kein Termin und keine Frist ab heute.' }));
        if (s.vergangen.length) termine.push(h('details', { class: 'gs-klapp' }, h('summary', { text: mehrzahl(s.vergangen.length, 'vergangener Termin', 'vergangene Termine') }),
          h('div', { class: 'gs-liste blass' }, s.vergangen.map(terminZeile))));
      }
      innen.append(abschnitt('gs-termine', 'Termine und Fristen', m.laedt ? '' : s.termine.length || '', termine));

      /* Offene Punkte je Partei */
      const punkte = [];
      if (warte || fehlt) punkte.push(warte ? h('div', { class: 'gs-leer', text: 'Wird geladen …' }) : h('div', { class: 'gs-leer', text: 'Keine Angabe.' }));
      else if (!s.punkte.length) punkte.push(h('div', { class: 'gs-leer', text: 'Kein offener Punkt.' }));
      else {
        for (const gr of s.punkte) {
          punkte.push(h('div', { class: 'gs-gruppe' + (gr.eigen ? ' eigen' : '') },
            h('h3', null, h('span', { text: 'Bei ' + gr.bei }), h('span', { class: 'anz', text: String(gr.eintraege.length) })),
            h('div', { class: 'gs-liste' }, gr.eintraege.map((p) => h('div', { class: 'gs-punkt' },
              h('span', { class: 'gs-titel', text: p.text }),
              h('span', { class: 'gs-neben', text: (p.seit ? 'offen seit ' + Aw.tagOhneWt(p.seit) : p.eingang ? 'Mail vom ' + Aw.tagOhneWt(p.eingang) : '') + (p.betreff ? ' · ' + kurz(p.betreff, 56) : '') }))))));
        }
      }
      innen.append(abschnitt('gs-punkte', 'Offene Punkte', m.laedt ? '' : s.punkteAnzahl || '', punkte));

      /* Verlauf, zugeklappt */
      if (!m.laedt && s.verlauf.length) {
        innen.append(h('details', { class: 'gs-klapp gs-verlauf', id: 'gs-verlauf' },
          h('summary', null, h('span', { text: 'Verlauf' }), h('span', { class: 'anz', text: mehrzahl(s.verlauf.length, 'Mail', 'Mails') })),
          h('div', { class: 'gs-liste' }, s.verlauf.map((x) => {
            const haupt = x.haupt || {};
            const neben = h('span', { class: 'gs-neben' }, [x.von_name, x.partei].filter(Boolean).join(' · '));
            if (haupt.art) neben.append(' · ', h('span', { class: 'ampel ' + (['gruen', 'gelb', 'rot'].includes(haupt.ampel) ? haupt.ampel : 'rot'), title: String(haupt.grund || '') }), ARTEN[haupt.art] || String(haupt.art));
            if (Number(x.neuere_im_verlauf) > 0) neben.append(' · neuere Nachricht im Verlauf');
            return h('div', { class: 'gs-zeile zwei' },
              h('span', { class: 'gs-wann' }, h('b', { text: Aw.tagOhneWt(x.eingang) }), h('i', { text: Aw.uhrzeit(x.eingang) })),
              h('span', { class: 'gs-was' }, h('span', { class: 'gs-titel', text: String(x.thema || x.betreff || '') }),
                x.zusammenfassung ? h('span', { class: 'gs-text', text: String(x.zusammenfassung) }) : null, neben));
          }))));
      }
    }

    /* Fuß: Herkunft und Stand, ehrlich benannt */
    const fuss = h('footer', { class: 'gs-fuss' });
    if (m.fussnote) fuss.append(h('p', { class: 'warn', text: m.fussnote }));
    if (!istOhne) {
      fuss.append(h('p', { text: 'Aufgaben kommen aus den FP-Aufgaben. Termine, offene Punkte und Verlauf hat der Helfer aus den Mails gelesen. Von Hand geprüft sind sie nicht.' + (m.datei && m.datei.stand ? ' Stand der Auswertung ' + Aw.standText(m.datei.stand) + '.' : '') }));
      const daten = [g.planer ? 'Planung ' + g.planer : '', g.status].filter(Boolean).join(' · ');
      if (daten) fuss.append(h('p', { text: daten }));
    } else {
      fuss.append(h('p', { text: 'Das Gebäude ergibt sich aus der Spalte „Gebäude“ der Aufgabe, solange sie leer ist aus dem Namen der Mappe. Hier steht, was so keinem Gebäude zufällt.' }));
    }
    innen.append(fuss);

    root.replaceChildren(kopf, innen);
    root.scrollTop = o.behalten ? oben : 0;
  }

  wurzel.FPGebaeudeseite = { zeichnen, filterLeiste, firmenWahl, h };
})(typeof window !== 'undefined' ? window : this);
