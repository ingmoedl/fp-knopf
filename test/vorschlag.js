/*
  FP-Knopf – Reiter „Vorschlag“
  ------------------------------
  Zeigt zu der geöffneten Mail: Einordnung (Gebäude, Partei), einen Hauptvorschlag und höchstens drei weitere.
  Jede Angabe trägt eine Ampel mit Begründung. Bei Rot wird nichts mit einem Klick ausgeführt.

  Zwei Quellen:
    1. sofort die Einordnung nach festen Regeln (einordnung.js)
    2. sobald vorhanden die Bewertung des Helfers aus dem Speicher FP-Wissen

  Der Reiter kennt weder Outlook noch SharePoint. Alles, was nach außen geht, läuft über „dienste“,
  die taskpane.js stellt. So lässt sich die Anzeige im Browser mit Beispieldaten prüfen
  (werkzeuge/test_reiter.html).

  Inhalte aus Mails gelten als fremd: Sie werden nur als Text eingesetzt, nie als HTML.
*/
(function (wurzel) {
  'use strict';

  const AMPEL_WORT = { gruen: 'sicher', gelb: 'prüfen', rot: 'geraten' };
  const RANG = { gruen: 0, gelb: 1, rot: 2 };
  const besser = (a, b) => (RANG[a] <= RANG[b] ? a : b);
  const schlechter = (a, b) => (RANG[a] >= RANG[b] ? a : b);

  // Baut ein Element. Texte kommen über textContent hinein.
  function h(tag, attrs, ...kinder) {
    const e = document.createElement(tag);
    for (const k of Object.keys(attrs || {})) {
      if (k === 'class') e.className = attrs[k];
      else if (k === 'onclick') e.onclick = attrs[k];
      else if (attrs[k] !== undefined && attrs[k] !== null && attrs[k] !== false) e.setAttribute(k, attrs[k]);
    }
    for (const kind of kinder.flat()) {
      if (kind === null || kind === undefined || kind === false) continue;
      e.appendChild(typeof kind === 'string' ? document.createTextNode(kind) : kind);
    }
    return e;
  }
  const ampel = (a) => h('span', { class: 'vs-ampel vs-' + a, title: AMPEL_WORT[a] || '' });
  const deDatum = (iso) => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ''); return m ? m[3] + '.' + m[2] + '.' : (iso || ''); };
  const ohneKennung = (mappe) => String(mappe || '').replace(/\s*\[[0-9a-f]{8}\]\s*$/i, '');
  // "robert.kalusok@…" -> "Robert Kalusok"
  const kurzAdresse = (a) => String(a || '').split('@')[0].split(/[._]+/).filter(Boolean).map((w) => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
  const modellKurz = (m) => { const t = /(haiku|sonnet|opus|fable)/i.exec(m || ''); return t ? t[1].charAt(0).toUpperCase() + t[1].slice(1).toLowerCase() : (m || ''); };

  /*
    Führt Regeln und Bewertung zu einem Anzeigemodell zusammen.
    regeln:      Ergebnis von FPEinordnung.einordnen
    gespeichert: Datei des Helfers ({ vorschlag, bewertet, pruefung, neuere_im_verlauf }) oder null
    reg:         Register, für die Gebäudenamen
  */
  function zusammenfuehren(regeln, gespeichert, reg) {
    const namen = {};
    for (const g of (reg && reg.gebaeude) || []) namen[String(g.nr).toLowerCase()] = g.name;
    const v = gespeichert && gespeichert.vorschlag;

    // Gebäude: beide Quellen vereinen. Nennen beide dasselbe Gebäude, gilt die bessere Ampel.
    const geb = {};
    for (const g of (v && v.gebaeude) || []) geb[String(g.nr).toLowerCase()] = { nr: g.nr, name: namen[String(g.nr).toLowerCase()] || '', ampel: g.ampel, grund: g.grund };
    for (const g of regeln.gebaeude || []) {
      const k = String(g.nr).toLowerCase();
      if (!geb[k]) geb[k] = { nr: g.nr, name: g.name, ampel: g.ampel, grund: g.grund };
      else if (RANG[g.ampel] < RANG[geb[k].ampel]) { geb[k].ampel = g.ampel; geb[k].grund = g.grund; }
    }
    const gebaeude = Object.values(geb).sort((a, b) => RANG[a.ampel] - RANG[b.ampel] || String(a.nr).localeCompare(String(b.nr)));

    let bereich = null;
    if (!gebaeude.length) {
      if (regeln.bereich) bereich = regeln.bereich;
      else if (v && v.bereich) bereich = { name: v.bereich, ampel: 'gelb', grund: 'Kein Gebäude genannt, Bereich laut Bewertung.' };
    }

    // Partei: Die Absenderdomain ist der härtere Beleg. Sonst gilt die Bewertung.
    let partei = regeln.partei;
    if (partei.ampel !== 'gruen' && v && v.partei && v.partei.ordner_kom) partei = { ordner_kom: v.partei.ordner_kom, ordner_ext: '', ampel: v.partei.ampel, grund: v.partei.grund };

    const ausRegeln = (art) => (regeln.vorschlaege || []).find((s) => s.art === art);
    let liste;
    if (v) {
      liste = (v.vorschlaege || []).map((s) => Object.assign({}, s));
      // Die Regeln kennen die Aufgabenmappen und die Ordner für Unterlagen, die Bewertung nicht.
      const zu = ausRegeln('aufgabe_zuordnen');
      if (zu) {
        const neu = liste.findIndex((s) => s.art === 'aufgabe_neu');
        if (neu >= 0 && zu.ampel === 'gruen') liste[neu] = Object.assign({}, zu, { haupt: liste[neu].haupt });   // Aufgabe gibt es schon
        else liste.push(Object.assign({ haupt: false }, zu));
      }
      const anhRegel = ausRegeln('anhang_ablegen');
      for (const s of liste) {
        if (s.art === 'anhang_ablegen') {
          s.ziele = anhRegel ? anhRegel.ziele : [];
          s.ampel = anhRegel ? schlechter(s.ampel, anhRegel.ampel) : 'rot';
          s.grund = anhRegel ? anhRegel.grund : 'Kein passender Ordner für Unterlagen gefunden.';
        }
        if (s.art === 'ablegen' && partei.ordner_kom && partei.ampel === 'gruen') { s.ziel_kom = partei.ordner_kom; s.ampel = 'gruen'; }
      }
    } else {
      liste = (regeln.vorschlaege || []).map((s) => Object.assign({ haupt: false }, s));
      const rang = ['pruefaufgabe', 'aufgabe_zuordnen', 'ablegen', 'anhang_ablegen'];
      liste.sort((a, b) => rang.indexOf(a.art) - rang.indexOf(b.art));
    }
    if (liste.length && !liste.some((s) => s.haupt)) liste[0].haupt = true;
    const haupt = liste.find((s) => s.haupt) || null;
    let weitere = liste.filter((s) => s !== haupt);
    if (weitere.length > 3) {                               // Ablegen bleibt immer sichtbar
      const fest = weitere.filter((s) => s.art === 'ablegen');
      weitere = weitere.filter((s) => s.art !== 'ablegen').slice(0, 3 - fest.length).concat(fest);
    }

    return {
      quelle: v ? 'claude' : 'regeln',
      gebaeude, bereich, partei,
      unbekannteNummern: [...new Set([...(regeln.unbekannteNummern || []), ...((v && v.unbekannte_nummern) || [])])],
      zusammenfassung: (v && v.zusammenfassung) || '',
      haupt, weitere,
      neuere: (gespeichert && gespeichert.neuere_im_verlauf) || [],
      vermerke: (gespeichert && gespeichert.pruefung) || [],
      bewertet: (gespeichert && gespeichert.bewertet) || null,
      offen: v ? [] : (regeln.offen || []),
    };
  }

  // Überschrift und Zusatzzeile eines Vorschlags
  // ablage: { kom, ext } aus dem Register, also die Namen der beiden Ablageordner
  function beschriftung(s, ablage) {
    const kom = (ablage && ablage.kom) ? ablage.kom + ' / ' : '';
    const ext = (ablage && ablage.ext) ? ablage.ext + ' / ' : '';
    switch (s.art) {
      case 'antworten': return ['Antworten', s.an && s.an.length ? 'Entwurf an ' + s.an.map(kurzAdresse).join(', ') : 'Entwurf an den Absender'];
      case 'entscheidung': return ['Zur Entscheidung an den Bauherrn', s.an && s.an.length ? 'Entwurf an ' + s.an.map(kurzAdresse).join(', ') : 'Empfänger fehlt'];
      case 'weiterleiten': return ['Weiterleiten', s.an && s.an.length ? 'Entwurf an ' + s.an.map(kurzAdresse).join(', ') : 'Empfänger fehlt'];
      case 'zuweisen': return ['Intern zuweisen' + (s.kuerzel ? ': ' + s.kuerzel : ''), ''];
      case 'aufgabe_neu': return ['Aufgabe anlegen', (s.titel || '') + (s.frist ? ', fällig ' + deDatum(s.frist) : '')];
      case 'aufgabe_zuordnen': return ['Zu Aufgabe „' + ohneKennung(s.mappe) + '“', ''];
      case 'ablegen': return [s.ziel_kom ? 'Ablegen in ' + kom + s.ziel_kom : 'Ablegen: Ordner wählen', ''];
      case 'anhang_ablegen': return [s.ziele && s.ziele.length ? 'Anhang nach ' + ext + s.ziele[0] : 'Anhang ablegen: Ordner wählen', ''];
      case 'pruefaufgabe': return [(s.name ? s.name + '-Prüfung' : 'Planprüfung') + (s.frist ? ' bis ' + deDatum(s.frist) : ''), 'Aufgabe mit Frist anlegen'];
      case 'kenntnis': return ['Nur zur Kenntnis', 'Keine Handlung nötig'];
      default: return [s.art, ''];
    }
  }

  function erzeuge(container, dienste) {
    let lauf = 0, reg = null, letzteMail = null, trotzdem = null;

    const meldung = (text, art) => h('div', { class: 'vs-meldung' + (art ? ' vs-' + art : '') }, text);

    async function ausfuehren(s, knopf, zeile, wahl) {
      const alt = knopf.textContent;
      knopf.disabled = true; knopf.textContent = '…';
      try {
        const text = await dienste.aktion(s, wahl || {});
        zeile.querySelectorAll('.vs-ergebnis').forEach((x) => x.remove());
        zeile.appendChild(h('div', { class: 'vs-ergebnis vs-ok' }, text || 'Erledigt.'));
        knopf.textContent = alt; knopf.disabled = false;
      } catch (e) {
        zeile.querySelectorAll('.vs-ergebnis').forEach((x) => x.remove());
        zeile.appendChild(h('div', { class: 'vs-ergebnis vs-fehler' }, 'Nicht ausgeführt: ' + (e && e.message ? e.message : e)));
        knopf.textContent = alt; knopf.disabled = false;
      }
    }

    // Knopf oder Auswahl je Art. Bei Rot und bei fehlendem Ziel wählt der Mensch selbst.
    function bedienung(s, zeile) {
      const entwurf = s.art === 'antworten' || s.art === 'entscheidung' || s.art === 'weiterleiten';
      if (s.art === 'kenntnis') return null;
      if (s.art === 'ablegen' || s.art === 'anhang_ablegen' || s.art === 'zuweisen') {
        const ziel = s.art === 'ablegen' ? s.ziel_kom : s.art === 'anhang_ablegen' ? (s.ziele || [])[0] : s.kuerzel;
        const werte = s.art === 'ablegen' ? (reg.parteien || []).map((p) => p.ordner_kom)
          : s.art === 'anhang_ablegen' ? (reg.ordnerExt || []) : (reg.kuerzel || []);
        const wort = s.art === 'zuweisen' ? 'Zuweisen' : 'Ablegen';
        if (s.ampel !== 'rot' && ziel) {
          const k = h('button', { class: 'vs-knopf' }, wort);
          k.onclick = () => ausfuehren(s, k, zeile, { ziel });
          return k;
        }
        const sel = h('select', { class: 'vs-wahl', 'aria-label': 'Ziel wählen' }, h('option', { value: '' }, 'wählen …'), werte.map((w) => h('option', { value: w }, w)));
        if (ziel) sel.value = ziel;
        const k = h('button', { class: 'vs-knopf' }, wort);
        k.onclick = () => {
          if (!sel.value) { zeile.querySelectorAll('.vs-ergebnis').forEach((x) => x.remove()); zeile.appendChild(h('div', { class: 'vs-ergebnis vs-fehler' }, 'Bitte zuerst ein Ziel wählen.')); return; }
          ausfuehren(s, k, zeile, { ziel: sel.value, selbstGewaehlt: true });
        };
        return h('div', { class: 'vs-wahlzeile' }, sel, k);
      }
      const wort = entwurf ? (s.ampel === 'rot' ? 'Entwurf prüfen' : 'Entwurf öffnen') : s.art === 'aufgabe_zuordnen' ? 'Zuordnen' : 'Vorbelegen';
      const k = h('button', { class: 'vs-knopf' }, wort);
      k.onclick = () => ausfuehren(s, k, zeile, {});
      return k;
    }

    function vorschlagZeile(s, istHaupt) {
      const [titel, zusatz] = beschriftung(s, reg && reg.ablage);
      const zeile = h('div', { class: istHaupt ? 'vs-haupt' : 'vs-zeile' });
      if (istHaupt) zeile.appendChild(h('div', { class: 'vs-label' }, 'Hauptvorschlag'));
      zeile.appendChild(h('div', { class: 'vs-titel' }, titel));
      if (zusatz) zeile.appendChild(h('div', { class: 'vs-zusatz' }, zusatz));
      zeile.appendChild(h('div', { class: 'vs-grund' }, ampel(s.ampel), (AMPEL_WORT[s.ampel] || '') + '. ' + (s.grund || '')));
      if (s.entwurf) zeile.appendChild(h('details', { class: 'vs-entwurf' }, h('summary', {}, 'Entwurf ansehen'), h('pre', {}, s.entwurf)));
      const b = bedienung(s, zeile);
      if (b) zeile.appendChild(h('div', { class: 'vs-bedienung' }, b));
      return zeile;
    }

    function male(m) {
      container.textContent = '';
      if (m.quelle === 'regeln') container.appendChild(meldung('Erste Einordnung nach festen Regeln. Die Bewertung folgt, sobald der Helfer die Mail gelesen hat.'));
      for (const n of m.neuere) {
        const wer = n.art === 'gesendet' ? 'aus dem eigenen Postfach gesendet' : 'von ' + n.von;
        container.appendChild(meldung('Im selben Verlauf gibt es eine neuere Nachricht: ' + deDatum(n.zeit) + ' ' + String(n.zeit).slice(11, 16) + ', ' + wer + '.', 'hinweis'));
      }

      container.appendChild(h('div', { class: 'vs-label' }, 'Einordnung'));
      const chips = h('div', { class: 'vs-chips' });
      const gruende = h('div', { class: 'vs-gruende' });
      const chip = (a, text, grund) => {
        const c = h('button', { class: 'vs-chip', title: grund }, ampel(a), text);
        c.onclick = () => { gruende.textContent = (AMPEL_WORT[a] || '') + '. ' + grund; };
        return c;
      };
      for (const g of m.gebaeude) chips.appendChild(chip(g.ampel, 'Geb. ' + g.nr + (g.name ? ' ' + g.name : ''), g.grund));
      if (!m.gebaeude.length && m.bereich) chips.appendChild(chip(m.bereich.ampel, m.bereich.name, m.bereich.grund));
      if (!m.gebaeude.length && !m.bereich) chips.appendChild(chip('rot', 'kein Gebäude erkannt', 'Weder Nummer noch Name noch Bereich gefunden.'));
      chips.appendChild(chip(m.partei.ampel, m.partei.ordner_kom || 'Partei unbekannt', m.partei.grund));
      container.appendChild(chips);
      container.appendChild(gruende);
      if (m.unbekannteNummern.length) container.appendChild(meldung('Genannt, aber nicht im Register: ' + m.unbekannteNummern.join(', ') + '.', 'hinweis'));
      if (m.zusammenfassung) container.appendChild(h('div', { class: 'vs-zusammenfassung' }, m.zusammenfassung));

      if (m.haupt) container.appendChild(vorschlagZeile(m.haupt, true));
      if (m.weitere.length) {
        container.appendChild(h('div', { class: 'vs-label' }, 'Außerdem'));
        for (const s of m.weitere) container.appendChild(vorschlagZeile(s, false));
      }
      for (const t of m.vermerke) container.appendChild(meldung('Prüfung: ' + t, 'hinweis'));
      if (m.offen.length) container.appendChild(h('div', { class: 'vs-offen' }, 'Das sagen die Regeln nicht: ' + m.offen.join(', ') + '.'));

      const fuss = h('div', { class: 'vs-fuss' });
      if (m.bewertet) {
        fuss.appendChild(h('span', {}, String(m.bewertet.wann || '').slice(11, 16) + ' · ' + modellKurz(m.bewertet.modell) + ' · Trefferquote: noch keine Messung'));
        const neu = h('button', { class: 'vs-link' }, 'Neu bewerten');
        neu.onclick = () => ausfuehren({ art: 'neu_bewerten' }, neu, fuss, {});
        fuss.appendChild(neu);
      } else fuss.appendChild(h('span', {}, 'Nur Regeln, ohne Claude'));
      container.appendChild(fuss);
    }

    /*
      mail: { schluessel, von, vonName, betreff, text, eingang (Date), anhaenge: [name], eigeneDomain, postfach }
      Zeigt zuerst die Regeln, dann die Bewertung, falls sie im Speicher liegt.
    */
    async function zeige(mail) {
      const hier = ++lauf;
      letzteMail = mail;
      container.textContent = '';
      if (!mail) { container.appendChild(meldung('Keine Mail ausgewählt.')); return; }
      try {
        if (!reg) reg = await dienste.ladeRegister();
      } catch (e) {
        if (hier !== lauf) return;
        container.appendChild(meldung('Die Register lassen sich nicht laden. Der Speicher FP-Wissen fehlt oder ist nicht erreichbar. (' + (e && e.message ? e.message : e) + ')', 'fehler'));
        return;
      }
      if (hier !== lauf) return;
      // Vorschläge gelten nur für das Postfach aus dem Register. Fehlt eine der beiden Angaben, wird nicht gesperrt.
      // Die Meldung nennt das erkannte Postfach, und die Sperre lässt sich je Mail aufheben.
      const soll = String(reg.postfach || '').toLowerCase();
      const ist = String(mail.postfach || '').toLowerCase();
      if (soll && ist && soll !== ist && trotzdem !== mail.schluessel) {
        container.appendChild(meldung('Diese Mail liegt im Postfach ' + ist + '. Vorschläge gibt es nur für ' + soll + '.', 'hinweis'));
        const k = h('button', { class: 'vs-knopf' }, 'Trotzdem einordnen');
        k.onclick = () => { trotzdem = mail.schluessel; zeige(mail); };
        container.appendChild(h('div', { class: 'vs-bedienung' }, k));
        return;
      }
      const regeln = wurzel.FPEinordnung.einordnen(mail, reg, (dienste.mappen && dienste.mappen()) || []);
      male(zusammenfuehren(regeln, null, reg));
      let gespeichert = null;
      try { gespeichert = await dienste.ladeVorschlag(mail); } catch (e) { gespeichert = null; }
      if (hier !== lauf || !gespeichert) return;
      male(zusammenfuehren(regeln, gespeichert, reg));
    }

    return { zeige, neuLaden: () => zeige(letzteMail) };
  }

  wurzel.FPVorschlag = { erzeuge, zusammenfuehren, beschriftung };
})(window);
