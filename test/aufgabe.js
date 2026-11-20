/*
  FP-Karte – Aufgabenseite
  ------------------------
  Zeigt eine Aufgabe ganz: Stand, Firma, Zuweisung, Fälligkeit und darunter den Verlauf. Im Verlauf stehen Notizen
  und Mails zusammen: Eine Notiz trägt die Mail, zu der sie gehört, direkt unter sich. So wie bisher von Hand in den
  Notizen der Outlook-Aufgabe: ein Satz, darunter die Mail.

  Woher die Zuordnung kommt. Das Format ist das des Entwurfs der Übersicht (knopf/uebersicht.js im privaten Repo),
  damit beide später zusammenpassen. Die veröffentlichte Übersicht kennt es noch nicht: Sie zeigt den Verlauf als
  reinen Text, die Zeile mit dem Dateinamen steht dort lesbar unter der Notiz.
    Der Verlauf ist der Text der Spalte FPVerlauf, neuester Eintrag oben. Ein Eintrag beginnt mit einer Zeile
    „TT.MM.“ oder „TT.MM.JJJJ“, alle Zeilen bis zum nächsten Eintrag gehören dazu.
    Eine Zeile „✉ Dateiname“ ordnet den Eintrag einer Mail in der Mappe zu.
    Ohne diese Zeile gilt „Mail „Betreff““ im Text als Zuordnung. So schreibt es der FP-Knopf beim Anlegen.
  Mails, zu denen es keine Notiz gibt, stehen an ihrem Datum allein im Verlauf.

  Kurzfassung einer Mail: aus der Bewertung des Helfers (gebaeude/_mails.json im Speicher FP-Wissen). Zugeordnet wird
  über Eingang (auf drei Minuten genau) und Betreff. Mails, die der Helfer nie gesehen hat, haben keine Kurzfassung.

  „In Outlook öffnen“: Link fp-mail:<Pfad, base64url>. Den richtet der FP-Abgleich oder der PC-Auftrag 002 am PC ein.

  Fremde Inhalte (Betreff, Notiz, Kurzfassung) werden nur als Text eingesetzt, nie als HTML.

  FPKarteAufgabe.zeichnen(root, modell, rueckrufe)
    modell     { a (Aufgabe), team, heute, stand { status, von, geaendert, erledigtAm }, laedt, fehler,
                 verlauf (Text), dateien [{ Name, ServerRelativeUrl, Length }], mails (Bewertungen des Helfers),
                 schreiben, arbeit, meldung { text, art, rueck }, zurueck (Wort für die Taste zurück), outlook, firmen,
                 host (für Verweise auf Dateien), notizMail (vorgewählte Mail im Feld für den neuen Eintrag) }
    rueckrufe  { beiZurueck(), beiZu(), beiErledigt(a), beiOeffnen(a), beiFaellig(a, iso), beiFirma(a, text),
                 beiZuweisung(a, text), beiNotiz(a, text, datei), beiLoeschen(a, { roh, gleiche, mitMail }), beiRueck(),
                 beiOutlook(rel), beiNotizMail(name) }
*/
(function (wurzel) {
  'use strict';

  const z2 = (n) => String(n).padStart(2, '0');
  const MAILDATEI = /\.(msg|eml)$/i;
  const EINTRAG = /^\s*\d{1,2}\.\d{1,2}\.(?:\d{4}|\d{2})?(?=[\s:,;–-]|$)/;
  const MAILZEILE = /^\s*✉\s*(\S.*?)\s*$/;
  const KOPF = /^(\s*(\d{1,2}\.\d{1,2}\.(?:\d{4}|\d{2})?)\s*(?:([A-ZÄÖÜ]{2,4}):\s*)?)([\s\S]*)$/;
  const woerter = (s) => String(s || '').split(/[^A-Za-zÄÖÜäöüß]+/).filter(Boolean);
  const kern = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9äöüß]+/g, ' ').trim();
  const gleicherBetreff = (a, b) => !!a && !!b && (a === b || (Math.min(a.length, b.length) >= 10 && (a.startsWith(b) || b.startsWith(a))));

  /* Teilt den Verlauf in Blöcke von Zeilen, je Eintrag einer. Jede Zeile bleibt, wie sie ist: Mit Zeilenumbrüchen
     wieder zusammengesetzt ergeben die Blöcke genau den Text. Das Löschen in bearbeiten.js teilt mit derselben Regel. */
  function bloecke(text) {
    const aus = [];
    String(text || '').replace(/\r\n?/g, '\n').split('\n').forEach((z) => {
      if (!aus.length || EINTRAG.test(z)) aus.push([]);
      aus[aus.length - 1].push(z);
    });
    return aus;
  }

  /* Zerlegt den Verlauf in Einträge. roh ist der Eintrag, wie er im Verlauf steht, mit allen seinen Zeilen. */
  function eintraege(text) {
    return bloecke(text).map((zeilen, i) => {
      const inhalt = zeilen.filter((z) => !MAILZEILE.test(z));
      while (inhalt.length && !inhalt[inhalt.length - 1].trim()) inhalt.pop();
      const m = zeilen.map((z) => MAILZEILE.exec(z)).find(Boolean);
      const t = inhalt.join('\n'), k = KOPF.exec(t);
      return { i, zeilen, roh: zeilen.join('\n'), text: t, datei: m ? m[1] : '', datum: k ? k[2] : '', kz: k ? k[3] || '' : '', rest: k ? k[4] : t };
    }).filter((e) => e.text.trim() || e.datei);
  }

  /* „07.10.“ ist der letzte vergangene 7.10., mit Jahr gilt das Jahr. heute als JJJJ-MM-TT. */
  function eintragTag(datum, heute) {
    const m = /^(\d{1,2})\.(\d{1,2})\.(\d{2,4})?$/.exec(datum || '');
    if (!m) return '';
    const j = m[3] ? (m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3])) : Number(String(heute).slice(0, 4));
    const t = `${j}-${z2(m[2])}-${z2(m[1])}`;
    return !m[3] && t > heute ? `${j - 1}-${z2(m[2])}-${z2(m[1])}` : t;
  }

  /* Datei in der Mappe. Mails heißen „JJJJ-MM-TT_HHMM Betreff.eml“ oder „.msg“. */
  function dateiInfo(name, rel, laenge) {
    name = String(name || '');
    const m = /^(\d{4})-(\d{2})-(\d{2})_(\d{2})(\d{2}) (.*)$/.exec(name);
    const mail = MAILDATEI.test(name) && !name.startsWith('_');
    return {
      name, rel: String(rel || ''), groesse: Number(laenge) || 0, mail,
      tag: m ? `${m[1]}-${m[2]}-${m[3]}` : '', uhr: m ? `${m[4]}:${m[5]}` : '',
      zeit: m ? `${m[3]}.${m[2]}.${m[1].slice(2)} ${m[4]}:${m[5]}` : '',
      betreff: (m ? m[6] : name).replace(MAILDATEI, ''),
    };
  }

  /*
    Verlauf und Mails zu einer Folge von Posten, neueste oben.
      { art: 'notiz', tag, datum, kz, text, mails: [Datei], fehlt: Dateiname, der nicht mehr in der Mappe liegt,
        roh: der Eintrag wie im Verlauf, gleiche: Zahl gleichlautender Einträge darüber (beides für das Löschen) }
      { art: 'mail', tag, mail: Datei }
    dateien: was sonst in der Mappe liegt.
  */
  function zeitstrahl(verlauf, dateien, heute) {
    const infos = (dateien || []).map((d) => dateiInfo(d.Name, d.ServerRelativeUrl, d.Length));
    const mails = infos.filter((d) => d.mail);
    const vergeben = new Set();
    let vorher = '';
    const gesehen = new Map();
    const notizen = eintraege(verlauf).map((e, nr) => {
      const gleiche = gesehen.get(e.roh) || 0;
      gesehen.set(e.roh, gleiche + 1);
      const zu = [];
      if (e.datei) {
        const m = mails.find((x) => x.name.toLowerCase() === e.datei.toLowerCase());
        if (m) zu.push(m);
      } else {
        const q = /Mail\s+[„"]([^“”"]+)[“”"]/.exec(e.text);
        const k = q ? kern(q[1]) : '';
        const m = k ? mails.find((x) => !vergeben.has(x.name) && gleicherBetreff(k, kern(x.betreff))) : null;
        if (m) zu.push(m);
      }
      zu.forEach((m) => vergeben.add(m.name));
      const tag = eintragTag(e.datum, heute) || vorher;       // Text ohne Datum bleibt bei dem Eintrag davor
      vorher = tag;
      return { art: 'notiz', tag, datum: e.datum, kz: e.kz, text: e.rest.trim(), mails: zu, fehlt: e.datei && !zu.length ? e.datei : '', nr, roh: e.roh, gleiche };
    });
    const allein = mails.filter((m) => !vergeben.has(m.name))
      .sort((a, b) => (a.tag + a.uhr < b.tag + b.uhr ? 1 : -1))
      .map((m, nr) => ({ art: 'mail', tag: m.tag, mail: m, nr }));
    // Neueste oben. Am selben Tag erst die Notizen in ihrer Reihenfolge, dann die Mails ohne Notiz.
    const posten = notizen.concat(allein).map((p, i) => ({ p, i }))
      .sort((x, y) => (x.p.tag === y.p.tag ? x.i - y.i : x.p.tag < y.p.tag ? 1 : -1)).map((x) => x.p);
    return { posten, dateien: infos.filter((d) => !d.mail), mails };
  }

  /* Kurzfassung einer Mail aus den Bewertungen des Helfers: gleicher Eingang auf drei Minuten, gleicher Betreff. */
  function kurzFinden(info, bewertungen) {
    if (!info || !info.tag || !info.uhr) return null;
    const t = new Date(`${info.tag}T${info.uhr}:00`).getTime(), k = kern(info.betreff);
    let beste = null, abstand = Infinity;
    for (const b of bewertungen || []) {
      if (!b || !b.eingang || !b.kurz) continue;
      const d = Math.abs(new Date(b.eingang).getTime() - t);
      if (!(d <= 3 * 60000) || d >= abstand || !gleicherBetreff(k, kern(b.betreff))) continue;
      beste = b; abstand = d;
    }
    return beste ? { kurz: String(beste.kurz), art: String(beste.art || ''), von: String(beste.von || '') } : null;
  }

  /* Link, den der FP-Abgleich am PC kennt: holt die Mail und öffnet sie in Outlook. */
  function outlookLink(rel) {
    const b = new TextEncoder().encode(String(rel));
    let s = '';
    b.forEach((x) => { s += String.fromCharCode(x); });
    return 'fp-mail:' + btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }

  /* Taste mit einem Kürzel an der Zuweisung: schaltet es ein oder aus. Steht dort ein fremder Name, ersetzt das Kürzel ihn. */
  function zuweisungUmschalten(text, kz, team) {
    const w = woerter(text), eigene = (team || []).map((b) => b.kuerzel);
    const nurEigene = w.length > 0 && w.every((t) => eigene.includes(t));
    return w.includes(kz) ? w.filter((t) => t !== kz).join('/') : nurEigene ? w.concat(kz).join('/') : kz;
  }

  /* Schnelltasten für die Fälligkeit, ab heute gerechnet */
  function schnell(heute) {
    const plus = (n) => { const d = new Date(heute + 'T12:00:00'); d.setDate(d.getDate() + n); return `${d.getFullYear()}-${z2(d.getMonth() + 1)}-${z2(d.getDate())}`; };
    return [['heute', plus(0)], ['morgen', plus(1)], ['+3 Tage', plus(3)], ['+1 Woche', plus(7)], ['+2 Wochen', plus(14)], ['ohne', '']];
  }

  /* ---------- Anzeige ---------- */
  const deLang = (iso) => (iso ? `${iso.slice(8, 10)}.${iso.slice(5, 7)}.${iso.slice(0, 4)}` : '');
  const deKurz = (iso) => (iso ? `${iso.slice(8, 10)}.${iso.slice(5, 7)}.` : '');
  const groesse = (n) => (n >= 1048576 ? (n / 1048576).toFixed(1).replace('.', ',') + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB');

  function wannGeaendert(iso, heute) {
    const d = new Date(iso);
    if (!iso || isNaN(d)) return '';
    const tag = `${d.getFullYear()}-${z2(d.getMonth() + 1)}-${z2(d.getDate())}`;
    return (tag === heute ? 'heute' : deKurz(tag)) + ' ' + z2(d.getHours()) + ':' + z2(d.getMinutes());
  }

  function zeichnen(root, m, o) {
    o = o || {};
    const h = wurzel.FPGebaeudeseite.h;
    const a = m.a, st = m.stand || {};
    const zu = !!m.arbeit, darf = !!m.schreiben && !zu;
    const erledigt = st.status === 'Erledigt';
    const oben = root.scrollTop;
    const dateiUrl = (d) => (m.host && d.rel ? m.host + d.rel.split('/').map(encodeURIComponent).join('/') : '');

    /* Kopf */
    const zustand = erledigt ? 'Erledigt' + (st.erledigtAm ? ' am ' + deLang(st.erledigtAm) : '')
      : a.faellig && a.faellig < m.heute ? 'Überfällig seit ' + deLang(a.faellig) : 'Offen';
    const info = h('p', { class: 'as-info' }, h('b', { class: erledigt ? 'gut' : a.faellig && a.faellig < m.heute ? 'rot' : '', text: zustand }));
    if (st.geaendert) info.append(' · zuletzt geändert' + (st.von ? ' von ' + st.von : '') + ', ' + wannGeaendert(st.geaendert, m.heute));
    const kopf = h('header', { class: 'gs-kopf as-kopf' },
      h('button', { type: 'button', class: 'gs-zu', 'aria-label': 'Aufgabe schließen, zurück zur Karte', onclick: () => o.beiZu && o.beiZu() },
        h('span', { class: 'nur-schmal', text: '‹ Karte' }), h('span', { class: 'nur-breit', text: 'Schließen' })),
      h('button', { type: 'button', class: 'as-zurueck', onclick: () => o.beiZurueck && o.beiZurueck() }, '‹ ' + (m.zurueck || 'Zurück')),
      h('h1', { class: 'as-titel', text: a.titel }), info);

    const innen = h('div', { class: 'gs-innen as-innen' });
    if (m.meldung) {
      const mel = h('div', { class: 'gs-meldung oben ' + (m.meldung.art === 'warn' ? 'warn' : 'ok'), role: 'status' }, h('span', { text: String(m.meldung.text || '') }));
      if (m.meldung.rueck) mel.append(h('button', { type: 'button', onclick: () => o.beiRueck && o.beiRueck() }, 'Rückgängig'));
      innen.append(mel);
    } else if (zu) innen.append(h('div', { class: 'gs-meldung oben', role: 'status' }, h('span', { text: 'Wird gespeichert …' })));

    /* Erledigen und Verweis auf die Mappe */
    const aktionen = h('div', { class: 'as-reihe' });
    if (m.schreiben) {
      aktionen.append(h('button', { type: 'button', class: 'as-taste erledigt', disabled: zu || !m.stand,
        onclick: () => (erledigt ? o.beiOeffnen && o.beiOeffnen(a) : o.beiErledigt && o.beiErledigt(a)) }, erledigt ? 'Wieder öffnen' : '✓ Erledigt'));
    }
    if (/^https:\/\//.test(a.url || '')) aktionen.append(h('a', { class: 'as-link', href: a.url, target: '_blank', rel: 'noopener' }, 'Mappe in SharePoint ↗'));
    innen.append(aktionen);

    /* Firma und Zuweisung */
    const firmenId = 'as-firmen';
    const firma = h('input', { type: 'text', class: 'as-feld firma', value: a.firma || '', maxlength: 100, list: firmenId, autocomplete: 'off', disabled: !darf, 'aria-label': 'Firma' });
    const zuweisung = h('input', { type: 'text', class: 'as-feld zuweisung', value: a.ansprech || '', maxlength: 255, autocomplete: 'off', placeholder: 'Person oder Kürzel', disabled: !darf, 'aria-label': 'Zuweisung' });
    const beiEnter = (feld) => feld.addEventListener('keydown', (ev) => { if (ev.key === 'Enter') { ev.preventDefault(); feld.blur(); } });
    firma.addEventListener('change', () => { const v = firma.value.trim(); if (v !== (a.firma || '') && o.beiFirma) o.beiFirma(a, v); });
    zuweisung.addEventListener('change', () => { const v = zuweisung.value.trim(); if (v !== (a.ansprech || '') && o.beiZuweisung) o.beiZuweisung(a, v); });
    beiEnter(firma); beiEnter(zuweisung);
    const kz = h('div', { class: 'as-kz', role: 'group', 'aria-label': 'Im Team zuweisen' }, (m.team || []).map((b) => {
      const an = woerter(a.ansprech).includes(b.kuerzel);
      const t = h('button', { type: 'button', class: 'kz', 'aria-pressed': String(an), title: b.name, disabled: !darf,
        onclick: () => o.beiZuweisung && o.beiZuweisung(a, zuweisungUmschalten(a.ansprech, b.kuerzel, m.team)) }, b.kuerzel);
      t.style.background = b.hinter; t.style.color = b.schrift;
      return t;
    }));
    innen.append(h('div', { class: 'as-zwei' },
      h('label', null, h('span', { text: 'Firma' }), firma),
      h('label', null, h('span', { text: 'Zuweisung' }), zuweisung)),
      h('datalist', { id: firmenId }, (m.firmen || []).map((f) => h('option', { value: f }))), kz);

    /* Fälligkeit */
    const datum = h('input', { type: 'date', class: 'as-feld faellig', value: a.faellig || '', disabled: !darf || erledigt, 'aria-label': 'Fälligkeit' });
    datum.addEventListener('change', () => { if (datum.value !== (a.faellig || '') && o.beiFaellig) o.beiFaellig(a, datum.value); });
    innen.append(h('div', { class: 'as-block' }, h('label', null, h('span', { text: 'Fälligkeit' }), datum),
      h('div', { class: 'as-schnell' }, schnell(m.heute).map(([wort, iso]) => h('button', { type: 'button', disabled: !darf || erledigt || iso === (a.faellig || ''),
        onclick: () => o.beiFaellig && o.beiFaellig(a, iso) }, wort)))));

    /* Verlauf: Notizen und Mails zusammen */
    const z = zeitstrahl(m.verlauf, m.dateien, m.heute);
    const mailKarte = (d) => {
      const k = kurzFinden(d, m.mails);
      const karte = h('div', { class: 'as-mail', 'data-datei': d.name },
        h('div', { class: 'as-mailkopf' }, h('span', { class: 'umschlag', 'aria-hidden': 'true', text: '✉' }), h('span', { class: 'zeit', text: d.zeit }), h('span', { class: 'betreff', text: d.betreff })));
      if (k) karte.append(h('p', { class: 'as-kurz', text: k.kurz }));
      const tasten = h('div', { class: 'as-mailtasten' });
      if (m.outlook) tasten.append(h('button', { type: 'button', class: 'outlook', onclick: () => o.beiOutlook && o.beiOutlook(d.rel) }, 'In Outlook öffnen'));
      const url = dateiUrl(d);
      if (url) tasten.append(h('a', { href: url, target: '_blank', rel: 'noopener' }, 'Datei ↗'));
      if (darf && !erledigt) tasten.append(h('button', { type: 'button', class: 'dazu', onclick: () => o.beiNotizMail && o.beiNotizMail(d.name) }, 'Notiz dazu'));
      karte.append(tasten);
      return karte;
    };
    const verlauf = h('div', { class: 'as-verlauf' });
    if (m.laedt) verlauf.append(h('div', { class: 'gs-leer', text: 'Wird geladen …' }));
    else if (m.fehler) verlauf.append(h('div', { class: 'gs-leer warn', text: 'Verlauf und Mails ließen sich nicht laden: ' + m.fehler }));
    else if (!z.posten.length) verlauf.append(h('div', { class: 'gs-leer', text: 'Noch kein Eintrag und keine Mail.' }));
    else {
      for (const p of z.posten) {
        if (p.art === 'mail') { verlauf.append(h('div', { class: 'as-posten nur-mail' }, mailKarte(p.mail))); continue; }
        const kopfzeile = h('div', { class: 'as-notizkopf' }, h('span', { class: 'tag', text: p.datum || '' }), p.kz ? h('span', { class: 'wer', text: p.kz }) : null);
        // Löschen nimmt den Eintrag aus dem Verlauf. Eine zugeordnete Mail bleibt in der Mappe und steht danach allein.
        if (darf && !erledigt) kopfzeile.append(h('button', { type: 'button', class: 'weg', 'aria-label': 'Eintrag löschen' + (p.datum ? ' vom ' + p.datum : ''),
          onclick: () => o.beiLoeschen && o.beiLoeschen(a, { roh: p.roh, gleiche: p.gleiche, mitMail: p.mails.length > 0 }) }, 'Löschen'));
        const posten = h('div', { class: 'as-posten' }, kopfzeile, h('p', { class: 'as-notiz', text: p.text }));
        for (const d of p.mails) posten.append(mailKarte(d));
        if (p.fehlt) posten.append(h('div', { class: 'as-fehlt', text: 'Die Mail „' + p.fehlt + '“ liegt nicht mehr in der Mappe.' }));
        verlauf.append(posten);
      }
    }
    const neu = h('div', { class: 'as-neu' });
    if (m.schreiben && !erledigt) {
      const text = h('input', { type: 'text', class: 'as-feld notiz', maxlength: 500, placeholder: 'Neuer Eintrag, Enter', disabled: zu, 'aria-label': 'Neuer Eintrag im Verlauf' });
      const wahl = h('select', { class: 'as-feld zu-mail', disabled: zu || !z.mails.length, 'aria-label': 'Eintrag einer Mail zuordnen' },
        h('option', { value: '' }, z.mails.length ? 'ohne Mail' : 'keine Mail in der Mappe'),
        z.mails.map((d) => h('option', { value: d.name }, 'zu: ' + (d.zeit ? d.zeit + ' ' : '') + d.betreff)));
      if (m.notizMail && z.mails.some((d) => d.name === m.notizMail)) wahl.value = m.notizMail;
      const sende = () => { const t = text.value.trim(); if (t && o.beiNotiz) o.beiNotiz(a, t, wahl.value); };
      text.addEventListener('keydown', (ev) => { if (ev.key === 'Enter') { ev.preventDefault(); sende(); } });
      neu.append(text, wahl, h('button', { type: 'button', class: 'as-taste eintragen', disabled: zu, onclick: sende }, 'Eintragen'));
    }
    const anzahl = m.laedt || m.fehler ? '' : z.posten.length || '';
    const abschnitt = h('section', { class: 'gs-abschnitt', id: 'as-verlauf' }, h('h2', null, h('span', { text: 'Verlauf und Mails' }), anzahl !== '' ? h('span', { class: 'anz', text: String(anzahl) }) : null), neu, verlauf);
    innen.append(abschnitt);

    if (!m.laedt && !m.fehler && z.dateien.length) {
      innen.append(h('section', { class: 'gs-abschnitt', id: 'as-dateien' }, h('h2', null, h('span', { text: 'Weitere Dateien' }), h('span', { class: 'anz', text: String(z.dateien.length) })),
        h('div', { class: 'gs-liste' }, z.dateien.map((d) => {
          const url = dateiUrl(d), eig = { class: 'as-datei' };
          if (url) { eig.href = url + (MAILDATEI.test(d.name) ? '' : '?web=1'); eig.target = '_blank'; eig.rel = 'noopener'; }
          return h(url ? 'a' : 'div', eig, h('span', { class: 'name', text: d.name }), h('span', { class: 'gr', text: groesse(d.groesse) }));
        }))));
    }
    innen.append(h('p', { class: 'gs-fuss', text: 'Kurzfassungen stammen aus der Bewertung des Helfers und können irren. Gesendet, verschoben oder gelöscht wird von hier nichts.' }));

    root.replaceChildren(kopf, innen);
    root.scrollTop = o.behalten ? oben : 0;
    // Nach „Notiz dazu“ steht der Zeiger im Feld für den neuen Eintrag
    if (m.notizMail && o.fokus) { const f = root.querySelector('.as-neu input.notiz'); if (f) f.focus({ preventScroll: false }); }
  }

  wurzel.FPKarteAufgabe = { bloecke, eintraege, eintragTag, dateiInfo, zeitstrahl, kurzFinden, outlookLink, zuweisungUmschalten, schnell, zeichnen };
})(typeof window !== 'undefined' ? window : this);
