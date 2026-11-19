/*
  FP-Karte – Auswertung
  ---------------------
  Reine Funktionen ohne Zugriff auf Seite oder SharePoint: was auf der Karte hervortritt und was die
  Gebäudeseite in welcher Reihenfolge zeigt. Prüfbar in werkzeuge/test_karte.html.

  Filter: 'alle', 'meine' (angemeldete Person) oder das Kürzel eines Teammitglieds.
  Firmen: Liste gewählter Firmen (Spalte Firma der Aufgabe), leer heißt alle. '' steht für Aufgaben ohne Firma.
  Beide Filter gelten zusammen und wirken auf die Aufgaben. Termine und offene Punkte gehören zum Gebäude
  und bleiben stehen.
*/
(function (wurzel) {
  'use strict';

  const z2 = (n) => String(n).padStart(2, '0');
  const WT = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
  const tagTeile = (t) => /^(\d{4})-(\d{2})-(\d{2})/.exec(String(t || ''));
  const tagZahl = (t) => { const m = tagTeile(t); return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) / 864e5 : null; };
  const tageBis = (t, heute) => { const a = tagZahl(t), b = tagZahl(heute); return a === null || b === null ? null : Math.round(a - b); };
  const heuteIso = () => { const d = new Date(); return `${d.getFullYear()}-${z2(d.getMonth() + 1)}-${z2(d.getDate())}`; };
  const tagKurz = (t) => { const m = tagTeile(t); if (!m) return ''; return `${WT[new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])).getUTCDay()]} ${m[3]}.${m[2]}.`; };
  const tagOhneWt = (t) => { const m = tagTeile(t); return m ? `${m[3]}.${m[2]}.` : ''; };
  const uhrzeit = (t) => { const m = /T(\d{2}):(\d{2})/.exec(String(t || '')); return m ? `${m[1]}:${m[2]}` : ''; };
  const standText = (t) => { const m = tagTeile(t); return m ? `${m[3]}.${m[2]}.${m[1]}${uhrzeit(t) ? ' ' + uhrzeit(t) : ''}` : ''; };

  /* Wie weit ein Datum weg ist, in Worten. */
  function abstand(tage) {
    if (tage === null) return '';
    if (tage === 0) return 'heute';
    if (tage === 1) return 'morgen';
    if (tage === -1) return 'seit gestern überfällig';
    if (tage < 0) return `${-tage} Tage überfällig`;
    return `in ${tage} Tagen`;
  }

  const hatFirma = (a, firmen) => !firmen || !firmen.length || firmen.includes(a.firma || '');
  const nachFirmen = (aufgaben, firmen) => (firmen && firmen.length ? aufgaben.filter((a) => hatFirma(a, firmen)) : aufgaben);

  function passt(a, filter, ich, firmen) {
    if (!hatFirma(a, firmen)) return false;
    if (!filter || filter === 'alle') return true;
    if (filter === 'meine') return !!ich && ((!!ich.kuerzel && a.kuerzel.includes(ich.kuerzel)) || a.leute.includes(ich.mail));
    return a.kuerzel.includes(filter);
  }

  /* Anzahl offener Aufgaben je Filter, für die Zahlen an den Tasten. */
  function zaehler(aufgaben, team, ich) {
    const z = { alle: aufgaben.length, meine: aufgaben.filter((a) => passt(a, 'meine', ich)).length };
    for (const b of team) z[b.kuerzel] = aufgaben.filter((a) => a.kuerzel.includes(b.kuerzel)).length;
    return z;
  }

  /* Firmen mit der Anzahl ihrer Aufgaben, für die Firmenwahl. Die Zahlen folgen dem Personenfilter.
     Reihenfolge: meiste Aufgaben zuerst, bei Gleichstand nach Name, Aufgaben ohne Firma zuletzt. */
  function firmenListe(aufgaben, filter, ich) {
    const zahl = new Map();
    for (const a of aufgaben) {
      if (!passt(a, filter, ich)) continue;
      const f = a.firma || '';
      zahl.set(f, (zahl.get(f) || 0) + 1);
    }
    return [...zahl].map(([name, anzahl]) => ({ name, anzahl }))
      .sort((x, y) => (!x.name - !y.name) || (y.anzahl - x.anzahl) || x.name.localeCompare(y.name, 'de'));
  }

  /* Was auf der Karte hervortritt. Mit Personen- oder Firmenfilter zählen nur die passenden Aufgaben,
     ohne Filter auch die offenen Punkte und kommenden Fristen aus den Mails. */
  function lagebild(gebaeude, aufgaben, index, filter, ich, heute, firmen) {
    const aus = new Map();
    const neu = () => ({ aufgaben: 0, ueberfaellig: 0, punkte: 0, frist: '', relevant: false, dringend: false });
    for (const nr of gebaeude.keys()) aus.set(nr, neu());
    const ohne = { aufgaben: 0, ueberfaellig: 0 };
    for (const a of aufgaben) {
      if (!passt(a, filter, ich, firmen)) continue;
      const t = a.faellig ? tageBis(a.faellig, heute) : null;
      const ueber = t !== null && t <= 0;
      if (!a.gebaeude.length) { ohne.aufgaben++; if (ueber) ohne.ueberfaellig++; continue; }
      for (const g of a.gebaeude) {
        const e = aus.get(g.nr); if (!e) continue;
        e.aufgaben++; if (ueber) e.ueberfaellig++;
      }
    }
    const alle = (!filter || filter === 'alle') && !(firmen && firmen.length);
    for (const [nr, e] of aus) {
      const i = index && index.get(nr);
      if (i && alle) { e.punkte = i.offene_punkte; e.frist = i.naechste_frist && tageBis(i.naechste_frist, heute) >= 0 ? i.naechste_frist : ''; }
      e.relevant = e.aufgaben > 0 || e.punkte > 0 || !!e.frist;
      e.dringend = e.ueberfaellig > 0;
    }
    return { gebaeude: aus, ohne };
  }

  const nachFaelligkeit = (ich) => (a, b) => {
    const fa = a.faellig || '9999', fb = b.faellig || '9999';
    if (fa !== fb) return fa < fb ? -1 : 1;
    const ma = passt(a, 'meine', ich) ? 0 : 1, mb = passt(b, 'meine', ich) ? 0 : 1;
    return ma - mb || a.titel.localeCompare(b.titel, 'de');
  };

  /*
    Inhalt der Gebäudeseite.
      nr       Gebäudenummer, oder null für die Aufgaben ohne Gebäude
      datei    Gebäudedatei aus der Ladeschicht (oder null)
      o        { filter, ich, heute, eigeneParteien, firmen }
    Reihenfolge:
      Aufgaben   nach Fälligkeit, Überfälliges oben, ohne Termin unten. Bei gleichem Tag die eigenen zuerst.
      Termine    überfällige Fristen, dann alles ab heute nach Datum. Vergangene Termine getrennt.
      Punkte     je Partei, die eigene zuerst, dann die mit den meisten Punkten. In der Gruppe das Neueste oben.
      Verlauf    Mails, neueste zuerst.
  */
  function seite(nr, aufgaben, datei, o) {
    const heute = o.heute || heuteIso();
    const zumGebaeude = aufgaben.filter((a) => (nr === null ? !a.gebaeude.length : a.gebaeude.some((g) => g.nr === nr)));
    const sicht = zumGebaeude.filter((a) => passt(a, o.filter, o.ich, o.firmen)).sort(nachFaelligkeit(o.ich)).map((a) => {
      const tage = a.faellig ? tageBis(a.faellig, heute) : null;
      const z = nr === null ? null : a.gebaeude.find((g) => g.nr === nr);
      return Object.assign({}, a, { tage, ueberfaellig: tage !== null && tage < 0, heute: tage === 0, meine: passt(a, 'meine', o.ich), zuordnung: z });
    });

    const fristen = ((datei && datei.fristen) || []).filter((f) => tagTeile(f.datum)).map((f) => {
      const tage = tageBis(f.datum, heute);
      return { datum: f.datum, text: String(f.text || ''), art: f.art === 'frist' ? 'frist' : 'termin', tage, betreff: String(f.betreff || ''), eingang: String(f.eingang || '') };
    });
    const ueber = fristen.filter((f) => f.tage < 0 && f.art === 'frist').sort((a, b) => (a.datum < b.datum ? -1 : 1)).map((f) => Object.assign(f, { ueberfaellig: true }));
    const kommend = fristen.filter((f) => f.tage >= 0).sort((a, b) => (a.datum < b.datum ? -1 : a.datum > b.datum ? 1 : 0));
    const vergangen = fristen.filter((f) => f.tage < 0 && f.art !== 'frist').sort((a, b) => (a.datum > b.datum ? -1 : 1));

    const eigen = (o.eigeneParteien || []).map((s) => String(s).toLowerCase());
    const gruppen = new Map();
    for (const p of (datei && datei.offene_punkte) || []) {
      const bei = String(p.bei || '').trim() || 'Ohne Angabe';
      if (!gruppen.has(bei)) gruppen.set(bei, { bei, eigen: eigen.includes(bei.toLowerCase()), eintraege: [] });
      const seit = tagTeile(p.seit) ? String(p.seit).slice(0, 10) : '';
      gruppen.get(bei).eintraege.push({ text: String(p.text || ''), seit, eingang: String(p.eingang || ''), betreff: String(p.betreff || '') });
    }
    const punkte = [...gruppen.values()].sort((a, b) => (b.eigen - a.eigen) || (b.eintraege.length - a.eintraege.length) || a.bei.localeCompare(b.bei, 'de'));
    for (const g of punkte) g.eintraege.sort((a, b) => (a.eingang > b.eingang ? -1 : a.eingang < b.eingang ? 1 : a.text.localeCompare(b.text, 'de')));
    const punkteAnzahl = punkte.reduce((s, g) => s + g.eintraege.length, 0);
    const eigene = punkte.filter((g) => g.eigen);

    const verlauf = ((datei && datei.mails) || []).slice().sort((a, b) => (String(a.eingang) > String(b.eingang) ? -1 : 1));

    return {
      aufgaben: sicht, aufgabenAlle: zumGebaeude.length,
      firmen: firmenListe(zumGebaeude, o.filter, o.ich),
      termine: ueber.concat(kommend), vergangen,
      punkte, punkteAnzahl,
      verlauf,
      zahlen: {
        aufgaben: sicht.length, ueberfaellig: sicht.filter((a) => a.ueberfaellig).length, heute: sicht.filter((a) => a.heute).length,
        termine: ueber.length + kommend.length, fristUeber: ueber.length, naechster: kommend.length ? kommend[0].datum : '',
        punkte: punkteAnzahl, punkteEigen: eigene.reduce((s, g) => s + g.eintraege.length, 0), eigenName: eigene.length ? eigene[0].bei : '',
      },
    };
  }

  /*
    Aufgabenliste der Auswahl: alle Aufgaben, die zu Person und Firmen passen, gruppiert nach Gebäude.
    Aufgaben ohne Gebäude stehen zuerst, dann die Gebäude nach Nummer. Eine Aufgabe an zwei Gebäuden steht bei
    beiden, in den Zahlen zählt sie einmal.
      gebaeude   Map der Gebäude aus der Ladeschicht (für die Namen)
      o          { filter, ich, heute, firmen }
  */
  function liste(aufgaben, gebaeude, o) {
    const heute = o.heute || heuteIso();
    const passend = aufgaben.filter((a) => passt(a, o.filter, o.ich, o.firmen));
    const reich = (a, nr) => {
      const tage = a.faellig ? tageBis(a.faellig, heute) : null;
      return Object.assign({}, a, { tage, ueberfaellig: tage !== null && tage < 0, heute: tage === 0, meine: passt(a, 'meine', o.ich), zuordnung: nr === null ? null : a.gebaeude.find((g) => g.nr === nr) });
    };
    const nach = new Map();
    for (const a of passend) {
      for (const nr of (a.gebaeude.length ? a.gebaeude.map((g) => g.nr) : [null])) {
        if (!nach.has(nr)) nach.set(nr, []);
        nach.get(nr).push(reich(a, nr));
      }
    }
    const ordnung = (x, y) => (x === null ? -1 : y === null ? 1 : x.localeCompare(y, 'de', { numeric: true }));
    const gruppen = [...nach.keys()].sort(ordnung).map((nr) => {
      const g = nr === null ? null : gebaeude.get(nr);
      return { nr, name: nr === null ? 'Ohne Gebäude' : (g && g.name) || '', aufgaben: nach.get(nr).sort(nachFaelligkeit(o.ich)) };
    });
    const tage = passend.map((a) => (a.faellig ? tageBis(a.faellig, heute) : null));
    return {
      gruppen,
      zahlen: {
        aufgaben: passend.length, ueberfaellig: tage.filter((t) => t !== null && t < 0).length, heute: tage.filter((t) => t === 0).length,
        gebaeude: gruppen.filter((g) => g.nr !== null).length, ohne: nach.has(null) ? nach.get(null).length : 0,
      },
      firmen: firmenListe(aufgaben, o.filter, o.ich),
    };
  }

  wurzel.FPKarteAuswertung = { passt, nachFirmen, firmenListe, zaehler, lagebild, seite, liste, tageBis, tagKurz, tagOhneWt, uhrzeit, standText, abstand, heuteIso };
})(typeof window !== 'undefined' ? window : this);
