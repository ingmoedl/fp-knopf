/*
  FP-Knopf – Ablage per Klick
  ---------------------------
  Legt die geöffnete Mail als Datei (.eml) in den Ordner einer Partei. Ausgelöst wird das nur durch einen Klick
  im Reiter „Vorschlag“. Die Mail selbst bleibt im Postfach, wo sie ist: Hier wird nichts verschoben, markiert
  oder gelöscht.

  Wohin:
    Probebetrieb (Standard)   <FP-Wissen>/ablage-test/<Partei>
    echte Ablage              <Pfad des Projektordners>/<Ablageordner>/<Partei>   (Namen aus dem Register)
  Echt abgelegt wird nur für Personen, deren Kürzel im Register unter ablage_echt steht, und nur, wenn das
  Register den Pfad des Projektordners nennt (ablage.pfad). Sonst bleibt es beim Testordner.

  Ablauf: prüfen, holen, hochladen, festhalten.
    prüfen      Ist die Mail schon abgelegt (Eintrag in <FP-Wissen>/ablage)? Gibt es den Zielordner?
                Geprüft wird mit Proben auf einen Pfad. Ein Ordner wird nie aufgelistet: Parteiordner können
                mehr als 5.000 Dateien enthalten, und ab dieser Zahl lehnt SharePoint Listen ab.
    holen       die Mail als Datei, vom Aufrufer
    hochladen   unter „JJMMTT Betreff.eml“. Liegt dort schon eine andere Mail gleichen Namens, mit „ (2)“.
                In der echten Ablage wird kein Ordner angelegt, und liegt die Mail dort schon als .msg, bleibt es dabei.
    festhalten  <FP-Wissen>/ablage/<JJJJ-MM des Eingangs>/<Kurz-Hash>.json: Ziel, Datei, Vorschlag, Ampel,
                eigene Wahl, Kürzel, Zeit, Weg. Daraus liest der Reiter „liegt schon in …“.

  FPAblage.erzeugen({ sp, web, wissenRel, register(), kuerzel(), hash(schluessel), mailBytes(), jetzt() })
    liefert { ablegen(mail, partei, vorschlag, wahl), schonAbgelegt(mail) }
    mail: { schluessel, betreff, eingang (Date) }
  FPAblage.meldung(ergebnis, register) macht aus dem Ergebnis von ablegen den Satz für den Reiter.
  Braucht FPEinordnung.dateiname (einordnung.js).
*/
(function (wurzel) {
  'use strict';

  const pfad = (s) => encodeURIComponent(String(s).replace(/'/g, "''"));
  const ohneSchraeg = (s) => String(s || '').replace(/\/+$/, '');

  /* Zielordner für diese Partei und Person. Der Name muss wörtlich ein Ordner aus dem Register sein. */
  function ziel(reg, partei, kuerzel, wissenRel) {
    const bekannt = ((reg && reg.parteien) || []).some((p) => p.ordner_kom === partei);
    if (!partei || !bekannt) throw new Error('„' + partei + '“ ist kein Ordner der Ablage.');
    const abl = (reg && reg.ablage) || {};
    const echt = Array.isArray(reg.ablage_echt) && !!kuerzel && reg.ablage_echt.includes(kuerzel) && !!abl.pfad && !!abl.kom;
    return echt
      ? { art: 'echt', web: abl.web || '', ordner: ohneSchraeg(abl.pfad) + '/' + abl.kom + '/' + partei }
      : { art: 'test', web: '', ordner: ohneSchraeg(wissenRel) + '/ablage-test/' + partei };
  }

  /* Der Satz, den der Mensch nach dem Klick liest. */
  function meldung(r, reg) {
    const e = (r && r.eintrag) || {};
    const kom = (reg && reg.ablage && reg.ablage.kom) || 'der Ablage';
    const ort = e.art === 'echt' ? 'in ' + kom + ' / ' + e.ziel : 'im Testordner unter ' + e.ziel;
    const tag = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(e.wann || ''));
    if (r.schonDa && r.alsMsg) return 'Liegt schon als MSG ' + ort + ': ' + e.datei + '. Es wurde nichts hochgeladen.';
    if (r.schonDa) return 'Liegt schon ' + ort + ': ' + e.datei + (e.von ? ' (' + e.von + (tag ? ', ' + tag[3] + '.' + tag[2] + '.' : '') + ')' : '') + '. Es wurde nichts hochgeladen.';
    if (r.nachgetragen) return 'Die Datei lag schon ' + ort + ': ' + e.datei + '. Der Eintrag ist nachgetragen.';
    return 'Abgelegt ' + ort + ': ' + e.datei + '. Die Mail bleibt im Postfach.';
  }

  function erzeugen(o) {
    const wissen = ohneSchraeg(o.wissenRel);
    const api = (web) => (web || o.web) + '/_api/web';
    const istWeg = (e) => !!e && (e.status === 404 || /not ?found|nicht gefunden/i.test(e.message || ''));

    async function ordnerDa(web, rel) {
      try { const r = await o.sp(`${api(web)}/GetFolderByServerRelativePath(decodedurl='${pfad(rel)}')?$select=Exists`); return !!(r && r.Exists); }
      catch (e) { if (istWeg(e)) return false; throw e; }
    }
    /* Länge der Datei, oder null, wenn es sie nicht gibt */
    async function dateiLaenge(web, rel) {
      try { const r = await o.sp(`${api(web)}/GetFileByServerRelativePath(decodedurl='${pfad(rel)}')?$select=Exists,Length`); return r && r.Exists !== false ? Number(r.Length) : null; }
      catch (e) { if (istWeg(e)) return null; throw e; }
    }
    async function lies(rel) {
      try { return await o.sp(`${api('')}/GetFileByServerRelativePath(decodedurl='${pfad(rel)}')/$value`); }
      catch (e) { if (istWeg(e)) return null; throw e; }
    }
    /* Legt einen Ordner in FP-Wissen an, mit den fehlenden Ordnern darüber. Nie außerhalb von FP-Wissen. */
    async function wissenOrdner(rel) {
      if (rel !== wissen && !rel.startsWith(wissen + '/')) throw new Error('Angelegt wird nur in FP-Wissen.');
      if (rel === wissen || (await ordnerDa('', rel))) return;
      await wissenOrdner(rel.slice(0, rel.lastIndexOf('/')));
      try { await o.sp(`${api('')}/folders/AddUsingPath(DecodedUrl='${pfad(rel)}',overwrite=false)`, { method: 'POST' }); }
      catch (e) { if (!(await ordnerDa('', rel))) throw e; }        // legt ihn gerade jemand anderes an, ist das kein Fehler
    }
    const hochladen = (web, ordner, name, bytes, ersetzen) =>
      o.sp(`${api(web)}/GetFolderByServerRelativePath(decodedurl='${pfad(ordner)}')/Files/AddUsingPath(decodedurl='${pfad(name)}',overwrite=${ersetzen ? 'true' : 'false'})`, { method: 'POST', body: bytes });

    function eintragOrt(mail, hash) {
      if (!/^[0-9a-f]{20}$/.test(String(hash))) throw new Error('Die Kennung der Mail fehlt.');
      const d = mail.eingang instanceof Date ? mail.eingang : new Date(mail.eingang);
      if (isNaN(d)) throw new Error('Der Eingang der Mail ist nicht bekannt.');
      return { ordner: wissen + '/ablage/' + d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'), name: hash + '.json', eingang: d };
    }

    async function schonAbgelegt(mail) {
      const ort = eintragOrt(mail, await o.hash(mail.schluessel));
      return lies(ort.ordner + '/' + ort.name);
    }

    async function ablegen(mail, partei, vorschlag, wahl) {
      const k = String((o.kuerzel && o.kuerzel()) || '');
      if (!k) throw new Error('Dein Kürzel ist nicht bekannt. Ohne Kürzel wird nichts abgelegt.');
      const z = ziel(o.register(), partei, k, wissen);
      const ort = eintragOrt(mail, await o.hash(mail.schluessel));

      // prüfen
      const vorher = await lies(ort.ordner + '/' + ort.name);
      if (vorher) return { schonDa: true, eintrag: vorher };
      if (z.art === 'echt') {
        if (!(await ordnerDa(z.web, z.ordner))) throw new Error('Den Ordner „' + partei + '“ gibt es in der Ablage nicht. Es wird nichts angelegt.');
      } else await wissenOrdner(z.ordner);
      const basis = wurzel.FPEinordnung.dateiname(mail.betreff, ort.eingang, 100);
      if (z.art === 'echt' && (await dateiLaenge(z.web, z.ordner + '/' + basis + '.msg')) !== null) {
        return { schonDa: true, alsMsg: true, eintrag: { ziel: partei, ordner: z.ordner, datei: basis + '.msg', art: z.art } };
      }

      // holen
      const bytes = await o.mailBytes();
      if (!(bytes instanceof Uint8Array) || !bytes.length) throw new Error('Die Mail ließ sich nicht als Datei lesen.');

      // hochladen: freien Namen suchen. Liegt unter dem Namen schon genau diese Mail (gleiche Länge), blieb ein
      // früherer Versuch nach dem Hochladen stehen. Dann wird nur der Eintrag nachgetragen.
      let name = '', nachgetragen = false;
      for (let i = 1; i <= 9 && !name; i++) {
        const n = basis + (i > 1 ? ' (' + i + ')' : '') + '.eml';
        const laenge = await dateiLaenge(z.web, z.ordner + '/' + n);
        if (laenge === null) name = n;
        else if (laenge === bytes.length) { name = n; nachgetragen = true; }
      }
      if (!name) throw new Error('Unter diesem Namen liegen schon neun Dateien. Bitte von Hand ablegen.');
      if (!nachgetragen) await hochladen(z.web, z.ordner, name, bytes, false);

      // festhalten
      const s = vorschlag || {}, w = wahl || {};
      const satz = {
        schluessel: String(mail.schluessel || ''), eingang: ort.eingang.toISOString(),
        ziel: partei, ordner: z.ordner, datei: name, art: z.art,
        vorgeschlagen: String(s.ziel_kom || ''), ampel: String(s.ampel || ''), selbstGewaehlt: !!w.selbstGewaehlt,
        von: k, wann: o.jetzt().toISOString(), weg: 'klick',
      };
      await wissenOrdner(ort.ordner);
      await hochladen('', ort.ordner, ort.name, new TextEncoder().encode(JSON.stringify(satz, null, 1)), true);
      return nachgetragen ? { schonDa: false, nachgetragen: true, eintrag: satz } : { schonDa: false, eintrag: satz };
    }

    return { ablegen, schonAbgelegt };
  }

  wurzel.FPAblage = { ziel, erzeugen, meldung };
})(typeof window !== 'undefined' ? window : this);
