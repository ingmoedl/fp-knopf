/*
  FP-Karte – Aufgaben bearbeiten
  ------------------------------
  Die einzige Stelle, an der die Karte schreibt. An einer Aufgabe: erledigen (und wieder öffnen), Fälligkeit,
  Firma, Zuweisung, Notiz im Verlauf. Geschrieben wird immer erst auf einen Klick der angemeldeten Person.

  Felder und Formate sind dieselben wie in der Übersicht (uebersicht.js):
    FPStatus „Erledigt“ oder „Offen“, FPErledigtAm und FPFaelligkeit als TT.MM.JJJJ,
    FPVerlauf mit der neuen Zeile „TT.MM. KÜRZEL: Text“ vor dem bisherigen Text. Gehört die Notiz zu einer Mail
    in der Mappe, folgt die Zeile „✉ Dateiname“.
    FPFirma als Text, FPAnsprechperson als Text (Person bei der Firma oder Kürzel aus dem Team), dazu FPBearbeiter mit
    den Konten zu diesen Kürzeln.

  Schutz vor dem Überschreiben, je Feld: Vor dem Schreiben liest die Karte den Stand der Aufgabe.
    erledigen       nur, wenn die Aufgabe noch offen ist
    wiederOeffnen   nur, wenn sie erledigt ist
    faelligSetzen   nur, wenn die Fälligkeit noch die ist, die die Karte zeigt
    firmaSetzen, zuweisungSetzen   nur, wenn das Feld noch den Wert hat, den die Karte zeigt
    notiz           immer, die neue Zeile kommt vor den Verlauf, wie er in diesem Augenblick ist
  Dass sich an der Aufgabe sonst etwas getan hat (eine Mail kam in die Mappe, jemand hat eine Notiz angefügt),
  ist kein Hindernis. Die erste Fassung verglich den Zeitpunkt der letzten Änderung und lehnte am 19.11.2026
  im Betrieb eine Änderung ab, obwohl niemand dazwischen war.

  FPKarteBearbeiten.erzeugen({ sp, cfg, ich(), heute() }) liefert
    stand(a), lesen(a), erledigen(a), wiederOeffnen(a), faelligSetzen(a, 'JJJJ-MM-TT' | ''),
    firmaSetzen(a, text), zuweisungSetzen(a, text, team), notiz(a, text, datei)
  Jede Funktion gibt den Stand zurück: { status, faellig, erledigtAm, firma, ansprech, von, geaendert }, lesen dazu
  den Verlauf. Bei einem Konflikt trägt der Fehler .konflikt = true und .stand, damit die Anzeige nachziehen kann.
  a ist eine Aufgabe aus der Ladeschicht mit { id, faellig, firma, ansprech }.
*/
(function (wurzel) {
  'use strict';

  const FELDER = ['FPStatus', 'FPErledigtAm', 'FPFaelligkeit', 'FPVerlauf', 'FPFirma', 'FPAnsprechperson', 'FPBearbeiter'];
  const woerter = (s) => String(s || '').split(/[^A-Za-zÄÖÜäöüß]+/).filter(Boolean);
  const einzeilig = (t, max) => String(t == null ? '' : t).replace(/\s+/g, ' ').trim().slice(0, max);
  const z2 = (n) => String(n).padStart(2, '0');
  const deLang = (iso) => (iso ? `${iso.slice(8, 10)}.${iso.slice(5, 7)}.${iso.slice(0, 4)}` : '');
  const deKurz = (iso) => (iso ? `${iso.slice(8, 10)}.${iso.slice(5, 7)}.` : '');
  // SharePoint speichert einen Zeitpunkt in UTC, gemeint ist der Tag in Ortszeit (wie in daten.js)
  const ausSp = (iso) => { if (!iso) return ''; const d = new Date(iso); return isNaN(d) ? '' : `${d.getFullYear()}-${z2(d.getMonth() + 1)}-${z2(d.getDate())}`; };
  const feld = (name, wert) => ({ FieldName: name, FieldValue: wert });

  function gueltigerTag(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso));
    if (!m) return false;
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    return d.getUTCFullYear() === +m[1] && d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3];
  }

  function erzeugen(o) {
    const liste = `${o.cfg.web}/_api/web/lists(guid'${o.cfg.listId}')`;
    const eintrag = (a) => {
      if (!a || !Number.isInteger(a.id) || a.id <= 0) throw new Error('Die Aufgabe hat keine gültige Nummer.');
      return `${liste}/items(${a.id})`;
    };

    const SPALTEN = 'Modified,FPStatus,FPFaelligkeit,FPErledigtAm,FPFirma,FPAnsprechperson,Editor/Title';
    const ausZeile = (x) => ({
      status: String((x && x.FPStatus) || 'Offen'), faellig: ausSp(x && x.FPFaelligkeit), erledigtAm: ausSp(x && x.FPErledigtAm),
      firma: String((x && x.FPFirma) || '').trim(), ansprech: String((x && x.FPAnsprechperson) || '').trim(),
      von: String((x && x.Editor && x.Editor.Title) || ''), geaendert: String((x && x.Modified) || ''),
    });
    async function stand(a) { return ausZeile(await o.sp(`${eintrag(a)}?$select=${SPALTEN}&$expand=Editor`)); }
    /* Stand und Verlauf in einer Anfrage, für die Aufgabenseite */
    async function lesen(a) {
      const x = await o.sp(`${eintrag(a)}?$select=${SPALTEN},FPVerlauf&$expand=Editor`);
      return Object.assign(ausZeile(x), { verlauf: String((x && x.FPVerlauf) || '').replace(/\r\n?/g, '\n') });
    }

    function konflikt(s, text) {
      const e = new Error((s.von || 'Jemand') + ' ' + text + ' Es wurde nichts überschrieben.');
      e.konflikt = true; e.stand = s;
      return e;
    }

    async function schreiben(a, werte) {
      const r = await o.sp(`${eintrag(a)}/ValidateUpdateListItem`, { method: 'POST', body: { formValues: werte, bNewDocumentUpdate: false } });
      const fehler = ((r && r.value) || []).filter((x) => x.HasException).map((x) => x.FieldName + ': ' + x.ErrorMessage);
      if (fehler.length) throw new Error('SharePoint hat die Änderung abgelehnt. ' + fehler.join(', '));
      return stand(a);
    }

    return {
      stand, lesen,
      erledigen: async (a) => {
        const s = await stand(a);
        if (s.status !== 'Offen') throw konflikt(s, 'hat die Aufgabe schon erledigt.');
        return schreiben(a, [feld('FPStatus', 'Erledigt'), feld('FPErledigtAm', deLang(o.heute()))]);
      },
      wiederOeffnen: async (a) => {
        const s = await stand(a);
        if (s.status !== 'Erledigt') throw konflikt(s, 'hat die Aufgabe schon wieder geöffnet.');
        return schreiben(a, [feld('FPStatus', 'Offen'), feld('FPErledigtAm', '')]);
      },
      faelligSetzen: async (a, iso) => {
        if (iso !== '' && !gueltigerTag(iso)) throw new Error('Das ist kein gültiges Datum.');
        const s = await stand(a);
        if (s.status !== 'Offen') throw konflikt(s, 'hat die Aufgabe inzwischen erledigt.');
        if (s.faellig !== (a.faellig || '')) throw konflikt(s, s.faellig ? `hat die Fälligkeit inzwischen auf den ${deKurz(s.faellig)} gesetzt.` : 'hat die Aufgabe inzwischen ohne Termin gestellt.');
        return schreiben(a, [feld('FPFaelligkeit', deLang(iso))]);
      },
      firmaSetzen: async (a, text) => {
        const v = einzeilig(text, 100);
        const s = await stand(a);
        if (s.firma !== (a.firma || '')) throw konflikt(s, 'hat die Firma inzwischen geändert.');
        return schreiben(a, [feld('FPFirma', v)]);
      },
      /* team: [{ kuerzel, mail }]. Kürzel aus dem Team in der Zuweisung werden zu Bearbeitern, wie in der Übersicht. */
      zuweisungSetzen: async (a, text, team) => {
        const v = einzeilig(text, 255);
        const s = await stand(a);
        if (s.ansprech !== (a.ansprech || '')) throw konflikt(s, 'hat die Zuweisung inzwischen geändert.');
        const w = woerter(v);
        const konten = (team || []).filter((b) => b && w.includes(b.kuerzel) && b.mail).map((b) => ({ Key: 'i:0#.f|membership|' + b.mail }));
        return schreiben(a, [feld('FPAnsprechperson', v), feld('FPBearbeiter', konten.length ? JSON.stringify(konten) : '')]);
      },
      /* datei: Name einer Mail in der Mappe, zu der die Notiz gehört, oder leer */
      notiz: async (a, text, datei) => {
        const t = einzeilig(text, 500);
        if (!t) throw new Error('Die Notiz ist leer.');
        const d = String(datei == null ? '' : datei).trim();
        if (d && (/[\r\n\/\\]/.test(d) || d.length > 255)) throw new Error('Der Name der Mail ist ungültig.');
        const x = await o.sp(`${eintrag(a)}?$select=FPVerlauf`);
        const ich = o.ich && o.ich();
        const heute = o.heute();
        const zeile = `${heute.slice(8, 10)}.${heute.slice(5, 7)}. ${(ich && ich.kuerzel) || '?'}: ${t}` + (d ? '\n✉ ' + d : '');
        return schreiben(a, [feld('FPVerlauf', zeile + (x && x.FPVerlauf ? '\n' + x.FPVerlauf : ''))]);
      },
    };
  }

  /* Für die Sperre der Betriebsquelle: Ist das genau eine Bearbeitung, wie sie hier entsteht? */
  function istBearbeitung(pfad, opts, listId) {
    if (String((opts && opts.method) || '').toUpperCase() !== 'POST' || !listId) return false;
    const m = /\/_api\/web\/lists\(guid'([0-9a-f-]+)'\)\/items\(\d+\)\/ValidateUpdateListItem$/i.exec(String(pfad));
    if (!m || m[1].toLowerCase() !== String(listId).toLowerCase()) return false;
    const b = opts.body;
    if (!b || b.bNewDocumentUpdate !== false || !Array.isArray(b.formValues) || !b.formValues.length) return false;
    return b.formValues.every((w) => w && FELDER.includes(w.FieldName) && typeof w.FieldValue === 'string');
  }

  wurzel.FPKarteBearbeiten = { erzeugen, istBearbeitung, deLang, FELDER };
})(typeof window !== 'undefined' ? window : this);
