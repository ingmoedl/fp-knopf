/*
  FP-Karte – Aufgaben bearbeiten
  ------------------------------
  Die einzige Stelle, an der die Karte schreibt. Drei Dinge an einer Aufgabe: erledigen (und wieder öffnen),
  Fälligkeit setzen, Notiz anfügen. Geschrieben wird immer erst auf einen Klick der angemeldeten Person.

  Felder, Formate und Schutz sind dieselben wie in der Übersicht (uebersicht.js):
    FPStatus „Erledigt“ oder „Offen“, FPErledigtAm und FPFaelligkeit als TT.MM.JJJJ,
    FPVerlauf mit der neuen Zeile „TT.MM. KÜRZEL: Text“ vor dem bisherigen Text.
  Vor dem Schreiben wird der Stand der Aufgabe geprüft. Hat jemand anderes sie seit dem Laden geändert,
  wird nichts überschrieben.

  FPKarteBearbeiten.erzeugen({ sp, cfg, ich(), heute() }) liefert
    erledigen(a), wiederOeffnen(a), faelligSetzen(a, 'JJJJ-MM-TT' | ''), notiz(a, text)
  Jede Funktion gibt { geaendert } zurück, den neuen Stand der Aufgabe für die nächste Änderung.
  a ist eine Aufgabe aus der Ladeschicht mit { id, geaendert }.
*/
(function (wurzel) {
  'use strict';

  const FELDER = ['FPStatus', 'FPErledigtAm', 'FPFaelligkeit', 'FPVerlauf'];
  const deLang = (iso) => (iso ? `${iso.slice(8, 10)}.${iso.slice(5, 7)}.${iso.slice(0, 4)}` : '');
  const feld = (name, wert) => ({ FieldName: name, FieldValue: wert });

  function gueltigerTag(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso));
    if (!m) return false;
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    return d.getUTCFullYear() === +m[1] && d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3];
  }

  function erzeugen(o) {
    const liste = `${o.cfg.web}/_api/web/lists(guid'${o.cfg.listId}')`;

    async function speichern(a, werte) {
      if (!a || !Number.isInteger(a.id) || a.id <= 0) throw new Error('Die Aufgabe hat keine gültige Nummer.');
      const eintrag = `${liste}/items(${a.id})`;
      const stand = await o.sp(`${eintrag}?$select=Modified,Editor/Title&$expand=Editor`);
      if (stand && a.geaendert && stand.Modified !== a.geaendert) {
        const e = new Error(`${(stand.Editor && stand.Editor.Title) || 'Jemand'} hat die Aufgabe inzwischen geändert. Es wurde nichts überschrieben.`);
        e.konflikt = true;
        throw e;
      }
      const r = await o.sp(`${eintrag}/ValidateUpdateListItem`, { method: 'POST', body: { formValues: werte, bNewDocumentUpdate: false } });
      const fehler = ((r && r.value) || []).filter((x) => x.HasException).map((x) => x.FieldName + ': ' + x.ErrorMessage);
      if (fehler.length) throw new Error('SharePoint hat die Änderung abgelehnt. ' + fehler.join(', '));
      const neu = await o.sp(`${eintrag}?$select=Modified`);
      return { geaendert: String((neu && neu.Modified) || '') };
    }

    return {
      erledigen: (a) => speichern(a, [feld('FPStatus', 'Erledigt'), feld('FPErledigtAm', deLang(o.heute()))]),
      wiederOeffnen: (a) => speichern(a, [feld('FPStatus', 'Offen'), feld('FPErledigtAm', '')]),
      faelligSetzen: async (a, iso) => {
        if (iso !== '' && !gueltigerTag(iso)) throw new Error('Das ist kein gültiges Datum.');
        return speichern(a, [feld('FPFaelligkeit', deLang(iso))]);
      },
      notiz: async (a, text) => {
        const t = String(text == null ? '' : text).replace(/\s+/g, ' ').trim().slice(0, 500);
        if (!t) throw new Error('Die Notiz ist leer.');
        if (!a || !Number.isInteger(a.id) || a.id <= 0) throw new Error('Die Aufgabe hat keine gültige Nummer.');
        const x = await o.sp(`${liste}/items(${a.id})?$select=FPVerlauf`);
        const ich = o.ich && o.ich();
        const heute = o.heute();
        const zeile = `${heute.slice(8, 10)}.${heute.slice(5, 7)}. ${(ich && ich.kuerzel) || '?'}: ${t}`;
        return speichern(a, [feld('FPVerlauf', zeile + (x && x.FPVerlauf ? '\n' + x.FPVerlauf : ''))]);
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
