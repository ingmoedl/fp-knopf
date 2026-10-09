/* FP-Knopf – Einstellungen. Nur diese Datei muss man für neue Kollegen oder eine neue App ändern.
 * Danach in taskpane.html und uebersicht.html die Zahl hinter „einstellungen.js?v=“ um 1 erhöhen. */

window.FP_EINSTELLUNGEN = {
  // Anwendungs-ID (Client) der App-Registrierung „FP-Knopf“ in Entra
  clientId: "ffce7fb6-79a5-4127-a921-1a8fe45390da",
  tenantId: "1571141a-75a9-43a3-ad47-8d613cfbb3e6",

  // Wo die Aufgaben liegen: Bibliothek „FP-Aufgaben“ im Team „2024“
  sharepoint: {
    host: "https://ingburghausengmbh.sharepoint.com",
    web: "/sites/2024",
    listId: "a83c2f28-6212-4181-b8ce-3445d7cd81ab",
    libRel: "/sites/2024/FPAufgaben",
  },

  // Wer als Bearbeiter wählbar ist; Farben wie die Outlook-Kategorien
  bearbeiter: [
    { kuerzel: "SMO", name: "Mödl, Samuel", mail: "moedl@ing-burghausen.de", hinter: "#dff6dd", schrift: "#0b6a0b" },
    { kuerzel: "HWE", name: "Weidacher, Helmut", mail: "weidacher@ing-burghausen.de", hinter: "#d6f1f2", schrift: "#005b70" },
    { kuerzel: "LSC", name: "Schweiger, Lena", mail: "schweiger@ing-burghausen.de", hinter: "#fbe3f1", schrift: "#9b0062" },
  ],

  // Gelbe Kategorie an der Mail, sobald der FP-Knopf eine Aufgabe dazu anlegt oder die Mail ablegt.
  // In der Testphase aus: Die Kollegen würden sonst eine Outlook-Aufgabe suchen, die es nicht gibt.
  gelbMarkieren: false,
  gelbKategorie: "Aufgabe erstellt",
};
