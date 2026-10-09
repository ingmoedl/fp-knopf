/* FP-Knopf – Einstellungen. Diese Datei ist öffentlich (GitHub Pages): keine Namen, Mailadressen oder Geheimnisse.
 * Das Team (Kürzel, Namen, Mails, Farben) steht in SharePoint: Bibliothek FP-Aufgaben → _Import/team.json.
 * Nach Änderungen hier in taskpane.html und uebersicht.html die Zahl hinter „einstellungen.js?v=“ um 1 erhöhen. */

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

  // Nur Teammitglieder mit dieser Maildomain werden aus team.json übernommen
  maildomain: "ing-burghausen.de",

  // Gelbe Kategorie an der Mail, sobald der FP-Knopf eine Aufgabe dazu anlegt oder die Mail ablegt.
  // In der Testphase aus: Die Kollegen würden sonst eine Outlook-Aufgabe suchen, die es nicht gibt.
  gelbMarkieren: false,
  gelbKategorie: "Aufgabe erstellt",
};
