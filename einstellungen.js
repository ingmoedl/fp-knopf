/* FP-Knopf – Einstellungen. Nur diese Datei muss man für neue Kollegen oder eine neue App ändern.
 * Danach in taskpane.html die Zahl hinter „einstellungen.js?v=“ um 1 erhöhen. */

window.FP_EINSTELLUNGEN = {
  // Anwendungs-ID (Client) der App-Registrierung „FP-Knopf“ in Entra
  clientId: "ffce7fb6-79a5-4127-a921-1a8fe45390da",

  // Wer als Bearbeiter wählbar ist. Testphase: nur SMO.
  // Nach dem Test ergänzen:
  //   { kuerzel: "HWE", name: "Weidacher, Helmut", mail: "weidacher@ing-burghausen.de" },
  //   { kuerzel: "LSC", name: "Schweiger, Lena", mail: "schweiger@ing-burghausen.de" },
  bearbeiter: [
    { kuerzel: "SMO", name: "Mödl, Samuel", mail: "moedl@ing-burghausen.de" },
  ],
};
