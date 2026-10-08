/* FP-Knopf – Outlook-Add-in (ing Burghausen GmbH)
 * Legt aus der geöffneten Mail eine FP-Aufgabe (Mappe) in der Bibliothek „FP-Aufgaben“ an
 * oder legt die Mail in eine bestehende Mappe. Die Mail kommt als .eml-Datei in die Mappe.
 * Auth: Nested App Authentication (MSAL.js), keine Server-Komponente; SharePoint-REST mit Bearer-Token.
 * Gegenstück zu A_Einrichtung.js / B_Import.js im Ordner FP-Aufgaben (gleiche Spalten, gleiche Dateinamen). */

"use strict";

const E = window.FP_EINSTELLUNGEN || {};
const CONFIG = {
  version: "0.1",
  clientId: E.clientId,                             // App-Registrierung „FP-Knopf“ (einstellungen.js)
  tenantId: "1571141a-75a9-43a3-ad47-8d613cfbb3e6",
  spHost: "https://ingburghausengmbh.sharepoint.com",
  web: "/sites/2024",
  listId: "a83c2f28-6212-4181-b8ce-3445d7cd81ab",   // Bibliothek „FP-Aufgaben“
  libRel: "/sites/2024/FPAufgaben",
  contentType: "FP-Aufgabe",
  bearbeiter: E.bearbeiter || [],                   // einstellungen.js
  ohneFirma: ["ing"],                               // eigene Firma nicht automatisch vorschlagen
};
CONFIG.scopes = [CONFIG.spHost + "/AllSites.Write"];

let pca = null;
let standalone = false;
let firmen = [];
let mappen = [];          // offene Mappen [{id, name, rel, firma, faellig}]
let gewaehlt = null;      // gewählte Mappe im Modus „Zu bestehender“
let modus = "neu";
let laeuft = false;
const el = (id) => document.getElementById(id);
const mailItem = () => (Office.context && Office.context.mailbox) ? Office.context.mailbox.item : null;
const warte = (ms) => new Promise((s) => setTimeout(s, ms));

/* ---------- Start ---------- */

Office.onReady(async () => {
  try {
    if (!CONFIG.clientId || !/^[0-9a-f-]{36}$/i.test(CONFIG.clientId)) {
      oberflaeche();
      status("Noch nicht eingerichtet: In einstellungen.js fehlt die Client-ID der App „FP-Knopf“.", "err");
      return;
    }
    const inOutlook = !!(Office.context && Office.context.mailbox);
    const naa = !!(inOutlook && Office.context.requirements && Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"));
    const msalConfig = {
      auth: {
        clientId: CONFIG.clientId,
        authority: "https://login.microsoftonline.com/" + CONFIG.tenantId,
        redirectUri: window.location.origin + window.location.pathname,
      },
    };
    pca = naa
      ? await msal.createNestablePublicClientApplication(msalConfig)
      : await msal.PublicClientApplication.createPublicClientApplication(msalConfig);
    standalone = !inOutlook;
    if (standalone) {
      try {
        const rr = await pca.handleRedirectPromise();
        if (rr && rr.account) pca.setActiveAccount(rr.account);
      } catch (e) { console.warn("[FP] handleRedirectPromise:", msg(e)); }
    }

    oberflaeche();
    el("version").textContent = "FP-Knopf v" + CONFIG.version;
    if (inOutlook) {
      ausMail();
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => { status(""); ausMail(); });
      if (!Office.context.requirements.isSetSupported("Mailbox", "1.14")) {
        status("Dieses Outlook kann Mails nicht als Datei weitergeben (Mailbox 1.14 fehlt). Bitte Outlook aktualisieren.", "err");
      }
    } else {
      el("mail").textContent = "Browser-Test: keine Mail geöffnet";
    }

    const token = await stillesToken(8000);
    if (token) nachAnmeldung();
    else el("login").style.display = "block";
  } catch (e) {
    status("Startfehler: " + msg(e), "err");
  }
});

/* ---------- Anmeldung ---------- */

async function stillesToken(timeoutMs) {
  const konto = pca.getActiveAccount() || pca.getAllAccounts()[0];
  const versuch = pca.acquireTokenSilent({ scopes: CONFIG.scopes, account: konto })
    .then((r) => { pca.setActiveAccount(r.account); return r.accessToken; });
  const zeit = new Promise((res) => setTimeout(() => res(null), timeoutMs || 8000));
  try { return await Promise.race([versuch, zeit]); } catch (e) { return null; }
}

async function anmelden() {
  if (standalone) {
    await pca.acquireTokenRedirect({ scopes: CONFIG.scopes });
    return new Promise(() => {});
  }
  const r = await pca.acquireTokenPopup({ scopes: CONFIG.scopes });
  pca.setActiveAccount(r.account);
  return r.accessToken;
}

async function token() {
  const t = await stillesToken(15000);
  if (t) return t;
  el("login").style.display = "block";
  throw new Error("Anmeldung erforderlich – bitte oben auf „Bei Microsoft anmelden“ klicken.");
}

function konto() { return pca.getActiveAccount() || pca.getAllAccounts()[0] || null; }

function ich() {
  const k = konto();
  const mail = ((k && k.username) || "").toLowerCase();
  const b = CONFIG.bearbeiter.find((x) => x.mail.toLowerCase() === mail);
  if (b) return b;
  const n = (k && k.name) || "";
  const teile = n.split(/[,\s]+/).filter(Boolean);
  return { kuerzel: teile.map((t) => t[0]).join("").toUpperCase().slice(0, 3) || "?", name: n, mail };
}

async function nachAnmeldung() {
  el("login").style.display = "none";
  const mein = ich();
  const sel = el("bearbeiter");
  if (CONFIG.bearbeiter.some((b) => b.mail === mein.mail)) sel.value = mein.mail;
  try {
    await Promise.all([ladeFirmen(), ladeMappen()]);
  } catch (e) {
    status("Laden fehlgeschlagen: " + msg(e), "err");
  }
  pruefeKnopf();
}

/* ---------- SharePoint ---------- */

const pfad = (s) => encodeURIComponent(s.replace(/'/g, "''"));

async function sp(path, opts = {}) {
  for (let v = 1; v <= 4; v++) {
    const t = await token();
    const binaer = opts.body instanceof Uint8Array;
    const headers = { Authorization: "Bearer " + t, Accept: "application/json;odata=nometadata" };
    if (opts.body !== undefined) headers["Content-Type"] = binaer ? "application/octet-stream" : "application/json;odata=nometadata";
    const r = await fetch(CONFIG.spHost + path, {
      method: opts.method || "GET",
      headers,
      body: opts.body === undefined ? undefined : (binaer ? opts.body : JSON.stringify(opts.body)),
    });
    if (r.status === 429 || r.status === 503) { await warte((Number(r.headers.get("Retry-After")) || 3 * v) * 1000); continue; }
    const txt = await r.text();
    let j = null; try { j = txt ? JSON.parse(txt) : null; } catch (_) { }
    if (!r.ok) {
      const m = (j && j["odata.error"] && j["odata.error"].message && j["odata.error"].message.value) || (j && j.error && j.error.message) || txt.slice(0, 200);
      const e = new Error(r.status + " " + m); e.status = r.status; throw e;
    }
    return j;
  }
  throw new Error("SharePoint ist gerade überlastet. Bitte gleich nochmal versuchen.");
}

const liste = () => `${CONFIG.web}/_api/web/lists(guid'${CONFIG.listId}')`;

async function ladeFirmen() {
  const f = await sp(liste() + "/fields/getbyinternalnameortitle('FPFirma')?$select=Choices");
  firmen = (f && f.Choices) || [];
  el("firmen").innerHTML = firmen.map((x) => `<option value="${html(x)}"></option>`).join("");
  if (!el("firma").value) el("firma").value = firmaRaten();
}

async function ladeMappen() {
  const r = await sp(liste() + "/items?$select=Id,FileLeafRef,FileRef,FPFirma,FPFaelligkeit&$filter=FSObjType eq 1 and FPStatus eq 'Offen'&$top=5000");
  mappen = ((r && r.value) || []).map((x) => ({ id: x.Id, name: x.FileLeafRef, rel: x.FileRef, firma: x.FPFirma || "", faellig: x.FPFaelligkeit || "" }));
  zeigeMappen();
}

async function ordnerInfo(rel) {
  try {
    return await sp(`${CONFIG.web}/_api/web/GetFolderByServerRelativePath(decodedurl='${pfad(rel)}')?$select=Exists&$expand=ListItemAllFields&$select=ListItemAllFields/Id`);
  } catch (e) {
    if (e.status === 404) return null;
    throw e;
  }
}

async function felderSetzen(id, werte, neueVersion) {
  const r = await sp(liste() + `/items(${id})/ValidateUpdateListItem`, { method: "POST", body: { formValues: werte, bNewDocumentUpdate: !!neueVersion } });
  const fe = ((r && r.value) || []).filter((x) => x.HasException).map((x) => x.FieldName + ": " + x.ErrorMessage);
  if (fe.length) throw new Error(fe.join("; "));
}

/* ---------- Mail ---------- */

function ausMail() {
  const it = mailItem();
  gewaehlt = null;
  if (!it) { el("mail").textContent = "Keine Mail ausgewählt"; pruefeKnopf(); return; }
  const von = (it.from && (it.from.displayName || it.from.emailAddress)) || "";
  el("mail").textContent = (it.subject || "(ohne Betreff)") + (von ? " · " + von : "");
  el("mail").title = el("mail").textContent;
  el("name").value = mappenName(it.subject || "");
  el("ansprech").value = (it.from && it.from.displayName) || "";
  el("firma").value = firmaRaten();
  el("notiz").value = "";
  el("suche").value = "";
  zeigeMappen();
  pruefeKnopf();
}

/* Betreff ohne AW/WG/RE/FW und ohne vorangestellte Datumskürzel („261008 WG: 261007 …“) */
function mappenName(betreff) {
  let s = betreff;
  for (let i = 0; i < 6; i++) s = s.replace(/^\s*((AW|WG|RE|FW|FWD|Antw|Fw)\s*:|\d{6}(\s+|$))/i, "");
  return kuerzen(sauber(s), 80);
}

function sauber(s) {
  return (s || "").replace(/[\x00-\x1F"*:<>?\/\\|]/g, " ").replace(/\s+/g, " ").replace(/^[\s.~]+|[\s.]+$/g, "");
}

function kuerzen(s, max) {
  if (s.length <= max) return s;
  const t = s.slice(0, max);
  const i = t.lastIndexOf(" ");
  return (i > max * 0.6 ? t.slice(0, i) : t).replace(/[\s.]+$/, "");
}

function firmaRaten() {
  const it = mailItem();
  if (!it || !firmen.length) return "";
  const mail = (it.from && it.from.emailAddress) || "";
  const domain = mail.split("@")[1] || "";
  const text = " " + [(it.subject || ""), domain.replace(/\.[a-z]+$/i, "")].join(" ").toLowerCase().replace(/[^a-z0-9äöüß]+/g, " ") + " ";
  const treffer = firmen.filter((f) => !CONFIG.ohneFirma.includes(f.toLowerCase()) &&
    text.includes(" " + f.toLowerCase().replace(/[^a-z0-9äöüß]+/g, " ").trim() + " "));
  const eindeutig = [...new Set(treffer.map((f) => f.toLowerCase()))];
  return eindeutig.length === 1 ? treffer[0] : "";
}

function mailAlsDatei() {
  return new Promise((ok, nein) => {
    const it = mailItem();
    if (!it || typeof it.getAsFileAsync !== "function") { nein(new Error("Dieses Outlook kann die Mail nicht als Datei weitergeben.")); return; }
    it.getAsFileAsync((r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) ok(r.value);
      else nein(new Error("Mail ließ sich nicht lesen: " + (r.error && r.error.message)));
    });
  });
}

function zweistellig(n) { return String(n).padStart(2, "0"); }

/* gleiches Schema wie der Export: „JJJJ-MM-TT_HHMM Betreff.eml“ */
function dateiname() {
  const it = mailItem();
  const d = new Date(it.dateTimeCreated || Date.now());
  const vorn = `${d.getFullYear()}-${zweistellig(d.getMonth() + 1)}-${zweistellig(d.getDate())}_${zweistellig(d.getHours())}${zweistellig(d.getMinutes())}`;
  const betreff = kuerzen(sauber(it.subject || "") || "Ohne Betreff", 80);
  return { name: `${vorn} ${betreff}.eml`, erstellt: `${zweistellig(d.getDate())}.${zweistellig(d.getMonth() + 1)}.${d.getFullYear()} ${zweistellig(d.getHours())}:${zweistellig(d.getMinutes())}` };
}

async function mailAblegen(rel) {
  const { name, erstellt } = dateiname();
  const da = await sp(`${CONFIG.web}/_api/web/GetFolderByServerRelativePath(decodedurl='${pfad(rel)}')/Files?$select=Name`);
  if (((da && da.value) || []).some((f) => f.Name.toLowerCase() === name.toLowerCase())) return { name, schonDa: true };
  const b64 = await mailAlsDatei();
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  await sp(`${CONFIG.web}/_api/web/GetFolderByServerRelativePath(decodedurl='${pfad(rel)}')/Files/AddUsingPath(decodedurl='${pfad(name)}',overwrite=false)`, { method: "POST", body: bytes });
  // Spalte „Maildatum“: danach sortieren die Ansichten die Mails in der Mappe (wie beim Import)
  try {
    const li = await sp(`${CONFIG.web}/_api/web/GetFileByServerRelativePath(decodedurl='${pfad(rel + "/" + name)}')/ListItemAllFields?$select=Id`);
    await felderSetzen(li.Id, [{ FieldName: "FPMaildatum", FieldValue: erstellt }], true);
  } catch (e) { console.warn("[FP] Maildatum:", msg(e)); }
  return { name, schonDa: false };
}

/* ---------- Aktionen ---------- */

function heute(lang) {
  const d = new Date();
  return `${zweistellig(d.getDate())}.${zweistellig(d.getMonth() + 1)}.` + (lang ? d.getFullYear() : "");
}

function deDatum(iso) { // „2026-10-15“ → „15.10.2026“
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
  return m ? `${m[3]}.${m[2]}.${m[1]}` : "";
}

async function neueAufgabe() {
  const it = mailItem();
  const name = kuerzen(sauber(el("name").value), 100);
  if (!name) throw new Error("Bitte einen Namen eingeben.");
  const rel = CONFIG.libRel + "/" + name;
  const info = await ordnerInfo(rel);
  if (info && info.Exists) throw new Error(`Die Mappe „${name}“ gibt es schon. Unter „Zu bestehender“ auswählen oder den Namen ändern.`);

  await sp(`${CONFIG.web}/_api/web/folders/AddUsingPath(DecodedUrl='${pfad(rel)}',overwrite=false)`, { method: "POST" });
  let neu = await ordnerInfo(rel);
  if (!neu || !neu.ListItemAllFields) { await warte(2000); neu = await ordnerInfo(rel); }
  if (!neu || !neu.ListItemAllFields) throw new Error("Mappe wurde angelegt, ist aber noch nicht lesbar. Bitte in SharePoint prüfen.");

  const kz = ich().kuerzel;
  const von = (it.from && (it.from.displayName || it.from.emailAddress)) || "";
  const verlauf = [];
  const notiz = el("notiz").value.trim();
  if (notiz) verlauf.push(`${heute()} ${kz}: ${notiz}`);
  verlauf.push(`${heute(true)} ${kz}: angelegt aus Mail „${it.subject || ""}“${von ? " von " + von : ""}`);

  const werte = [{ FieldName: "ContentType", FieldValue: CONFIG.contentType }, { FieldName: "FPStatus", FieldValue: "Offen" }];
  const add = (k, v) => { if (v) werte.push({ FieldName: k, FieldValue: v }); };
  add("FPFirma", el("firma").value.trim());
  add("FPAnsprechperson", el("ansprech").value.trim());
  const b = el("bearbeiter").value;
  if (b) add("FPBearbeiter", JSON.stringify([{ Key: "i:0#.f|membership|" + b }]));
  add("FPFaelligkeit", deDatum(el("faellig").value));
  add("FPVerlauf", verlauf.join("\n"));
  await felderSetzen(neu.ListItemAllFields.Id, werte, false);

  const m = await mailAblegen(rel);
  return { name, rel, mail: m };
}

async function zuBestehender() {
  if (!gewaehlt) throw new Error("Bitte eine Aufgabe auswählen.");
  const it = mailItem();
  const m = await mailAblegen(gewaehlt.rel);
  const notiz = el("notiz").value.trim();
  if (!m.schonDa || notiz) {
    const alt = await sp(liste() + `/items(${gewaehlt.id})?$select=FPVerlauf`);
    const zeile = `${heute()} ${ich().kuerzel}: ` + (notiz ? notiz + " (Mail „" + (it.subject || "") + "“)" : "Mail „" + (it.subject || "") + "“ abgelegt");
    await felderSetzen(gewaehlt.id, [{ FieldName: "FPVerlauf", FieldValue: zeile + (alt && alt.FPVerlauf ? "\n" + alt.FPVerlauf : "") }], false);
  }
  return { name: gewaehlt.name, rel: gewaehlt.rel, mail: m };
}

async function los() {
  if (laeuft) return;
  laeuft = true; pruefeKnopf();
  status("Wird angelegt …");
  try {
    const r = modus === "neu" ? await neueAufgabe() : await zuBestehender();
    const link = `${CONFIG.spHost}${CONFIG.libRel}/Forms/Offen.aspx?id=${encodeURIComponent(r.rel)}`;
    const zur = r.mail.schonDa ? "Die Mail lag schon in der Mappe." : "Mail abgelegt.";
    el("status").className = "ok";
    el("status").innerHTML = `${modus === "neu" ? "Aufgabe angelegt" : "Zur Aufgabe hinzugefügt"}: <a href="${link}" target="_blank" rel="noopener">${html(r.name)}</a>\n${zur}`;
    if (modus === "neu") ladeMappen().catch(() => {});
  } catch (e) {
    status(msg(e), "err");
  } finally {
    laeuft = false; pruefeKnopf();
  }
}

/* ---------- Oberfläche ---------- */

function oberflaeche() {
  el("bearbeiter").innerHTML = `<option value="">– niemand –</option>` +
    CONFIG.bearbeiter.map((b) => `<option value="${html(b.mail)}">${html(b.kuerzel + " · " + b.name)}</option>`).join("");
  el("tabNeu").onclick = () => setzeModus("neu");
  el("tabAlt").onclick = () => setzeModus("alt");
  el("loginBtn").onclick = async () => {
    try { await anmelden(); await nachAnmeldung(); status(""); } catch (e) { status("Anmeldung fehlgeschlagen: " + msg(e), "err"); }
  };
  el("los").onclick = los;
  el("name").oninput = pruefeKnopf;
  el("suche").oninput = zeigeMappen;
  document.querySelectorAll(".quick button").forEach((b) => {
    b.onclick = () => {
      const d = new Date(); d.setDate(d.getDate() + Number(b.dataset.tage));
      el("faellig").value = `${d.getFullYear()}-${zweistellig(d.getMonth() + 1)}-${zweistellig(d.getDate())}`;
    };
  });
}

function setzeModus(m) {
  modus = m;
  document.body.className = m;
  el("tabNeu").classList.toggle("on", m === "neu");
  el("tabAlt").classList.toggle("on", m === "alt");
  el("los").textContent = m === "neu" ? "Aufgabe anlegen" : "Mail zur Aufgabe legen";
  status("");
  pruefeKnopf();
}

/* Vorschläge: Mappen, deren Name Wörter aus dem Betreff enthält, zuerst */
function zeigeMappen() {
  const box = el("liste");
  if (!mappen.length) { box.innerHTML = `<div class="leer">${konto() ? "Keine offenen Aufgaben gefunden." : "Bitte zuerst anmelden."}</div>`; return; }
  const q = el("suche").value.trim().toLowerCase();
  const it = mailItem();
  const woerter = ((it && it.subject) || "").toLowerCase().split(/[^a-z0-9äöüß]+/).filter((w) => w.length >= 4 && !/^\d+$/.test(w));
  const treffer = mappen
    .filter((m) => !q || (m.name + " " + m.firma).toLowerCase().includes(q))
    .map((m) => ({ m, p: woerter.filter((w) => m.name.toLowerCase().includes(w)).length }))
    .sort((a, b) => b.p - a.p || a.m.name.localeCompare(b.m.name, "de"))
    .slice(0, 60);
  box.innerHTML = treffer.length ? "" : `<div class="leer">Nichts gefunden.</div>`;
  for (const { m } of treffer) {
    const d = document.createElement("div");
    if (gewaehlt && gewaehlt.id === m.id) d.className = "on";
    const f = m.faellig ? "fällig " + new Date(m.faellig).toLocaleDateString("de-DE") : "";
    d.innerHTML = `${html(m.name)}<small>${html([m.firma, f].filter(Boolean).join(" · "))}</small>`;
    d.onclick = () => { gewaehlt = m; zeigeMappen(); pruefeKnopf(); };
    box.appendChild(d);
  }
}

function pruefeKnopf() {
  const bereit = !laeuft && !!mailItem() && !!konto() && (modus === "neu" ? !!sauber(el("name").value) : !!gewaehlt);
  el("los").disabled = !bereit;
}

function status(text, art) {
  el("status").className = art || "";
  el("status").textContent = text;
}

function html(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function msg(e) {
  return (e && (e.errorMessage || e.message)) || String(e);
}
