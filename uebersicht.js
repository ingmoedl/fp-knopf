/* FP-Übersicht – alle FP-Aufgaben im Blick (ing Burghausen GmbH)
 * Eine Komponente für drei Orte: eigene Seite (uebersicht.html: Browser, Edge-App, Handy) und der
 * Outlook-Seitenbereich des FP-Knopfs (schmal). Aufbau wie die Outlook-Aufgabenliste von fp@:
 * Fälligkeit › Firma, darin nach Zuweisung. Zuweisung = Spalte FPAnsprechperson (in Outlook „Abrechnungsinfo“):
 * eine Person bei der Firma oder ein internes Kürzel (SMO, HWE, LSC); interne Kürzel erscheinen farbig und
 * bestimmen die Bearbeiter (FPBearbeiter, für „Meine“ und Benachrichtigungen). Outlook-Kategorien zählen nicht.
 * Daten nur in der Bibliothek „FP-Aufgaben“ (eine Mappe = eine Aufgabe), keine eigene Kopie:
 * jede Änderung geht sofort nach SharePoint, Änderungen der Kollegen holt die Liste alle 20 Sekunden
 * und sofort beim Zurückkehren ins Fenster. Vor dem Speichern prüft sie, ob jemand anderes die Aufgabe
 * inzwischen geändert hat, und überschreibt dann nichts.
 *
 * FPUebersicht.start(element, {
 *   sp(pfad, {method, body}) → JSON,  wirft Error mit .status   (Anmeldung macht die Seite drumherum)
 *   cfg: {host, web, listId, libRel}, ich: {kuerzel, name, mail}, bearbeiter: [...] (einstellungen.js),
 *   schmal: true erzwingt die schmale Darstellung, oeffnen(url) für Links (Outlook: eigener Browser)
 * }) */
(function () {
  "use strict";

  const TAKT = 20000;
  const SEL = "Id,FileLeafRef,FileRef,FPFirma,FPAnsprechperson,FPFaelligkeit,FPStatus,FPErledigtAm,Modified," +
    "Editor/Title,Editor/EMail,FPBearbeiter/EMail,FPBearbeiter/Title,Folder/ItemCount";
  const EXP = "Editor,FPBearbeiter,Folder";
  const GRUPPEN = [["faellig-firma", "Fälligkeit › Firma"], ["faellig", "Fälligkeit"], ["firma", "Firma"], ["ansprech", "Zuweisung"]];
  const WT = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];

  const z2 = (n) => String(n).padStart(2, "0");
  const html = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const tagIso = (d) => `${d.getFullYear()}-${z2(d.getMonth() + 1)}-${z2(d.getDate())}`;      // lokales Datum
  const ausSp = (iso) => (iso ? tagIso(new Date(iso)) : "");                                     // SharePoint speichert UTC
  const deLang = (t) => (t ? `${t.slice(8, 10)}.${t.slice(5, 7)}.${t.slice(0, 4)}` : "");        // Eingabeformat der Website
  const deKurz = (t) => (t ? `${t.slice(8, 10)}.${t.slice(5, 7)}.` : "");
  const plusTage = (n) => { const d = new Date(); d.setDate(d.getDate() + n); return tagIso(d); };
  const uhr = (iso) => { const d = new Date(iso); return `${z2(d.getHours())}:${z2(d.getMinutes())}`; };
  const wann = (iso) => { if (!iso) return ""; const t = ausSp(iso); return (t === tagIso(new Date()) ? "heute" : deKurz(t)) + " " + uhr(iso); };
  const pfad = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const msg = (e) => (e && (e.errorMessage || e.message)) || String(e);
  const sauber = (s) => (s || "").replace(/[\x00-\x1F"*:<>?\/\\|]/g, " ").replace(/\s+/g, " ").replace(/^[\s.~]+|[\s.]+$/g, "");
  const vergleich = (a, b) => a.localeCompare(b, "de", { sensitivity: "base" });
  const leuteWert = (mails) => (mails.length ? JSON.stringify(mails.map((m) => ({ Key: "i:0#.f|membership|" + m }))) : "");
  const woerter = (s) => String(s || "").split(/[^A-Za-zÄÖÜäöüß]+/).filter(Boolean);
  const lies = (k) => { try { return localStorage.getItem(k); } catch (e) { return null; } };
  const schreib = (k, v) => { try { localStorage.setItem(k, v); } catch (e) { } };

  window.FPUebersicht = { start: (root, o) => { const a = new App(root, o); a.init(); return a; } };

  function App(root, o) {
    this.root = root;
    this.o = o;
    this.c = o.cfg;
    this.ich = Object.assign({ kuerzel: "?", name: "", mail: "" }, o.ich);
    this.ich.mail = (this.ich.mail || "").toLowerCase();
    this.leute = (o.bearbeiter || []).map((b) => Object.assign({}, b, { mail: b.mail.toLowerCase() }));
    this.items = new Map();
    this.punkte = new Set();          // „neu für mich“ seit dem letzten Besuch
    this.blink = new Map();           // gerade von anderen geändert: Id → hervorheben bis
    this.stamm = { firmen: [], personen: [] };
    this.suche = "";
    this.wahl = null;                 // Id in der Detailansicht, "neu" für eine neue Aufgabe
    this.zuletzt = 0;
    this.fehler = null;
    this.mitOrdner = true;            // Anzahl Mails je Mappe (Folder/ItemCount)
    this.schluessel = "fpu-" + (o.speicher || "standard");
    let st = {};
    try { st = JSON.parse(lies(this.schluessel) || "{}"); } catch (e) { }
    this.st = Object.assign({ filter: "alle", gruppe: "faellig-firma", heute: true, zu: [] }, st);
    if (!GRUPPEN.some((g) => g[0] === this.st.gruppe)) this.st.gruppe = "faellig-firma";
    this.id = "fpu" + Math.random().toString(36).slice(2, 7);
  }

  App.prototype.liste = function () { return `${this.c.web}/_api/web/lists(guid'${this.c.listId}')`; };
  /* interne Kürzel in einer Zuweisung („LSC“, „HWE/LSC“) → Mail-Adressen aus einstellungen.js */
  App.prototype.intern = function (text) { const w = woerter(text); return this.leute.filter((b) => w.includes(b.kuerzel)).map((b) => b.mail); };
  /* wer ist zuständig: interne Kürzel der Zuweisung, dazu eingetragene Bearbeiter (z. B. aus dem FP-Knopf) */
  App.prototype.wer = function (a) { return [...new Set(this.intern(a.ansprech).concat(a.leute))]; };
  App.prototype.zuweisungWerte = function (text) {
    return [{ FieldName: "FPAnsprechperson", FieldValue: text }, { FieldName: "FPBearbeiter", FieldValue: leuteWert(this.intern(text)) }];
  };
  App.prototype.merke = function () { this.st.zu = this.st.zu.slice(-300); schreib(this.schluessel, JSON.stringify(this.st)); };
  App.prototype.$ = function (sel) { return this.root.querySelector(sel); };

  /* ---------- Start ---------- */

  App.prototype.init = async function () {
    this.geruest();
    this.letzterBesuch = Number(lies("fpu-besuch-" + this.ich.mail) || 0);
    await this.laden(true);
    schreib("fpu-besuch-" + this.ich.mail, String(Date.now()));
    this.stammdaten().catch(() => { });
    this.takt = setInterval(() => { if (!document.hidden) this.laden(false); }, TAKT);
    this.wieder = () => { if (!document.hidden && Date.now() - this.zuletzt > 5000) this.laden(false); };
    document.addEventListener("visibilitychange", this.wieder);
    window.addEventListener("focus", this.wieder);
  };

  /* von außen: z. B. nachdem der FP-Knopf eine Aufgabe angelegt hat */
  App.prototype.aktualisieren = function () { return this.laden(false); };

  App.prototype.stop = function () {
    clearInterval(this.takt);
    document.removeEventListener("visibilitychange", this.wieder);
    window.removeEventListener("focus", this.wieder);
  };

  App.prototype.geruest = function () {
    const r = this.root;
    r.classList.add("fpu");
    r.innerHTML = `
      <div class="fpu-kopf">
        <span class="fpu-titel">FP-Aufgaben<small data-r="anzahl"></small></span>
        <input class="fpu-suche" type="search" placeholder="Suchen: Name, Firma, Person" aria-label="Suchen">
        <select class="fpu-gruppe" aria-label="Gruppieren nach">${GRUPPEN.map(([k, t]) => `<option value="${k}">${t}</option>`).join("")}</select>
        <label class="fpu-check" title="Heute erledigte Aufgaben durchgestrichen anzeigen"><input type="checkbox" data-r="heute"> heute erledigte</label>
        <span class="fpu-live" title="Aktualisiert sich alle 20 Sekunden. Klicken = jetzt aktualisieren"><i></i><span data-r="live">lädt …</span></span>
        <a class="fpu-archiv" data-extern target="_blank" rel="noopener" title="Erledigte und in Outlook gelöschte Aufgaben (SharePoint, Unterordner „Archiv“)"
          href="${html(`${this.c.host}${this.c.libRel}/Forms/Archiv.aspx?id=${encodeURIComponent(this.c.libRel + "/Archiv")}`)}">Archiv ↗</a>
        <div class="fpu-filter"></div>
      </div>
      <div class="fpu-haupt">
        <div class="fpu-liste" tabindex="-1"></div>
        <div class="fpu-detail"></div>
      </div>
      <div class="fpu-toast" role="status"></div>
      <datalist id="${this.id}-firmen"></datalist><datalist id="${this.id}-personen"></datalist>`;
    const suche = this.$(".fpu-suche");
    suche.oninput = () => { this.suche = suche.value.trim(); this.zeichne(); };
    const gr = this.$(".fpu-gruppe");
    gr.value = this.st.gruppe;
    gr.onchange = () => { this.st.gruppe = gr.value; this.merke(); this.zeichne(); };
    const heute = this.$('[data-r="heute"]');
    heute.checked = !!this.st.heute;
    heute.onchange = () => { this.st.heute = heute.checked; this.merke(); this.zeichne(); };
    this.$(".fpu-live").onclick = () => this.laden(false);
    this.$(".fpu-archiv").onclick = (ev) => { if (this.o.oeffnen) { ev.preventDefault(); this.o.oeffnen(ev.currentTarget.href); } };
    this.$(".fpu-filter").onclick = (ev) => {
      const b = ev.target.closest("[data-filter]");
      if (!b) return;
      this.st.filter = b.dataset.filter; this.merke(); this.zeichne();
    };
    const liste = this.$(".fpu-liste");
    liste.addEventListener("click", (ev) => this.klickListe(ev));
    liste.addEventListener("keydown", (ev) => {
      if (ev.target.matches(".fpu-neu input") && ev.key === "Enter" && ev.target.value.trim()) {
        const name = ev.target.value.trim(); ev.target.value = ""; this.zeigeDetail("neu", name);
      }
    });
    this.ziehen(liste);
    r.addEventListener("keydown", (ev) => {
      if (ev.key === "Escape") { if (this.pop) this.schliessePop(); else if (this.wahl) this.schliesseDetail(); }
      if (ev.key === "/" && !/^(INPUT|TEXTAREA|SELECT)$/.test(ev.target.tagName)) { ev.preventDefault(); suche.focus(); }
    });
    document.addEventListener("mousedown", (ev) => { if (this.pop && !this.pop.contains(ev.target)) this.schliessePop(); });
    // schmal: Outlook-Seitenbereich, Handy
    const pruefe = () => r.classList.toggle("schmal", !!this.o.schmal || r.clientWidth < 760);
    pruefe();
    if (window.ResizeObserver) new ResizeObserver(pruefe).observe(r); else window.addEventListener("resize", pruefe);
  };

  /* ---------- Daten ---------- */

  App.prototype.abfrage = function () {
    const h = new Date(); h.setHours(0, 0, 0, 0);
    const sel = this.mitOrdner ? SEL : SEL.replace(",Folder/ItemCount", "");
    const exp = this.mitOrdner ? EXP : EXP.replace(",Folder", "");
    return this.liste() + `/items?$select=${sel}&$expand=${exp}&$top=5000` +
      `&$filter=FSObjType eq 1 and (FPStatus eq 'Offen' or FPErledigtAm ge datetime'${h.toISOString().replace(/\.\d+Z$/, "Z")}')`;
  };

  App.prototype.aufgabe = function (x) {
    const p = Array.isArray(x.FPBearbeiter) ? x.FPBearbeiter : [];
    return {
      id: x.Id, name: x.FileLeafRef, rel: x.FileRef,
      firma: (x.FPFirma || "").trim(), ansprech: (x.FPAnsprechperson || "").trim(),
      leute: p.map((q) => (q.EMail || "").toLowerCase()).filter(Boolean),
      namen: p.map((q) => q.Title || q.EMail),
      faellig: ausSp(x.FPFaelligkeit), status: x.FPStatus || "Offen", erledigtAm: ausSp(x.FPErledigtAm),
      geaendert: x.Modified,
      von: { name: (x.Editor && x.Editor.Title) || "", mail: ((x.Editor && x.Editor.EMail) || "").toLowerCase() },
      mails: x.Folder && typeof x.Folder.ItemCount === "number" ? x.Folder.ItemCount : null,
    };
  };

  App.prototype.laden = async function (erstes) {
    if (this.laeuft) return;
    this.laeuft = true;
    this.live();
    try {
      let r;
      try { r = await this.o.sp(this.abfrage()); }
      catch (e) { if (this.mitOrdner && e.status === 400) { this.mitOrdner = false; r = await this.o.sp(this.abfrage()); } else throw e; }
      const neu = new Map();
      // nur die aktive Liste: Mappen direkt in der Bibliothek, nicht im Unterordner „Archiv“
      for (const x of (r && r.value) || []) if ((x.FileRef || "").slice(0, x.FileRef.lastIndexOf("/")) === this.c.libRel) neu.set(x.Id, this.aufgabe(x));
      for (const [id, a] of neu) {
        const alt = this.items.get(id);
        const vonAnderen = a.von.mail && a.von.mail !== this.ich.mail;
        const meine = this.wer(a).includes(this.ich.mail);
        if (erstes) {
          if (vonAnderen && meine && this.letzterBesuch && Date.parse(a.geaendert) > this.letzterBesuch) this.punkte.add(id);
        } else if (vonAnderen && (!alt || alt.geaendert !== a.geaendert)) {
          this.blink.set(id, Date.now() + 8000);
          if (meine) this.punkte.add(id);
        }
      }
      for (const [id, bis] of this.blink) if (bis < Date.now()) this.blink.delete(id);
      const offen = this.wahl && this.wahl !== "neu" ? this.items.get(this.wahl) : null;
      this.items = neu;
      this.zuletzt = Date.now();
      this.fehler = null;
      this.zeichne();
      if (offen) {
        // Detail offen und von jemand anderem geändert: neu aufbauen, außer man tippt gerade darin
        const jetzt = neu.get(this.wahl);
        const tippt = this.$(".fpu-detail").contains(document.activeElement) && /^(INPUT|TEXTAREA)$/.test(document.activeElement.tagName);
        if (jetzt && jetzt.geaendert !== offen.geaendert && !tippt) this.zeigeDetail(this.wahl); else this.detailKopf();
      }
    } catch (e) {
      this.fehler = msg(e);
      if (erstes) this.$(".fpu-liste").innerHTML = `<div class="fpu-leer">Laden fehlgeschlagen: ${html(this.fehler)}</div>`;
    } finally {
      this.laeuft = false;
      this.live();
    }
  };

  App.prototype.einzeln = async function (id) {
    const x = await this.o.sp(this.liste() + `/items(${id})?$select=${this.mitOrdner ? SEL : SEL.replace(",Folder/ItemCount", "")}&$expand=${this.mitOrdner ? EXP : EXP.replace(",Folder", "")}`);
    const a = this.aufgabe(x);
    this.items.set(id, a);
    this.zeichne();
    return a;
  };

  /* Firmen und Personen aus allen Mappen für die Eingabehilfen: häufige zuerst */
  App.prototype.stammdaten = async function () {
    const [f, r] = await Promise.all([
      this.o.sp(this.liste() + "/fields/getbyinternalnameortitle('FPFirma')?$select=Choices"),
      this.o.sp(this.liste() + "/items?$select=FPFirma,FPAnsprechperson&$filter=FSObjType eq 1&$top=5000"),
    ]);
    const firmen = {}, personen = {};
    ((f && f.Choices) || []).forEach((c) => { firmen[c] = firmen[c] || 0; });
    for (const x of (r && r.value) || []) {
      if (x.FPFirma) firmen[x.FPFirma.trim()] = (firmen[x.FPFirma.trim()] || 0) + 1;
      for (const p of (x.FPAnsprechperson || "").split(/[\/_;]+/).map((s) => s.trim()).filter(Boolean)) personen[p] = (personen[p] || 0) + 1;
    }
    const sortiert = (o) => Object.entries(o).sort((a, b) => b[1] - a[1] || vergleich(a[0], b[0])).map((e) => e[0]);
    // interne Kürzel zuerst
    const team = this.leute.map((b) => b.kuerzel);
    this.stamm = { firmen: sortiert(firmen), personen: team.concat(sortiert(personen).filter((p) => !team.includes(p))) };
    document.getElementById(this.id + "-firmen").innerHTML = this.stamm.firmen.map((x) => `<option value="${html(x)}">`).join("");
    document.getElementById(this.id + "-personen").innerHTML = this.stamm.personen.slice(0, 400).map((x) => `<option value="${html(x)}">`).join("");
  };

  /* Speichern mit Schutz: wurde die Aufgabe seit dem Laden von jemand anderem geändert, wird nichts überschrieben */
  App.prototype.speichern = async function (id, werte) {
    const a = this.items.get(id);
    if (a && a.geaendert) {
      const j = await this.o.sp(this.liste() + `/items(${id})?$select=Modified,Editor/Title&$expand=Editor`);
      if (j && j.Modified !== a.geaendert) {
        await this.einzeln(id);
        const e = new Error(`${(j.Editor && j.Editor.Title) || "Jemand"} hat „${a.name}“ gerade geändert. Die Anzeige ist jetzt aktuell, bitte nochmal.`);
        e.konflikt = true;
        throw e;
      }
    }
    const r = await this.o.sp(this.liste() + `/items(${id})/ValidateUpdateListItem`, { method: "POST", body: { formValues: werte, bNewDocumentUpdate: false } });
    const fe = ((r && r.value) || []).filter((x) => x.HasException).map((x) => x.FieldName + ": " + x.ErrorMessage);
    if (fe.length) throw new Error(fe.join("; "));
    return this.einzeln(id);
  };

  App.prototype.erledigen = async function (id, an) {
    const a = this.items.get(id);
    if (!a) return;
    const vorher = { status: a.status, erledigtAm: a.erledigtAm };
    a.status = an ? "Erledigt" : "Offen"; a.erledigtAm = an ? tagIso(new Date()) : "";
    this.zeichne();
    try {
      await this.speichern(id, [{ FieldName: "FPStatus", FieldValue: a.status }, { FieldName: "FPErledigtAm", FieldValue: deLang(a.erledigtAm) }]);
      this.toast((an ? "Erledigt: " : "Wieder offen: ") + a.name, "Rückgängig", () => this.erledigen(id, !an));
      if (this.wahl === id) this.detailKopf();
    } catch (e) {
      const b = this.items.get(id);
      if (b && !e.konflikt) { b.status = vorher.status; b.erledigtAm = vorher.erledigtAm; this.zeichne(); }
      this.toast(msg(e));
    }
  };

  App.prototype.setze = async function (id, werte, text) {
    try {
      await this.speichern(id, werte);
      if (text) this.toast(text);
      if (this.wahl === id) { this.detailKopf(); this.gespeichert("Gespeichert"); }
      return true;
    } catch (e) {
      if (this.wahl === id) { this.gespeichert(msg(e), true); this.zeigeDetail(id); } else this.toast(msg(e));
      return false;
    }
  };

  App.prototype.verlaufEintrag = async function (id, text) {
    const x = await this.o.sp(this.liste() + `/items(${id})?$select=FPVerlauf`);
    const d = new Date();
    const zeile = `${z2(d.getDate())}.${z2(d.getMonth() + 1)}. ${this.ich.kuerzel}: ${text}`;
    return this.setze(id, [{ FieldName: "FPVerlauf", FieldValue: zeile + (x && x.FPVerlauf ? "\n" + x.FPVerlauf : "") }]);
  };

  App.prototype.anlegen = async function (w) {
    const name = sauber(w.name).slice(0, 100).replace(/[\s.]+$/, "");
    if (!name) throw new Error("Bitte einen Namen eingeben.");
    const rel = this.c.libRel + "/" + name;
    const info = async () => {
      try { return await this.o.sp(`${this.c.web}/_api/web/GetFolderByServerRelativePath(decodedurl='${pfad(rel)}')?$select=Exists&$expand=ListItemAllFields&$select=ListItemAllFields/Id`); }
      catch (e) { if (e.status === 404) return null; throw e; }
    };
    const da = await info();
    if (da && da.Exists) throw new Error(`Die Mappe „${name}“ gibt es schon.`);
    await this.o.sp(`${this.c.web}/_api/web/folders/AddUsingPath(DecodedUrl='${pfad(rel)}',overwrite=false)`, { method: "POST" });
    let neu = await info();
    if (!neu || !neu.ListItemAllFields) { await new Promise((s) => setTimeout(s, 2000)); neu = await info(); }
    if (!neu || !neu.ListItemAllFields) throw new Error("Mappe angelegt, aber noch nicht lesbar. Gleich nochmal aktualisieren.");
    const id = neu.ListItemAllFields.Id;
    const d = new Date();
    const werte = [{ FieldName: "ContentType", FieldValue: "FP-Aufgabe" }, { FieldName: "FPStatus", FieldValue: "Offen" },
      { FieldName: "FPVerlauf", FieldValue: `${z2(d.getDate())}.${z2(d.getMonth() + 1)}.${d.getFullYear()} ${this.ich.kuerzel}: angelegt` }];
    if (w.firma) werte.push({ FieldName: "FPFirma", FieldValue: w.firma });
    if (w.ansprech) werte.push(...this.zuweisungWerte(w.ansprech));
    if (w.faellig) werte.push({ FieldName: "FPFaelligkeit", FieldValue: deLang(w.faellig) });
    const r = await this.o.sp(this.liste() + `/items(${id})/ValidateUpdateListItem`, { method: "POST", body: { formValues: werte, bNewDocumentUpdate: false } });
    const fe = ((r && r.value) || []).filter((x) => x.HasException).map((x) => x.FieldName + ": " + x.ErrorMessage);
    if (fe.length) throw new Error("Mappe angelegt, aber: " + fe.join("; "));
    await this.einzeln(id);
    return id;
  };

  /* ---------- Anzeige ---------- */

  App.prototype.kuerzel = function (mail, name) {
    const b = this.leute.find((x) => x.mail === mail);
    if (b) return b;
    const t = (name || mail || "?").split(/[,\s@.]+/).filter(Boolean);
    return { kuerzel: t.slice(0, 2).map((s) => s[0]).join("").toUpperCase(), name: name || mail, mail };
  };

  /* Zuweisung wie in Outlook; interne Kürzel als farbige Marke */
  App.prototype.zuweisung = function (text) {
    if (!text) return `<span class="fpu-ohne">ohne Zuweisung</span>`;
    const team = new Map(this.leute.map((b) => [b.kuerzel, b]));
    const teile = String(text).split(/([^A-Za-zÄÖÜäöüß]+)/);
    if (!teile.some((t) => team.has(t))) return html(text);
    return teile.map((t) => {
      const b = team.get(t);
      if (b) return `<span class="fpu-k" title="${html(b.name)}" style="background:${b.hinter};color:${b.schrift}">${html(t)}</span>`;
      return /^[^A-Za-zÄÖÜäöüß]+$/.test(t) ? " " : html(t);
    }).join("");
  };

  App.prototype.sichtbar = function () {
    const q = this.suche.toLowerCase();
    const f = this.st.filter;
    return [...this.items.values()].filter((a) => {
      if (a.status === "Erledigt" && !this.st.heute) return false;
      if (f === "meine" && !this.wer(a).includes(this.ich.mail)) return false;
      if (f === "ohne" && a.ansprech) return false;
      if (f.includes("@") && !this.wer(a).includes(f)) return false;
      if (q && !(a.name + " " + a.firma + " " + a.ansprech).toLowerCase().includes(q)) return false;
      return true;
    });
  };

  App.prototype.gruppen = function (liste) {
    const heute = tagIso(new Date());
    const g = this.st.gruppe;
    const nachPerson = (a, b) => vergleich(a.ansprech, b.ansprech) || vergleich(a.name, b.name);
    const nachDatum = (a, b) => (a.faellig || "9999").localeCompare(b.faellig || "9999") || nachPerson(a, b);
    const datumTitel = (t) => {
      if (!t) return "Ohne Fälligkeit";
      const d = new Date(t + "T12:00:00");
      const s = `${WT[d.getDay()]} ${deLang(t)}`;
      return s + (t < heute ? " · überfällig" : t === heute ? " · heute" : t === plusTage(1) ? " · morgen" : "");
    };
    const bilde = (aufgaben, schluessel, titel, sortSchluessel, extra) => {
      const m = new Map();
      for (const a of aufgaben) for (const k of [].concat(schluessel(a))) { if (!m.has(k)) m.set(k, []); m.get(k).push(a); }
      return [...m.entries()].sort((x, y) => sortSchluessel(x[0], y[0])).map(([k, l]) => Object.assign({ k, titel: titel(k), aufgaben: l }, extra(k)));
    };
    const leerZuletzt = (x, y) => (x === "") - (y === "") || vergleich(x, y);
    const datumSort = (x, y) => (x || "9999").localeCompare(y || "9999");
    const firmaTitel = (k) => k || "Ohne Firma";
    if (g === "faellig" || g === "faellig-firma") {
      return bilde(liste, (a) => a.faellig, datumTitel, datumSort, (k) => ({ art: "faellig", faellig: k, rot: !!k && k < heute })).map((gr) => {
        if (g === "faellig") { gr.aufgaben.sort(nachPerson); return gr; }
        gr.kinder = bilde(gr.aufgaben, (a) => a.firma, firmaTitel, leerZuletzt, (k) => ({ art: "firma", firma: k, faellig: gr.faellig }));
        gr.kinder.forEach((k) => k.aufgaben.sort(nachPerson));
        return gr;
      });
    }
    if (g === "firma") return bilde(liste, (a) => a.firma, firmaTitel, leerZuletzt, (k) => ({ art: "firma", firma: k })).map((gr) => { gr.aufgaben.sort(nachDatum); return gr; });
    // Zuweisung: interne Kürzel zuerst, dann Personen alphabetisch, „Ohne Zuweisung“ zuletzt
    const intern = (k) => (k && this.intern(k).length ? 0 : 1);
    return bilde(liste, (a) => a.ansprech, (k) => k || "Ohne Zuweisung", (x, y) => (x === "") - (y === "") || intern(x) - intern(y) || vergleich(x, y),
      (k) => ({ art: "ansprech", zuweisung: k })).map((gr) => { gr.aufgaben.sort(nachDatum); return gr; });
  };

  App.prototype.spalten = function () {
    const g = this.st.gruppe;
    const mitFirma = g !== "firma" && g !== "faellig-firma";
    const mitDatum = g !== "faellig" && g !== "faellig-firma";
    return { mitFirma, mitDatum, vorlage: ["26px", "minmax(0, 1fr)", mitFirma ? "110px" : null, "180px", mitDatum ? "86px" : null, this.mitOrdner ? "38px" : null].filter(Boolean).join(" ") };
  };

  App.prototype.zeile = function (a, sp, gk) {
    const heute = tagIso(new Date());
    const ueber = a.status === "Offen" && a.faellig && a.faellig < heute;
    const kl = ["fpu-zeile", ueber ? "ueber" : "", a.status === "Erledigt" ? "erledigt" : "", (this.blink.get(a.id) || 0) > Date.now() ? "neu" : "", this.wahl === a.id ? "gewaehlt" : ""].filter(Boolean).join(" ");
    const titel = a.status === "Erledigt" ? `erledigt${a.von.name ? " von " + a.von.name : ""}` : `zuletzt geändert von ${a.von.name || "?"}, ${wann(a.geaendert)}`;
    const unter = [sp.mitFirma ? a.firma : "", sp.mitDatum && a.faellig ? deKurz(a.faellig) : ""].filter(Boolean).join(" · ");
    return `<div class="${kl}" data-id="${a.id}" data-gk="${html(gk)}" draggable="true" title="${html(titel)}">
      <span class="box"><button class="fpu-box${a.status === "Erledigt" ? " an" : ""}" data-akt="erledigt" aria-label="${a.status === "Erledigt" ? "Wieder öffnen" : "Als erledigt markieren"}"></button></span>
      <div class="n">${this.punkte.has(a.id) ? `<span class="punkt" title="Neu oder geändert für dich"></span>` : ""}${html(a.name)}<div class="unter">${html(unter)}</div></div>
      ${sp.mitFirma ? `<span class="f breit">${html(a.firma)}</span>` : ""}
      <span class="p">${this.zuweisung(a.ansprech)}</span>
      ${sp.mitDatum ? `<span class="d breit">${a.faellig ? deLang(a.faellig) : ""}</span>` : ""}
      ${this.mitOrdner ? `<span class="m breit" title="Mails in der Mappe">${a.mails ? "✉ " + a.mails : ""}</span>` : ""}
    </div>`;
  };

  App.prototype.zeichne = function () {
    const liste = this.$(".fpu-liste");
    const alle = [...this.items.values()];
    const offen = alle.filter((a) => a.status === "Offen");
    this.$('[data-r="anzahl"]').textContent = `${offen.length} offen`;
    // Filter mit Anzahl (offene)
    const zahl = (fn) => offen.filter(fn).length;
    const filter = [["alle", "Alle", offen.length], ["meine", "Meine", zahl((a) => this.wer(a).includes(this.ich.mail))],
      ...this.leute.filter((b) => b.mail !== this.ich.mail).map((b) => [b.mail, b.kuerzel, zahl((a) => this.wer(a).includes(b.mail))]),
      ["ohne", "Ohne Zuweisung", zahl((a) => !a.ansprech)]];
    if (!filter.some((f) => f[0] === this.st.filter)) this.st.filter = "alle";
    this.$(".fpu-filter").innerHTML = filter.map(([k, t, n]) => `<button class="fpu-chip${this.st.filter === k ? " an" : ""}" data-filter="${html(k)}">${html(t)}<b>${n}</b></button>`).join("");

    const sp = this.spalten();
    liste.style.setProperty("--spalten", sp.vorlage);
    const sichtbar = this.sichtbar();
    const zu = new Set(this.st.zu);
    const teile = [`<div class="fpu-spalten"><span></span><span>Aufgabe</span>${sp.mitFirma ? "<span>Firma</span>" : ""}<span>Zuweisung</span>${sp.mitDatum ? "<span>Fällig</span>" : ""}${this.mitOrdner ? "<span></span>" : ""}</div>`,
      `<div class="fpu-neu"><span>+</span><input type="text" placeholder="Neue Aufgabe: Namen eintippen, Enter" aria-label="Neue Aufgabe" maxlength="100"></div>`];
    const kopf = (stufe, gr, schluessel, anzahl) => {
      const daten = (gr.art === "faellig" || gr.art === "firma") ? ` data-faellig="${html(gr.faellig == null ? "" : gr.faellig)}" data-hat-faellig="${gr.faellig != null ? 1 : 0}"` +
        (gr.art === "firma" ? ` data-firma="${html(gr.firma)}"` : "") : gr.art === "ansprech" ? ` data-zuweisung="${html(gr.zuweisung)}"` : "";
      return `<div class="fpu-g${stufe}${gr.rot ? " rot" : ""}${zu.has(schluessel) ? " fpu-zu" : ""}" data-gk="${html(schluessel)}"${daten}><span class="pfeil">▼</span>${html(gr.titel)} <small>${anzahl}</small></div>`;
    };
    for (const gr of this.gruppen(sichtbar)) {
      const k1 = this.st.gruppe + ":" + gr.k;
      teile.push(kopf(1, gr, k1, gr.aufgaben.length));
      if (zu.has(k1)) continue;
      if (gr.kinder) {
        for (const ki of gr.kinder) {
          const k2 = k1 + "|" + ki.k;
          teile.push(kopf(2, ki, k2, ki.aufgaben.length));
          if (!zu.has(k2)) teile.push(...ki.aufgaben.map((a) => this.zeile(a, sp, k2)));
        }
      } else teile.push(...gr.aufgaben.map((a) => this.zeile(a, sp, k1)));
    }
    if (!sichtbar.length) teile.push(`<div class="fpu-leer">${this.items.size ? "Keine Aufgaben für diese Auswahl." : "Keine offenen Aufgaben."}</div>`);
    // Eingabe in der „Neue Aufgabe“-Zeile und Bildlauf überstehen das Neuzeichnen
    const alt = liste.querySelector(".fpu-neu input");
    const merk = alt ? { wert: alt.value, fokus: document.activeElement === alt } : null;
    const oben = liste.scrollTop;
    liste.innerHTML = teile.join("");
    if (merk) { const n = liste.querySelector(".fpu-neu input"); n.value = merk.wert; if (merk.fokus) n.focus(); }
    liste.scrollTop = oben;
  };

  App.prototype.live = function () {
    const el = this.$(".fpu-live");
    if (!el) return;
    el.classList.toggle("laedt", !!this.laeuft);
    el.classList.toggle("fehler", !!this.fehler);
    const z = this.zuletzt ? uhr(new Date(this.zuletzt).toISOString()) : "";
    this.$('[data-r="live"]').textContent = this.laeuft ? "aktualisiert …" : this.fehler ? `offline${z ? " · Stand " + z : ""}` : `live · ${z}`;
    el.title = this.fehler ? "Letzter Versuch: " + this.fehler : "Aktualisiert sich alle 20 Sekunden. Klicken = jetzt aktualisieren";
  };

  App.prototype.toast = function (text, knopf, fn) {
    const t = this.$(".fpu-toast");
    t.innerHTML = `<span>${html(text)}</span>` + (knopf ? `<button>${html(knopf)}</button>` : "");
    t.classList.add("da");
    if (knopf) t.querySelector("button").onclick = () => { t.classList.remove("da"); fn(); };
    clearTimeout(this.toastZeit);
    this.toastZeit = setTimeout(() => t.classList.remove("da"), knopf ? 7000 : 5000);
  };

  /* ---------- Bedienung in der Liste ---------- */

  App.prototype.klickListe = function (ev) {
    const g = ev.target.closest(".fpu-g1, .fpu-g2");
    if (g) {
      const k = g.dataset.gk;
      this.st.zu = this.st.zu.includes(k) ? this.st.zu.filter((x) => x !== k) : this.st.zu.concat(k);
      this.merke(); this.zeichne();
      return;
    }
    const z = ev.target.closest(".fpu-zeile");
    if (!z) return;
    const id = Number(z.dataset.id);
    const akt = ev.target.closest("[data-akt]");
    if (akt && akt.dataset.akt === "erledigt") { ev.stopPropagation(); const a = this.items.get(id); if (a) this.erledigen(id, a.status !== "Erledigt"); return; }
    this.zeigeDetail(id);
  };

  App.prototype.schliessePop = function () { if (this.pop) { this.pop.remove(); this.pop = null; } };

  /* Zeile auf eine Datums-, Firmen- oder Zuweisungsgruppe ziehen = neue Fälligkeit, Firma bzw. Zuweisung (nur breite Ansicht) */
  App.prototype.ziehen = function (liste) {
    let id = null;
    const ziel = (ev) => { const g = ev.target.closest(".fpu-g1[data-hat-faellig='1'], .fpu-g2[data-hat-faellig='1'], .fpu-g1[data-firma], .fpu-g2[data-firma], .fpu-g1[data-zuweisung]"); return g; };
    liste.addEventListener("dragstart", (ev) => {
      const z = ev.target.closest(".fpu-zeile");
      if (!z || this.root.classList.contains("schmal")) { ev.preventDefault(); return; }
      id = Number(z.dataset.id);
      ev.dataTransfer.effectAllowed = "move";
      ev.dataTransfer.setData("text/plain", String(id));
    });
    liste.addEventListener("dragover", (ev) => { const g = id && ziel(ev); if (g) { ev.preventDefault(); liste.querySelectorAll(".fpu-ziel").forEach((x) => x.classList.remove("fpu-ziel")); g.classList.add("fpu-ziel"); } });
    liste.addEventListener("dragleave", (ev) => { const g = ziel(ev); if (g && !g.contains(ev.relatedTarget)) g.classList.remove("fpu-ziel"); });
    liste.addEventListener("dragend", () => { id = null; liste.querySelectorAll(".fpu-ziel").forEach((x) => x.classList.remove("fpu-ziel")); });
    liste.addEventListener("drop", (ev) => {
      const g = ziel(ev);
      if (!g || !id) return;
      ev.preventDefault();
      const a = this.items.get(id);
      const werte = [], text = [];
      if (g.dataset.hatFaellig === "1" && g.dataset.faellig !== a.faellig) {
        werte.push({ FieldName: "FPFaelligkeit", FieldValue: deLang(g.dataset.faellig) });
        text.push(g.dataset.faellig ? "fällig " + deLang(g.dataset.faellig) : "ohne Fälligkeit");
      }
      if (g.dataset.firma !== undefined && g.dataset.firma !== a.firma) { werte.push({ FieldName: "FPFirma", FieldValue: g.dataset.firma }); text.push("Firma " + (g.dataset.firma || "leer")); }
      if (g.dataset.zuweisung !== undefined && g.dataset.zuweisung !== a.ansprech) { werte.push(...this.zuweisungWerte(g.dataset.zuweisung)); text.push("Zuweisung " + (g.dataset.zuweisung || "leer")); }
      id = null;
      if (werte.length) this.setze(a.id, werte, `${a.name}: ${text.join(", ")}`);
    });
  };

  /* ---------- Detail ---------- */

  App.prototype.schliesseDetail = function () {
    this.wahl = null;
    this.root.classList.remove("mit-detail");
    this.zeichne();
  };

  App.prototype.gespeichert = function (text, fehler) {
    const g = this.$(".fpu-gespeichert");
    if (!g) return;
    g.textContent = text; g.classList.toggle("fehler", !!fehler);
    clearTimeout(this.gespZeit);
    if (!fehler) this.gespZeit = setTimeout(() => { g.textContent = ""; }, 2500);
  };

  App.prototype.detailKopf = function () {
    const a = this.items.get(this.wahl);
    const i = this.$(".fpu-info");
    if (!a || !i) return;
    const heute = tagIso(new Date());
    const zust = a.status === "Erledigt" ? `Erledigt${a.erledigtAm ? " am " + deLang(a.erledigtAm) : ""}` : a.faellig && a.faellig < heute ? `Überfällig seit ${deLang(a.faellig)}` : "Offen";
    i.innerHTML = `<b style="color:${a.status === "Erledigt" ? "var(--ok)" : a.faellig && a.faellig < heute ? "var(--red)" : "inherit"}">${html(zust)}</b> · zuletzt geändert von ${html(a.von.name || "?")}, ${html(wann(a.geaendert))}`;
    const k = this.$('[data-d="erledigt"]');
    if (k) k.textContent = a.status === "Erledigt" ? "Wieder öffnen" : "✓ Erledigt";
  };

  App.prototype.zeigeDetail = async function (id, neuName) {
    this.schliessePop();
    const neu = id === "neu";
    const a = neu ? { name: neuName || "", firma: "", ansprech: "", leute: [], faellig: "", status: "Offen" } : this.items.get(id);
    if (!a) return;
    this.wahl = id;
    this.punkte.delete(id);
    this.root.classList.add("mit-detail");
    this.zeichne();
    const d = this.$(".fpu-detail");
    // Ansicht „Offen“ zeigt die Mails in der Mappe nach Maildatum, neueste oben
    const ordnerUrl = neu ? "" : `${this.c.host}${this.c.libRel}/Forms/Offen.aspx?id=${encodeURIComponent(a.rel)}`;
    d.innerHTML = `
      <div class="fpu-dkopf"><input data-d="name" value="${html(a.name)}" maxlength="100" aria-label="Name der Aufgabe" placeholder="Name der Aufgabe"><button class="fpu-x" data-d="zu" aria-label="Schließen">×</button></div>
      <div class="fpu-info">${neu ? "Neue Aufgabe" : ""}</div>
      ${neu ? "" : `<div class="fpu-zeilen"><button class="fpu-knopf" data-d="erledigt"></button><a class="fpu-link" data-extern href="${html(ordnerUrl)}" target="_blank" rel="noopener">Mappe in SharePoint ↗</a></div>`}
      <div class="fpu-zwei">
        <div><label>Firma</label><input type="text" data-d="firma" list="${this.id}-firmen" value="${html(a.firma)}" autocomplete="off"></div>
        <div><label>Zuweisung</label><input type="text" data-d="ansprech" list="${this.id}-personen" value="${html(a.ansprech)}" maxlength="255" autocomplete="off" placeholder="Person oder Kürzel"></div>
      </div>
      <div class="fpu-wahl" data-d="kz" title="Intern zuweisen">${this.leute.map((b) => `<button type="button" data-kz="${html(b.kuerzel)}" class="${woerter(a.ansprech).includes(b.kuerzel) ? "an" : ""}" style="background:${b.hinter};color:${b.schrift}" title="${html(b.name)}">${html(b.kuerzel)}</button>`).join("")}</div>
      <label>Fälligkeit</label>
      <input type="date" data-d="faellig" value="${a.faellig || ""}">
      <div class="fpu-schnell" data-d="schnell"><button data-t="0">heute</button><button data-t="1">morgen</button><button data-t="3">+3 Tage</button><button data-t="7">+1 Woche</button><button data-t="14">+2 Wochen</button><button data-t="">ohne</button></div>
      ${neu ? `<label>Verlauf (optional)</label><textarea data-d="notiz" rows="2"></textarea>
        <div class="fpu-zeilen"><button class="fpu-knopf haupt" data-d="anlegen">Aufgabe anlegen</button><button class="fpu-knopf" data-d="zu">Abbrechen</button></div>` :
        `<label>Verlauf</label><div style="display:flex;gap:6px"><input type="text" data-d="notiz" placeholder="Neuer Eintrag, Enter" style="flex:1"><button class="fpu-knopf" data-d="eintragen">Eintragen</button></div>
        <div class="fpu-verlauf" data-d="verlauf">lädt …</div>
        <label data-d="mailtitel">Mails</label><div class="fpu-mails" data-d="mails"><div class="leer">lädt …</div></div>`}
      <div class="fpu-gespeichert"></div>`;
    const q = (n) => d.querySelector(`[data-d="${n}"]`);
    d.querySelectorAll('[data-d="zu"]').forEach((b) => { b.onclick = () => this.schliesseDetail(); });
    d.querySelectorAll("[data-extern]").forEach((x) => { x.onclick = (ev) => { if (this.o.oeffnen) { ev.preventDefault(); this.o.oeffnen(x.href); } }; });
    if (!neu) this.detailKopf();

    // Zuweisung: Freitext (Person bei der Firma) oder interne Kürzel; speichert Bearbeiter gleich mit
    const zfeld = q("ansprech");
    const team = this.leute.map((b) => b.kuerzel);
    const zSpeichern = () => {
      const v = zfeld.value.trim();
      q("kz").querySelectorAll("button").forEach((b) => b.classList.toggle("an", woerter(v).includes(b.dataset.kz)));
      const x = this.items.get(id);
      if (!neu && x && v !== x.ansprech) this.setze(id, this.zuweisungWerte(v));
    };
    zfeld.onchange = zSpeichern;
    zfeld.onkeydown = (ev) => { if (ev.key === "Enter") zfeld.blur(); };
    q("kz").onclick = (ev) => {
      const b = ev.target.closest("button[data-kz]");
      if (!b) return;
      const w = woerter(zfeld.value), kz = b.dataset.kz;
      const nurIntern = w.length && w.every((t) => team.includes(t));
      zfeld.value = w.includes(kz) ? w.filter((t) => t !== kz).join("/") : nurIntern ? w.concat(kz).join("/") : kz;
      zSpeichern();
    };
    const datum = q("faellig");
    const datumSetzen = () => { if (!neu && datum.value !== (this.items.get(id) || {}).faellig) this.setze(id, [{ FieldName: "FPFaelligkeit", FieldValue: deLang(datum.value) }]); };
    datum.onchange = datumSetzen;
    q("schnell").onclick = (ev) => {
      const b = ev.target.closest("button[data-t]");
      if (!b) return;
      datum.value = b.dataset.t === "" ? "" : plusTage(Number(b.dataset.t));
      datumSetzen();
    };
    if (neu) {
      q("name").focus();
      q("anlegen").onclick = async () => {
        const knopf = q("anlegen");
        knopf.disabled = true; this.gespeichert("Wird angelegt …");
        try {
          const nid = await this.anlegen({ name: q("name").value, firma: q("firma").value.trim(), ansprech: q("ansprech").value.trim(), faellig: datum.value });
          const notiz = q("notiz").value.trim();
          if (notiz) await this.verlaufEintrag(nid, notiz);
          this.toast("Angelegt: " + this.items.get(nid).name);
          this.zeigeDetail(nid);
        } catch (e) { knopf.disabled = false; this.gespeichert(msg(e), true); }
      };
      return;
    }
    q("erledigt").onclick = () => { const x = this.items.get(id); if (x) this.erledigen(id, x.status !== "Erledigt"); };
    const textFeld = (n, feld, wert) => {
      const el = q(n);
      el.onchange = () => {
        const x = this.items.get(id);
        const v = n === "name" ? sauber(el.value).slice(0, 100) : el.value.trim();
        if (!x || v === wert(x)) return;
        if (n === "name" && !v) { el.value = x.name; return; }
        this.setze(id, [{ FieldName: feld, FieldValue: v }]);
      };
      el.onkeydown = (ev) => { if (ev.key === "Enter") el.blur(); };
    };
    textFeld("name", "FileLeafRef", (x) => x.name);
    textFeld("firma", "FPFirma", (x) => x.firma);
    const eintragen = async () => {
      const n = q("notiz"); const t = n.value.trim();
      if (!t) return;
      n.disabled = true;
      if (await this.verlaufEintrag(id, t)) { n.value = ""; this.ladeVerlauf(id); }
      n.disabled = false; n.focus();
    };
    q("eintragen").onclick = eintragen;
    q("notiz").onkeydown = (ev) => { if (ev.key === "Enter") eintragen(); };
    this.ladeVerlauf(id);
    this.ladeMails(id);
  };

  App.prototype.ladeVerlauf = async function (id) {
    const el = this.$('[data-d="verlauf"]');
    try {
      const x = await this.o.sp(this.liste() + `/items(${id})?$select=FPVerlauf`);
      if (this.wahl === id && el.isConnected) el.textContent = (x && x.FPVerlauf) || "Noch kein Eintrag.";
    } catch (e) { if (el.isConnected) el.textContent = "Verlauf ließ sich nicht laden: " + msg(e); }
  };

  App.prototype.ladeMails = async function (id) {
    const a = this.items.get(id);
    const el = this.$('[data-d="mails"]');
    try {
      const f = await this.o.sp(`${this.c.web}/_api/web/GetFolderByServerRelativePath(decodedurl='${pfad(a.rel)}')/Files?$select=Name,ServerRelativeUrl&$top=500`);
      if (this.wahl !== id || !el.isConnected) return;
      const dateien = ((f && f.value) || []).sort((x, y) => y.Name.localeCompare(x.Name));
      this.$('[data-d="mailtitel"]').textContent = `Mails und Dateien (${dateien.length})`;
      el.innerHTML = dateien.length ? dateien.map((x) => {
        const m = /^(\d{4})-(\d{2})-(\d{2})_(\d{2})(\d{2}) (.*)$/.exec(x.Name);
        const zeit = m ? `${m[3]}.${m[2]}.${m[1].slice(2)} ${m[4]}:${m[5]}` : "";
        return `<a data-extern href="${html(this.c.host + x.ServerRelativeUrl.split("/").map(encodeURIComponent).join("/"))}" target="_blank" rel="noopener" title="${html(x.Name)}">${zeit ? `<small>${zeit}</small>` : ""}${html(m ? m[6] : x.Name)}</a>`;
      }).join("") : `<div class="leer">Keine Mails in der Mappe.</div>`;
      el.querySelectorAll("[data-extern]").forEach((x) => { x.onclick = (ev) => { if (this.o.oeffnen) { ev.preventDefault(); this.o.oeffnen(x.href); } }; });
    } catch (e) { if (el.isConnected) el.innerHTML = `<div class="leer">Mails ließen sich nicht laden: ${html(msg(e))}</div>`; }
  };
})();
