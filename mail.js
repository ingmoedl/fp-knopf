/* FP-Mail – liest abgelegte Mails im Browser, damit die Übersicht sie selbst anzeigen kann: ohne Herunterladen,
 * ohne Sicherheitsabfrage, auch am Handy. Zwei Formate:
 *   .msg – Outlook (Import, FP-Abgleich): Verbunddatei (Compound File) mit MAPI-Eigenschaften; der Text steckt als
 *          HTML, als komprimiertes RTF mit eingebettetem HTML (\fromhtml1) oder als reiner Text darin.
 *   .eml – FP-Knopf (Outlook gibt Add-ins nur MIME): Kopfzeilen, multipart, base64/quoted-printable.
 * FPMail.lesen(ArrayBuffer|Uint8Array, dateiname) → {
 *   betreff, von: {name, mail}, an: [...], cc: [...], datum: Date|null, html | text,
 *   anhaenge: [{ name, typ, daten: Uint8Array, cid, versteckt, mail? (eingebettete Mail, gleiche Form) }] }
 * Nur lesen, keine Abhängigkeiten. */
(function () {
  "use strict";

  const ENDE = 0xFFFFFFFA;   // ab hier: Kettenende/Sonderwerte im Compound File

  /* ---------- Hilfen ---------- */

  function decoder(cs) {
    try { return new TextDecoder(cs || "windows-1252"); } catch (e) { return new TextDecoder("windows-1252"); }
  }
  function codepage(cp) {
    cp = Number(cp) || 1252;
    if (cp === 65001) return "utf-8";
    if (cp >= 1250 && cp <= 1258) return "windows-" + cp;
    if (cp >= 28591 && cp <= 28605) return "iso-8859-" + (cp - 28590);
    return { 20127: "us-ascii", 932: "shift_jis", 936: "gbk", 949: "euc-kr", 950: "big5", 20866: "koi8-r", 21866: "koi8-u", 50220: "iso-2022-jp", 51932: "euc-jp" }[cp] || "windows-1252";
  }
  const utf16 = (b) => new TextDecoder("utf-16le").decode(b).replace(/\0+$/, "");
  const bytes = (s) => { const b = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i) & 255; return b; };
  /* Bytes → Zeichen 1:1 (für Suchen in Rohdaten; Position = Byte) */
  const roh = (b) => { let s = ""; for (let i = 0; i < b.length; i += 32768) s += String.fromCharCode.apply(null, b.subarray(i, i + 32768)); return s; };

  /* ---------- Compound File (.msg) ---------- */

  function verbund(buf) {
    const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    if (dv.getUint32(0, true) !== 0xE011CFD0 || dv.getUint32(4, true) !== 0xE11AB1A1) throw new Error("Keine Outlook-Mail (.msg)");
    const ss = 1 << dv.getUint16(30, true), ms = 1 << dv.getUint16(32, true);
    const nFat = dv.getUint32(44, true), dir0 = dv.getUint32(48, true), grenze = dv.getUint32(56, true);
    const mf0 = dv.getUint32(60, true), nMf = dv.getUint32(64, true);
    let dif = dv.getUint32(68, true);
    const nDif = dv.getUint32(72, true);
    const off = (s) => (s + 1) * ss;
    const u32 = (o) => (o + 4 <= u8.length ? dv.getUint32(o, true) : ENDE);
    const fatSek = [];
    for (let i = 0; i < 109 && fatSek.length < nFat; i++) fatSek.push(u32(76 + i * 4));
    for (let k = 0; k < nDif && dif < ENDE; k++) {
      const n = ss / 4 - 1;
      for (let i = 0; i < n && fatSek.length < nFat; i++) fatSek.push(u32(off(dif) + i * 4));
      dif = u32(off(dif) + n * 4);
    }
    const proSek = ss / 4;
    const fat = new Uint32Array(fatSek.length * proSek);
    fatSek.forEach((s, i) => { for (let j = 0; j < proSek; j++) fat[i * proSek + j] = u32(off(s) + j * 4); });
    const kette = (start, tab) => { const r = []; for (let s = start; s < ENDE && s < tab.length && r.length <= tab.length; s = tab[s]) r.push(s); return r; };
    const lang = (start, groesse) => {
      const out = new Uint8Array(groesse);
      let p = 0;
      for (const s of kette(start, fat)) { const n = Math.min(ss, groesse - p); if (n <= 0) break; out.set(u8.subarray(off(s), off(s) + n), p); p += n; }
      return out;
    };
    const eintraege = [];
    for (const s of kette(dir0, fat)) {
      for (let i = 0; i < ss / 128; i++) {
        const o = off(s) + i * 128;
        if (o + 128 > u8.length) break;
        const nl = dv.getUint16(o + 64, true);
        eintraege.push({
          name: nl > 2 ? utf16(u8.subarray(o, o + Math.min(nl, 64) - 2)) : "", typ: u8[o + 66],
          l: dv.getUint32(o + 68, true), r: dv.getUint32(o + 72, true), kind: dv.getUint32(o + 76, true),
          start: dv.getUint32(o + 116, true), groesse: dv.getUint32(o + 120, true),
        });
      }
    }
    const wurzel = eintraege[0];
    const mfb = nMf ? lang(mf0, kette(mf0, fat).length * ss) : new Uint8Array(0);
    const minifat = new Uint32Array(mfb.buffer, 0, mfb.length >> 2);
    const mini = lang(wurzel.start, wurzel.groesse);
    const kurz = (start, groesse) => {
      const out = new Uint8Array(groesse);
      let p = 0;
      for (const s of kette(start, minifat)) { const n = Math.min(ms, groesse - p); if (n <= 0) break; out.set(mini.subarray(s * ms, s * ms + n), p); p += n; }
      return out;
    };
    return {
      eintraege,
      inhalt: (e) => (e.groesse < grenze ? kurz(e.start, e.groesse) : lang(e.start, e.groesse)),
      /* Kinder eines Speichers: Name → Index (Rot-Schwarz-Baum über l/r, ab „kind“) */
      kinder(i) {
        const r = {}, offen = [eintraege[i].kind], gesehen = new Set();
        while (offen.length) {
          const j = offen.pop();
          if (j >= eintraege.length || gesehen.has(j)) continue;
          gesehen.add(j);
          const e = eintraege[j];
          r[e.name] = j;
          offen.push(e.l, e.r);
        }
        return r;
      },
    };
  }

  /* Eigenschaften eines Speichers: Hauptnachricht (Kopf 32 Byte), eingebettete Nachricht (24), Empfänger/Anhang (8) */
  function eigenschaften(c, idx, kopf) {
    const k = c.kinder(idx);
    const P = {};
    const st = k["__properties_version1.0"];
    if (st != null) {
      const b = c.inhalt(c.eintraege[st]);
      const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
      for (let o = kopf; o + 16 <= b.length; o += 16) {
        const tag = dv.getUint32(o, true), typ = tag & 0xFFFF, id = tag >>> 16;
        if (typ === 0x0003) P[id] = dv.getInt32(o + 8, true);
        else if (typ === 0x000B) P[id] = dv.getUint16(o + 8, true) !== 0;
        else if (typ === 0x0040) {
          const ms = (dv.getUint32(o + 12, true) * 4294967296 + dv.getUint32(o + 8, true)) / 10000 - 11644473600000;
          if (ms > 0) P[id] = new Date(ms);
        }
      }
    }
    for (const [name, j] of Object.entries(k)) {
      const m = /^__substg1\.0_([0-9A-F]{4})([0-9A-F]{4})$/i.exec(name);
      if (!m) continue;
      const id = parseInt(m[1], 16), typ = parseInt(m[2], 16), e = c.eintraege[j];
      if (typ === 0x001F) P[id] = utf16(c.inhalt(e));
      else if (typ === 0x001E) P[id] = { s8: c.inhalt(e) };
      else if (typ === 0x0102) P[id] = c.inhalt(e);
      else if (typ === 0x000D) P[id] = { speicher: j };
    }
    return { P, k };
  }

  function msgLesen(c, idx, kopf) {
    const { P, k } = eigenschaften(c, idx, kopf);
    const cp = codepage(P[0x3FFD] || P[0x3FDE]);
    const s = (id) => (typeof P[id] === "string" ? P[id] : P[id] && P[id].s8 ? decoder(cp).decode(P[id].s8).replace(/\0+$/, "") : "");
    const smtp = (x) => (/@/.test(x) && !/^\//.test(x) ? x : "");
    // Kopfzeilen der Übertragung als Rückfall für Absender und Datum
    const kz = s(0x007D);
    const kopfWert = (n) => { const m = new RegExp("^" + n + ":[ \\t]*(.*(?:\\r?\\n[ \\t].*)*)", "im").exec(kz); return m ? m[1].replace(/\r?\n[ \t]+/g, " ").trim() : ""; };
    const mail = {
      betreff: s(0x0037),
      von: { name: s(0x0042) || s(0x0C1A), mail: s(0x5D02) || s(0x5D01) || smtp(s(0x0065)) || smtp(s(0x0C1F)) },
      an: [], cc: [],
      datum: P[0x0039] || P[0x0E06] || (kopfWert("Date") && new Date(kopfWert("Date"))) || P[0x3007] || null,
      anhaenge: [],
    };
    if (!mail.von.mail && kz) { const v = adressen(kopfWert("From"))[0]; if (v) { mail.von.mail = v.mail; mail.von.name = mail.von.name || v.name; } }
    if (mail.datum && isNaN(mail.datum)) mail.datum = null;
    // Empfänger
    const empf = Object.keys(k).filter((n) => /^__recip_version1\.0_#/i.test(n)).sort();
    for (const n of empf) {
      const r = eigenschaften(c, k[n], 8).P;
      const rs = (id) => (typeof r[id] === "string" ? r[id] : r[id] && r[id].s8 ? decoder(cp).decode(r[id].s8).replace(/\0+$/, "") : "");
      const x = { name: rs(0x3001), mail: rs(0x39FE) || smtp(rs(0x3003)) };
      if (r[0x0C15] === 2) mail.cc.push(x); else if (r[0x0C15] !== 3) mail.an.push(x);
    }
    if (!empf.length) {
      mail.an = s(0x0E04).split(";").map((t) => ({ name: t.trim(), mail: "" })).filter((x) => x.name);
      mail.cc = s(0x0E03).split(";").map((t) => ({ name: t.trim(), mail: "" })).filter((x) => x.name);
    }
    // Text: HTML, sonst HTML aus dem RTF, sonst reiner Text
    const h = P[0x1013];
    if (typeof h === "string") mail.html = h;
    else if (h instanceof Uint8Array && h.length) mail.html = htmlDekodieren(h, P[0x3FDE] ? codepage(P[0x3FDE]) : null);
    if (!mail.html && P[0x1009] instanceof Uint8Array) {
      try { const r = rtfEntpacken(P[0x1009]); if (r) mail.html = rtfZuHtml(r); } catch (e) { }
    }
    if (!mail.html) mail.text = s(0x1000);
    // Anhänge
    const anh = Object.keys(k).filter((n) => /^__attach_version1\.0_#/i.test(n)).sort();
    for (const n of anh) {
      const a = eigenschaften(c, k[n], 8);
      const A = a.P;
      const as = (id) => (typeof A[id] === "string" ? A[id] : A[id] && A[id].s8 ? decoder(cp).decode(A[id].s8).replace(/\0+$/, "") : "");
      const x = {
        name: as(0x3707) || as(0x3704) || as(0x3001) || "Anhang",
        typ: (as(0x370E) || "").toLowerCase(), cid: as(0x3712).replace(/^<|>$/g, ""),
        versteckt: A[0x7FFE] === true,
      };
      const d = A[0x3701];
      if (d instanceof Uint8Array) x.daten = d;
      else if (d && d.speicher != null) {
        try { x.mail = msgLesen(c, d.speicher, 24); } catch (e) { continue; }
        if (!/\.(msg|eml)$/i.test(x.name)) x.name = (x.mail.betreff || x.name).replace(/[\\/:*?"<>|]/g, " ") + ".msg";
      } else continue;
      mail.anhaenge.push(x);
    }
    return mail;
  }

  /* HTML-Bytes: Zeichensatz aus <meta charset>, sonst Codepage der Mail, sonst UTF-8 wenn gültig */
  function htmlDekodieren(b, cs) {
    const kopf = roh(b.subarray(0, 4096));
    const m = /<meta[^>]+charset\s*=\s*["']?([\w-]+)/i.exec(kopf);
    if (m) cs = m[1];
    if (!cs) { try { return new TextDecoder("utf-8", { fatal: true }).decode(b); } catch (e) { cs = "windows-1252"; } }
    return decoder(cs).decode(b);
  }

  /* ---------- komprimiertes RTF (MS-OXRTFCP) ---------- */

  const RTF_VORLAGE = "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx";

  function rtfEntpacken(b) {
    const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    const comp = dv.getUint32(0, true), groesse = dv.getUint32(4, true), art = dv.getUint32(8, true);
    if (art === 0x414C454D) return b.subarray(16, 16 + groesse);       // „MELA“: nicht komprimiert
    if (art !== 0x75465A4C) return null;                                // „LZFu“
    const dict = new Uint8Array(4096);
    for (let i = 0; i < RTF_VORLAGE.length; i++) dict[i] = RTF_VORLAGE.charCodeAt(i);
    let w = RTF_VORLAGE.length;
    const out = new Uint8Array(groesse);
    let o = 0, p = 16;
    const ende = Math.min(b.length, comp + 4);
    while (p < ende && o < groesse) {
      const ctl = b[p++];
      for (let bit = 0; bit < 8 && p < ende && o < groesse; bit++) {
        if (ctl & (1 << bit)) {
          const ref = (b[p] << 8) | b[p + 1];
          p += 2;
          const von = ref >> 4, n = (ref & 15) + 2;
          if (von === w) return out.subarray(0, o);
          for (let i = 0; i < n && o < groesse; i++) { const z = dict[(von + i) & 4095]; out[o++] = z; dict[w] = z; w = (w + 1) & 4095; }
        } else {
          const z = b[p++];
          out[o++] = z; dict[w] = z; w = (w + 1) & 4095;
        }
      }
    }
    return out.subarray(0, o);
  }

  /* HTML aus RTF zurückgewinnen (MS-OXRTFEX): Inhalt der {\*\htmltag}-Gruppen und Text außerhalb von \htmlrtf */
  const RTF_UEBERSPRINGEN = new Set(["fonttbl", "colortbl", "stylesheet", "info", "pict", "object", "header", "footer", "headerl", "headerr", "footerl", "footerr",
    "listtable", "listoverridetable", "rsidtbl", "generator", "themedata", "colorschememapping", "latentstyles", "datastore", "xmlnstbl", "filetbl", "revtbl", "fldinst", "pntext", "pntxta", "pntxtb"]);
  const RTF_ZEICHEN = { par: "\r\n", line: "\r\n", tab: "\t", lquote: 0x2018, rquote: 0x2019, ldblquote: 0x201C, rdblquote: 0x201D, bullet: 0x2022,
    endash: 0x2013, emdash: 0x2014, enspace: 0x2002, emspace: 0x2003, qmspace: 0x2005 };

  function rtfZuHtml(b) {
    const anfang = roh(b.subarray(0, 2048));
    if (!/\\fromhtml1/.test(anfang)) return null;
    const dec = decoder(codepage((/\\ansicpg(\d+)/.exec(anfang) || [])[1]));
    const teile = [];
    let puffer = [];
    const leeren = () => { if (puffer.length) { teile.push(dec.decode(new Uint8Array(puffer))); puffer = []; } };
    let st = { weg: false, aus: false, uc: 1 };
    const stapel = [];
    let neueGruppe = false, stern = false, auslassen = 0;
    const istBuchstabe = (z) => (z >= 65 && z <= 90) || (z >= 97 && z <= 122);
    const istZiffer = (z) => z >= 48 && z <= 57;
    let i = 0;
    while (i < b.length) {
      const z = b[i];
      if (z === 123) { stapel.push(st); st = Object.assign({}, st); neueGruppe = true; stern = false; i++; continue; }         // {
      if (z === 125) { st = stapel.pop() || st; neueGruppe = false; stern = false; i++; continue; }                            // }
      if (z === 13 || z === 10) { i++; continue; }
      if (z === 92) {                                                                                                          // \
        const n = b[i + 1];
        if (n === 39) {                                                                                                        // \'hh
          const wert = parseInt(String.fromCharCode(b[i + 2], b[i + 3]), 16);
          i += 4;
          if (auslassen) { auslassen--; continue; }
          if (!st.weg && !st.aus && !isNaN(wert)) puffer.push(wert);
          continue;
        }
        if (n === 42) { stern = true; i += 2; continue; }                                                                     // \*
        if (istBuchstabe(n)) {
          let j = i + 1;
          while (j < b.length && istBuchstabe(b[j])) j++;
          const wort = roh(b.subarray(i + 1, j));
          let k = j;
          if (b[k] === 45) k++;
          while (k < b.length && istZiffer(b[k])) k++;
          const zahl = k > j && !(k === j + 1 && b[j] === 45) ? Number(roh(b.subarray(j, k))) : null;
          if (b[k] === 32) k++;
          i = k;
          if (neueGruppe) {
            neueGruppe = false;
            if (wort === "htmltag") { stern = false; continue; }
            if (stern || RTF_UEBERSPRINGEN.has(wort)) { st.weg = true; stern = false; continue; }
          }
          if (wort === "htmlrtf") { st.aus = zahl !== 0; continue; }
          if (wort === "uc") { st.uc = zahl == null ? 1 : zahl; continue; }
          if (st.weg || st.aus) { if (wort === "u") auslassen = st.uc; continue; }
          if (wort === "u" && zahl != null) { leeren(); teile.push(String.fromCharCode(zahl < 0 ? zahl + 65536 : zahl)); auslassen = st.uc; continue; }
          const zch = RTF_ZEICHEN[wort];
          if (zch) { leeren(); teile.push(typeof zch === "number" ? String.fromCharCode(zch) : zch); }
          continue;
        }
        // Steuerzeichen: \{ \} \\ \~ \- \_ und \Zeilenumbruch
        i += 2;
        neueGruppe = false;
        if (st.weg || st.aus) continue;
        if (n === 123 || n === 125 || n === 92) { if (auslassen) { auslassen--; continue; } puffer.push(n); }
        else if (n === 126) { leeren(); teile.push(String.fromCharCode(0xA0)); }
        else if (n === 95) puffer.push(45);
        else if (n === 13 || n === 10) { leeren(); teile.push("\r\n"); }
        continue;
      }
      neueGruppe = false;
      i++;
      if (auslassen) { auslassen--; continue; }
      if (!st.weg && !st.aus) puffer.push(z);
    }
    leeren();
    const html = teile.join("");
    return /<\w/.test(html) ? html : null;
  }

  /* ---------- MIME (.eml) ---------- */

  /* =?charset?B|Q?…?= in Kopfzeilen; nebeneinanderliegende kodierte Wörter ohne Leerraum dazwischen */
  function woerter(s) {
    return String(s || "").replace(/(=\?[^?]+\?[BbQq]\?[^?]*\?=)\s+(?==\?)/g, "$1").replace(/=\?([^?*]+)(?:\*[^?]*)?\?([BbQq])\?([^?]*)\?=/g, (_, cs, art, t) => {
      try {
        const b = art.toUpperCase() === "B" ? bytes(atob(t.replace(/[^A-Za-z0-9+/=]/g, ""))) : bytes(t.replace(/_/g, " ").replace(/=([0-9A-Fa-f]{2})/g, (x, h) => String.fromCharCode(parseInt(h, 16))));
        return decoder(cs).decode(b);
      } catch (e) { return t; }
    });
  }

  function adressen(s) {
    const r = [];
    let cur = "", q = false, w = 0;
    for (const ch of String(s || "")) {
      if (ch === '"') q = !q;
      else if (ch === "<" && !q) w++;
      else if (ch === ">" && !q) w--;
      if ((ch === "," || ch === ";") && !q && w <= 0) { r.push(cur); cur = ""; } else cur += ch;
    }
    r.push(cur);
    // erst trennen, dann Namen dekodieren: „=?utf-8?Q?M=C3=B6dl=2C_Samuel?=“ enthält ein Komma
    return r.map((t) => t.trim()).filter(Boolean).map((t) => {
      const m = /^(.*?)<([^>]*)>\s*$/.exec(t);
      const ohne = (x) => woerter(x.trim().replace(/^"(.*)"$/, "$1").replace(/\\(.)/g, "$1"));
      return m ? { name: ohne(m[1]), mail: m[2].trim() } : { name: "", mail: ohne(t) };
    });
  }

  /* Content-Type/-Disposition: Wert und Parameter, auch RFC 2231 (name*=utf-8''…, name*0*=…) */
  function parameter(v) {
    const t = [];
    let cur = "", q = false;
    for (const ch of String(v || "")) { if (ch === '"') q = !q; if (ch === ";" && !q) { t.push(cur); cur = ""; } else cur += ch; }
    t.push(cur);
    const wert = t.shift().trim().toLowerCase();
    const p = {}, stuecke = {};
    for (const x of t) {
      const m = /^\s*([^=\s]+)\s*=\s*(.*?)\s*$/.exec(x);
      if (!m) continue;
      const n = m[1].toLowerCase();
      let w = m[2];
      if (/^".*"$/.test(w)) w = w.slice(1, -1).replace(/\\(.)/g, "$1");
      const f = /^(.+?)\*(\d+)(\*?)$/.exec(n);
      if (f) { (stuecke[f[1]] = stuecke[f[1]] || [])[Number(f[2])] = { w, kod: !!f[3] }; continue; }
      if (n.endsWith("*")) { p[n.slice(0, -1)] = rfc2231([{ w, kod: true }]); continue; }
      if (!(n in p)) p[n] = woerter(w);
    }
    for (const [n, liste] of Object.entries(stuecke)) p[n] = rfc2231(liste.filter(Boolean));
    return { wert, p };
  }
  function rfc2231(liste) {
    let cs = "utf-8", s = "";
    liste.forEach((x, i) => {
      let w = x.w;
      if (i === 0 && x.kod) { const m = /^([^']*)'[^']*'(.*)$/.exec(w); if (m) { cs = m[1] || cs; w = m[2]; } }
      s += x.kod ? w.replace(/%([0-9A-Fa-f]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16))) : w;
    });
    try { return decoder(cs).decode(bytes(s)); } catch (e) { return s; }
  }

  function mimeTeil(b) {
    const s = roh(b.subarray(0, Math.min(b.length, 262144)));
    let k = s.indexOf("\r\n\r\n"), sep = 4;
    const k2 = s.indexOf("\n\n");
    if (k < 0 || (k2 >= 0 && k2 < k)) { k = k2; sep = 2; }
    const kopfBytes = k < 0 ? b : b.subarray(0, k);
    let kopfText;
    try { kopfText = new TextDecoder("utf-8", { fatal: true }).decode(kopfBytes); } catch (e) { kopfText = decoder("windows-1252").decode(kopfBytes); }
    const kopf = {};
    kopfText.replace(/\r?\n[ \t]+/g, " ").split(/\r?\n/).forEach((z) => {
      const m = /^([^:\s]+)\s*:\s*(.*)$/.exec(z);
      if (m && !(m[1].toLowerCase() in kopf)) kopf[m[1].toLowerCase()] = m[2];
    });
    const ct = parameter(kopf["content-type"] || "text/plain");
    const cd = parameter(kopf["content-disposition"] || "");
    return { kopf, typ: ct.wert, ct: ct.p, cd, koerper: k < 0 ? new Uint8Array(0) : b.subarray(k + sep) };
  }

  function entschluesseln(t) {
    const art = (t.kopf["content-transfer-encoding"] || "").trim().toLowerCase();
    if (art === "base64") {
      const s = roh(t.koerper).replace(/[^A-Za-z0-9+/]/g, "");
      try { return bytes(atob(s.slice(0, s.length - (s.length % 4)))); } catch (e) { return new Uint8Array(0); }
    }
    if (art === "quoted-printable") {
      const s = roh(t.koerper).replace(/=\r?\n/g, "");
      const out = new Uint8Array(s.length);
      let o = 0;
      for (let i = 0; i < s.length; i++) {
        if (s[i] === "=" && /^[0-9A-Fa-f]{2}$/.test(s.substr(i + 1, 2))) { out[o++] = parseInt(s.substr(i + 1, 2), 16); i += 2; }
        else out[o++] = s.charCodeAt(i);
      }
      return out.subarray(0, o);
    }
    return t.koerper;
  }

  function mehrteilig(b, grenze) {
    const s = roh(b), g = "--" + grenze, r = [];
    let pos = s.indexOf(g);
    if (pos < 0) return r;
    for (;;) {
      if (s.substr(pos + g.length, 2) === "--") break;
      let start = s.indexOf("\n", pos);
      if (start < 0) break;
      start++;
      const nl = s.indexOf("\n" + g, start - 1);
      if (nl < 0) { r.push(b.subarray(start)); break; }
      r.push(b.subarray(start, s[nl - 1] === "\r" ? nl - 1 : nl));
      pos = nl + 1;
    }
    return r;
  }

  function emlLesen(b) {
    const t = mimeTeil(b);
    const mail = {
      betreff: woerter(t.kopf.subject), von: adressen(t.kopf.from)[0] || { name: "", mail: "" },
      an: adressen(t.kopf.to), cc: adressen(t.kopf.cc),
      datum: t.kopf.date ? new Date(t.kopf.date.replace(/\s*\([^)]*\)\s*$/, "")) : null, anhaenge: [],
    };
    if (mail.datum && isNaN(mail.datum)) mail.datum = null;
    const text = (x) => decoder(x.ct.charset || "utf-8").decode(entschluesseln(x));
    const gehe = (x, alternativ) => {
      if (/^multipart\//.test(x.typ) && x.ct.boundary) {
        const kinder = mehrteilig(x.koerper, x.ct.boundary).map(mimeTeil);
        // bei „alternative“ die reichste Fassung (HTML) für den Text, der Rest zählt nicht als Anhang
        if (x.typ === "multipart/alternative") {
          const h = kinder.filter((k) => k.typ === "text/html" || /^multipart\//.test(k.typ)).pop() || kinder[kinder.length - 1];
          for (const k of kinder) if (k === h || (!h && k.typ === "text/plain")) gehe(k, true);
          if (!mail.html && !mail.text) { const p = kinder.find((k) => k.typ === "text/plain"); if (p) mail.text = text(p); }
          return;
        }
        kinder.forEach((k) => gehe(k, alternativ));
        return;
      }
      const name = x.cd.p.filename || x.ct.name || "";
      const anhang = x.cd.wert === "attachment" || (name && !/^text\/(html|plain)$/.test(x.typ));
      if (!anhang && x.typ === "text/html" && !mail.html) { mail.html = text(x); return; }
      if (!anhang && x.typ === "text/plain" && !mail.html && !mail.text) { mail.text = text(x); return; }
      if (!anhang && /^text\/(html|plain)$/.test(x.typ)) return;
      const daten = entschluesseln(x);
      const a = { name: name || (x.typ === "message/rfc822" ? "Mail.eml" : "Anhang"), typ: x.typ, daten, cid: (x.kopf["content-id"] || "").trim().replace(/^<|>$/g, ""), versteckt: false };
      if (x.typ === "message/rfc822") { try { a.mail = emlLesen(daten); if (a.name === "Mail.eml") a.name = (a.mail.betreff || "Mail") + ".eml"; } catch (e) { } }
      mail.anhaenge.push(a);
    };
    gehe(t, false);
    return mail;
  }

  /* ---------- Einstieg ---------- */

  function lesen(daten, name) {
    const b = daten instanceof Uint8Array ? daten : new Uint8Array(daten);
    const ole = b.length > 8 && b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0;
    if (ole) { const c = verbund(b); return msgLesen(c, 0, 32); }
    if (/\.msg$/i.test(name || "")) throw new Error("Die Datei ist keine gültige Outlook-Mail.");
    return emlLesen(b);
  }

  window.FPMail = { lesen, _intern: { verbund, eigenschaften, rtfEntpacken, rtfZuHtml, emlLesen, woerter, parameter } };
})();
