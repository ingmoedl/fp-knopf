/*
  FP-Karte – Kartenansicht
  ------------------------
  Zeichnet den Werksplan auf eine Zeichenfläche, ohne fremde Bibliothek (die Seite lädt Code nur aus dem
  eigenen Ordner). Schräge Parallelansicht: Der Boden ist eine einfache lineare Abbildung, Höhen gehen
  senkrecht nach oben. Baukörper mit Einträgen treten in Bernstein hervor, alle anderen bleiben flach.
  An jedem Gebäude steht nur ein Nummernschild. Beim Überfahren erscheint kein Textfenster.
  Hat ein Gebäude überfällige Aufgaben, trägt das Schild deren Zahl in einem roten Zeichen.

  Einheiten der Geometrie: Dezimeter, x nach Osten, y nach Norden.

  FPKarteAnsicht.erzeugen(el, karte, gebaeude, { beiWahl(nr), beiLeer(), beiAnsicht({ az, draufsicht }), ohneBewegung })
    el         Behälter mit <canvas> und <div class="schilder">
    karte      Geometrie (flaechen, linien, neubau, bestand_ids, texte)
    gebaeude   Map nr → { nr, name, hoehe (m), lage: { art: polygon | pin | unbekannt, polys, p, sicher } }
  Rückgabe: { zustand(Map nr → { relevant, dringend, ueberfaellig }), waehle(nr), zeige(nr), rand({l, r, o, u}),
              zoom(faktor), ansichtWechseln(), start(), hat(nr), stand(), groesse(), ordnen() }
  Gezeichnet wird auf Anforderung. Bleibt der Bildtakt des Browsers aus (Seite im Hintergrund), zeichnet ein Zeitgeber.
  Nach jeder Größenänderung, jedem neuen Zustand und jeder Auswahl wird sofort neu gerechnet.
*/
(function (wurzel) {
  'use strict';

  const klemme = (v, a, b) => Math.min(b, Math.max(a, v));
  /* Zahl der überfälligen Aufgaben am Schild. Ohne Zahl bleibt es beim Punkt. */
  const ueberText = (n) => (n > 99 ? '99+' : n > 0 ? String(n) : '');
  const weich = (p) => (p < 0.5 ? 4 * p * p * p : 1 - Math.pow(-2 * p + 2, 3) / 2);
  const SCHRAEG = 0.86;            // Neigung aus der Senkrechten, im Bogenmaß
  const NEIG_MAX = 1.15;
  const K_NAH = 0.16;              // Bildpunkte je Dezimeter, ab hier stehen die Höhen ohne Überhöhung
  const K_MAX = 1.0;
  const K_ALLE = 0.085;            // ab diesem Maßstab stehen auch die Schilder der Gebäude ohne Einträge
  const ABSTAND = 3;               // freier Rand um jedes Schild, in Bildpunkten

  function erzeugen(el, karte, gebaeude, o) {
    o = o || {};
    const canvas = el.querySelector('canvas');
    const ctx = canvas.getContext('2d', { alpha: false });
    const schilderEl = el.querySelector('.schilder');
    const stil = getComputedStyle(el);
    const farbe = (n, ersatz) => stil.getPropertyValue(n).trim() || ersatz;
    const reduziert = !!o.ohneBewegung || (typeof matchMedia === 'function' && matchMedia('(prefers-reduced-motion: reduce)').matches);
    const F = {
      grund: farbe('--grund', '#141619'), nachbar: farbe('--karte-nachbar', '#1d2024'), gruen: farbe('--karte-gruen', '#1c2420'),
      befestigt: farbe('--karte-befestigt', '#22262b'), dunkel: farbe('--karte-dunkel', '#272b31'), bestand: farbe('--karte-bestand', '#2c3036'),
      fein: farbe('--karte-fein', '#262a2f'), mittel: farbe('--karte-mittel', '#343941'), kante: farbe('--karte-kante', '#3e444c'),
      gruenlinie: farbe('--karte-gruenlinie', '#36483c'), zaun: farbe('--karte-zaun', '#6f8f74'), ring: farbe('--karte-ring', '#b59d63'),
      bauFlach: farbe('--bau-flach', '#353a41'), bau: farbe('--bau', '#8c939c'), stahl: farbe('--stahl', '#c6d0da'),
      text: farbe('--karte-text', '#59606a'), schrift: farbe('--font-text', 'sans-serif'),
    };
    const BERNSTEIN = rgb(farbe('--bernstein', '#f1a23b'));
    const TIEF = [38, 30, 22];

    function rgb(hex) { const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex); return m ? [parseInt(m[1], 16), parseInt(m[2], 16), parseInt(m[3], 16)] : [241, 162, 59]; }
    const misch = (a, b, t) => `rgb(${Math.round(a[0] + (b[0] - a[0]) * t)},${Math.round(a[1] + (b[1] - a[1]) * t)},${Math.round(a[2] + (b[2] - a[2]) * t)})`;

    /* ---------- Geometrie vorbereiten ---------- */
    const fl = karte.flaechen || {}, li = karte.linien || {};
    const pfadLinien = (listen) => {
      const p = new Path2D();
      for (const L of listen || []) { if (L.length < 4) continue; p.moveTo(L[0], L[1]); for (let i = 2; i < L.length - 1; i += 2) p.lineTo(L[i], L[i + 1]); }
      return p;
    };
    const pfadFlaechen = (polys) => {
      const p = new Path2D();
      for (const poly of polys || []) for (const r of poly) { if (r.length < 6) continue; p.moveTo(r[0], r[1]); for (let i = 2; i < r.length - 1; i += 2) p.lineTo(r[i], r[i + 1]); p.closePath(); }
      return p;
    };
    // Bekannte Lagen in fester Reihenfolge. Jede weitere Fläche gilt als Nachbargelände und liegt zuunterst.
    const FLAECHEN_BEKANNT = { gruen: F.gruen, befestigt: F.befestigt, dunkel: F.dunkel };
    const flaechen = [];
    for (const n of Object.keys(fl)) if (n !== 'bestand' && !FLAECHEN_BEKANNT[n]) flaechen.push({ pfad: pfadFlaechen(fl[n]), farbe: F.nachbar });
    for (const n of Object.keys(FLAECHEN_BEKANNT)) if (fl[n]) flaechen.push({ pfad: pfadFlaechen(fl[n]), farbe: FLAECHEN_BEKANNT[n] });
    const LINIEN_BEKANNT = { fein: [F.fein, 1], mittel: [F.mittel, 1], kanten: [F.kante, 1], gruenlinie: [F.gruenlinie, 1], werkzaun: [F.zaun, 1.2] };
    const linien = [];
    for (const n of Object.keys(LINIEN_BEKANNT)) if (li[n]) linien.push({ name: n, pfad: pfadLinien(li[n]), farbe: LINIEN_BEKANNT[n][0], breite: LINIEN_BEKANNT[n][1] });
    const bestand = pfadFlaechen(fl.bestand);
    const ring = pfadLinien(li.ringstrasse);
    const texte = (karte.texte || []).filter((t) => t && Array.isArray(t.p) && ['strasse', 'gleis', 'bereich'].includes(t.art))
      .map((t) => ({ t: String(t.t || '').replace(/str\.$/i, 'straße'), art: t.art, x: t.p[0], y: t.p[1] }));

    // Ausdehnung des Werks: der Zaun, sonst alle Linien
    const grenze = { x0: Infinity, y0: Infinity, x1: -Infinity, y1: -Infinity };
    const dehne = (listen) => { for (const L of listen || []) for (let i = 0; i < L.length - 1; i += 2) { grenze.x0 = Math.min(grenze.x0, L[i]); grenze.x1 = Math.max(grenze.x1, L[i]); grenze.y0 = Math.min(grenze.y0, L[i + 1]); grenze.y1 = Math.max(grenze.y1, L[i + 1]); } };
    dehne(li.werkzaun);
    if (!isFinite(grenze.x0)) for (const n of Object.keys(li)) dehne(li[n]);
    if (!isFinite(grenze.x0)) Object.assign(grenze, { x0: -1000, y0: -1000, x1: 1000, y1: 1000 });

    function flaecheUndMitte(r) {
      let a = 0, cx = 0, cy = 0; const n = r.length >> 1;
      for (let i = 0; i < n; i++) {
        const x0 = r[2 * i], y0 = r[2 * i + 1], j = (i + 1) % n, x1 = r[2 * j], y1 = r[2 * j + 1], k = x0 * y1 - x1 * y0;
        a += k; cx += (x0 + x1) * k; cy += (y0 + y1) * k;
      }
      a /= 2;
      return Math.abs(a) > 1e-6 ? { a, x: cx / (6 * a), y: cy / (6 * a) } : { a: 0, x: r[0], y: r[1] };
    }

    const POLY = new Map();
    for (const n of karte.neubau || []) if (n && n.id) POLY.set(n.id, n.poly);
    (karte.bestand_ids || []).forEach((b, i) => { if (b && fl.bestand && fl.bestand[i]) POLY.set(b.id, fl.bestand[i]); });

    const bauten = new Map();
    const pinOrte = new Map();
    for (const g of gebaeude.values()) {
      const lage = g.lage || {};
      const b = { nr: g.nr, ringe: [], mitte: null, pin: false, sicher: lage.sicher !== false, hoehe: (Number(g.hoehe) || 8) * 10, hoch: 0, zielHoch: 0,
        relevant: false, dringend: false, gewicht: 0, versatz: 0, schild: null, sb: 44, sh: 22, box: null };
      if (lage.art === 'polygon') {
        let gross = null;
        for (const id of lage.polys || []) {
          const poly = POLY.get(id); if (!poly) continue;
          poly.forEach((r, i) => {
            if (!r || r.length < 6) return;
            const fm = flaecheUndMitte(r);
            b.ringe.push({ p: r, seite: (fm.a >= 0 ? 1 : -1) * (i > 0 ? -1 : 1) });
            if (i === 0 && (!gross || Math.abs(fm.a) > Math.abs(gross.a))) gross = fm;
          });
        }
        if (gross) {
          b.mitte = [gross.x, gross.y]; b.gewicht = Math.min(4, Math.log10(Math.abs(gross.a) + 1) - 2);
          b.box = { x0: Infinity, y0: Infinity, x1: -Infinity, y1: -Infinity };
          for (const r of b.ringe) for (let i = 0; i < r.p.length - 1; i += 2) { b.box.x0 = Math.min(b.box.x0, r.p[i]); b.box.x1 = Math.max(b.box.x1, r.p[i]); b.box.y0 = Math.min(b.box.y0, r.p[i + 1]); b.box.y1 = Math.max(b.box.y1, r.p[i + 1]); }
        }
      } else if (lage.art === 'pin' && Array.isArray(lage.p) && lage.p.length >= 2) {
        b.pin = true; b.mitte = [Number(lage.p[0]), Number(lage.p[1])];
        const k = b.mitte.join(','); if (!pinOrte.has(k)) pinOrte.set(k, []); pinOrte.get(k).push(b);
      }
      if (!b.mitte || !isFinite(b.mitte[0]) || !isFinite(b.mitte[1])) continue;        // ohne Lage: nicht auf der Karte
      bauten.set(g.nr, b);
    }
    // Rahmen für die Startansicht: der Zaun und alle Gebäude, auch die außerhalb des Zauns
    const passPunkte = [];
    for (const L of li.werkzaun || []) for (let i = 0; i < L.length - 1; i += 10) passPunkte.push([L[i], L[i + 1]]);
    for (const b of bauten.values()) {
      if (b.box) passPunkte.push([b.box.x0, b.box.y0], [b.box.x1, b.box.y0], [b.box.x1, b.box.y1], [b.box.x0, b.box.y1]); else passPunkte.push(b.mitte);
    }
    if (passPunkte.length < 3) passPunkte.push([grenze.x0, grenze.y0], [grenze.x1, grenze.y0], [grenze.x1, grenze.y1], [grenze.x0, grenze.y1]);
    for (const q of passPunkte) { grenze.x0 = Math.min(grenze.x0, q[0]); grenze.x1 = Math.max(grenze.x1, q[0]); grenze.y0 = Math.min(grenze.y0, q[1]); grenze.y1 = Math.max(grenze.y1, q[1]); }
    // Lange Achse des Werks (Winkel gegen Osten), für die Startansicht im Hochformat
    const HAUPT = (() => {
      let mx = 0, my = 0; for (const q of passPunkte) { mx += q[0]; my += q[1]; } mx /= passPunkte.length; my /= passPunkte.length;
      let xx = 0, yy = 0, xy = 0; for (const q of passPunkte) { const u = q[0] - mx, v = q[1] - my; xx += u * u; yy += v * v; xy += u * v; }
      return 0.5 * Math.atan2(2 * xy, xx - yy);
    })();
    function rahmen(az) {
      const sa = Math.sin(az), ca = Math.cos(az), r = { x0: Infinity, x1: -Infinity, y0: Infinity, y1: -Infinity };
      for (const q of passPunkte) { const x = q[0] * ca - q[1] * sa, y = q[0] * sa + q[1] * ca; r.x0 = Math.min(r.x0, x); r.x1 = Math.max(r.x1, x); r.y0 = Math.min(r.y0, y); r.y1 = Math.max(r.y1, y); }
      return r;
    }

    // Mehrere Platzhalter am selben Ort stehen nebeneinander statt übereinander
    for (const gruppe of pinOrte.values()) gruppe.forEach((b, i) => { b.versatz = i - (gruppe.length - 1) / 2; });

    /* ---------- Nummernschilder ---------- */
    for (const b of bauten.values()) {
      const g = gebaeude.get(b.nr);
      const s = document.createElement('button');
      s.type = 'button'; s.className = 'schild weg' + (b.pin ? ' pin' : '') + (b.sicher ? '' : ' unsicher');     // verborgen, bis die Lage gerechnet ist
      s.dataset.nr = b.nr;
      b.titel = 'Gebäude ' + b.nr + (g && g.name ? ', ' + g.name : '');
      s.setAttribute('aria-label', b.titel);
      const t = document.createElement('span'); t.textContent = b.nr; s.append(t);
      s.addEventListener('click', (ev) => { if (ev.detail === 0 && o.beiWahl) o.beiWahl(b.nr); });      // Tastatur. Der Zeiger läuft über die Karte
      s.addEventListener('focus', () => { if (s.matches(':focus-visible')) hoverSetzen(b.nr); });
      s.addEventListener('blur', () => { if (hoverNr === b.nr) hoverSetzen(null); });
      schilderEl.append(s);
      b.schild = s;
    }
    let gemessen = false;
    function schilderMessen() {
      gemessen = false;
      for (const b of bauten.values()) { const w = b.schild.offsetWidth, h = b.schild.offsetHeight; if (w && h) gemessen = true; b.sb = w || 44; b.sh = h || 22; }
    }

    /* ---------- Kamera ---------- */
    const K = { x: (grenze.x0 + grenze.x1) / 2, y: (grenze.y0 + grenze.y1) / 2, k: 0.06, az: 0, neig: SCHRAEG };
    let B = 0, H = 0, dpr = 1, kPass = 0.06, unberuehrt = true;
    const rand = { l: 0, r: 0, o: 0, u: 0 };
    const mitteX = () => rand.l + Math.max(40, B - rand.l - rand.r) / 2;
    const mitteY = () => rand.o + Math.max(40, H - rand.o - rand.u) / 2;

    function matrix(k) {
      k = k || K;
      const ca = Math.cos(k.az), sa = Math.sin(k.az), cn = Math.cos(k.neig);
      const a = k.k * ca, c = -k.k * sa, b = -k.k * cn * sa, d = -k.k * cn * ca;
      return { a, b, c, d, e: mitteX() - (a * k.x + c * k.y), f: mitteY() - (b * k.x + d * k.y), hz: k.k * Math.sin(k.neig), sa, ca };
    }
    const zuBoden = (m, sx, sy) => { const det = m.a * m.d - m.b * m.c, u = sx - m.e, v = sy - m.f; return [(m.d * u - m.c * v) / det, (-m.b * u + m.a * v) / det]; };
    const ueberhoehung = () => klemme(K_NAH / K.k, 1, 3);
    /* Verschiebt die Kamera so, dass der Bodenpunkt (gx, gy) am Bildpunkt (sx, sy) liegt. */
    function haltePunkt(gx, gy, sx, sy) {
      const m = matrix();
      const [wx, wy] = zuBoden(m, sx, sy);
      K.x += gx - wx; K.y += gy - wy;
    }
    function begrenzen() {
      K.k = klemme(K.k, kPass * 0.6, K_MAX);
      K.neig = klemme(K.neig, 0, NEIG_MAX);
      K.x = klemme(K.x, grenze.x0, grenze.x1); K.y = klemme(K.y, grenze.y0, grenze.y1);
    }
    /* Breit steht der Plan wie gezeichnet. Im Hochformat steht die lange Achse des Werks senkrecht, damit es die Höhe nutzt. */
    const startAz = () => (H > B * 1.15 ? Math.PI / 2 - HAUPT : 0);
    function passend() {
      const r = rahmen(startAz());
      const w = Math.max(40, B - rand.l - rand.r), h = Math.max(40, H - rand.o - rand.u);
      return Math.min(w / (r.x1 - r.x0), h / ((r.y1 - r.y0) * Math.cos(SCHRAEG))) * 0.94;
    }
    function startMitte() {
      const az = startAz(), r = rahmen(az), x = (r.x0 + r.x1) / 2, y = (r.y0 + r.y1) / 2, sa = Math.sin(az), ca = Math.cos(az);
      return [x * ca + y * sa, -x * sa + y * ca];
    }

    /* ---------- Abläufe ---------- */
    const ticker = new Map();
    let raf = 0, ersatz = 0;
    /* Neuzeichnen anfordern. Läuft die Seite im Hintergrund, liefert der Browser keinen Bildtakt. Dann zeichnet ein Zeitgeber,
       sonst blieben Karte und Schilder auf einem alten Stand stehen. */
    const fordern = () => {
      if (raf) return;
      raf = requestAnimationFrame(zeichnen);
      ersatz = setTimeout(() => { if (raf) { cancelAnimationFrame(raf); zeichnen(performance.now()); } }, 250);
    };
    /* Sofort zeichnen und die Schilder setzen, ohne auf den Bildtakt zu warten. */
    const sofort = () => { if (raf) cancelAnimationFrame(raf); zeichnen(performance.now()); };
    function animiere(schluessel, dauer, schritt, fertig) {
      if (reduziert || dauer <= 0) { ticker.delete(schluessel); schritt(1); if (fertig) fertig(); fordern(); return; }
      const t0 = performance.now();
      ticker.set(schluessel, (t) => { const p = klemme((t - t0) / dauer, 0, 1); schritt(weich(p)); if (p >= 1) { if (fertig) fertig(); return false; } return true; });
      fordern();
    }
    function fahre(ziel, dauer) {
      const von = Object.assign({}, K);
      const zu = Object.assign({}, K, ziel);
      let dAz = (zu.az - von.az) % (2 * Math.PI); if (dAz > Math.PI) dAz -= 2 * Math.PI; if (dAz < -Math.PI) dAz += 2 * Math.PI;
      animiere('kamera', dauer === undefined ? 600 : dauer, (p) => {
        K.x = von.x + (zu.x - von.x) * p; K.y = von.y + (zu.y - von.y) * p;
        K.k = von.k * Math.pow(zu.k / von.k, p);
        K.az = von.az + dAz * p; K.neig = von.neig + (zu.neig - von.neig) * p;
        begrenzen(); melden();
      });
    }
    function melden() { if (o.beiAnsicht) o.beiAnsicht({ az: K.az, draufsicht: K.neig < 0.12 }); }

    /* ---------- Zeichnen ---------- */
    let gewaehlt = null, hoverNr = null;

    function ringPfad(b, m, dy) {
      const p = new Path2D();
      for (const r of b.ringe) {
        const q = r.p;
        p.moveTo(m.a * q[0] + m.c * q[1] + m.e, m.b * q[0] + m.d * q[1] + m.f - dy);
        for (let i = 2; i < q.length - 1; i += 2) p.lineTo(m.a * q[i] + m.c * q[i + 1] + m.e, m.b * q[i] + m.d * q[i + 1] + m.f - dy);
        p.closePath();
      }
      return p;
    }

    function koerperZeichnen(b, m, h) {
      // Wände, die zum Betrachter zeigen, von hinten nach vorn. Licht von Westsüdwest.
      const waende = [];
      for (const r of b.ringe) {
        const q = r.p, n = q.length >> 1;
        for (let i = 0; i < n; i++) {
          const j = (i + 1) % n, x0 = q[2 * i], y0 = q[2 * i + 1], x1 = q[2 * j], y1 = q[2 * j + 1];
          const dx = x1 - x0, dy = y1 - y0, len = Math.hypot(dx, dy); if (len < 1e-6) continue;
          const nx = r.seite * dy / len, ny = -r.seite * dx / len;
          if (nx * m.sa + ny * m.ca >= -0.02) continue;                       // abgewandt
          const hell = Math.max(0, nx * -0.916 + ny * -0.4);
          waende.push({ x0, y0, x1, y1, tiefe: (x0 + x1) * m.sa + (y0 + y1) * m.ca, hell });
        }
      }
      waende.sort((u, v) => v.tiefe - u.tiefe);
      for (const w of waende) {
        const ax = m.a * w.x0 + m.c * w.y0 + m.e, ay = m.b * w.x0 + m.d * w.y0 + m.f;
        const bx = m.a * w.x1 + m.c * w.y1 + m.e, by = m.b * w.x1 + m.d * w.y1 + m.f;
        ctx.fillStyle = misch(TIEF, BERNSTEIN, 0.42 + 0.36 * w.hell);
        ctx.beginPath(); ctx.moveTo(ax, ay); ctx.lineTo(bx, by); ctx.lineTo(bx, by - h); ctx.lineTo(ax, ay - h); ctx.closePath(); ctx.fill();
        ctx.strokeStyle = ctx.fillStyle; ctx.lineWidth = 0.6; ctx.stroke();    // schließt Haarfugen zwischen den Wänden
      }
      const dach = ringPfad(b, m, h);
      const hervor = b.nr === gewaehlt || b.nr === hoverNr;
      ctx.fillStyle = misch(BERNSTEIN, [255, 255, 255], hervor ? 0.16 : 0.02);
      ctx.fill(dach, 'evenodd');
      ctx.strokeStyle = b.nr === gewaehlt ? '#ffffff' : 'rgba(255,255,255,0.34)';
      ctx.lineWidth = b.nr === gewaehlt ? 2 : 1;
      ctx.stroke(dach);
    }

    function zeichnen(jetzt) {
      raf = 0; clearTimeout(ersatz);
      let laeuft = false;
      for (const [k, f] of [...ticker]) { if (f(jetzt)) laeuft = true; else ticker.delete(k); }
      if (!B || !H) { groesse(true); if (!B || !H) return; }
      if (!gemessen) schilderMessen();
      const m = matrix();
      const dm = new DOMMatrix([m.a, m.b, m.c, m.d, m.e, m.f]);
      const um = (p) => { const q = new Path2D(); q.addPath(p, dm); return q; };

      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx.fillStyle = F.grund; ctx.fillRect(0, 0, B, H);
      ctx.lineJoin = 'round'; ctx.lineCap = 'round'; ctx.setLineDash([]);
      for (const s of flaechen) { ctx.fillStyle = s.farbe; ctx.fill(um(s.pfad), 'evenodd'); }
      for (const s of linien) {
        if (s.name === 'fein' && K.k < kPass * 0.8) continue;                  // weit herausgezoomt wäre es nur Rauschen
        ctx.strokeStyle = s.farbe; ctx.lineWidth = s.breite; ctx.stroke(um(s.pfad));
      }
      ctx.fillStyle = F.bestand; ctx.fill(um(bestand), 'evenodd');
      ctx.globalAlpha = 0.85; ctx.strokeStyle = F.ring; ctx.lineWidth = 1.8; ctx.stroke(um(ring)); ctx.globalAlpha = 1;

      const faktor = ueberhoehung() * m.hz;
      const belegt = schilderSetzen(m, faktor);

      // Straßen, Gleise, Bereiche: leise Schrift auf dem Boden. Sie weicht den Schildern und einander aus.
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillStyle = F.text;
      for (const t of texte) {
        const x = m.a * t.x + m.c * t.y + m.e, y = m.b * t.x + m.d * t.y + m.f;
        if (x < -80 || x > B + 80 || y < -20 || y > H + 20) continue;
        const gross = t.art !== 'strasse';
        ctx.font = gross ? '600 ' + (t.art === 'bereich' ? 11 : 10) + 'px ' + F.schrift : 'italic 500 10.5px ' + F.schrift;
        try { ctx.letterSpacing = gross ? '1.6px' : '0.5px'; } catch (e) { }
        const wort = gross ? t.t.toUpperCase() : t.t;
        if (!t.b || t.bk !== gross) { t.b = ctx.measureText(wort).width; t.bk = gross; }
        const r = { x0: x - t.b / 2 - 4, x1: x + t.b / 2 + 4, y0: y - 9, y1: y + 9 };
        if (belegt.some((q) => !(r.x1 < q.x0 || r.x0 > q.x1 || r.y1 < q.y0 || r.y0 > q.y1))) continue;
        belegt.push(r);
        ctx.fillText(wort, x, y);
      }
      try { ctx.letterSpacing = '0px'; } catch (e) { }

      // Gebäude flach, darauf die Platzhalter
      const hoch = [];
      for (const b of bauten.values()) {
        if (b.pin) continue;
        const hervor = b.nr === gewaehlt || b.nr === hoverNr;
        const p = ringPfad(b, m, 0);
        ctx.fillStyle = F.bauFlach; ctx.fill(p, 'evenodd');
        ctx.setLineDash(b.sicher ? [] : [4, 3]);
        ctx.strokeStyle = hervor ? F.stahl : b.relevant ? misch(TIEF, BERNSTEIN, 0.8) : F.bau; ctx.lineWidth = hervor ? 1.6 : 1; ctx.stroke(p);
        ctx.setLineDash([]);
        if (b.hoch > 0.002) hoch.push(b);
      }
      for (const b of bauten.values()) {
        if (!b.pin) continue;
        const x = m.a * b.mitte[0] + m.c * b.mitte[1] + m.e, y = m.b * b.mitte[0] + m.d * b.mitte[1] + m.f;
        const hervor = b.nr === gewaehlt || b.nr === hoverNr;
        ctx.strokeStyle = hervor ? F.stahl : b.relevant ? misch(TIEF, BERNSTEIN, 0.9) : F.bau; ctx.lineWidth = 1.4;
        ctx.beginPath(); ctx.ellipse(x, y, 5, 5 * Math.max(0.35, Math.cos(K.neig)), 0, 0, 2 * Math.PI); ctx.stroke();
      }
      hoch.sort((u, v) => (v.mitte[0] * m.sa + v.mitte[1] * m.ca) - (u.mitte[0] * m.sa + u.mitte[1] * m.ca));
      for (const b of hoch) koerperZeichnen(b, m, b.hoehe * b.hoch * faktor);
      if (laeuft) fordern();
    }

    /* Nummernschilder setzen. In der Übersicht stehen nur die Schilder der Gebäude mit Einträgen, die übrigen kommen beim
       Hineinzoomen dazu. Überdecken sich zwei, bleibt das wichtigere stehen. Gerechnet wird bei jedem Zeichnen neu. */
    function schilderSetzen(m, faktor) {
      const liste = [];
      const alle = K.k >= Math.max(kPass * 1.6, K_ALLE);
      for (const b of bauten.values()) {
        let x = m.a * b.mitte[0] + m.c * b.mitte[1] + m.e, y = m.b * b.mitte[0] + m.d * b.mitte[1] + m.f;
        let l, t;
        if (b.pin) { x += b.versatz * (b.sb + 6); l = x - b.sb / 2; t = y + 10; }
        else { y -= b.hoehe * b.hoch * faktor; l = x - b.sb / 2; t = y - b.sh / 2 - b.hoch * (b.sh / 2 + 2); }      // flach: mittig. Erhoben: über der Dachmitte
        l = Math.round(l); t = Math.round(t);
        b.schild.style.transform = `translate(${l}px, ${t}px)`;
        const wichtig = b.relevant || b.nr === gewaehlt || b.nr === hoverNr;
        if ((!alle && !wichtig) || l > B || t > H || l + b.sb < 0 || t + b.sh < 0) { b.schild.classList.add('weg'); continue; }
        const p = (b.nr === gewaehlt ? 100 : 0) + (b.nr === hoverNr ? 60 : 0) + (b.dringend ? 36 : b.relevant ? 30 : 10) - (b.pin ? 3 : 0) + b.gewicht;
        liste.push({ b, p, x0: l - ABSTAND, x1: l + b.sb + ABSTAND, y0: t - ABSTAND, y1: t + b.sh + ABSTAND });
      }
      liste.sort((u, v) => v.p - u.p);
      const belegt = [];
      for (const e of liste) {
        const frei = !belegt.some((q) => !(e.x1 < q.x0 || e.x0 > q.x1 || e.y1 < q.y0 || e.y0 > q.y1));
        e.b.schild.classList.toggle('weg', !frei);
        if (frei) belegt.push(e);
      }
      return belegt;
    }

    /* ---------- Treffen ---------- */
    function imRing(ringe, x, y) {
      let drin = false;
      for (const r of ringe) {
        const q = r.p, n = q.length >> 1;
        for (let i = 0, j = n - 1; i < n; j = i++) {
          const xi = q[2 * i], yi = q[2 * i + 1], xj = q[2 * j], yj = q[2 * j + 1];
          if ((yi > y) !== (yj > y) && x < (xj - xi) * (y - yi) / (yj - yi) + xi) drin = !drin;
        }
      }
      return drin;
    }
    function treffer(sx, sy) {
      const m = matrix(), faktor = ueberhoehung() * m.hz;
      let best = null, bestTiefe = Infinity, pin = null;
      for (const b of bauten.values()) {
        const tiefe = b.mitte[0] * m.sa + b.mitte[1] * m.ca;                    // kleiner ist näher am Betrachter
        if (b.pin) {
          // Platzhalter zählen nur, wenn dort kein Baukörper getroffen ist. Sie sitzen oft auf ihrem Bezugsgebäude.
          const x = m.a * b.mitte[0] + m.c * b.mitte[1] + m.e, y = m.b * b.mitte[0] + m.d * b.mitte[1] + m.f;
          if (!pin && Math.hypot(sx - x, sy - y) <= 11) pin = b.nr;
          continue;
        }
        const h = b.hoehe * b.hoch * faktor;
        const stufen = h > 2 ? [1, 0.75, 0.5, 0.25, 0] : [0];
        for (const s of stufen) {
          const [gx, gy] = zuBoden(m, sx, sy + h * s);
          if (gx < b.box.x0 || gx > b.box.x1 || gy < b.box.y0 || gy > b.box.y1) continue;
          if (imRing(b.ringe, gx, gy)) { const t = tiefe - (h > 2 ? 1e7 : 0); if (t < bestTiefe) { bestTiefe = t; best = b.nr; } break; }
        }
      }
      return best || pin;
    }

    function hoverSetzen(nr) {
      if (hoverNr === nr) return;
      const alt = hoverNr && bauten.get(hoverNr); if (alt) alt.schild.classList.remove('aktiv');
      hoverNr = nr;
      const neu = nr && bauten.get(nr); if (neu) neu.schild.classList.add('aktiv');
      el.classList.toggle('zeigt', !!nr);
      fordern();
    }

    /* ---------- Bedienung ---------- */
    const zeiger = new Map();
    let geste = null, hoverRaf = 0;
    const ort = (ev) => { const r = el.getBoundingClientRect(); return [ev.clientX - r.left, ev.clientY - r.top]; };
    const zwei = () => { const [p, q] = [...zeiger.values()]; return { x: (p[0] + q[0]) / 2, y: (p[1] + q[1]) / 2, d: Math.hypot(q[0] - p[0], q[1] - p[1]) || 1, w: Math.atan2(q[1] - p[1], q[0] - p[0]) }; };

    el.addEventListener('contextmenu', (ev) => ev.preventDefault());
    el.addEventListener('pointerdown', (ev) => {
      const p = ort(ev);
      zeiger.set(ev.pointerId, p);
      try { el.setPointerCapture(ev.pointerId); } catch (e) { }
      ticker.delete('kamera');
      if (zeiger.size === 1) {
        const s = ev.target.closest && ev.target.closest('.schild');
        geste = { art: ev.button === 2 || ev.shiftKey || ev.ctrlKey ? 'drehen' : 'schieben', x: p[0], y: p[1], t: performance.now(), bewegt: false,
          schild: s ? s.dataset.nr : null, K0: Object.assign({}, K), boden: zuBoden(matrix(), p[0], p[1]) };
      } else if (zeiger.size === 2) {
        const z = zwei();
        geste = { art: 'zwei', bewegt: true, z0: z, K0: Object.assign({}, K), boden: zuBoden(matrix(), z.x, z.y) };
      }
    });
    el.addEventListener('pointermove', (ev) => {
      const p = ort(ev);
      if (!zeiger.has(ev.pointerId)) {
        if (ev.pointerType === 'touch') return;
        cancelAnimationFrame(hoverRaf);
        const s = ev.target.closest && ev.target.closest('.schild');
        hoverRaf = requestAnimationFrame(() => hoverSetzen(s ? s.dataset.nr : treffer(p[0], p[1])));
        return;
      }
      zeiger.set(ev.pointerId, p);
      if (!geste) return;
      if (geste.art === 'zwei' && zeiger.size >= 2) {
        const z = zwei();
        K.k = geste.K0.k * z.d / geste.z0.d; K.az = geste.K0.az - (z.w - geste.z0.w);
        begrenzen(); haltePunkt(geste.boden[0], geste.boden[1], z.x, z.y); begrenzen();
      } else if (geste.art !== 'zwei') {
        const dx = p[0] - geste.x, dy = p[1] - geste.y;
        if (!geste.bewegt && Math.hypot(dx, dy) < 6) return;
        geste.bewegt = true;
        if (geste.art === 'drehen') { K.az = geste.K0.az + dx * 0.006; K.neig = klemme(geste.K0.neig - dy * 0.005, 0, NEIG_MAX); }
        else { haltePunkt(geste.boden[0], geste.boden[1], p[0], p[1]); begrenzen(); }
      }
      unberuehrt = false; hoverSetzen(null); melden(); fordern();
    });
    const ende = (ev) => {
      if (!zeiger.has(ev.pointerId)) return;
      const p = zeiger.get(ev.pointerId);
      zeiger.delete(ev.pointerId);
      const g = geste;
      if (zeiger.size === 0) geste = null;
      else if (zeiger.size === 1 && g && g.art === 'zwei') { const [q] = [...zeiger.values()]; geste = { art: 'schieben', x: q[0], y: q[1], t: 0, bewegt: true, schild: null, K0: Object.assign({}, K), boden: zuBoden(matrix(), q[0], q[1]) }; }
      if (ev.type !== 'pointerup' || !g || g.art === 'zwei' || g.bewegt || performance.now() - g.t > 700) return;
      const nr = g.schild || treffer(p[0], p[1]);
      if (nr) { if (o.beiWahl) o.beiWahl(nr); } else if (o.beiLeer) o.beiLeer();
    };
    el.addEventListener('pointerup', ende);
    el.addEventListener('pointercancel', ende);
    el.addEventListener('pointerleave', (ev) => { if (!zeiger.size && ev.pointerType !== 'touch') hoverSetzen(null); });
    el.addEventListener('wheel', (ev) => {
      ev.preventDefault();
      const p = ort(ev), boden = zuBoden(matrix(), p[0], p[1]);
      ticker.delete('kamera');
      K.k *= Math.exp(-(ev.deltaMode === 1 ? ev.deltaY * 30 : ev.deltaY) * 0.0016);
      begrenzen(); haltePunkt(boden[0], boden[1], p[0], p[1]); begrenzen();
      unberuehrt = false; fordern();
    }, { passive: false });
    el.addEventListener('keydown', (ev) => {
      if (ev.target !== el) return;
      const m = matrix();
      const schiebe = (sx, sy) => { const [gx, gy] = zuBoden(m, mitteX() + sx, mitteY() + sy); fahre({ x: gx, y: gy }, 160); };
      if (ev.key === 'ArrowLeft') schiebe(-70, 0); else if (ev.key === 'ArrowRight') schiebe(70, 0);
      else if (ev.key === 'ArrowUp') schiebe(0, -70); else if (ev.key === 'ArrowDown') schiebe(0, 70);
      else if (ev.key === '+' || ev.key === '=') zoom(1.4); else if (ev.key === '-') zoom(1 / 1.4);
      else return;
      unberuehrt = false; ev.preventDefault();
    });

    /* ---------- Größe ---------- */
    /* Größe der Fläche übernehmen. Danach wird sofort neu gezeichnet, damit die Schilder zur neuen Größe passen. */
    function groesse(nurMessen) {
      const w = el.clientWidth, h = el.clientHeight;
      if (!w || !h) return;
      const neu = w !== B || h !== H;
      B = w; H = h; dpr = Math.min(window.devicePixelRatio || 1, 2);
      if (neu || canvas.width !== Math.round(B * dpr)) { canvas.width = Math.round(B * dpr); canvas.height = Math.round(H * dpr); }
      kPass = passend();
      if (unberuehrt) { const mi = startMitte(); K.k = kPass; K.x = mi[0]; K.y = mi[1]; K.az = startAz(); K.neig = SCHRAEG; melden(); }
      begrenzen();
      if (nurMessen !== true) sofort();
    }
    if (typeof ResizeObserver === 'function') new ResizeObserver(() => groesse()).observe(el);
    window.addEventListener('resize', () => groesse());

    /* ---------- Schnittstelle ---------- */
    function zoom(f) { unberuehrt = false; fahre({ k: klemme(K.k * f, kPass * 0.6, K_MAX) }, 220); }
    function zustand(map) {
      for (const b of bauten.values()) {
        const z = (map && map.get(b.nr)) || {};
        b.relevant = !!z.relevant; b.dringend = !!z.dringend;
        b.schild.classList.toggle('relevant', b.relevant); b.schild.classList.toggle('dringend', b.dringend);
        const ueber = b.dringend ? ueberText(z.ueberfaellig) : '';
        if (ueber) b.schild.dataset.ueber = ueber; else delete b.schild.dataset.ueber;
        b.schild.setAttribute('aria-label', b.titel + (ueber ? ', ' + ueber + ' überfällig' : ''));
        const ziel = b.relevant && !b.pin ? 1 : 0;
        if (ziel !== b.zielHoch) {
          b.zielHoch = ziel; const von = b.hoch;
          animiere('hoch:' + b.nr, 520 + Math.random() * 260, (p) => { b.hoch = von + (ziel - von) * p; });
        }
      }
      sofort();
    }
    function waehle(nr) {
      const alt = gewaehlt && bauten.get(gewaehlt); if (alt) alt.schild.classList.remove('gewaehlt');
      gewaehlt = nr && bauten.has(nr) ? nr : null;
      if (gewaehlt) bauten.get(gewaehlt).schild.classList.add('gewaehlt');
      sofort();
    }
    function zeige(nr) {
      const b = bauten.get(nr); if (!b) return;
      unberuehrt = false;
      fahre({ x: b.mitte[0], y: b.mitte[1], k: klemme(K.k, K_NAH, 0.45) }, 650);
    }

    schilderMessen(); groesse(); melden();

    return {
      zustand, waehle, zeige, zoom,
      rand: (r) => { Object.assign(rand, r); groesse(); },
      groesse: () => groesse(),
      ordnen: sofort,
      ansichtWechseln: () => fahre({ neig: K.neig < 0.12 ? SCHRAEG : 0 }, 600),
      start: () => { unberuehrt = true; kPass = passend(); const mi = startMitte(); fahre({ x: mi[0], y: mi[1], k: kPass, az: startAz(), neig: SCHRAEG }, 650); },
      hat: (nr) => bauten.has(nr),
      messen: () => { schilderMessen(); sofort(); },
      stand: () => ({ kamera: Object.assign({}, K), breite: B, hoehe: H, schilder: [...bauten.values()].filter((b) => !b.schild.classList.contains('weg')).length,
        erhoben: [...bauten.values()].filter((b) => b.zielHoch === 1).map((b) => b.nr), gewaehlt, laeuft: ticker.size > 0 }),
      bildpunkt: (nr) => { const b = bauten.get(nr); if (!b) return null; const m = matrix(); return [m.a * b.mitte[0] + m.c * b.mitte[1] + m.e, m.b * b.mitte[0] + m.d * b.mitte[1] + m.f - b.hoehe * b.hoch * ueberhoehung() * m.hz]; },
    };
  }

  wurzel.FPKarteAnsicht = { ueberText, erzeugen };
})(typeof window !== 'undefined' ? window : this);
