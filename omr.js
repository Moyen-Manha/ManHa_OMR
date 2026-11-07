/* OMR কোর (OpenCV.js) — ব্রাউজার ও Node দুটোতেই চলে। গ্লোবাল `cv` দরকার। */
const OMR = (() => {
  const TW = 1600, TH = 2211, OPT = ['A', 'B', 'C', 'D'];
  const LAYOUT = { b1: [.141, .182, .222, .263], b2: [.355, .395, .436, .477], y0: .309, y1: .928, rows: 15,
    setx: .897, sety: [.314, .339, .364, .389],
    rx: [.523, .557, .591, .625, .659, .693], ry: [.311, .330, .351, .371, .391, .411, .431, .451, .471, .490] };
  const BN = { 'ক': 'A', 'খ': 'B', 'গ': 'C', 'ঘ': 'D' };
  const enDigits = t => String(t).replace(/[০-৯]/g, d => '০১২৩৪৫৬৭৮৯'.indexOf(d));
  const normOpt = s => { s = String(s || '').trim().toUpperCase(); s = BN[s] || s; return OPT.includes(s) ? s : null; };

  /* ---------- সঠিক উত্তর ---------- */
  function parseKey(text) {
    text = enDigits(text); const key = {}; const re = /(\d+)\s*[-:.=)\s]\s*([A-Da-dকখগঘ])/g; let m;
    while ((m = re.exec(text))) { const o = normOpt(m[2]); if (o) key[+m[1]] = o; }
    if (!Object.keys(key).length) text.split(/[,\s]+/).map(normOpt).filter(Boolean).forEach((o, i) => key[i + 1] = o);
    return key;
  }
  function parseKeyRows(rows) {  // Excel/CSV: কলাম A = প্রশ্ন, B = উত্তর
    const key = {};
    rows.forEach(r => { const q = parseInt(enDigits(r[0])), o = normOpt(r[1]); if (q && o) key[q] = o; });
    return key;
  }

  /** Excel/CSV থেকে উত্তরমালা: হেডারে সেট কলাম (A/B/C/D, ক/খ/গ/ঘ, "সেট ক", "Set A", "ডিফল্ট") থাকলে সব সেট একসাথে;
   *  না থাকলে প্রথম দুই কলাম (প্রশ্ন, উত্তর) → {single}। রিটার্ন: { multi, sets } */
  function parseKeySets(rows) {
    const label = c => { const t = String(c ?? '').trim().replace(/set|সেট|[\s:：.\-_()]/gi, ''); if (/^(default|ডিফল্ট|সব)$/i.test(t)) return 'default'; return normOpt(t); };
    for (let i = 0; i < Math.min(rows.length, 6); i++) {
      const cols = {}; (rows[i] || []).forEach((c, j) => { if (j >= 1) { const l = label(c); if (l && !(l in cols)) cols[l] = j; } });
      if (Object.keys(cols).length >= 2) {
        const sets = {}; Object.keys(cols).forEach(s => sets[s] = {});
        rows.slice(i + 1).forEach(r => { const q = parseInt(enDigits(r[0])); if (!q) return; for (const s in cols) { const o = normOpt(r[cols[s]]); if (o) sets[s][q] = o; } });
        return { multi: true, sets };
      }
    }
    return { multi: false, sets: { single: parseKeyRows(rows) } };
  }
  const rollKey = r => { const n = enDigits(r).replace(/\D/g, ''); return n ? String(parseInt(n, 10)) : ''; };
  function parseStudents(rows) {  // কলাম: নাম, শ্রেণি, রোল, গ্রুপ (হেডার থাকলে নাম দেখে, না থাকলে এই ক্রমে)
    const re = { name: /নাম|name/i, class: /শ্রেণি|শ্রেণী|class/i, roll: /রোল|roll/i, group: /গ্রুপ|বিভাগ|group|dept|section/i };
    let map = { name: 0, class: 1, roll: 2, group: 3 }, start = 0;
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
      const m = {}; (rows[i] || []).forEach((c, j) => { for (const k in re) if (!(k in m) && re[k].test(String(c))) m[k] = j; });
      if ('roll' in m || 'name' in m) { map = Object.assign({ name: -1, class: -1, roll: -1, group: -1 }, m); start = i + 1; break; }
    }
    const v = (r, k) => map[k] >= 0 && r[map[k]] != null ? String(r[map[k]]).trim() : '';
    return rows.slice(start).map(r => ({ name: v(r, 'name'), class: v(r, 'class'), roll: enDigits(v(r, 'roll')), group: v(r, 'group') })).filter(x => x.roll || x.name);
  }
  const findStudent = (list, roll) => { const k = rollKey(roll); return k ? list.find(x => rollKey(x.roll) === k) : null; };

  /* ---------- কাগজ খোঁজা ---------- */
  const orderPts = p => {
    const s = p.map(a => a.x + a.y), d = p.map(a => a.y - a.x);
    return [p[s.indexOf(Math.min(...s))], p[d.indexOf(Math.min(...d))], p[s.indexOf(Math.max(...s))], p[d.indexOf(Math.max(...d))]];
  };
  function quadFromMask(mask, minRatio) {
    const cs = new cv.MatVector(), hi = new cv.Mat();
    cv.findContours(mask, cs, hi, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE);
    let best = -1, ba = 0;
    for (let i = 0; i < cs.size(); i++) { const a = cv.contourArea(cs.get(i)); if (a > ba) { ba = a; best = i; } }
    let pts = null;
    if (best >= 0 && ba >= minRatio * mask.rows * mask.cols) {
      const hull = new cv.Mat(); cv.convexHull(cs.get(best), hull, false, true);
      const peri = cv.arcLength(hull, true);
      for (const e of [.02, .03, .05, .08]) {
        const ap = new cv.Mat(); cv.approxPolyDP(hull, ap, e * peri, true);
        if (ap.rows === 4) { const d = ap.data32S; pts = [0, 1, 2, 3].map(i => ({ x: d[2 * i], y: d[2 * i + 1] })); }
        ap.delete(); if (pts) break;
      }
      if (!pts) pts = cv.RotatedRect.points(cv.minAreaRect(hull));
      hull.delete();
    }
    cs.delete(); hi.delete(); return pts;
  }
  /** সোর্স (RGBA Mat) থেকে কাগজের চার কোণা (মূল স্কেলে) বা null */
  function findQuad(src, maxSide = 1000) {
    const sc = maxSide / Math.max(src.cols, src.rows), small = new cv.Mat();
    cv.resize(src, small, new cv.Size(Math.round(src.cols * sc), Math.round(src.rows * sc)), 0, 0, cv.INTER_AREA);
    const g = new cv.Mat(); cv.cvtColor(small, g, cv.COLOR_RGBA2GRAY); cv.GaussianBlur(g, g, new cv.Size(7, 7), 0);
    const k = cv.Mat.ones(5, 5, cv.CV_8U), a = new cv.Point(-1, -1), gens = [
      m => { cv.threshold(g, m, 0, 255, cv.THRESH_BINARY + cv.THRESH_OTSU); cv.morphologyEx(m, m, cv.MORPH_CLOSE, k, a, 3); },
      m => { cv.adaptiveThreshold(g, m, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY, 101, -10); cv.morphologyEx(m, m, cv.MORPH_CLOSE, k, a, 5); },
      m => { cv.Canny(g, m, 40, 120); cv.dilate(m, m, k, a, 3); cv.morphologyEx(m, m, cv.MORPH_CLOSE, k, a, 3); }];
    let quad = null;
    for (const gen of gens) {
      const m = new cv.Mat(); gen(m); const q = quadFromMask(m, .2); m.delete();
      if (!q) continue;
      const r = orderPts(q), W = Math.hypot(r[1].x - r[0].x, r[1].y - r[0].y), H = Math.hypot(r[3].x - r[0].x, r[3].y - r[0].y);
      const ratio = Math.min(W, H) / Math.max(W, H);
      if (ratio > .5 && ratio < .95) { quad = r.map(p => ({ x: p.x / sc, y: p.y / sc })); break; }
    }
    small.delete(); g.delete(); k.delete(); return quad;
  }
  /** কাগজ সোজা করে TW×TH Mat ফেরত দেয়। রিটার্ন {mat, found} */
  function fullFrame(src, quad) {  // স্ক্যান করা পেজ পুরো ছবি জুড়ে থাকলে ভুল কোণা ধরে বাঁকিয়ে ফেলা ঠেকানো
    const sc = 400 / Math.max(src.cols, src.rows), sm = new cv.Mat(), g = new cv.Mat(), b = new cv.Mat();
    cv.resize(src, sm, new cv.Size(Math.round(src.cols * sc), Math.round(src.rows * sc)), 0, 0, cv.INTER_AREA); cv.cvtColor(sm, g, cv.COLOR_RGBA2GRAY);
    const th = cv.threshold(g, b, 0, 255, cv.THRESH_BINARY + cv.THRESH_OTSU), W = g.cols, H = g.rows, m = Math.max(3, Math.round(.025 * Math.min(W, H))); let on = 0, n = 0;
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) if (x < m || y < m || x >= W - m || y >= H - m) { n++; if (g.data[y * W + x] > th) on++; }
    sm.delete(); g.delete(); b.delete();
    const area = Math.abs(quad.reduce((a, p, i) => { const q = quad[(i + 1) % 4]; return a + p.x * q.y - q.x * p.y; }, 0)) / 2 / (src.cols * src.rows);
    const W0 = src.cols, H0 = src.rows, d = .04 * Math.hypot(W0, H0), cn = [[0, 0], [W0, 0], [W0, H0], [0, H0]];
    const near = quad.filter((p, i) => Math.hypot(p.x - cn[i][0], p.y - cn[i][1]) < d).length;  // ৪ কোণার ≥৩টি ছবির কোণায় = পেজ ফ্রেম জুড়ে
    return (on / n > .85 && area > .75) || (near >= 3 && area > .7);
  }
  function autoScan(src) {
    let quad = findQuad(src), dst = new cv.Mat();
    if (!quad || fullFrame(src, quad)) { cv.resize(src, dst, new cv.Size(TW, TH), 0, 0, cv.INTER_AREA); return { mat: dst, found: !!quad }; }
    const [tl, tr, br, bl] = quad;
    if (Math.hypot(tr.x - tl.x, tr.y - tl.y) > Math.hypot(bl.x - tl.x, bl.y - tl.y)) quad = [bl, tl, tr, br];
    const s = cv.matFromArray(4, 1, cv.CV_32FC2, quad.flatMap(p => [p.x, p.y]));
    const d = cv.matFromArray(4, 1, cv.CV_32FC2, [0, 0, TW - 1, 0, TW - 1, TH - 1, 0, TH - 1]);
    const M = cv.getPerspectiveTransform(s, d);
    cv.warpPerspective(src, dst, M, new cv.Size(TW, TH), cv.INTER_CUBIC, cv.BORDER_CONSTANT, new cv.Scalar());
    s.delete(); d.delete(); M.delete(); return { mat: dst, found: true };
  }

  /* ---------- ছাপা বৃত্ত থেকে গ্রিড নিজে মেলানো (স্কেল/সরানো/বাঁকা পেজ সামলায়) ---------- */
  const median = a => { const b = [...a].sort((x, y) => x - y); return b[Math.floor(b.length / 2)]; };
  function kmeans(v, c) {
    c = [...c]; const near = x => { let b = 0; for (let i = 1; i < c.length; i++) if (Math.abs(x - c[i]) < Math.abs(x - c[b])) b = i; return b; };
    for (let it = 0; it < 40; it++) {
      const s = c.map(() => 0), n = c.map(() => 0); v.forEach(x => { const b = near(x); s[b] += x; n[b]++; });
      const nc = c.map((x, i) => n[i] ? s[i] / n[i] : x), done = nc.every((x, i) => Math.abs(x - c[i]) < .01); c = nc; if (done) break;
    }
    return { c, near };
  }
  function polyfit(xs, ys, deg) {
    const n = deg + 1, A = Array.from({ length: n }, () => new Array(n + 1).fill(0));
    xs.forEach((x, k) => { for (let i = 0; i < n; i++) { for (let j = 0; j < n; j++) A[i][j] += x ** (i + j); A[i][n] += ys[k] * x ** i; } });
    for (let i = 0; i < n; i++) { let p = i; for (let r = i + 1; r < n; r++) if (Math.abs(A[r][i]) > Math.abs(A[p][i])) p = r; [A[i], A[p]] = [A[p], A[i]];
      for (let r = 0; r < n; r++) if (r !== i) { const f = A[r][i] / (A[i][i] || 1e-9); for (let c = i; c <= n; c++) A[r][c] -= f * A[i][c]; } }
    const co = A.map((r, i) => r[n] / (r[i] || 1e-9)); return x => co.reduce((a, c, i) => a + c * x ** i, 0);
  }
  function hough(gray, s, box, rmin, rmax) {  // box/ব্যাসার্ধ শিট-পিক্সেলে; gray = s গুণ ছোট ছবি
    const x0 = Math.max(0, Math.round(box[0] * s)), y0 = Math.max(0, Math.round(box[1] * s)), x1 = Math.min(gray.cols, Math.round(box[2] * s)), y1 = Math.min(gray.rows, Math.round(box[3] * s));
    if (x1 - x0 < 20 || y1 - y0 < 20) return [];
    const roi = gray.roi(new cv.Rect(x0, y0, x1 - x0, y1 - y0)).clone(), out = new cv.Mat();
    cv.HoughCircles(roi, out, cv.HOUGH_GRADIENT, 1, rmin * s * 1.7, 100, 18, Math.round(rmin * s), Math.round(rmax * s));
    const r = []; for (let i = 0; i < out.cols; i++) r.push({ x: (out.data32F[3 * i] + x0) / s, y: (out.data32F[3 * i + 1] + y0) / s, r: out.data32F[3 * i + 2] / s });
    roi.delete(); out.delete(); return r;
  }
  function calibrate(sheet) {
    const W = sheet.cols, H = sheet.rows, s = 1000 / Math.max(W, H), L = LAYOUT, small = new cv.Mat(), g = new cv.Mat();
    cv.cvtColor(sheet, g, cv.COLOR_RGBA2GRAY); cv.resize(g, small, new cv.Size(Math.round(W * s), Math.round(H * s)), 0, 0, cv.INTER_AREA); cv.GaussianBlur(small, small, new cv.Size(5, 5), 0); g.delete();
    try {
      const r0 = .0145 * W; let A = hough(small, s, [.03 * W, .24 * H, .488 * W, .97 * H], r0 * .75, r0 * 1.3);
      if (A.length < 50) return null;
      const mr = median(A.map(c => c.r)); A = A.filter(c => Math.abs(c.r - mr) < mr * .25); if (A.length < 50) return null;
      const xs = A.map(c => c.x), ys = A.map(c => c.y), x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...ys), y1 = Math.max(...ys), nom = [...L.b1, ...L.b2];
      const kc = kmeans(xs, nom.map(v => x0 + (v - nom[0]) / (nom[7] - nom[0]) * (x1 - x0))), kr = kmeans(ys, Array.from({ length: L.rows }, (_, i) => y0 + (y1 - y0) * i / (L.rows - 1)));
      const cnt = kc.c.map(() => 0); xs.forEach(x => cnt[kc.near(x)]++); if (cnt.some(n => n < 6)) return null;
      const per = kc.c.map(() => ({ r: [], x: [], y: [] }));
      A.forEach(c => { const p = per[kc.near(c.x)], r = kr.near(c.y); p.r.push(r); p.x.push(c.x); p.y.push(c.y); });
      const G = {};
      per.forEach((p, ci) => {
        const fx = p.r.length >= 3 ? polyfit(p.r, p.x, 1) : () => kc.c[ci], fy = p.r.length >= 6 ? polyfit(p.r, p.y, 2) : r => kr.c[r];
        for (let r = 0; r < L.rows; r++) { const q = (ci < 4 ? 0 : L.rows) + r + 1; (G[q] = G[q] || [])[ci % 4] = [fx(r), fy(r)]; }
      });
      // রোল নম্বর গ্রিড: প্রথম সারি = উত্তরের প্রথম সারির সমান উচ্চতা; বাকিটা নমিনাল লেআউট স্কেল করে, তারপর ছাপা বৃত্ত দিয়ে মেলানো
      const sc = (kc.c[7] - kc.c[0]) / (W * (nom[7] - nom[0])), yr1 = G[1].reduce((a, p) => a + p[1], 0) / 4, pr = (L.ry[9] - L.ry[0]) / 9 * H * sc;
      const rx0 = L.rx.map(v => kc.c[7] + (v - L.b2[3]) * W * sc), ry0 = Array.from({ length: 10 }, (_, i) => yr1 + i * pr);
      let roll = { xs: rx0, ys: ry0, r: (L.rx[1] - L.rx[0]) * W * sc * .3 };
      const R = hough(small, s, [.49 * W, yr1 - .6 * pr, .705 * W, yr1 + 9.6 * pr], mr * .55, mr * .95).filter(c => c.x > .49 * W && c.x < .705 * W && c.y > yr1 - .6 * pr);
      if (R.length >= 12) {
        // কলাম: যে কলামে ≥২ বৃত্ত পাওয়া গেছে সেটি নিজের অবস্থান নেয়, বাকিরা অন্যদের গড় সরণ অনুসরণ করে
        const cx = kmeans(R.map(c => c.x), rx0), nx = rx0.map(() => 0); R.forEach(c => nx[cx.near(c.x)]++);
        const offs = cx.c.map((v, i) => nx[i] >= 2 ? v - rx0[i] : null).filter(v => v !== null), off = offs.length ? median(offs) : 0;
        const xs = rx0.map((v, i) => nx[i] >= 2 ? cx.c[i] : v + off);
        // সারি: সব বৃত্তের y গ্রুপ করে সারির আসল দূরত্ব বের করা (শিট ভেদে ভিন্ন হয়); প্রথম সারি = উত্তরের প্রথম সারির উচ্চতা
        const ysR = R.map(c => c.y).sort((p, q) => p - q), gr = [];
        ysR.forEach(y => { const g = gr[gr.length - 1]; if (g && y - g.last < .4 * pr) { g.s += y; g.n++; g.last = y; } else gr.push({ s: y, n: 1, last: y }); });
        const ce = gr.map(g => g.s / g.n); let ys = ry0;
        if (ce.length >= 5) {
          const d = ce.slice(1).map((v, i) => v - ce[i]).filter(v => v > .6 * pr && v < 1.4 * pr), pit = d.length ? median(d) : pr, k = Math.round((ce[0] - yr1) / pit);
          ys = Array.from({ length: 10 }, (_, i) => ce[0] + (i - k) * pit);
        }
        roll = { xs, ys, r: median(R.map(c => c.r)) * .8 };
      }
      // সেট কোড (ক/খ/গ/ঘ): রোল গ্রিডের ডানে; প্রথম বাবল = প্রথম সারির সমান উচ্চতা। ছাপা বৃত্ত থেকে মেলানো
      const sx0 = roll.xs[5] + .204 * W * sc, sp0 = .0252 * H * sc; let set = { x: sx0, ys: [0, 1, 2, 3].map(i => yr1 + i * sp0), r: mr * .9 };
      const SR = hough(small, s, [sx0 - .05 * W, yr1 - .6 * sp0, sx0 + .05 * W, yr1 + 4.3 * sp0], mr * .7, mr * 1.15).filter(c => Math.abs(c.x - sx0) < .045 * W);
      if (SR.length >= 2) {
        const x = median(SR.map(c => c.x)), ys = SR.map(c => c.y).sort((a, b) => a - b), d = ys.slice(1).map((v, i) => v - ys[i]).filter(v => v > .6 * sp0), pit = d.length ? Math.min(...d) : sp0;
        const idx = ys.map(y => Math.max(0, Math.min(3, Math.round((y - yr1) / pit)))); if (new Set(idx).size >= 2) { const f = polyfit(idx, ys, 1); set = { x, ys: [0, 1, 2, 3].map(f), r: mr * .9 }; }
      }
      return { G, diam: 2 * mr, roll, set };
    } catch (e) { return null; } finally { small.delete(); }
  }

  /* ---------- হাতে সরানো সমন্বয় (শিট-ভিত্তিক): ends=উপর/নীচ প্রান্ত, rows=সারি, bubbles=একক বাবল ---------- */
  function applyAdj(G, adj, rows = LAYOUT.rows) {
    if (!adj) return G; const out = {};
    for (const q in G) {
      const n = +q, b = n > rows ? 1 : 0, r = (n - 1) % rows, t = rows > 1 ? r / (rows - 1) : 0, e = (adj.ends || {})[b] || {}, tp = e.top || [0, 0], bt = e.bottom || [0, 0];
      let dx = tp[0] * (1 - t) + bt[0] * t, dy = tp[1] * (1 - t) + bt[1] * t; const rw = (adj.rows || {})[q]; if (rw) { dx += rw[0]; dy += rw[1]; }
      out[q] = G[q].map(([x, y], i) => { const bb = (adj.bubbles || {})[q + '_' + i]; return [x + dx + (bb ? bb[0] : 0), y + dy + (bb ? bb[1] : 0)]; });
    }
    return out;
  }

  /* ---------- বাবল শনাক্তকরণ ---------- */
  function grid(w, h, L) {
    const g = {};
    for (let b = 0; b < 2; b++) for (let r = 0; r < L.rows; r++) {
      const y = (L.y0 + (L.y1 - L.y0) * r / (L.rows - 1)) * h;
      g[b * L.rows + r + 1] = (b ? L.b2 : L.b1).map(x => [x * w, y]);
    }
    return g;
  }
  function fill(bin, cx, cy, r) {
    const W = bin.cols, H = bin.rows, d = bin.data; let on = 0, n = 0;
    for (let y = Math.max(0, Math.floor(cy - r)); y <= Math.min(H - 1, Math.ceil(cy + r)); y++)
      for (let x = Math.max(0, Math.floor(cx - r)); x <= Math.min(W - 1, Math.ceil(cx + r)); x++)
        if ((x - cx) ** 2 + (y - cy) ** 2 <= r * r) { n++; if (d[y * W + x]) on++; }
    return n ? on / n : 0;
  }
  function detect(sheet, o = {}) {
    const L = o.layout || LAYOUT, minFill = o.minFill ?? .35, w = sheet.cols, h = sheet.rows;
    const g = new cv.Mat(), bin = new cv.Mat();
    const ch = new cv.MatVector(); cv.split(sheet, ch); cv.max(ch.get(0), ch.get(1), g); cv.max(g, ch.get(2), g); ch.delete(); // গোলাপি ছাপা রিং মুছে ফেলে
    cv.GaussianBlur(g, g, new cv.Size(5, 5), 0);
    const t = cv.threshold(g, bin, 0, 255, cv.THRESH_BINARY_INV + cv.THRESH_OTSU);
    cv.threshold(g, bin, Math.min(t, 140), 255, cv.THRESH_BINARY_INV);
    cv.morphologyEx(bin, bin, cv.MORPH_OPEN, cv.Mat.ones(3, 3, cv.CV_8U));
    // app.js থেকে আসা layout-এর অফসেট = ব্যবহারকারীর "গ্রিড সরান"; এটি অটো-ক্যালিব্রেশনের ওপর যোগ হয়
    const dx = o.layout ? (L.b1[0] - LAYOUT.b1[0]) * w : (o.shift ? o.shift.dx : 0), dy = o.layout ? (L.y0 - LAYOUT.y0) * h : (o.shift ? o.shift.dy : 0);
    const cal = o.noAuto ? null : calibrate(sheet); let G, diam, rp;
    let sp; if (cal) { G = cal.G; diam = cal.diam; rp = cal.roll; sp = cal.set; }
    else { G = grid(w, h, LAYOUT); diam = (L.b1[1] - L.b1[0]) * w * .85; sp = { x: LAYOUT.setx * w, ys: LAYOUT.sety.map(y => y * h), r: diam * .45 }; rp = { xs: LAYOUT.rx.map(x => x * w), ys: LAYOUT.ry.map(y => y * h), r: (LAYOUT.rx[1] - LAYOUT.rx[0]) * w * .3 }; }
    for (const q in G) G[q] = G[q].map(([x, y]) => [x + dx, y + dy]);
    rp = { xs: rp.xs.map(x => x + dx), ys: rp.ys.map(y => y + dy), r: rp.r };
    sp = { x: sp.x + dx, ys: sp.ys.map(y => y + dy), r: sp.r };
    const baseG = G; G = applyAdj(G, o.adj);
    const A = Math.PI * (diam / 2) ** 2, tol = diam * .75;
    const det = {}; Object.keys(G).forEach(q => det[q] = []);
    const cs = new cv.MatVector(), hi = new cv.Mat(); cv.findContours(bin, cs, hi, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE);
    for (let i = 0; i < cs.size(); i++) {
      const c = cs.get(i), a = cv.contourArea(c); if (a < A * .3 || a > A * 1.8) continue;
      const p = cv.arcLength(c, true), circ = 4 * Math.PI * a / (p * p + 1e-6), r = cv.boundingRect(c), ar = r.width / (r.height + 1e-6);
      if (circ < .55 || ar < .6 || ar > 1.6) continue;
      const m = cv.moments(c), cx = m.m10 / m.m00, cy = m.m01 / m.m00; let best = null;
      for (const q in G) G[q].forEach(([gx, gy], oi) => { const d = Math.hypot(cx - gx, cy - gy); if (d < tol && (!best || d < best[0])) best = [d, q, oi]; });
      if (best && fill(bin, ...G[best[1]][best[2]], diam * .4) >= minFill) det[best[1]].push(best[2]);
    }
    for (const q in G) if (!det[q].length)  // ফলব্যাক: সেলের কালো অনুপাত
      G[q].forEach(([gx, gy], oi) => { if (fill(bin, gx, gy, diam * .4) >= Math.max(minFill, .55)) det[q].push(oi); });
    const answers = {}, status = {};
    for (const q in G) {
      const m = det[q]; answers[q] = m.length === 1 ? OPT[m[0]] : null;
      status[q] = m.length === 1 ? 'OK' : m.length ? 'MULTIPLE' : 'BLANK';
    }
    let roll = '';  // রোল নম্বরের বাবল গ্রিড
    const rollFills = [];
    rp.xs.forEach(x => { const f = rp.ys.map(y => fill(bin, x, y, rp.r)), mx = Math.max(...f), i = f.indexOf(mx); rollFills.push(f);
      roll += mx >= .45 && f.filter(v => v >= .45).length === 1 ? i : '?'; });
    const sf = sp.ys.map(y => fill(bin, sp.x, y, diam * .4)), sm = sf.map((v, i) => v >= .45 ? i : -1).filter(i => i >= 0);  // সেট কোড
    const set = sm.length === 1 ? OPT[sm[0]] : null, setStatus = sm.length === 1 ? 'OK' : sm.length ? 'MULTIPLE' : 'BLANK';
    cs.delete(); hi.delete(); g.delete(); bin.delete();
    return { answers, status, grid: G, baseGrid: baseG, diam, roll, set, setStatus, rollFills, setPts: sp, rollPts: rp, calibrated: !!cal };
  }

  /* ---------- মূল্যায়ন ও মার্জ ---------- */
  function evaluate(ans, key, mark = 1, neg = 0) {
    let ok = 0, bad = 0, blank = 0; const rows = [];
    Object.keys(key).map(Number).sort((a, b) => a - b).forEach(q => {
      const s = ans[q] || '-', k = key[q]; let res;
      if (s === 'A' || s === 'B' || s === 'C' || s === 'D') { if (s === k) { res = 'CORRECT'; ok++; } else { res = 'WRONG'; bad++; } }
      else if (s === 'MULTI') { res = 'MULTIPLE (ভুল)'; bad++; } else { res = 'BLANK'; blank++; }
      rows.push([q, s, k, res]);
    });
    const total = Object.keys(key).length;
    return { rows, sum: { 'মোট প্রশ্ন': total, 'সঠিক': ok, 'ভুল': bad, 'ফাঁকা': blank, 'প্রাপ্ত নম্বর': +(ok * mark - bad * neg).toFixed(2), 'পূর্ণমান': +(total * mark).toFixed(2) } };
  }
  function merge(la, ls, ga, gs) {
    const fa = {}, fs = {}, review = [];
    for (const q in la) {
      if (la[q] === ga[q] && ls[q] === gs[q]) { fa[q] = la[q]; fs[q] = ls[q]; }
      else { review.push(+q); if (ls[q] === 'OK') { fa[q] = la[q]; fs[q] = 'OK'; } else { fa[q] = ga[q]; fs[q] = gs[q]; } }
    }
    return { fa, fs, review };
  }

  /* ---------- Gemini ---------- */
  const PROMPT = `You are an expert OMR (optical mark recognition) reader. The image is a Bangladeshi MCQ answer sheet.
Layout: 30 questions in two table columns. Left table = questions 1-15 (top to bottom), right table = questions 16-30.
Each question row has 4 bubbles in left-to-right order: ক (=A), খ (=B), গ (=C), ঘ (=D).
A bubble is FILLED only if solidly darkened with dark pen ink (the letter inside is hidden). Empty bubbles are printed pink rings with a visible pink letter - these are NOT answers. Ignore thin pen strokes, stray lines, the roll-number grid, registration grid, subject/set code, department and board grids.
For each question return exactly one of: "A","B","C","D" (one filled bubble), "MULTI" (two or more filled), or null (none filled).
Also read the handwritten header: student name, class, roll number (convert Bengali digits to English), department.
Return ONLY JSON: {"answers": {"1": "C", "2": null, ...}, "name": "", "class": "", "roll": "", "department": ""}`;
  async function gemini(jpegBase64, apiKey, model = 'gemini-2.5-flash', n = 30) {
    const body = { contents: [{ parts: [{ text: PROMPT }, { inline_data: { mime_type: 'image/jpeg', data: jpegBase64 } }] }], generationConfig: { temperature: 0, responseMimeType: 'application/json' } };
    let r, last;
    for (let i = 0; i < 3; i++) {
      r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`, { method: 'POST', headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey }, body: JSON.stringify(body) });
      if ([429, 500, 503].includes(r.status)) { last = (await r.text()).slice(0, 200); await new Promise(s => setTimeout(s, 2000 * (i + 1))); continue; }
      if (!r.ok) throw new Error(`Gemini ত্রুটি ${r.status}: ${(await r.text()).slice(0, 200)}`);
      last = null; break;
    }
    if (last) throw new Error('Gemini ব্যস্ত/লিমিট: ' + last);
    let t = (await r.json()).candidates[0].content.parts[0].text.replace(/^```(?:json)?|```$/gm, '').trim();
    const d = JSON.parse(t), answers = {}, status = {};
    for (let q = 1; q <= n; q++) {
      const v = (d.answers || {})[q];
      if (typeof v === 'string' && /^MULTI/i.test(v)) { answers[q] = null; status[q] = 'MULTIPLE'; }
      else if (normOpt(v)) { answers[q] = normOpt(v); status[q] = 'OK'; } else { answers[q] = null; status[q] = 'BLANK'; }
    }
    return { answers, status, info: { name: d.name || '', class: d.class || '', roll: enDigits(d.roll || ''), department: d.department || '' } };
  }
  const api = { parseKeySets, applyAdj, rollKey, parseStudents, findStudent, TW, TH, OPT, LAYOUT, parseKey, parseKeyRows, findQuad, autoScan, detect, evaluate, merge, gemini };
  if (typeof module !== 'undefined') module.exports = api;
  return api;
})();
