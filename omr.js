/* OMR কোর (OpenCV.js) — ব্রাউজার ও Node দুটোতেই চলে। গ্লোবাল `cv` দরকার। */
const OMR = (() => {
  const TW = 1600, TH = 2211, OPT = ['A', 'B', 'C', 'D'];
  const LAYOUT = { b1: [.141, .182, .222, .263], b2: [.355, .395, .436, .477], y0: .309, y1: .928, rows: 15,
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

  /* ---------- একাধিক সেট (ক/খ/গ/ঘ) এক Excel-এ ---------- */
  const normSet = s => normOpt(String(s == null ? '' : s).replace(/সেট|কোড|set|code|[\s_\-:]/gi, ''));
  // sheets = [{name, rows}] → { A:{1:'C',…}, B:{…}, C:{…}, D:{…} } (যে সেট পাওয়া গেছে শুধু সেগুলো)
  function parseKeySets(sheets) {
    const res = {}, put = (set, key) => { if (set && Object.keys(key).length) res[set] = Object.assign(res[set] || {}, key); };
    for (const { name, rows } of sheets) {
      let hi = -1, cols = {};
      for (let i = 0; i < Math.min(rows.length, 8) && hi < 0; i++) {   // হেডার সারি: কমপক্ষে ২টি সেট-কলাম
        const m = {}; (rows[i] || []).forEach((c, j) => { const o = normSet(c); if (o && !(o in m)) m[o] = j; });
        if (Object.keys(m).length >= 2 && !parseInt(enDigits((rows[i] || [])[0]))) { hi = i; cols = m; }
      }
      if (hi >= 0) {                                                    // একটি শিটে একাধিক সেট-কলাম
        const qc = Math.max(0, Math.min(...Object.values(cols)) - 1);
        for (const set in cols) { const key = {};
          rows.slice(hi + 1).forEach(r => { const q = parseInt(enDigits(r[qc])), o = normOpt(r[cols[set]]); if (q && o) key[q] = o; });
          put(set, key); }
        continue;
      }
      const sn = normSet(name);
      if (sn) { put(sn, parseKeyRows(rows)); continue; }               // শিটের নামই সেট (ক / Set A …)
      const wide = rows.some(r => r.length >= 5 && parseInt(enDigits(r[0])) && normOpt(r[1]) && normOpt(r[4]));
      if (wide) ['A', 'B', 'C', 'D'].forEach((o, j) => { const key = {};   // হেডার ছাড়া: প্রশ্ন | ক | খ | গ | ঘ
        rows.forEach(r => { const q = parseInt(enDigits(r[0])), a = normOpt(r[j + 1]); if (q && a) key[q] = a; }); put(o, key); });
      else put('A', parseKeyRows(rows));                                // শুধু প্রশ্ন|উত্তর → ক সেট
    }
    return res;
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
  function autoScan(src) {
    let quad = findQuad(src), dst = new cv.Mat();
    if (!quad) { cv.resize(src, dst, new cv.Size(TW, TH), 0, 0, cv.INTER_AREA); return { mat: dst, found: false }; }
    const [tl, tr, br, bl] = quad;
    if (Math.hypot(tr.x - tl.x, tr.y - tl.y) > Math.hypot(bl.x - tl.x, bl.y - tl.y)) quad = [bl, tl, tr, br];
    const s = cv.matFromArray(4, 1, cv.CV_32FC2, quad.flatMap(p => [p.x, p.y]));
    const d = cv.matFromArray(4, 1, cv.CV_32FC2, [0, 0, TW - 1, 0, TW - 1, TH - 1, 0, TH - 1]);
    const M = cv.getPerspectiveTransform(s, d);
    cv.warpPerspective(src, dst, M, new cv.Size(TW, TH), cv.INTER_CUBIC, cv.BORDER_CONSTANT, new cv.Scalar());
    s.delete(); d.delete(); M.delete(); return { mat: dst, found: true };
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
    const G = grid(w, h, L), diam = (L.b1[1] - L.b1[0]) * w * .85, A = Math.PI * (diam / 2) ** 2, tol = diam * .75;
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
    const rr = (L.rx[1] - L.rx[0]) * w * .3; let roll = '';  // রোল নম্বরের বাবল গ্রিড
    L.rx.forEach(x => { const f = L.ry.map(y => fill(bin, x * w, y * h, rr)), mx = Math.max(...f), i = f.indexOf(mx);
      roll += mx >= .45 && f.filter(v => v >= .45).length === 1 ? i : '?'; });
    cs.delete(); hi.delete(); g.delete(); bin.delete();
    return { answers, status, grid: G, diam, roll };
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
  const api = { rollKey, parseStudents, findStudent, TW, TH, OPT, LAYOUT, parseKey, parseKeyRows, parseKeySets, normSet, findQuad, autoScan, detect, evaluate, merge, gemini };
  if (typeof module !== 'undefined') module.exports = api;
  return api;
})();
