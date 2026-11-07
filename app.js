const $ = id => document.getElementById(id);
const S = { key: {}, students: [], batch: [], lastCam: false, mode: 'tap', cfg: { mode: 'hybrid', gkey: '', gmodel: 'gemini-2.5-flash', mark: 1, neg: 0, minfill: .35, auto: true, dx: 0, dy: 0 }, r: null };
let stream = null, loop = null, stable = 0, prevQ = null, busy = false, cvReady = false;
try { Object.assign(S.cfg, JSON.parse(localStorage.cfg || '{}')); S.keys = JSON.parse(localStorage.getItem('keys') || 'null') || { default: JSON.parse(localStorage.getItem('key') || '{}') }; S.keytxts = JSON.parse(localStorage.getItem('keytxts') || 'null') || { default: localStorage.keytxt || '' }; S.students = JSON.parse(localStorage.students || '[]'); S.batch = JSON.parse(localStorage.batch || '[]'); $('ptitle').value = localStorage.ptitle || ''; } catch (e) {}
const saveCfg = () => localStorage.cfg = JSON.stringify(S.cfg);
const toast = (t, ok) => { $('msg').innerHTML = t ? `<div class="toast" ${ok ? 'style="background:#16a34a"' : ''}>${t}</div>` : ''; if (ok) setTimeout(() => toast(''), 3000); };

/* ---- নেভিগেশন ---- */
document.querySelectorAll('nav button').forEach(b => b.onclick = () => show(b.dataset.v));
function show(v) { document.querySelectorAll('.view').forEach(e => e.classList.toggle('on', e.id === 'v-' + v)); document.querySelectorAll('nav button').forEach(b => b.classList.toggle('on', b.dataset.v === v)); if (v !== 'scan') camStop(); if (v === 'batch') batchRender(); }
const SET_BN = { A: 'ক', B: 'খ', C: 'গ', D: 'ঘ' }, hasKey = k => k && Object.keys(k).length > 0;
const keyFor = set => hasKey(S.keys[set]) ? S.keys[set] : (S.keys.default || {});   // সেটের নিজস্ব উত্তরমালা না থাকলে ডিফল্ট
const keyInfo = () => { const p = Object.entries(S.keys).filter(([, k]) => hasKey(k)).map(([s, k]) => `${s === 'default' ? 'ডিফল্ট' : 'সেট ' + SET_BN[s] + ' (' + s + ')'}: ${Object.keys(k).length} টি`); $('keyinfo').textContent = p.length ? '✅ সংরক্ষিত — ' + p.join(' · ') : ''; };
const keyTxtLoad = () => $('keytxt').value = S.keytxts[$('keyset').value] || '';
const keyPersist = () => { localStorage.setItem('keys', JSON.stringify(S.keys)); localStorage.setItem('keytxts', JSON.stringify(S.keytxts)); };
$('keyset').onchange = keyTxtLoad; keyTxtLoad();
keyInfo();

/* ---- সেটিংস ---- */
for (const k of ['mode', 'gkey', 'gmodel', 'mark', 'neg', 'minfill']) { $(k).value = S.cfg[k]; $(k).onchange = () => { S.cfg[k] = $(k).type === 'number' ? +$(k).value : $(k).value; saveCfg(); }; }
$('auto').checked = S.cfg.auto; $('auto').onchange = () => { S.cfg.auto = $('auto').checked; saveCfg(); };

/* ---- সঠিক উত্তর ---- */
$('keysave').onclick = () => { const ks = $('keyset').value; S.keys[ks] = OMR.parseKey($('keytxt').value); S.keytxts[ks] = $('keytxt').value; keyPersist(); keyInfo(); toast(hasKey(S.keys[ks]) ? 'সংরক্ষিত ✅' : 'কোনো উত্তর পাওয়া যায়নি', 1); };
const readRows = async f => { const wb = XLSX.read(await f.arrayBuffer()); return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: '' }); };
$('keyfile').onchange = async e => {
  const f = e.target.files[0]; if (!f) return; const rows = await readRows(f); e.target.value = '';   // একই ফাইল আবার বাছলেও কাজ করবে
  const fmt = k => Object.entries(k).map(([q, a]) => `${q}-${a}`).join(', '), m = OMR.parseKeySets(rows);
  if (m.multi) {
    for (const s in m.sets) { S.keys[s] = m.sets[s]; S.keytxts[s] = fmt(m.sets[s]); }
    keyTxtLoad(); keyPersist(); keyInfo();
    toast(`✅ ${Object.keys(m.sets).map(s => s === 'default' ? 'ডিফল্ট' : 'সেট ' + SET_BN[s]).join(', ')} — ${Object.keys(m.sets).length} টি উত্তরমালা লোড হয়েছে`, 1);
  } else {
    const ks = $('keyset').value; S.keys[ks] = m.sets.single; S.keytxts[ks] = fmt(m.sets.single); keyTxtLoad(); keyPersist(); keyInfo();
    toast(hasKey(S.keys[ks]) ? `✅ ${Object.keys(S.keys[ks]).length} টি উত্তর লোড হয়েছে` : 'ফাইলে কোনো উত্তর পাওয়া যায়নি (কলাম: প্রশ্ন, উত্তর)', 1);
  }
};

/* ---- শিক্ষার্থী তালিকা ---- */
function stRefresh() {
  $('stinfo').textContent = S.students.length ? `✅ ${S.students.length} জন শিক্ষার্থী লোড হয়েছে` : '';
  $('stpick').innerHTML = '<option value="">— তালিকা থেকে শিক্ষার্থী বাছুন —</option>' + S.students.map((s, i) => `<option value="${i}">${s.roll} — ${s.name}</option>`).join('');
  $('stpick').hidden = !S.students.length;
}
$('stfile').onchange = async e => {
  const f = e.target.files[0]; if (!f) return; S.students = OMR.parseStudents(await readRows(f)); localStorage.students = JSON.stringify(S.students); stRefresh();
  toast(S.students.length ? `${S.students.length} জন শিক্ষার্থী যোগ হয়েছে` : 'তালিকা পড়া যায়নি (কলাম: নাম, শ্রেণি, রোল, গ্রুপ)', 1); e.target.value = '';
};
$('stclear').onclick = () => { S.students = []; localStorage.students = '[]'; stRefresh(); };
const fillStudent = s => { $('i-name').value = s.name; $('i-class').value = s.class; $('i-roll').value = s.roll; $('i-dept').value = s.group; };
$('stpick').onchange = () => { const s = S.students[$('stpick').value]; if (s) fillStudent(s); };
$('i-roll').onchange = () => { const s = OMR.findStudent(S.students, $('i-roll').value); if (s) { fillStudent(s); toast(`রোল ${s.roll} মিলেছে: ${s.name}`, 1); } };
stRefresh();

/* ---- OpenCV প্রস্তুত? (opencv.js ব্রাউজারে Promise হিসেবে লোড হয়) ---- */
const cvDone = m => { if (m && m !== window.cv) window.cv = m; cvReady = true; $('cvst').textContent = '✔ প্রস্তুত'; toast(''); };
const waitCv = setInterval(() => {
  if (!window.cv) return; clearInterval(waitCv);
  if (typeof cv.then === 'function') cv.then(m => { delete m.then; cvDone(m); });
  else if (cv.Mat) cvDone(); else cv.onRuntimeInitialized = () => cvDone();
}, 200);

/* ---- ক্যামেরা ও লাইভ ডিটেকশন ---- */
$('camon').onclick = async () => {
  try {
    stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' }, width: { ideal: 1920 }, height: { ideal: 1080 } }, audio: false });
    $('vid').srcObject = stream; await $('vid').play(); $('camwrap').hidden = $('camctl').hidden = false; stable = 0; prevQ = null; loop = setInterval(live, 400);
  } catch (e) { toast('ক্যামেরা চালু হয়নি (HTTPS লাগে ও অনুমতি দিতে হবে): ' + e.message); }
};
function camStop() { clearInterval(loop); loop = null; if (stream) stream.getTracks().forEach(t => t.stop()); stream = null; $('camwrap').hidden = $('camctl').hidden = true; }
$('camoff').onclick = camStop;
const grab = (src, w, h, max) => { const s = Math.min(1, max / Math.max(w, h)), c = document.createElement('canvas'); c.width = Math.round(w * s); c.height = Math.round(h * s); c.getContext('2d').drawImage(src, 0, 0, c.width, c.height); return c; };
function live() {
  if (!cvReady || busy || !stream) return; const v = $('vid'); if (!v.videoWidth) return;
  const c = grab(v, v.videoWidth, v.videoHeight, 640), src = cv.imread(c), q = OMR.findQuad(src, 640); src.delete();
  const ov = $('ov'); ov.width = c.width; ov.height = c.height; const x = ov.getContext('2d'); x.clearRect(0, 0, ov.width, ov.height);
  if (q) {
    x.strokeStyle = '#22c55e'; x.lineWidth = 5; x.beginPath(); q.forEach((p, i) => i ? x.lineTo(p.x, p.y) : x.moveTo(p.x, p.y)); x.closePath(); x.stroke();
    const cx = q.reduce((a, p) => a + p.x, 0) / 4, cy = q.reduce((a, p) => a + p.y, 0) / 4, d = Math.hypot(c.width, c.height);
    stable = prevQ && Math.hypot(cx - prevQ[0], cy - prevQ[1]) < d * .02 ? stable + 1 : 0; prevQ = [cx, cy];
    $('scanst').textContent = stable >= 1 ? '✅ কাগজ পেয়েছি — স্থির ধরে রাখুন…' : 'কাগজ পেয়েছি';
    if (S.cfg.auto && stable >= 3) snap();
  } else { stable = 0; prevQ = null; $('scanst').textContent = 'কাগজ খুঁজছি… পুরো কাগজ ও চার কোণা ফ্রেমে আনুন'; }
}
$('snap').onclick = () => snap();
function snap() { const v = $('vid'); if (!v.videoWidth) return; const c = grab(v, v.videoWidth, v.videoHeight, 2400); camStop(); S.lastCam = true; process(c); }
$('imgfile').onchange = async e => {
  const f = e.target.files[0]; if (!f) return; const bm = await createImageBitmap(f, { imageOrientation: 'from-image' });
  S.lastCam = false; process(grab(bm, bm.width, bm.height, 2400)); e.target.value = '';
};

/* ---- প্রসেসিং ---- */
const layout = () => { const L = OMR.LAYOUT, w = OMR.TW, h = OMR.TH, dx = S.cfg.dx / w, dy = S.cfg.dy / h;  // সংরক্ষিত গ্রিড-সরানো অফসেট
  return { ...L, b1: L.b1.map(x => x + dx), b2: L.b2.map(x => x + dx), rx: L.rx.map(x => x + dx), ry: L.ry.map(y => y + dy), y0: L.y0 + dy, y1: L.y1 + dy }; };
function applyEngines(r, loc) {  // লোকাল (+Gemini) মিলিয়ে চূড়ান্ত উত্তর বানায়
  r.loc = loc; if (!r.setEdited) r.set = loc.set || null; r.grid = loc.grid; r.baseGrid = loc.baseGrid; r.diam = loc.diam; r.manual = r.manual || {}; r.fa = loc.answers; r.fs = loc.status; r.review = [];
  if (r.ga) { if (S.cfg.mode === 'gemini') { r.fa = r.ga.answers; r.fs = r.ga.status; } else { const m = OMR.merge(loc.answers, loc.status, r.ga.answers, r.ga.status); r.fa = m.fa; r.fs = m.fs; r.review = m.review; } }
  r.final = {}; for (const q in r.fa) r.final[q] = r.manual[q] ?? (r.fa[q] || (r.fs[q] === 'MULTIPLE' ? 'MULTI' : '-'));
}
async function process(canvas) {
  if (!cvReady) return toast('OpenCV এখনো লোড হচ্ছে, একটু অপেক্ষা করুন…');
  if (!Object.values(S.keys).some(hasKey)) { toast('আগে "উত্তর" ট্যাবে সঠিক উত্তর দিন'); return show('key'); }
  if (busy) return; busy = true; toast('প্রসেসিং চলছে…');
  try {
    const src = cv.imread(canvas), { mat, found } = OMR.autoScan(src); src.delete();
    const loc = OMR.detect(mat, { minFill: S.cfg.minfill, layout: layout() }), sc = document.createElement('canvas'); cv.imshow(sc, mat); mat.delete();
    const r = { sheet: sc, found, info: {}, ga: null }; 
    if (S.cfg.mode !== 'local' && S.cfg.gkey) {
      try { const g = grab(sc, sc.width, sc.height, 2000); r.ga = await OMR.gemini(g.toDataURL('image/jpeg', .92).split(',')[1], S.cfg.gkey, S.cfg.gmodel); r.info = { ...r.ga.info }; }
      catch (e) { toast('Gemini ব্যর্থ — লোকাল ফলাফল দেখানো হচ্ছে। ' + e.message); }
    } else if (S.cfg.mode !== 'local') toast('Gemini key নেই — শুধু লোকাল ইঞ্জিন ব্যবহৃত হলো');
    applyEngines(r, loc); matchStudent(r, loc.roll);
    S.r = r; render(true);
    if (!r.found) toast('⚠️ কাগজের কোণা শনাক্ত হয়নি — গাঢ় ব্যাকগ্রাউন্ডে পুরো কাগজ দেখিয়ে আবার তুলুন।');
  } catch (e) { toast('ত্রুটি: ' + e.message); console.error(e); } finally { busy = false; }
}
function matchStudent(r, bubbleRoll) {  // রোল বাবল → Gemini রোল, এই ক্রমে তালিকায় খোঁজা
  const cands = [bubbleRoll && !bubbleRoll.includes('?') ? bubbleRoll : '', r.info.roll].filter(Boolean);
  r.rollHint = cands.length ? `শিট থেকে পড়া রোল: বাবল ${bubbleRoll || '-'}${r.info.roll ? ' · হাতের লেখা ' + r.info.roll : ''}` : '';
  for (const c of cands) { const s = OMR.findStudent(S.students, c); if (s) { r.info = { name: s.name, class: s.class, roll: s.roll, department: s.group }; toast(`রোল ${s.roll} মিলেছে: ${s.name}`, 1); return; } }
  if (cands.length) { r.info.roll = r.info.roll || String(parseInt(cands[0], 10) || cands[0]); if (S.students.length) toast(`রোল ${r.info.roll} তালিকায় পাওয়া যায়নি — তালিকা থেকে বেছে নিন`); }
}
function redetect() {  // গ্রিড সরানোর পর আবার পড়া
  const r = S.r, m = cv.imread(r.sheet), loc = OMR.detect(m, { minFill: S.cfg.minfill, layout: layout(), adj: r.adj }); m.delete();
  applyEngines(r, loc); render(false);
}

/* ---- ফলাফল রেন্ডার ---- */
function render(full, shift, tmpAdj) {
  const r = S.r; if (!r) return; $('resempty').hidden = true; $('resbody').hidden = false;
  const KEY = keyFor(r.set), { rows, sum } = OMR.evaluate(r.final, KEY, S.cfg.mark, S.cfg.neg); r.rows = rows; r.sum = sum;
  $('s-m').textContent = `${sum['প্রাপ্ত নম্বর']}/${sum['পূর্ণমান']}`; $('s-ok').textContent = sum['সঠিক']; $('s-bad').textContent = sum['ভুল']; $('s-bl').textContent = sum['ফাঁকা'];
  const c = $('sheet'); c.width = r.sheet.width; c.height = r.sheet.height; const x = c.getContext('2d'); x.drawImage(r.sheet, 0, 0);
  const GG = tmpAdj ? OMR.applyAdj(r.baseGrid, tmpAdj) : r.grid, tx = shift ? shift.tx : 0, ty = shift ? shift.ty : 0, R = r.diam * .45, ring = (q, i, col, lw) => { x.strokeStyle = col; x.lineWidth = lw; x.beginPath(); x.arc(GG[q][i][0] + tx, GG[q][i][1] + ty, R, 0, 7); x.stroke(); };
  for (const q in r.grid) {
    if (S.mode !== 'tap') for (let i = 0; i < 4; i++) ring(q, i, 'rgba(59,130,246,.9)', 3);  // সারিবদ্ধতা দেখার গাইড
    const k = KEY[q], f = r.final[q];
    if (OMR.OPT.includes(k)) ring(q, OMR.OPT.indexOf(k), '#16a34a', 6);
    if (OMR.OPT.includes(f) && f !== k) ring(q, OMR.OPT.indexOf(f), '#dc2626', 6);
    else if (f === k && OMR.OPT.includes(f)) ring(q, OMR.OPT.indexOf(f), '#16a34a', 10);
  }
  $('i-set').value = r.set || '-';
  { const used = hasKey(S.keys[r.set]) ? 'সেট ' + SET_BN[r.set] + '-এর উত্তরমালা' : hasKey(S.keys.default) ? 'ডিফল্ট উত্তরমালা' : null, st = r.loc.setStatus;
    $('sethint').textContent = !used ? `⚠️ ${r.set ? 'সেট ' + SET_BN[r.set] + '-এর' : 'এই শিটের'} উত্তরমালা নেই — ড্রপডাউন থেকে সঠিক সেট বাছুন` : r.setEdited ? `হাতে বাছা — ${used} ব্যবহৃত` : st === 'OK' ? `শিট থেকে পড়া: ${SET_BN[r.set]} (${r.set}) — ${used} ব্যবহৃত` : st === 'MULTIPLE' ? 'একাধিক সেট ভরাট — হাতে বাছুন' : `সেট কোড ভরাট নেই — ${used} ব্যবহৃত`; }
  $('tbl').innerHTML = '<tr><th>#</th><th>উত্তর</th><th>সঠিক</th><th>ফল</th></tr>' + rows.map(([q, s, k, res]) =>
    `<tr class="${r.review.includes(q) ? 'rev' : ''}"><td>${q}</td><td><select data-q="${q}">${['A', 'B', 'C', 'D', '-', 'MULTI'].map(o => `<option ${o === s ? 'selected' : ''}>${o}</option>`).join('')}</select></td><td>${k}</td><td class="${res.split(' ')[0]}">${res}</td></tr>`).join('');
  $('tbl').querySelectorAll('select').forEach(s => s.onchange = () => { r.final[s.dataset.q] = r.manual[s.dataset.q] = s.value; render(); });
  $('revbox').hidden = !r.review.length; $('crops').innerHTML = '';
  r.review.forEach(q => {
    const g = r.grid[q], d = r.diam, cv2 = document.createElement('canvas'), x0 = g[0][0] - d, w = g[3][0] - g[0][0] + 2 * d;
    cv2.width = w; cv2.height = 2 * d; cv2.getContext('2d').drawImage(r.sheet, x0, g[0][1] - d, w, 2 * d, 0, 0, w, 2 * d);
    const f = document.createElement('figure'); f.append(cv2); f.insertAdjacentHTML('beforeend', `<figcaption class="mut">প্রশ্ন ${q} · লোকাল: ${r.loc.answers[q] || '-'} · Gemini: ${r.ga.answers[q] || '-'}</figcaption>`); $('crops').append(f);
  });
  if (full) {
    for (const [id, k] of [['i-name', 'name'], ['i-class', 'class'], ['i-roll', 'roll'], ['i-dept', 'department']]) $(id).value = r.info[k] || '';
    $('rollhint').textContent = r.rollHint || ''; $('stpick').value = ''; show('res');
  }
}

/* ---- ক্যানভাসে ট্যাপ করে উত্তর বদলানো / বাবল-গ্রিড মাউস বা আঙুলে টেনে ঠিক করা ---- */
const cvs = $('sheet'); let drag = null;
const HINT = {
  tap: '👆 কোনো প্রশ্নের বৃত্তে ট্যাপ/ক্লিক করলে সেটি উত্তর হবে (আবার চাপলে মুছে যাবে)।',
  all: '↔️ নীল বৃত্তগুলো ছাপা বৃত্তের সাথে না মিললে শিটে যেকোনো জায়গা ধরে টানুন — পুরো গ্রিড একসাথে সরবে (পরের স্ক্যানেও থাকবে)।',
  ends: '↕️ গ্রিডের ওপরের অংশের কোনো বাবল ধরে টানলে ওপরের দিক, নীচের অংশ ধরলে নীচের দিক সরবে; মাঝের সারিগুলো আপনাআপনি অনুপাতে সরবে (স্কেল/ঢাল ঠিক করতে)।',
  row: '▬ কোনো একটি প্রশ্নের বাবল ধরে টানুন — ওই প্রশ্নের ৪টি বাবল একসাথে সরবে।',
  bubble: '● যেকোনো একটি বাবল ধরে টানুন — শুধু সেটি সরবে।',
};
const setMode = m => { S.mode = m; document.querySelectorAll('#modes [data-m]').forEach(b => b.className = b.dataset.m === m ? '' : 'sec'); cvs.style.touchAction = m === 'tap' ? 'pan-y' : 'none'; $('canhint').textContent = HINT[m]; if (S.r) render(); };
document.querySelectorAll('#modes [data-m]').forEach(b => b.onclick = () => setMode(b.dataset.m)); setMode('tap');
$('i-set').onchange = () => { const r = S.r; if (!r) return; r.set = $('i-set').value === '-' ? null : $('i-set').value; r.setEdited = true; render(); };
$('resetgrid').onclick = () => { S.cfg.dx = S.cfg.dy = 0; saveCfg(); if (S.r) { S.r.adj = null; redetect(); toast('গ্রিড রিসেট হয়েছে', 1); } };
const k = () => cvs.width / cvs.getBoundingClientRect().width;
const cloneAdj = a => JSON.parse(JSON.stringify(a || { ends: {}, rows: {}, bubbles: {} }));
const nearest = (px, py, tol) => { const r = S.r; let best = null; for (const q in r.grid) r.grid[q].forEach(([gx, gy], i) => { const d = Math.hypot(px - gx, py - gy); if (d < tol && (!best || d < best[0])) best = [d, +q, i]; }); return best; };
const tmpFor = d => {  // টানার সময়ের সমন্বয় (প্রিভিউ ও ছাড়ার পর কমিট দুটোতেই একই)
  const t = cloneAdj(S.r.adj), add = (o, key) => { const p = o[key] || [0, 0]; o[key] = [p[0] + d.tx, p[1] + d.ty]; };
  if (d.t.kind === 'end') { t.ends[d.t.b] = t.ends[d.t.b] || {}; add(t.ends[d.t.b], d.t.part); }
  else if (d.t.kind === 'row') add(t.rows, d.t.q); else if (d.t.kind === 'bubble') add(t.bubbles, d.t.q + '_' + d.t.i);
  return t;
};
cvs.onpointerdown = e => {
  const b = cvs.getBoundingClientRect(), px = (e.clientX - b.left) * k(), py = (e.clientY - b.top) * k(), r = S.r; if (!r) return;
  let t = null;
  if (S.mode === 'all') t = { kind: 'all' };
  else if (S.mode !== 'tap') {
    const n = nearest(px, py, r.diam * 1.6); if (!n) return;
    const q = n[1], rows = OMR.LAYOUT.rows, row = (q - 1) % rows;
    t = S.mode === 'ends' ? { kind: 'end', b: q > rows ? 1 : 0, part: row / (rows - 1) < .5 ? 'top' : 'bottom' } : S.mode === 'row' ? { kind: 'row', q } : { kind: 'bubble', q, i: n[2] };
  }
  drag = { x: e.clientX, y: e.clientY, tx: 0, ty: 0, moved: false, t }; cvs.setPointerCapture(e.pointerId);
};
cvs.onpointermove = e => {
  if (!drag) return; drag.tx = (e.clientX - drag.x) * k(); drag.ty = (e.clientY - drag.y) * k();
  if (Math.hypot(e.clientX - drag.x, e.clientY - drag.y) > 6) drag.moved = true;
  if (S.mode !== 'tap' && drag.t && drag.moved) { if (drag.t.kind === 'all') render(false, drag); else render(false, null, tmpFor(drag)); }
};
cvs.onpointerup = e => {
  const d = drag; drag = null; if (!d || !S.r) return; const r = S.r;
  if (S.mode !== 'tap') {
    if (!d.t || !d.moved) return;
    if (d.t.kind === 'all') { S.cfg.dx += d.tx; S.cfg.dy += d.ty; saveCfg(); } else r.adj = tmpFor(d);
    redetect(); toast(d.t.kind === 'all' ? 'গ্রিড সরিয়ে আবার পড়া হয়েছে (পরের স্ক্যানেও এই অবস্থান থাকবে)' : 'বাবল সরিয়ে আবার পড়া হয়েছে — হাতে ঠিক করা উত্তর অপরিবর্তিত', 1); return;
  }
  if (d.moved) return;
  const b = cvs.getBoundingClientRect(), n = nearest((e.clientX - b.left) * k(), (e.clientY - b.top) * k(), r.diam * .7); if (!n) return;
  const o = OMR.OPT[n[2]]; r.final[n[1]] = r.manual[n[1]] = r.final[n[1]] === o ? '-' : o; render();
};
cvs.onpointercancel = () => { drag = null; };

/* ---- জমাকৃত ফলাফল সীট (একের পর এক স্ক্যান → প্রিন্ট/Excel) ---- */
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const saveBatch = () => localStorage.batch = JSON.stringify(S.batch);
const sortedBatch = () => [...S.batch].sort((a, b) => String(a.roll).localeCompare(String(b.roll), 'en', { numeric: true }));
const sheetTitle = () => $('ptitle').value.trim() || 'ফলাফল সীট';
function batchRender() {
  const list = sortedBatch(), tot = Math.max(0, ...S.batch.map(x => x.total || 0)), hs = S.batch.some(x => x.set);
  $('bcnt').textContent = S.batch.length ? ` (${S.batch.length})` : ''; $('bempty').hidden = !!list.length; $('ptbl').hidden = !list.length;
  $('phead').textContent = sheetTitle();
  $('pmeta').textContent = `তারিখ: ${new Date().toLocaleDateString('bn-BD')}${tot ? ' · পূর্ণমান: ' + tot : ''} · মোট শিক্ষার্থী: ${list.length}`;
  $('ptbl').innerHTML = '<tr><th>ক্রম</th><th>নাম</th><th>শ্রেণি</th><th>রোল</th><th>গ্রুপ</th>' + (hs ? '<th>সেট</th>' : '') + '<th>প্রাপ্ত নম্বর</th><th class="noprint"></th></tr>' +
    list.map((x, i) => `<tr><td>${i + 1}</td><td class="l">${esc(x.name)}</td><td>${esc(x.class)}</td><td>${esc(x.roll)}</td><td>${esc(x.group)}</td>${hs ? `<td>${x.set ? SET_BN[x.set] : ''}</td>` : ''}<td><b>${x.marks}</b></td><td class="noprint"><button class="sec" data-del="${S.batch.indexOf(x)}">✕</button></td></tr>`).join('');
  $('ptbl').querySelectorAll('[data-del]').forEach(b => b.onclick = () => { S.batch.splice(+b.dataset.del, 1); saveBatch(); batchRender(); });
}
$('addres').onclick = () => {
  const r = S.r; if (!r) return;
  const e = { name: $('i-name').value.trim(), class: $('i-class').value.trim(), roll: $('i-roll').value.trim(), group: $('i-dept').value.trim(), set: r.set || '', marks: r.sum['প্রাপ্ত নম্বর'], total: r.sum['পূর্ণমান'] };
  if (!e.roll && !e.name) return toast('আগে রোল বা নাম দিন');
  const i = e.roll ? S.batch.findIndex(x => OMR.rollKey(x.roll) === OMR.rollKey(e.roll)) : -1;
  if (i >= 0) S.batch[i] = e; else S.batch.push(e);
  saveBatch(); batchRender(); S.r = null; $('resbody').hidden = true; $('resempty').hidden = false;
  toast(`✅ ${e.name || e.roll} — ${e.marks} নম্বর ${i >= 0 ? 'হালনাগাদ' : 'জমা'} হয়েছে (মোট ${S.batch.length} জন)। পরবর্তী শিট স্ক্যান করুন`, 1);
  show('scan'); if (S.lastCam) $('camon').click();
};
$('ptitle').oninput = () => { localStorage.ptitle = $('ptitle').value; batchRender(); };
$('bprint').onclick = () => { if (!S.batch.length) return toast('কোনো ফলাফল জমা নেই'); batchRender(); window.print(); };
$('bclr').onclick = () => { if (S.batch.length && confirm('সব জমাকৃত ফলাফল মুছে ফেলবেন?')) { S.batch = []; saveBatch(); batchRender(); } };
$('bxl').onclick = () => {
  if (!S.batch.length) return toast('কোনো ফলাফল জমা নেই');
  const hs = S.batch.some(x => x.set), aoa = [[sheetTitle()], [], ['ক্রম', 'নাম', 'শ্রেণি', 'রোল', 'গ্রুপ', ...(hs ? ['সেট'] : []), 'প্রাপ্ত নম্বর'], ...sortedBatch().map((x, i) => [i + 1, x.name, x.class, x.roll, x.group, ...(hs ? [x.set ? SET_BN[x.set] : ''] : []), x.marks]), [], ['ManHa IT']];
  const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = [{ wch: 7 }, { wch: 26 }, { wch: 12 }, { wch: 12 }, { wch: 18 }, { wch: 14 }];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'ফলাফল'); XLSX.writeFile(wb, 'result_sheet.xlsx');
};
batchRender();

/* ---- Excel ---- */
$('xl').onclick = () => {
  const r = S.r; if (!r) return;
  const info = [['বিবরণ', 'মান'], ['নাম', $('i-name').value], ['শ্রেণি', $('i-class').value], ['রোল নং', $('i-roll').value], ['গ্রুপ', $('i-dept').value], ['সেট', r.set ? SET_BN[r.set] + ' (' + r.set + ')' : ''], ...Object.entries(r.sum)];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(info), 'সারাংশ');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['প্রশ্ন', 'শিক্ষার্থীর উত্তর', 'সঠিক উত্তর', 'ফলাফল'], ...r.rows]), 'বিস্তারিত');
  wb['!cols'] = [{ wch: 20 }, { wch: 20 }]; XLSX.writeFile(wb, `result_${$('i-roll').value || 'student'}.xlsx`);
};

if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
