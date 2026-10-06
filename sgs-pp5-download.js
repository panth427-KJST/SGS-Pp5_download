/*!
 * KJST · ดาวน์โหลด ปพ.5 จาก SGS (sgs-pp5-download.js) v1.2 — 6 ต.ค. 2569
 * เฟส 13 KJST e-Score · เครื่องมือ admin (bookmarklet)
 * v1.2 (ภาค 17): เลือกชนิดไฟล์ "7. ผลการเรียน ปพ.5" (ตรวจวิธีที่ 1) หรือ "48. คะแนนรายวิชา" (ตรวจวิธีที่ 2)
 *   หาเมนูจาก value ขึ้นต้น "48," หรือข้อความขึ้นต้น "48." · ตรวจหัวตารางตามชนิดไฟล์ · จำตัวเลือกใน kjst_pp5_opts.kind
 *
 * ทำงานบนหน้าใดก็ได้ของ SGS ที่ล็อกอินแล้ว (แนะนำหน้า สารสนเทศ /sgs/TblClassRoom/Universal.aspx)
 * ไม่คลิกบนหน้าจริง — จำลองการส่งฟอร์ม ASP.NET ด้วย fetch ทีละขั้น (หน้าไม่โหลดใหม่ กล่องไม่หาย):
 *   GET Universal.aspx → เลือก "7. ผลการเรียน ปพ.5" (DropDownListMode)
 *   → dropdown "ระดับชั้น" (ม.1–ม.6) + "ห้องที่" (1–7) เหนือตาราง ใน UpdatePanel1
 *   ⚠ ไม่แตะตัวเลือกชั้นบนหัวหน้า SGS (_PageHeader$_DropDownListLevel) — เปลี่ยนแล้ว SGS ออกจากเมนูไปหน้าหลัก
 *   → POST ปุ่ม Excel (ctl00$PageContent$ButtonExcel$_Button) → ได้ไฟล์ .xls
 *   → ตรวจไฟล์ (ตาราง MyGrid · ชั้น/ห้อง · ปี/ภาค) → เขียนลงโฟลเดอร์ที่เลือก (File System Access API)
 * ชื่อไฟล์ = ชื่อที่ SGS ตั้ง (Content-Disposition) · ไฟล์ชื่อซ้ำเขียนทับ
 */
(function () {
  'use strict';
  var VER = '1.2';
  if (window.__kjstPp5 && window.__kjstPp5.show) { window.__kjstPp5.show(); return; }

  // ---------- ค่าตั้งต้น ----------
  var ROOMS = { 1: 7, 2: 7, 3: 7, 4: 6, 5: 6, 6: 6 };   // ม.1–3 ชั้นละ 7 · ม.4–6 ชั้นละ 6 = 39 ห้อง
  // ชนิดไฟล์ (เมนูข้อมูล DropDownListMode) · prefix = ต้น value ("7,Usp_Stat") · text = ต้นข้อความ option (สำรอง)
  // need = หัวคอลัมน์ที่ต้องมี · deny = หัวคอลัมน์ที่ต้องไม่มี (กันไฟล์ผิดชนิด)
  var KINDS = {
    '7':  { prefix: '7,',  text: /^7\./,  label: '7. ผลการเรียน ปพ.5', fb: '7.ผลการเรียนปพ.5', need: ['กลุ่มที่'], deny: [] },
    '48': { prefix: '48,', text: /^48\./, label: '48. คะแนนรายวิชา',   fb: '48. คะแนนรายวิชา', need: ['รวมกลางภาค', 'S1'], deny: ['กลุ่มที่'] }
  };
  function K() { return KINDS[S.opts.kind] || KINDS['7']; }
  var ID_YEAR = 'ctl00__PageHeader__DropDownListYr';
  var ID_TERM = 'ctl00__PageHeader__DropDownListTr';
  var ID_MODE = 'ctl00_PageContent_DropDownListMode';
  var ID_PANEL = 'ctl00_PageContent_UpdatePanel1';
  var EXCEL_TARGET = 'ctl00$PageContent$ButtonExcel$_Button';
  var FETCH_TIMEOUT = 180000;   // server ช้าช่วงคนใช้เยอะ
  var MAX_TRY = 3;
  var SPEED = { normal: 400, slow: 2000 };
  var IDB_NAME = 'kjst_pp5_dl', IDB_STORE = 'h', LS_OPTS = 'kjst_pp5_opts';

  var S = {
    url: '', doc: null, yr: '', tr: '',
    dir: null, running: false, stop: false,
    rooms: {},            // key "g-r" → {g, r, st:'wait|run|ok|warn|fail|skip', msg, file, rows, students, subjects, bytes}
    opts: loadOpts()
  };

  // ---------- util ----------
  function $(id) { return document.getElementById(id); }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function norm(s) { return String(s == null ? '' : s).replace(/[\s\u00a0]+/g, '').trim(); }
  function kb(n) { return n >= 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.round(n / 1024) + ' KB'; }
  function loadOpts() {
    var o = { grades: [1, 2, 3, 4, 5, 6], speed: 'normal', kind: '7' };
    try { var j = JSON.parse(localStorage.getItem(LS_OPTS) || 'null'); if (j) { if (Array.isArray(j.grades)) o.grades = j.grades; if (SPEED[j.speed]) o.speed = j.speed; if (j.kind === '7' || j.kind === '48') o.kind = j.kind; } } catch (e) {}
    return o;
  }
  function saveOpts() { try { localStorage.setItem(LS_OPTS, JSON.stringify(S.opts)); } catch (e) {} }
  function stepWait() { return sleep(SPEED[S.opts.speed] || SPEED.normal); }

  function pageUrl() {
    var p = location.pathname;
    if (/\/TblClassRoom\/Universal\.aspx$/i.test(p)) return location.origin + p;
    var base = (p.match(/^\/[^\/]+\//) || ['/sgs/'])[0];
    return location.origin + base + 'TblClassRoom/Universal.aspx';
  }

  // ---------- HTTP ----------
  function xfetch(method, body) {
    var ctl = window.AbortController ? new AbortController() : null;
    var t = ctl ? setTimeout(function () { ctl.abort(); }, FETCH_TIMEOUT) : null;
    var init = { method: method, credentials: 'same-origin', cache: 'no-store', redirect: 'follow' };
    if (body) { init.body = body; init.headers = { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' }; }
    if (ctl) init.signal = ctl.signal;
    return fetch(S.url, init).then(function (res) {
      if (t) clearTimeout(t);
      if (!res.ok) throw new Error('SGS ตอบ HTTP ' + res.status);
      return res;
    }, function (e) {
      if (t) clearTimeout(t);
      throw new Error(e && e.name === 'AbortError' ? 'SGS ไม่ตอบภายใน ' + (FETCH_TIMEOUT / 1000) + ' วิ' : 'เชื่อมต่อ SGS ไม่ได้ (' + (e && e.message || e) + ')');
    });
  }

  function parsePage(html) {
    var d = new DOMParser().parseFromString(html, 'text/html');
    if (!d.getElementById('aspnetForm') || !d.getElementById(ID_MODE)) {
      var low = html.toLowerCase();
      if (low.indexOf('login') >= 0 || low.indexOf('signin') >= 0 || low.indexOf('รหัสผ่าน') >= 0) throw new Error('SESSION: หมดเวลาล็อกอิน SGS — ล็อกอินใหม่แล้วคลิกปุ่มเครื่องมืออีกครั้ง');
      throw new Error('ไม่ใช่หน้าสารสนเทศ SGS ที่คาดไว้ (ไม่พบ DropDownListMode)');
    }
    return d;
  }

  function getPage() {
    return xfetch('GET').then(function (r) { return r.text(); }).then(function (h) { S.doc = parsePage(h); return S.doc; });
  }

  // จำลองการ submit ฟอร์มของเบราว์เซอร์จาก DOM ที่ parse มา
  function formBody(doc, over) {
    var f = doc.getElementById('aspnetForm'), p = new URLSearchParams(), seen = {};
    var els = f.querySelectorAll('input, select, textarea');
    for (var i = 0; i < els.length; i++) {
      var el = els[i], n = el.getAttribute('name'); if (!n || el.hasAttribute('disabled')) continue;
      var tag = el.tagName.toLowerCase(), ty = (el.getAttribute('type') || 'text').toLowerCase();
      if (tag === 'input') {
        if (/^(image|submit|button|reset|file)$/.test(ty)) continue;
        if ((ty === 'checkbox' || ty === 'radio') && !el.hasAttribute('checked')) continue;
        p.append(n, el.getAttribute('value') == null ? (ty === 'checkbox' || ty === 'radio' ? 'on' : '') : el.getAttribute('value'));
      } else if (tag === 'select') {
        var sel = el.querySelectorAll('option[selected]'), ops = el.querySelectorAll('option');
        if (el.hasAttribute('multiple')) { for (var k = 0; k < sel.length; k++) p.append(n, optVal(sel[k])); }
        else if (sel.length) p.append(n, optVal(sel[sel.length - 1]));
        else if (ops.length) p.append(n, optVal(ops[0]));
      } else { p.append(n, el.textContent || ''); }
      seen[n] = 1;
    }
    Object.keys(over || {}).forEach(function (k) { p.set(k, over[k]); });
    return p.toString();
  }
  function optVal(o) { return o.hasAttribute('value') ? o.getAttribute('value') : o.textContent; }
  function selVal(el) { var s = el.querySelectorAll('option[selected]'); if (s.length) return optVal(s[s.length - 1]); var o = el.querySelector('option'); return o ? optVal(o) : ''; }
  function pbTarget(el) {
    var a = (el.getAttribute('onchange') || '') + ' ' + (el.getAttribute('href') || '');
    var m = /__doPostBack\(\\?['"]([^'"\\]+)/.exec(a);
    return m ? m[1] : '';
  }

  function postback(target, fields) {
    var over = { __EVENTTARGET: target, __EVENTARGUMENT: '' };
    Object.keys(fields || {}).forEach(function (k) { over[k] = fields[k]; });
    return xfetch('POST', formBody(S.doc, over)).then(function (r) { return r.text(); })
      .then(function (h) { S.doc = parsePage(h); return S.doc; });
  }

  // ---------- ขั้นตอน SGS ----------
  function readHeader() {
    var y = S.doc.getElementById(ID_YEAR), t = S.doc.getElementById(ID_TERM);
    S.yr = y ? selVal(y) : ''; S.tr = t ? selVal(t) : '';
  }

  // select ใน UpdatePanel1 (ไม่ใช่เมนูข้อมูล / ไม่ใช่หัวหน้า)
  function panelSelects(doc) {
    var panel = doc.getElementById(ID_PANEL); if (!panel) return [];
    return Array.prototype.filter.call(panel.querySelectorAll('select'), function (s) { return s.id !== ID_MODE && !/_PageHeader/.test(s.id || ''); });
  }
  function levelOf(text) { var m = /^ม\.?(\d)$/.exec(norm(text)); return m ? +m[1] : 0; }
  // dropdown "ระดับชั้น" เหนือตาราง: option ข้อความ ม.1 … ม.6
  function findLevelSelect(doc) {
    var sels = panelSelects(doc);
    for (var i = 0; i < sels.length; i++) {
      var ops = sels[i].querySelectorAll('option'), n = 0;
      for (var k = 0; k < ops.length; k++) if (levelOf(ops[k].textContent)) n++;
      if (n >= 2) return sels[i];
    }
    return null;
  }
  // เลือก option ใน select: มี autopostback → POST · ไม่มี → ตั้ง selected ใน DOM ที่ parse ไว้
  async function choose(sel, opt) {
    var name = sel.getAttribute('name'), val = optVal(opt), tgt = pbTarget(sel);
    if (selVal(sel) === val && !tgt) return;
    if (tgt) { await postback(tgt, mapOf(name, val)); await stepWait(); }
    else {
      var ops = sel.querySelectorAll('option');
      for (var k = 0; k < ops.length; k++) ops[k].removeAttribute('selected');
      opt.setAttribute('selected', 'selected');
    }
  }

  async function ensureMode() {
    if (!S.doc) { await getPage(); readHeader(); }
    var m = S.doc.getElementById(ID_MODE), mv = '';
    var mo = m.querySelectorAll('option');
    var k = K();
    for (var i = 0; i < mo.length; i++) if (optVal(mo[i]).indexOf(k.prefix) === 0) { mv = optVal(mo[i]); break; }
    if (!mv) for (var i2 = 0; i2 < mo.length; i2++) if (k.text.test(norm(mo[i2].textContent))) { mv = optVal(mo[i2]); break; }
    if (!mv) throw new Error('ไม่พบรายการ "' + k.label + '" ในเมนูข้อมูล');
    if (selVal(m) !== mv || !findLevelSelect(S.doc)) { await postback(m.getAttribute('name'), mapOf(m.getAttribute('name'), mv)); await stepWait(); }
    if (!findLevelSelect(S.doc)) throw new Error('DIAG: ไม่พบ dropdown "ระดับชั้น" หลังเลือก "' + k.label + '"\n' + diagSelects(S.doc));
  }

  async function ensureContext(g) {
    await ensureMode();
    var ls = findLevelSelect(S.doc), ops = ls.querySelectorAll('option'), hit = null;
    for (var i = 0; i < ops.length; i++) if (levelOf(ops[i].textContent) === g) { hit = ops[i]; break; }
    if (!hit) throw new Error('NOROOM: ไม่พบระดับชั้น ม.' + g + ' ใน dropdown ระดับชั้น');
    if (selVal(ls) !== optVal(hit)) await choose(ls, hit);
    var ls2 = findLevelSelect(S.doc);
    if (!ls2 || selVal(ls2) !== optVal(hit)) throw new Error('SGS ไม่ยอมเปลี่ยนระดับชั้นเป็น ม.' + g);
    if (!findRoomSelect(S.doc)) throw new Error('DIAG: ไม่พบ dropdown "ห้องที่" หลังเลือก ม.' + g + '\n' + diagSelects(S.doc));
  }
  function mapOf(k, v) { var o = {}; o[k] = v; return o; }

  // แยกเลขห้องจากข้อความ option: "ม.1/3" "1/3" "ม. 1 / 3" → {g:1,r:3} · "3" "ห้อง 3" → {g:0,r:3}
  function roomOf(text) {
    var t = norm(text), m = /(\d{1,2})\/(\d{1,2})$/.exec(t);
    if (m) return { g: +m[1], r: +m[2] };
    m = /^(?:ห้อง)?(\d{1,2})$/.exec(t);
    if (m) return { g: 0, r: +m[1] };
    return null;
  }
  function findRoomSelect(doc) {
    var sels = panelSelects(doc), lv = findLevelSelect(doc), best = null, bestN = 0;
    for (var i = 0; i < sels.length; i++) {
      var s = sels[i]; if (s === lv) continue;
      var ops = s.querySelectorAll('option'), n = 0;
      for (var k = 0; k < ops.length; k++) if (roomOf(ops[k].textContent)) n++;
      if (n > bestN) { best = s; bestN = n; }
    }
    return bestN >= 1 ? best : null;
  }
  function roomsInSelect(sel, g) {
    var out = [], ops = sel.querySelectorAll('option');
    for (var i = 0; i < ops.length; i++) { var x = roomOf(ops[i].textContent); if (x && (x.g === 0 || x.g === g)) out.push(x.r); }
    return out;
  }
  function diagSelects(doc) {
    var panel = doc.getElementById(ID_PANEL) || doc, sels = panel.querySelectorAll('select'), lines = [];
    for (var i = 0; i < sels.length; i++) {
      var s = sels[i], ops = s.querySelectorAll('option'), tx = [];
      for (var k = 0; k < Math.min(ops.length, 8); k++) tx.push('"' + norm(ops[k].textContent) + '"=' + optVal(ops[k]));
      lines.push('  select ' + (s.getAttribute('name') || s.id) + ' (' + ops.length + ' ตัวเลือก) ' + (pbTarget(s) ? 'postback ' : '') + ': ' + tx.join(', '));
    }
    return lines.join('\n') || '  (ไม่พบ select ใน UpdatePanel1)';
  }

  async function selectRoom(g, r) {
    var sel = findRoomSelect(S.doc), ops = sel.querySelectorAll('option'), hit = null;
    for (var i = 0; i < ops.length; i++) { var x = roomOf(ops[i].textContent); if (x && x.r === r && (x.g === 0 || x.g === g)) { hit = ops[i]; break; } }
    if (!hit) throw new Error('NOROOM: ไม่พบห้อง ม.' + g + '/' + r + ' ในตัวเลือกของ SGS');
    var val = optVal(hit);
    await choose(sel, hit);   // เลือกห้องเดิมซ้ำก็ส่ง postback ให้ตารางดึงใหม่
    var s2 = findRoomSelect(S.doc);
    if (s2 && selVal(s2) !== val) throw new Error('SGS ไม่ยอมเปลี่ยนห้องเป็น ม.' + g + '/' + r);
  }

  function excelTarget() {
    var a = S.doc.querySelector('a[id$="ButtonExcel__Button"]');
    return (a && pbTarget(a)) || EXCEL_TARGET;
  }

  async function exportExcel() {
    var res = await xfetch('POST', formBody(S.doc, { __EVENTTARGET: excelTarget(), __EVENTARGUMENT: '' }));
    var buf = await res.arrayBuffer();
    var text = new TextDecoder('utf-8').decode(buf);
    if (/<form[^>]+aspnetForm/i.test(text)) {
      // ได้หน้าเว็บกลับมาแทนไฟล์ — อัปเดต state แล้วแจ้ง
      try { S.doc = parsePage(text); } catch (e) { throw e; }
      throw new Error('SGS ส่งหน้าเว็บกลับมาแทนไฟล์ Excel (อาจยังดึงข้อมูลไม่เสร็จ)');
    }
    return { buf: buf, text: text, cd: res.headers.get('Content-Disposition') || '', ct: res.headers.get('Content-Type') || '' };
  }

  function dispName(cd) {
    if (!cd) return '';
    var m = /filename\*\s*=\s*([^']*)'[^']*'([^;]+)/i.exec(cd);
    if (m) { try { return decodeURIComponent(m[2].trim()); } catch (e) {} }
    m = /filename\s*=\s*"?([^";]+)"?/i.exec(cd); if (!m) return '';
    var raw = m[1].trim();
    if (/%[0-9a-f]{2}/i.test(raw)) { try { return decodeURIComponent(raw.replace(/\+/g, ' ')); } catch (e) {} }
    try { // header เป็น UTF-8 ดิบที่ถูกอ่านเป็น latin1
      var b = new Uint8Array(raw.length); for (var i = 0; i < raw.length; i++) b[i] = raw.charCodeAt(i) & 255;
      if (/[\u0080-\u00ff]/.test(raw)) return new TextDecoder('utf-8', { fatal: true }).decode(b);
    } catch (e) {}
    return raw;
  }
  function fileNameFor(cd, g, r) {
    var fb = K().fb + '--' + S.yr + '_' + S.tr + '-' + g + '-' + r + '.xls';
    var n = dispName(cd).replace(/[\\\/:*?"<>|]/g, '_').trim();
    var ok = new RegExp('-' + g + '-' + r + '\\.xls$', 'i').test(n);
    return { name: ok ? n : fb, fromSgs: ok, raw: n };
  }

  // ตรวจไฟล์: ต้องมีตาราง MyGrid · ทุกแถว ชั้น/ห้อง = ม.g/r · ปี/ภาค = หัวหน้า SGS
  function checkFile(text, g, r) {
    if (text.indexOf('MyGrid') < 0) return { bad: 'ไฟล์ไม่มีตาราง MyGrid (ไม่ใช่ไฟล์ ' + K().label + ')' };
    var d = new DOMParser().parseFromString(text, 'text/html');
    var tb = d.getElementById('ctl00_PageContent_MyGrid') || d.querySelector('table');
    var trs = tb ? tb.querySelectorAll('tr') : [];
    if (!trs.length) return { bad: 'ไฟล์ว่าง (ไม่มีแถวหัวตาราง)' };
    var head = Array.prototype.map.call(trs[0].querySelectorAll('th,td'), function (c) { return norm(c.textContent); });
    function col(re) { for (var i = 0; i < head.length; i++) if (re.test(head[i])) return i; return -1; }
    var cRoom = col(/^ชั้น\/ห้อง$/), cYr = col(/^ปีการศึกษา$/), cTr = col(/^ภาคเรียนที่$/), cSub = col(/^รหัสวิชา$/), cSid = col(/^เลขประ/);
    if (cRoom < 0 || cSub < 0 || cSid < 0) return { bad: 'หัวตารางไม่ตรงรูปแบบ ปพ.5 (ไม่พบ ชั้น/ห้อง/รหัสวิชา/เลขประจำตัว)' };
    var kd = K();
    var lack = kd.need.filter(function (h) { return head.indexOf(h) < 0; }), extra = kd.deny.filter(function (h) { return head.indexOf(h) >= 0; });
    if (lack.length || extra.length) return { bad: 'ไม่ใช่ไฟล์ ' + kd.label + ' (หัวตาราง' + (lack.length ? ' ไม่มี ' + lack.join(', ') : '') + (extra.length ? ' มี ' + extra.join(', ') : '') + ')' };
    var want = 'ม.' + g + '/' + r, rooms = {}, yrs = {}, subs = {}, sids = {}, n = 0;
    for (var i = 1; i < trs.length; i++) {
      var c = trs[i].querySelectorAll('td'); if (!c.length) continue;
      var v = function (k) { return k >= 0 && c[k] ? norm(c[k].textContent) : ''; };
      if (!v(cSub) && !v(cSid)) continue;                    // แถวว่างท้ายตาราง (ไฟล์ 48 มี &nbsp; ทุกช่อง)
      n++; rooms[v(cRoom)] = (rooms[v(cRoom)] || 0) + 1; yrs[v(cYr) + '/' + v(cTr)] = 1; subs[v(cSub)] = 1; sids[v(cSid)] = 1;
    }
    var rk = Object.keys(rooms), yk = Object.keys(yrs);
    var res = { rows: n, subjects: Object.keys(subs).length, students: Object.keys(sids).length };
    if (!n) { res.warn = 'ไม่มีข้อมูลนักเรียน (0 แถว)'; return res; }
    var wrong = rk.filter(function (k) { return k !== want; });
    if (wrong.length) { res.bad = 'ข้อมูลเป็นห้อง ' + wrong.join(', ') + ' ไม่ใช่ ' + want + ' — ไม่บันทึก'; return res; }
    var wantYt = S.yr + '/' + S.tr;
    if (S.yr && (yk.length !== 1 || yk[0] !== wantYt)) { res.bad = 'ปี/ภาคในไฟล์ = ' + yk.join(', ') + ' ไม่ตรงกับ SGS ' + wantYt + ' — ไม่บันทึก'; return res; }
    return res;
  }

  // ---------- โฟลเดอร์ปลายทาง ----------
  function idb(mode, fn) {
    return new Promise(function (ok, no) {
      if (!window.indexedDB) return ok(null);
      var q = indexedDB.open(IDB_NAME, 1);
      q.onupgradeneeded = function () { q.result.createObjectStore(IDB_STORE); };
      q.onerror = function () { ok(null); };
      q.onsuccess = function () {
        try {
          var tx = q.result.transaction(IDB_STORE, mode), st = tx.objectStore(IDB_STORE), rq = fn(st);
          tx.oncomplete = function () { ok(rq && rq.result); }; tx.onerror = function () { ok(null); };
        } catch (e) { ok(null); }
      };
    });
  }
  function hasFsApi() { return typeof window.showDirectoryPicker === 'function'; }
  async function restoreDir() {
    if (!hasFsApi()) return;
    var h = await idb('readonly', function (st) { return st.get('dir'); });
    if (h && h.kind === 'directory') S.dir = h;
  }
  async function pickDir() {
    if (!hasFsApi()) { log('⚠ เบราว์เซอร์นี้เลือกโฟลเดอร์ไม่ได้ — ไฟล์จะลงโฟลเดอร์ดาวน์โหลดของเบราว์เซอร์ (ใช้ Chrome/Edge บนคอมพิวเตอร์เพื่อเลือก D:\\kjst-escore\\SGS)', 'warn'); return false; }
    try {
      var h = await window.showDirectoryPicker({ id: 'kjst-sgs-pp5', mode: 'readwrite' });
      S.dir = h; await idb('readwrite', function (st) { return st.put(h, 'dir'); });
      log('📁 เลือกโฟลเดอร์: ' + h.name + (h.name !== 'SGS' ? '  ⚠ ไม่ใช่โฟลเดอร์ชื่อ SGS — ตรวจว่าเป็น D:\\kjst-escore\\SGS' : ''), h.name === 'SGS' ? 'ok' : 'warn');
      renderDir(); return true;
    } catch (e) { if (e && e.name !== 'AbortError') log('เลือกโฟลเดอร์ไม่สำเร็จ: ' + e.message, 'err'); return false; }
  }
  async function ensureDirPermission() {
    if (!hasFsApi()) return true;
    if (!S.dir) return await pickDir();
    try {
      var o = { mode: 'readwrite' }, p = await S.dir.queryPermission(o);
      if (p !== 'granted') p = await S.dir.requestPermission(o);
      if (p === 'granted') return true;
    } catch (e) {}
    log('ไม่ได้รับสิทธิ์เขียนโฟลเดอร์ ' + S.dir.name + ' — กด "เลือกโฟลเดอร์" ใหม่', 'err');
    return false;
  }
  async function saveFile(name, buf) {
    var blob = new Blob([buf], { type: 'application/vnd.ms-excel' });
    if (S.dir) {
      var fh = await S.dir.getFileHandle(name, { create: true }), w = await fh.createWritable();
      await w.write(blob); await w.close(); return 'dir';
    }
    var a = document.createElement('a'), u = URL.createObjectURL(blob);
    a.href = u; a.download = name; a.style.display = 'none'; document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(u); a.remove(); }, 4000);
    return 'download';
  }

  // ---------- งานหลัก ----------
  function key(g, r) { return g + '-' + r; }
  function initRooms() {
    for (var g = 1; g <= 6; g++) for (var r = 1; r <= ROOMS[g]; r++) if (!S.rooms[key(g, r)]) S.rooms[key(g, r)] = { g: g, r: r, st: 'wait' };
  }

  async function doRoom(g, r) {
    var R = S.rooms[key(g, r)]; R.st = 'run'; R.msg = ''; renderRooms();
    var lastErr = null;
    for (var t = 1; t <= MAX_TRY; t++) {
      if (S.stop) { R.st = 'wait'; renderRooms(); return; }
      try {
        await ensureContext(g);
        await selectRoom(g, r);
        var x = await exportExcel();
        var chk = checkFile(x.text, g, r);
        if (chk.bad) throw new Error('BAD: ' + chk.bad);
        var fn = fileNameFor(x.cd, g, r);
        var how = await saveFile(fn.name, x.buf);
        R.file = fn.name; R.rows = chk.rows; R.students = chk.students; R.subjects = chk.subjects; R.bytes = x.buf.byteLength;
        R.st = chk.warn ? 'warn' : 'ok'; R.msg = chk.warn || '';
        log('ม.' + g + '/' + r + ' ' + (chk.warn ? '⚠' : '✓') + ' ' + chk.students + ' คน · ' + chk.subjects + ' วิชา · ' + chk.rows + ' แถว · ' + fn.name + ' (' + kb(R.bytes) + ')' +
          (how === 'download' ? ' → โฟลเดอร์ดาวน์โหลด' : '') + (fn.fromSgs ? '' : ' · ตั้งชื่อเอง (ชื่อจาก SGS: "' + fn.raw + '")') + (chk.warn ? ' · ' + chk.warn : ''), chk.warn ? 'warn' : 'ok');
        renderRooms(); return;
      } catch (e) {
        lastErr = e; var m = String(e && e.message || e);
        if (/^SESSION:/.test(m)) { S.stop = true; R.st = 'fail'; R.msg = m.slice(8).trim(); log(R.msg, 'err'); renderRooms(); return; }
        if (/^(NOROOM|BAD|DIAG):/.test(m)) break;              // ผิดเชิงข้อมูล ลองซ้ำไม่ช่วย
        log('ม.' + g + '/' + r + ' ครั้งที่ ' + t + ' ไม่สำเร็จ: ' + m + (t < MAX_TRY ? ' — ลองใหม่' : ''), 'warn');
        S.doc = null;                                          // เริ่มบริบทใหม่จาก GET
        await sleep(1500 * t);
      }
    }
    R.st = 'fail'; R.msg = String(lastErr && lastErr.message || lastErr).replace(/^(NOROOM|BAD|DIAG):\s*/, '');
    log('ม.' + g + '/' + r + ' ✗ ' + R.msg, 'err'); renderRooms();
  }

  async function run(onlyFailed) {
    if (S.running) return;
    S.running = true; S.stop = false; renderButtons();
    try {
      if (!(await ensureDirPermission()) && hasFsApi()) return;
      initRooms();
      var list = [];
      S.opts.grades.slice().sort().forEach(function (g) {
        for (var r = 1; r <= ROOMS[g]; r++) {
          var R = S.rooms[key(g, r)];
          if (onlyFailed ? R.st === 'fail' : true) list.push(R);
        }
      });
      if (!list.length) { log(onlyFailed ? 'ไม่มีห้องที่ล้มเหลว' : 'ยังไม่ได้เลือกชั้น', 'warn'); return; }
      if (!onlyFailed) list.forEach(function (R) { R.st = 'wait'; R.msg = ''; });
      renderRooms();
      log('▶ เริ่ม ' + list.length + ' ห้อง · ไฟล์ ' + K().label + (onlyFailed ? ' (เฉพาะที่ล้มเหลว)' : '') + ' · ความเร็ว ' + (S.opts.speed === 'slow' ? 'ช้า' : 'ปกติ'));
      var t0 = Date.now(), lastG = 0;
      for (var i = 0; i < list.length; i++) {
        if (S.stop) break;
        var R = list[i];
        if (R.g !== lastG) {           // เปลี่ยนชั้น → เทียบห้องที่ SGS มีจริงกับค่าตั้งต้น
          lastG = R.g;
          try { await ensureContext(R.g); noteExtraRooms(R.g); } catch (e) { /* doRoom จะจัดการ/รายงานเอง */ }
        }
        await doRoom(R.g, R.r);
      }
      summary(list, Date.now() - t0);
    } finally {
      S.running = false; renderButtons(); renderRooms();
    }
  }

  function noteExtraRooms(g) {
    var sel = S.doc && findRoomSelect(S.doc); if (!sel) return;
    var got = roomsInSelect(sel, g), extra = got.filter(function (r) { return r > ROOMS[g]; });
    var miss = []; for (var r = 1; r <= ROOMS[g]; r++) if (got.indexOf(r) < 0) miss.push(r);
    if (extra.length) log('ℹ SGS มีห้อง ม.' + g + '/' + extra.join(', ') + ' เพิ่มจากค่าตั้งต้น ' + ROOMS[g] + ' ห้อง — ไม่ได้ดาวน์โหลด', 'warn');
    if (miss.length) log('⚠ SGS ไม่มีห้อง ม.' + g + '/' + miss.join(', '), 'warn');
  }

  function summary(list, ms) {
    var c = { ok: 0, warn: 0, fail: 0, wait: 0 }, stu = 0;
    list.forEach(function (R) { c[R.st] = (c[R.st] || 0) + 1; stu += R.students || 0; });
    log('■ ' + (S.stop ? 'หยุดแล้ว' : 'เสร็จ') + ' · สำเร็จ ' + c.ok + (c.warn ? ' · มีข้อสังเกต ' + c.warn : '') + (c.fail ? ' · ล้มเหลว ' + c.fail : '') +
      (c.wait ? ' · ยังไม่ทำ ' + c.wait : '') + ' · นักเรียนรวม ' + stu + ' คน · ใช้เวลา ' + Math.round(ms / 1000) + ' วิ' +
      (c.fail ? ' — กด "ลองเฉพาะที่ล้มเหลว"' : ''), c.fail ? 'err' : 'ok');
    var per = {};
    list.forEach(function (R) { if (R.students) { per[R.g] = per[R.g] || []; per[R.g].push(R.r + ':' + R.students); } });
    Object.keys(per).forEach(function (g) { log('   ม.' + g + ' จำนวนคน/ห้อง  ' + per[g].join('  ')); });
  }

  // ---------- UI ----------
  var LOGS = [];
  function log(msg, cls) {
    var t = new Date(), ts = ('0' + t.getHours()).slice(-2) + ':' + ('0' + t.getMinutes()).slice(-2) + ':' + ('0' + t.getSeconds()).slice(-2);
    LOGS.push('[' + ts + '] ' + msg);
    var box = $('kpp5Log'); if (!box) return;
    var d = document.createElement('div'); d.className = 'kl ' + (cls || ''); d.textContent = '[' + ts + '] ' + msg;
    box.appendChild(d); box.scrollTop = box.scrollHeight;
  }

  var CSS = '#kpp5{position:fixed;top:12px;right:12px;width:430px;max-width:calc(100vw - 24px);z-index:2147483000;background:#fff;color:#1f2937;' +
    'border:1px solid #cbd5e1;border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.25);font:13px/1.45 Tahoma,"Sarabun",sans-serif}' +
    '#kpp5 *{box-sizing:border-box;font-family:inherit}' +
    '#kpp5 .hd{display:flex;align-items:center;gap:8px;padding:9px 12px;background:#0f3d6e;color:#fff;border-radius:12px 12px 0 0;cursor:move;user-select:none}' +
    '#kpp5 .hd b{flex:1;font-size:14px}#kpp5 .hd button{background:transparent;border:0;color:#fff;font-size:16px;cursor:pointer;padding:0 4px}' +
    '#kpp5 .bd{padding:10px 12px}#kpp5.min .bd{display:none}' +
    '#kpp5 .row{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin:0 0 8px}' +
    '#kpp5 .lb{color:#64748b;min-width:62px}' +
    '#kpp5 .chip{border:1px solid #94a3b8;background:#f8fafc;border-radius:999px;padding:2px 10px;cursor:pointer}' +
    '#kpp5 .chip.on{background:#0f3d6e;border-color:#0f3d6e;color:#fff}' +
    '#kpp5 button.b{border:1px solid #0f3d6e;background:#0f3d6e;color:#fff;border-radius:8px;padding:5px 12px;cursor:pointer}' +
    '#kpp5 button.b.g{background:#fff;color:#0f3d6e}#kpp5 button.b.r{background:#b91c1c;border-color:#b91c1c}' +
    '#kpp5 button.b:disabled{opacity:.45;cursor:default}' +
    '#kpp5 select{font-size:13px;padding:2px 4px}' +
    '#kpp5 .dir{flex:1;padding:3px 8px;border-radius:6px;background:#f1f5f9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}' +
    '#kpp5 .dir.ok{background:#dcfce7;color:#166534}#kpp5 .dir.no{background:#fef3c7;color:#92400e}' +
    '#kpp5 .grid{display:grid;grid-template-columns:34px repeat(7,1fr);gap:3px;margin:4px 0 8px}' +
    '#kpp5 .gl{color:#64748b;align-self:center;font-size:12px}' +
    '#kpp5 .rc{text-align:center;border-radius:5px;padding:3px 0;font-size:12px;background:#e2e8f0;color:#475569;cursor:default}' +
    '#kpp5 .rc.run{background:#fde68a;color:#78350f;animation:kpp5p 1s infinite}#kpp5 .rc.ok{background:#16a34a;color:#fff}' +
    '#kpp5 .rc.warn{background:#f59e0b;color:#fff}#kpp5 .rc.fail{background:#dc2626;color:#fff}#kpp5 .rc.off{opacity:.3}' +
    '@keyframes kpp5p{50%{opacity:.55}}' +
    '#kpp5 .prog{font-size:12px;color:#334155;margin-bottom:6px}' +
    '#kpp5Log{height:150px;overflow:auto;background:#0b1220;color:#cbd5e1;border-radius:8px;padding:6px 8px;font:11.5px/1.5 Consolas,monospace;white-space:pre-wrap}' +
    '#kpp5Log .ok{color:#86efac}#kpp5Log .warn{color:#fcd34d}#kpp5Log .err{color:#fca5a5}' +
    '#kpp5 .ft{display:flex;justify-content:space-between;margin-top:6px;font-size:11.5px;color:#94a3b8}#kpp5 .ft a{color:#0f3d6e;cursor:pointer}';

  function build() {
    var st = document.createElement('style'); st.id = 'kpp5Css'; st.textContent = CSS; document.head.appendChild(st);
    var w = document.createElement('div'); w.id = 'kpp5';
    var chips = ''; for (var g = 1; g <= 6; g++) chips += '<span class="chip" data-g="' + g + '">ม.' + g + '</span>';
    w.innerHTML =
      '<div class="hd" id="kpp5Hd"><b>KJST · ดาวน์โหลด ปพ.5 จาก SGS</b><span style="opacity:.7;font-size:11px">v' + VER + '</span>' +
      '<button id="kpp5Min" title="ย่อ">–</button><button id="kpp5X" title="ปิด">×</button></div>' +
      '<div class="bd">' +
      '<div class="row"><span class="lb">SGS</span><span id="kpp5Ctx">กำลังอ่านหน้าสารสนเทศ…</span></div>' +
      '<div class="row"><span class="lb">บันทึกที่</span><span class="dir" id="kpp5Dir"></span><button class="b g" id="kpp5Pick">เลือกโฟลเดอร์</button></div>' +
      '<div class="row"><span class="lb">ไฟล์</span><select id="kpp5Kind"><option value="7">7. ผลการเรียน ปพ.5 (ตรวจวิธีที่ 1)</option><option value="48">48. คะแนนรายวิชา (ตรวจวิธีที่ 2)</option></select></div>' +
      '<div class="row"><span class="lb">ชั้น</span>' + chips + '</div>' +
      '<div class="row"><span class="lb">ความเร็ว</span><select id="kpp5Spd"><option value="normal">ปกติ</option><option value="slow">ช้า (server มีคนใช้เยอะ)</option></select></div>' +
      '<div class="grid" id="kpp5Grid"></div>' +
      '<div class="prog" id="kpp5Prog"></div>' +
      '<div class="row"><button class="b" id="kpp5Go">▶ เริ่มดาวน์โหลด</button><button class="b r" id="kpp5Stop">■ หยุด</button>' +
      '<button class="b g" id="kpp5Retry">ลองเฉพาะที่ล้มเหลว</button></div>' +
      '<div id="kpp5Log"></div>' +
      '<div class="ft"><span>ไฟล์ชื่อซ้ำ = เขียนทับ · ไม่แตะข้อมูลใน SGS</span><a id="kpp5Copy">คัดลอก log</a></div>' +
      '</div>';
    document.body.appendChild(w);

    w.querySelectorAll('.chip').forEach(function (c) {
      c.addEventListener('click', function () {
        if (S.running) return;
        var g = +c.getAttribute('data-g'), i = S.opts.grades.indexOf(g);
        if (i >= 0) S.opts.grades.splice(i, 1); else S.opts.grades.push(g);
        saveOpts(); renderChips(); renderRooms();
      });
    });
    $('kpp5Kind').value = S.opts.kind;
    $('kpp5Kind').addEventListener('change', function () {
      S.opts.kind = this.value; saveOpts();
      Object.keys(S.rooms).forEach(function (k) { var R = S.rooms[k]; R.st = 'wait'; R.msg = ''; R.file = ''; R.students = 0; });
      renderRooms(); log('เปลี่ยนชนิดไฟล์เป็น "' + K().label + '"');
    });
    $('kpp5Spd').value = S.opts.speed;
    $('kpp5Spd').addEventListener('change', function () { S.opts.speed = this.value; saveOpts(); });
    $('kpp5Pick').addEventListener('click', function () { if (!S.running) pickDir(); });
    $('kpp5Go').addEventListener('click', function () { run(false); });
    $('kpp5Retry').addEventListener('click', function () { run(true); });
    $('kpp5Stop').addEventListener('click', function () { if (S.running) { S.stop = true; log('■ กำลังหยุดหลังห้องปัจจุบัน…', 'warn'); } });
    $('kpp5Min').addEventListener('click', function () { w.classList.toggle('min'); });
    $('kpp5X').addEventListener('click', function () {
      if (S.running && !confirm('กำลังดาวน์โหลดอยู่ — ปิดและหยุด?')) return;
      S.stop = true; w.style.display = 'none';
    });
    $('kpp5Copy').addEventListener('click', function () {
      var t = LOGS.join('\n');
      (navigator.clipboard && navigator.clipboard.writeText ? navigator.clipboard.writeText(t) : Promise.reject()).then(
        function () { log('คัดลอก log แล้ว', 'ok'); }, function () { prompt('คัดลอก log:', t); });
    });
    drag(w, $('kpp5Hd'));
  }
  function drag(w, h) {
    var sx, sy, ox, oy, on = false;
    h.addEventListener('mousedown', function (e) {
      if (e.target.tagName === 'BUTTON') return; on = true; var r = w.getBoundingClientRect();
      sx = e.clientX; sy = e.clientY; ox = r.left; oy = r.top; e.preventDefault();
    });
    document.addEventListener('mousemove', function (e) {
      if (!on) return; w.style.left = Math.max(0, ox + e.clientX - sx) + 'px'; w.style.top = Math.max(0, oy + e.clientY - sy) + 'px'; w.style.right = 'auto';
    });
    document.addEventListener('mouseup', function () { on = false; });
  }
  function renderChips() {
    document.querySelectorAll('#kpp5 .chip').forEach(function (c) { c.classList.toggle('on', S.opts.grades.indexOf(+c.getAttribute('data-g')) >= 0); });
  }
  function renderDir() {
    var d = $('kpp5Dir'); if (!d) return;
    if (!hasFsApi()) { d.className = 'dir no'; d.textContent = 'โฟลเดอร์ดาวน์โหลดของเบราว์เซอร์ (เลือกโฟลเดอร์ไม่ได้)'; return; }
    if (!S.dir) { d.className = 'dir no'; d.textContent = 'ยังไม่ได้เลือก → เลือก D:\\kjst-escore\\SGS'; return; }
    d.className = 'dir ' + (S.dir.name === 'SGS' ? 'ok' : 'no'); d.textContent = '📁 ' + S.dir.name + (S.dir.name === 'SGS' ? '' : '  (ไม่ใช่ SGS?)');
  }
  function renderRooms() {
    initRooms();
    var h = '', n = 0, ok = 0, warn = 0, fail = 0;
    for (var g = 1; g <= 6; g++) {
      var on = S.opts.grades.indexOf(g) >= 0;
      h += '<div class="gl">ม.' + g + '</div>';
      for (var r = 1; r <= 7; r++) {
        var R = S.rooms[key(g, r)];
        if (!R) { h += '<div></div>'; continue; }
        if (on) { n++; if (R.st === 'ok') ok++; if (R.st === 'warn') warn++; if (R.st === 'fail') fail++; }
        var tip = 'ม.' + g + '/' + r + (R.file ? '\n' + R.file + '\n' + R.students + ' คน · ' + R.subjects + ' วิชา' : '') + (R.msg ? '\n' + R.msg : '');
        h += '<div class="rc ' + R.st + (on ? '' : ' off') + '" title="' + esc(tip) + '">' + g + '/' + r + '</div>';
      }
    }
    var gEl = $('kpp5Grid'); if (gEl) gEl.innerHTML = h;
    var p = $('kpp5Prog');
    if (p) p.textContent = 'เลือก ' + n + ' ห้อง · สำเร็จ ' + (ok + warn) + '/' + n + (warn ? ' (ข้อสังเกต ' + warn + ')' : '') + (fail ? ' · ล้มเหลว ' + fail : '');
  }
  function renderButtons() {
    var r = S.running;
    ['kpp5Go', 'kpp5Retry', 'kpp5Pick', 'kpp5Spd', 'kpp5Kind'].forEach(function (id) { var e = $(id); if (e) e.disabled = r; });
    var s = $('kpp5Stop'); if (s) s.disabled = !r;
  }

  async function init() {
    build(); renderChips(); renderButtons(); initRooms(); renderRooms();
    S.url = pageUrl();
    await restoreDir(); renderDir();
    log('KJST ดาวน์โหลด ปพ.5 v' + VER + ' · หน้า ' + S.url.replace(location.origin, ''));
    try {
      await getPage(); readHeader();
      $('kpp5Ctx').textContent = 'ปีการศึกษา ' + S.yr + ' ภาคเรียน ' + S.tr + ' · ล็อกอินอยู่ ✓';
      log('อ่านหน้าสารสนเทศได้ · ปี/ภาค ' + S.yr + '/' + S.tr + ' — ตรวจให้ตรงภาคที่ต้องการก่อนเริ่ม (เปลี่ยนที่หัวหน้า SGS)', 'ok');
    } catch (e) {
      $('kpp5Ctx').textContent = '⚠ ' + e.message.replace(/^SESSION:\s*/, '');
      log(e.message.replace(/^SESSION:\s*/, ''), 'err');
    }
  }

  window.__kjstPp5 = {
    show: function () { var w = $('kpp5'); if (w) { w.style.display = ''; w.classList.remove('min'); } },
    _S: S, _t: { roomOf: roomOf, dispName: dispName, checkFile: checkFile, formBody: formBody, fileNameFor: fileNameFor }
  };
  init();
})();
