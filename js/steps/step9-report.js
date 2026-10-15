/*
 * 第 9 步：總表
 *   在活頁簿裡新增「總表」工作表。
 *
 *   版面（B 到 J 行、第 2 到 28 列）：
 *     第 2 列      大標題（B 到 J 合併，預設「製二課 (車用/光機)」，畫面上可以改）
 *     第 4 列      法定加班上限46hr/人（紅底白字，B、C 合併）
 *     第 5 列      表頭：B、C 合併畫斜線；D 到 I 是六個月；J 是備註（深藍底白字）
 *     第 6–8 列    MPS(PCS)：量產、試產、RMA（淺藍底）
 *     第 9–11 列   需求工時-量產、需求工時-試產、需求工時-重工（min/月）
 *     第 12–15 列  平日工作天數、假日工作天數、工作時間、加班時間
 *     第 16–19 列  綜合效率% - 量產、- 試產、- 重工、出勤率%
 *     第 20–23 列  需求人力：On line、新機種+重工、Off line(20%)、TTL（淺綠底）
 *     第 24–26 列  現有人力：On line、Off line、TTL（淺綠底）
 *     第 27 列     人力需求差異(現有-需求)
 *     第 28 列     Loading Rate%(含加班)（淺橘底）
 *
 *   月份表頭用公式參照 MPS 工作表「需求工時」的月份（第 5 步輸入的），例如 ='MPS'!N3。
 *
 *   從 MPS 工作表抓的列（都是公式；用標籤文字找出是哪一列）：
 *     MPS(PCS) 量產、試產    ← 第 2 步匯總的「量產」「試產」那一列（當月已含 WIP）
 *     MPS(PCS) RMA          ← 第 7 步 RMA 各項目數量的合計，例如 =120+300
 *     需求工時-量產、-試產   ← 第 6 步「量產總需求工時」「試產總需求工時」
 *     需求工時-重工         ← 第 7 步「RMA總需求工時」
 *
 *   使用者在畫面上輸入的列：
 *     平日工作天數     各月自己輸入（每個月不同，不記住）
 *     工作時間         各月自己輸入，預設 480 分
 *     加班時間         輸入每日加班幾小時、加班出勤人數 %，填公式 =2*60*80%，備註寫輸入的數值
 *     綜合效率% 量產、試產、重工，出勤率%   各月自己輸入，第一次預設 95%、85%、80%、95%，之後用上次輸入的值
 *     需求人力 新機種+重工   輸入各新機種（名稱可增減）和重工各月幾人，加總填公式 =2+3+1，備註寫是哪些機種
 *     需求人力 Off line      輸入領班、物料（含材製損）、文書、包裝幾人（每個月一樣），填公式 =1+2+1+3，備註寫各種人數
 *   工作時間、效率、出勤率、加班的兩個數字、大標題會記在這台電腦的瀏覽器裡。
 *   算出來的列：
 *     假日工作天數     =(46-D12*2)/8
 *     需求人力 On line      =(D9/D16+D10/D17+D11/D18)/(D12+D13)/(D14+D15)/D19，顯示到小數點第一位
 *     需求人力 TTL          =D20+D21+D22
 *   從「人力」工作表抓的列（第 8 步確認的列和當月那一行，例如 ='人力'!D4）：
 *     現有人力 On line ← Online；Off line ← Off line（直接）
 *     現有人力 TTL     =D24+D25
 *   最後兩列：
 *     人力需求差異(現有-需求)   =D26-D23
 *     Loading Rate%(含加班)     =D23/D26（需求 TTL ÷ 現有 TTL，顯示成 %）
 *
 *   第二張表（第 28 列下面空三列，第 32 到 56 列）：和第一張一樣，紅色格子寫「YO加班上限78hr/人」，
 *   假日工作天數改用 78。78 可以在畫面上改（會記住）。其他輸入和抓來的數值都參照第一張表。
 *
 *   來源有兩種：
 *     a. 剛才第 8 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 8 步匯出的檔案
 *
 * 做完之後把結果放在 IEApp.state.withReport。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var SHEET_NAME = XIO.REPORT_SHEET;
  var STEP1_KEY = IEApp.storeKey('step1');
  var TITLE = '需求工時';
  var SUFFIX = '_總表';
  var EXTRA = ['P003', '人力', SHEET_NAME];
  var STORE_KEY = IEApp.storeKey('report');
  var MONTHS = XIO.REPORT.months;
  var DEFAULT_WORK_MINUTES = '480';
  var EFF_VER = 2;            // 效率預設值的版本：預設值改過時加 1，舊的記憶就不沿用

  var S = {
    upload: null,         // 上傳的第 8 步檔案 { name, buffer, sheets }
    preferUpload: false,
    work: null,           // 拿來預覽的活頁簿 { src, loading, error, wb, main }
    inputs: null          // 輸入框裡的文字 { title, weekdays: [6], workMinutes: [6], effMass: [6], effTrial: [6], effRework: [6], attendance: [6], otHours, otRate, yoLimit }
  };
  var els = {};
  var timer = null;

  function blank6(v) { var a = []; for (var i = 0; i < MONTHS; i++) a.push(v); return a; }
  function strings6(list) {
    list = Array.isArray(list) ? list : [];
    return blank6('').map(function (v, i) { return list[i] === undefined || list[i] === null ? '' : String(list[i]); });
  }

  // 各月自己輸入、會記住上次的值的列（percent: 輸入的是 %，例如 80 代表 80%）
  var REMEMBERED = [
    { key: 'workMinutes', label: '工作時間（分/天）', what: '工作時間', def: DEFAULT_WORK_MINUTES },
    { key: 'effMass', label: '綜合效率% - 量產', what: '綜合效率% - 量產', def: '95', percent: true },
    { key: 'effTrial', label: '綜合效率% - 試產', what: '綜合效率% - 試產', def: '85', percent: true },
    { key: 'effRework', label: '綜合效率% - 重工', what: '綜合效率% - 重工', def: '80', percent: true },
    { key: 'attendance', label: '出勤率%', what: '出勤率%', def: '95', percent: true }
  ];

  function loadInputs() {
    var s = {};
    try { s = JSON.parse(localStorage.getItem(STORE_KEY) || '{}') || {}; } catch (e) { s = {}; }
    var out = {
      title: s.title === undefined || s.title === null ? XIO.REPORT_TITLE_DEFAULT : String(s.title),
      weekdays: blank6(''),
      otHours: s.otHours === undefined ? '' : String(s.otHours),
      otRate: s.otRate === undefined ? '' : String(s.otRate),
      yoLimit: s.yoLimit === undefined || s.yoLimit === '' ? String(XIO.YO_LIMIT_DEFAULT) : String(s.yoLimit)
    };
    // 需求人力 新機種+重工：新機種（名稱 + 各月人數，可增減）、重工各月人數
    out.newModels = (Array.isArray(s.newModels) ? s.newModels : []).map(function (md) {
      return { name: md && md.name ? String(md.name) : '', qty: strings6(md && md.qty) };
    });
    if (!out.newModels.length) out.newModels.push({ name: '', qty: blank6('') });
    out.reworkPeople = strings6(s.reworkPeople);
    // 需求人力 Off line：各種人員的人數（每個月一樣）
    var off = Array.isArray(s.offline) ? s.offline : [];
    out.offline = XIO.REPORT_OFFLINE.map(function (c, k) { return off[k] === undefined || off[k] === null ? '' : String(off[k]); });
    REMEMBERED.forEach(function (row) {
      // 效率的預設值改過（v5：量產 95、試產 85、重工 80），之前版本記住的效率不沿用
      var stale = /^eff/.test(row.key) && s.effVer !== EFF_VER;
      var saved = !stale && Array.isArray(s[row.key]) ? s[row.key].slice(0, MONTHS) : [];
      while (saved.length < MONTHS) saved.push(row.def);
      out[row.key] = saved.map(function (v) { return v === undefined || v === null ? row.def : String(v); });
    });
    return out;
  }
  function saveInputs() {
    var i = S.inputs;
    var keep = { title: i.title, otHours: i.otHours, otRate: i.otRate, yoLimit: i.yoLimit,
      newModels: i.newModels, reworkPeople: i.reworkPeople, offline: i.offline, effVer: EFF_VER };
    REMEMBERED.forEach(function (row) { keep[row.key] = i[row.key]; });
    try { localStorage.setItem(STORE_KEY, JSON.stringify(keep)); } catch (e) { /* 存不了就算了 */ }
  }

  // 輸入框的文字 → 數字。回傳 { values, error, blanks: [說明…] }
  function readInputs() {
    var i = S.inputs;
    var out = { weekdays: [], otHours: null, otRate: null };
    var blanks = [];
    function num(text, what, percent) {
      var t = String(text);
      if (percent) t = t.replace(/[%％]\s*$/, '');      // 允許打成 80%
      var n = XIO.toNumber(t);
      if (n !== null && isNaN(n)) throw new Error(what + '不是數字。');
      return n;
    }
    try {
      for (var m = 0; m < MONTHS; m++) {
        out.weekdays.push(num(i.weekdays[m], monthName(m) + '的平日工作天數'));
      }
      REMEMBERED.forEach(function (row) {
        out[row.key] = i[row.key].map(function (v, m) { return num(v, monthName(m) + '的' + row.what, row.percent); });
      });
      out.otHours = num(i.otHours, '每日加班小時');
      out.otRate = num(i.otRate, '加班出勤人數 %', true);
      out.yoLimit = num(i.yoLimit, 'YO 加班上限');
      out.newModels = i.newModels.map(function (md, k) {
        var name = String(md.name || '').trim();
        return { name: name, qty: md.qty.map(function (v, m) {
          return num(v, monthName(m) + '的新機種「' + (name || '第 ' + (k + 1) + ' 個') + '」人數');
        }) };
      });
      out.reworkPeople = i.reworkPeople.map(function (v, m) { return num(v, monthName(m) + '的重工人數'); });
      out.offline = XIO.REPORT_OFFLINE.map(function (c, k) { return num(i.offline[k], '需求人力 Off line 的' + c + '人數'); });
      if (out.yoLimit === null) throw new Error('請輸入 YO 加班上限（小時）。');
    } catch (e) {
      return { error: e.message };
    }
    var missWd = out.weekdays.filter(function (v) { return v === null; }).length;
    if (missWd) blanks.push('有 ' + missWd + ' 個月還沒輸入平日工作天數（假日工作天數會照 0 天算）。');
    REMEMBERED.forEach(function (row) {
      var miss = out[row.key].filter(function (v) { return v === null; }).length;
      if (miss) blanks.push('有 ' + miss + ' 個月沒有' + row.what + '。');
    });
    if (out.otHours === null || out.otRate === null) blanks.push('每日加班小時和加班出勤人數 % 都輸入後，才會填加班時間。');
    var anyNew = out.newModels.some(function (md) { return md.qty.some(function (v) { return v !== null; }); }) ||
      out.reworkPeople.some(function (v) { return v !== null; });
    if (!anyNew) blanks.push('需求人力「新機種+重工」還沒輸入人數，這一列先留空白。');
    out.newModels.forEach(function (md, k) {
      if (!md.name && md.qty.some(function (v) { return v !== null; })) blanks.push('第 ' + (k + 1) + ' 個新機種有人數但沒有名稱，備註會寫成「新機種' + (k + 1) + '」。');
    });
    if (!out.offline.some(function (v) { return v !== null; })) blanks.push('需求人力「Off line」還沒輸入人數，這一列先留空白。');
    if (missWd) blanks.push('平日工作天數沒填的月份，需求人力會是 #DIV/0!。');
    out.title = String(i.title || '').trim() || XIO.REPORT_TITLE_DEFAULT;
    return { values: out, blanks: blanks };
  }

  function monthName(m) {
    var r = S.work && S.work.refs && S.work.refs[m];
    return r && r.text ? r.text : '第 ' + (m + 1) + ' 個月';
  }

  function sameName(a, b) { return String(a).toLowerCase() === String(b).toLowerCase(); }
  function isExtra(name) { return EXTRA.some(function (n) { return sameName(n, name); }); }

  function activeSource() {
    var prev = IEApp.state.withManpower;
    if (S.upload && (S.preferUpload || !prev)) return { kind: 'file', key: S.upload, name: S.upload.name };
    if (prev) return { kind: 'prev', key: prev, name: prev.outName, prev: prev };
    return null;
  }

  function buildFrom(src) {
    return src.kind === 'prev' ? Promise.resolve(src.prev.build()) : XIO.load(src.key.buffer.slice(0));
  }

  function pickMain(wb, src) {
    var others = wb.worksheets.filter(function (w) { return !isExtra(w.name); });
    var wanted = src.kind === 'prev' ? src.prev.mainSheet : null;
    return others.filter(function (w) { return w.name === wanted; })[0] || others[0] || null;
  }

  // MPS 的資料從第幾列開始：沿用前面步驟的；上傳的檔案用第 1 步記住的設定（預設第 5 列）
  function dataStartRow(src) {
    if (src.kind === 'prev' && src.prev.dataStartRow) return src.prev.dataStartRow;
    var n = 0;
    try {
      var s1 = JSON.parse(localStorage.getItem(STEP1_KEY) || '{}') || {};
      n = parseInt(XIO.normalizeText(s1.startRow === undefined ? '' : s1.startRow), 10);
    } catch (e) { n = 0; }
    return n >= 1 ? n : 5;
  }

  // 月份表頭要參照的儲存格（MPS 工作表「需求工時」那一塊的月份）
  function monthRefs(main, src) {
    var start = dataStartRow(src);
    var keyCol = src.kind === 'prev' && src.prev.keyCol ? src.prev.keyCol : 1;
    var plan = XIO.planSummary(main, { dataStartRow: start, keyCol: keyCol, title: TITLE });
    if (!plan.found) return null;
    var headRow = start - 1 >= 3 ? 3 : start - 1;
    if (headRow < 1) return null;
    var refs = [];
    for (var i = 0; i < Math.min(plan.width, XIO.REPORT.months); i++) {
      var c = plan.startCol + i;
      var cell = main.findCell(headRow, c);
      refs.push({ sheet: main.name, col: c, row: headRow, text: cell ? XIO.displayText(cell) : '' });
    }
    return refs;
  }

  // 現有人力：從「人力」工作表抓第 8 步確認過的那幾列
  var MP_SHEET = '人力';
  var MP_STORE_KEY = IEApp.storeKey('manpower');
  var MP_LINKS = [
    { key: 'haveOnline', item: 'online', label: 'Online' },
    { key: 'haveOffline', item: 'offDirect', label: 'Off line（直接）' }
  ];

  // 第 8 步確認的位置：剛做完第 8 步就用它的；上傳的檔案用第 8 步記在瀏覽器裡的位置
  function manpowerPositions(src) {
    if (src.kind === 'prev' && src.prev.rows && src.prev.firstMonthCol) {
      return { rows: src.prev.rows, firstCol: src.prev.firstMonthCol, from: 'chain' };
    }
    var s = {};
    try { s = JSON.parse(localStorage.getItem(MP_STORE_KEY) || '{}') || {}; } catch (e) { s = {}; }
    var firstCol = parseInt(s.lastFirstCol, 10) || 0;
    if (!s.rows || !firstCol) return null;
    return { rows: s.rows, firstCol: firstCol, from: 'saved' };
  }

  // 回傳 { haveOnline: { found, label, sheet, row, refs, reason }, … }
  function manpowerLinks(wb, src) {
    var ws = wb.worksheets.filter(function (w) { return sameName(w.name, MP_SHEET); })[0];
    var pos = ws ? manpowerPositions(src) : null;
    var out = {};
    MP_LINKS.forEach(function (link) {
      var info = { found: false, label: link.label, sheet: MP_SHEET, refs: [], from: pos && pos.from };
      var r = pos ? parseInt(pos.rows[link.item], 10) : 0;
      if (!ws) info.reason = '檔案裡沒有「' + MP_SHEET + '」工作表（第 8 步）';
      else if (!pos || !r) info.reason = '不知道「' + link.label + '」在人力工作表的哪一列，請重做第 8 步';
      else {
        info.found = true;
        info.row = r;
        var row = ws.findRow(r);
        for (var i = 0; i < MONTHS; i++) {
          var c = pos.firstCol + i;
          var n = XIO.toNumber(row ? row.findCell(c) && XIO.plainValue(row.findCell(c).value) : null);
          info.refs.push({ sheet: ws.name, col: c, row: r, value: n === null || isNaN(n) ? null : n, numFmt: 'General' });
        }
      }
      out[link.key] = info;
    });
    return out;
  }

  // findReportLinks 的結果 → addReportSheet 要的 { key: refs }
  function linkRefs(links) {
    var out = {};
    Object.keys(links).forEach(function (k) { if (links[k].found) out[k] = links[k].refs; });
    return out;
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};

    els.srcInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.drop = ui.fileDrop({
      id: 'f9-base',
      title: '上傳第 8 步匯出的 Excel',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openUpload
    });
    els.switchBtn = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 8 步匯入人力報表',
      onclick: function () { IEApp.go('manpower'); }
    });
    els.srcMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    if (!S.inputs) S.inputs = loadInputs();
    els.entry = h('div', { class: 'entry-area' });
    els.monthHeads = [];
    buildEntry();

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.notes = h('ul', { class: 'notes' });
    els.result = h('div', { class: 'sheet-area' });
    els.preview = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '總表' }),
      h('p', { text: '在同一個檔案裡新增「總表」工作表，用公式從 MPS、人力工作表抓量產／試產／RMA 的產量、需求工時和現有人力，加上畫面上輸入的天數、工時、效率和出勤率，算出需求人力、人力差異和 Loading Rate。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 8 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg),
      h('li', { class: 'task' },
        h('h2', { text: '輸入各月的數值' }),
        h('p', { class: 'hint', text: '平日工作天數每個月都要輸入；工作時間預設 480 分，綜合效率預設 95%（量產）、85%（試產）、80%（重工），出勤率預設 95%，之後都用上次輸入的值。假日工作天數會用公式自動算：（法定加班上限 46 小時 − 平日工作天數 × 每天加班 2 小時）÷ 8。加班時間 = 每日加班小時 × 60 × 加班出勤人數 %。需求人力的「新機種+重工」和「Off line」輸入人數，加總填入，明細寫在備註。' }),
        els.entry),
      h('li', { class: 'task' },
        h('h2', { text: '總表' }),
        els.summary, els.notes, els.result, els.preview,
        h('div', { class: 'actions' }, els.exportBtn, els.exportNote))));

    refresh();
  }

  /* ---------- 輸入表 ---------- */
  function input(value, label, onInput, cls) {
    var el = h('input', {
      type: 'text', class: 'in ' + (cls || 'entry-num'), inputmode: 'decimal', autocomplete: 'off',
      value: value, 'aria-label': label
    });
    el.addEventListener('input', function () { onInput(el.value); onType(); });
    return el;
  }

  function buildEntry() {
    var inp = S.inputs;
    els.entry.appendChild(h('div', { class: 'field-row', style: 'margin-bottom:0.75rem' },
      h('label', { for: 'f9-title', text: '總表大標題' }),
      (function () {
        var el = h('input', { id: 'f9-title', type: 'text', class: 'in in-wide', autocomplete: 'off', value: inp.title });
        el.addEventListener('input', function () { inp.title = el.value; onType(); });
        return el;
      })()));
    var head = h('tr', null, h('th', { scope: 'col', text: '項目' }));
    for (var m = 0; m < MONTHS; m++) {
      var th = h('th', { scope: 'col', text: monthName(m) });
      els.monthHeads.push(th);
      head.appendChild(th);
    }
    var body = h('tbody');
    [{ key: 'weekdays', label: '平日工作天數（天/月）' }].concat(REMEMBERED).forEach(function (row) {
      var tr = h('tr', null, h('th', { scope: 'row', text: row.label }));
      inp[row.key].forEach(function (v, i) {
        tr.appendChild(h('td', null, input(v, row.label + '，' + monthName(i), function (val) { inp[row.key][i] = val; })));
      });
      body.appendChild(tr);
    });
    els.entry.appendChild(h('div', { class: 'entry-scroll' },
      h('table', { class: 'entry' }, h('thead', null, head), body)));
    els.entry.appendChild(h('div', { class: 'field-row entry-ot' },
      h('label', { for: 'f9-ot-hours', text: '加班時間：每日加班' }),
      (function () { var el = input(inp.otHours, '每日加班幾小時', function (v) { inp.otHours = v; }, 'in-tiny'); el.id = 'f9-ot-hours'; return el; })(),
      h('span', { text: '小時，加班出勤人數' }),
      (function () { var el = input(inp.otRate, '加班出勤人數 %', function (v) { inp.otRate = v; }, 'in-tiny'); el.id = 'f9-ot-rate'; return el; })(),
      h('span', { text: '%' })));

    // 需求人力 新機種+重工：各新機種各月幾人（可以增減機種）+ 重工各月幾人
    var nhead = h('tr', null, h('th', { scope: 'col', text: '新機種 / 重工' }));
    for (var nm = 0; nm < MONTHS; nm++) {
      var nth = h('th', { scope: 'col', text: monthName(nm) });
      els.monthHeads.push(nth);
      nhead.appendChild(nth);
    }
    nhead.appendChild(h('th', { scope: 'col' }));
    var nbody = h('tbody');
    inp.newModels.forEach(function (md, k) {
      var nameEl = h('input', { type: 'text', class: 'in entry-name', autocomplete: 'off', value: md.name,
        placeholder: '新機種名稱', 'aria-label': '第 ' + (k + 1) + ' 個新機種名稱' });
      nameEl.addEventListener('input', function () { md.name = nameEl.value; onType(); });
      var tr = h('tr', null, h('th', { scope: 'row' }, nameEl));
      md.qty.forEach(function (v, i) {
        tr.appendChild(h('td', null, input(v, (md.name || '新機種') + '，' + monthName(i) + '人數', function (val) { md.qty[i] = val; })));
      });
      tr.appendChild(h('td', null, h('button', { type: 'button', class: 'btn-link', text: '刪除',
        onclick: function () { inp.newModels.splice(k, 1); if (!inp.newModels.length) inp.newModels.push({ name: '', qty: blank6('') }); rebuildEntry(); } })));
      nbody.appendChild(tr);
    });
    var rtr = h('tr', null, h('th', { scope: 'row', text: '重工' }));
    inp.reworkPeople.forEach(function (v, i) {
      rtr.appendChild(h('td', null, input(v, '重工，' + monthName(i) + '人數', function (val) { inp.reworkPeople[i] = val; })));
    });
    rtr.appendChild(h('td'));
    nbody.appendChild(rtr);
    els.entry.appendChild(h('h3', { class: 'entry-title', text: '需求人力「新機種+重工」（人）' }));
    els.entry.appendChild(h('div', { class: 'entry-scroll' },
      h('table', { class: 'entry' }, h('thead', null, nhead), nbody)));
    els.entry.appendChild(h('div', { class: 'link-row' },
      h('button', { type: 'button', class: 'btn-link', text: '＋ 新增機種',
        onclick: function () { inp.newModels.push({ name: '', qty: blank6('') }); rebuildEntry(); } })));

    // 需求人力 Off line：各種人員的人數（每個月一樣）
    var offRow = h('div', { class: 'field-row entry-ot' }, h('span', { class: 'field-label', text: '需求人力「Off line」（每個月一樣）：' }));
    XIO.REPORT_OFFLINE.forEach(function (cat, k) {
      var id = 'f9-off-' + k;
      var el = input(inp.offline[k], '需求人力 Off line ' + cat + ' 人數', function (v) { inp.offline[k] = v; }, 'in-tiny');
      el.id = id;
      offRow.appendChild(h('label', { for: id, class: 'entry-cat', text: cat }));
      offRow.appendChild(el);
    });
    els.entry.appendChild(offRow);

    // 第二張表：YO 加班上限
    els.entry.appendChild(h('div', { class: 'field-row entry-ot' },
      h('label', { for: 'f9-yo-limit', text: 'YO 加班上限（第二張表）：' }),
      (function () { var el = input(inp.yoLimit, 'YO 加班上限幾小時', function (v) { inp.yoLimit = v; }, 'in-tiny'); el.id = 'f9-yo-limit'; return el; })(),
      h('span', { text: '小時/人（預設 ' + XIO.YO_LIMIT_DEFAULT + '）' })));

  }

  // 新增、刪除機種後整個輸入區重畫
  function rebuildEntry() {
    els.entry.textContent = '';
    els.monthHeads = [];
    buildEntry();
    onType();
  }

  function onType() {
    saveInputs();
    els.exportNote.textContent = '';
    clearTimeout(timer);
    timer = setTimeout(refresh, 250);
  }

  /* ---------- 來源 ---------- */
  function openUpload(file) {
    els.srcMsg.hidden = true;
    els.drop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        var names = res.workbook.worksheets.map(function (w) { return w.name; });
        if (!names.some(function (n) { return !isExtra(n); })) {
          throw new Error('這個檔案沒有 MPS 工作表。請上傳第 8 步匯出的檔案。');
        }
        S.upload = { name: file.name, buffer: res.buffer, sheets: names };
        S.preferUpload = true;
        els.exportNote.textContent = '';
        refresh();
      })
      .catch(function (err) {
        els.srcMsg.textContent = err && err.message ? err.message : String(err);
        els.srcMsg.hidden = false;
        refresh();
      });
  }

  function switchSource() {
    S.preferUpload = !S.preferUpload;
    els.exportNote.textContent = '';
    refresh();
  }

  // 先產生一份活頁簿，加上總表，給畫面預覽
  function ensureWork(src) {
    if (S.work && S.work.src === src.key) return;
    var work = { src: src.key, loading: true };
    S.work = work;
    buildFrom(src).then(function (wb) {
      if (S.work !== work) return;
      var main = pickMain(wb, src);
      if (!main) throw new Error('找不到 MPS 工作表。');
      var refs = monthRefs(main, src);
      work.refs = refs;
      work.links = XIO.findReportLinks(main, { dataStartRow: dataStartRow(src), months: XIO.REPORT.months });
      work.rma = XIO.readRmaQuantities(main, work.links.hoursRework);
      var mp = manpowerLinks(wb, src);
      Object.keys(mp).forEach(function (k) { work.links[k] = mp[k]; });
      work.wb = wb;
      work.main = main;
      work.loading = false;
      refresh();
    }).catch(function (err) {
      if (S.work !== work) return;
      work.loading = false;
      work.error = err && err.message ? err.message : String(err);
      refresh();
    });
  }

  /* ---------- 更新畫面 ---------- */
  function refresh() {
    var src = activeSource();
    var prev = IEApp.state.withManpower;
    IEApp.state.withReport = null;
    els.exportBtn.disabled = true;
    els.notes.textContent = '';
    els.result.textContent = '';
    els.preview.textContent = '';

    if (!src) {
      els.srcInfo.textContent = '還沒有第 8 步的結果。回到第 8 步，或上傳之前匯出的檔案。';
      els.drop.setIdle();
    } else if (src.kind === 'prev') {
      els.srcInfo.textContent = '沿用第 8 步剛做好的結果：' + prev.outName + '。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.upload) els.drop.setFile(S.upload.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.drop.setIdle();
    } else {
      els.srcInfo.textContent = '使用上傳的檔案，裡面的工作表：' + S.upload.sheets.join('、') + '。';
      els.drop.setFile(S.upload.name);
    }
    els.backBtn.hidden = !!src;
    els.switchBtn.hidden = !(S.upload && prev);
    if (src) els.switchBtn.textContent = src.kind === 'file' ? '改用第 8 步剛做好的結果' : '改用上傳的檔案';

    if (!src) { setSummary('先準備好第 8 步的 Excel。'); return; }
    ensureWork(src);
    if (S.work.loading) { setSummary('準備資料中…'); return; }
    if (S.work.error) { setSummary(S.work.error, true); return; }

    var work = S.work;
    // 輸入表的月份表頭
    els.monthHeads.forEach(function (th, m) { th.textContent = monthName(m % MONTHS); });   // 有兩個輸入表，各有六個月
    var read = readInputs();
    if (read.error) { setSummary(read.error, true); return; }
    // 每次都照目前的輸入重新產生總表（給預覽用）
    work.report = XIO.addReportSheet(work.wb, { title: read.values.title, monthRefs: work.refs || [], links: linkRefs(work.links),
      rmaQty: work.rma.qty, inputs: read.values, yoLimit: read.values.yoLimit });
    var ws = work.report.worksheet;
    var names = work.wb.worksheets.map(function (w) { return w.name; });
    var replacing = src.kind === 'file' && S.upload.sheets.some(function (n) { return sameName(n, SHEET_NAME); });
    var text = '匯出的檔案有 ' + names.length + ' 個工作表：' + names.join('、') + '。「' + SHEET_NAME +
      '」有兩張表：上面是法定加班上限 46hr，下面（第 ' + (XIO.REPORT.noteRow + XIO.REPORT_SECOND_OFFSET) + ' 列起）是 YO 加班上限 ' + read.values.yoLimit + 'hr。';
    if (replacing) text += ' 檔案裡原本的「' + SHEET_NAME + '」工作表會重新建立。';
    setSummary(text, false, replacing);

    if (work.refs && work.refs.length) {
      var r0 = work.refs[0];
      els.notes.appendChild(h('li', null, '月份表頭參照 ' + work.main.name + ' 工作表「需求工時」的月份，例如 D' + XIO.REPORT.headRow + ' ＝ ' +
        "='" + r0.sheet + "'!" + XIO.numToCol(r0.col) + r0.row + '（' + work.refs.map(function (r) { return r.text || '空白'; }).join('、') + '）。'));
    } else {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        work.main.name + ' 工作表裡找不到「需求工時」的月份，總表的月份表頭先留空白。'));
    }

    var warn = renderLinks(work);
    read.blanks.forEach(function (t) { warn = true; els.notes.appendChild(h('li', { class: 'is-warn' }, t)); });
    if (warn) els.summary.className = 'summary is-warn';

    ui.renderSheet(els.preview, ws, {
      headerEnd: XIO.REPORT.headRow,
      lastRow: XIO.REPORT.lastRow + XIO.REPORT_SECOND_OFFSET,
      colCount: XIO.REPORT.noteCol,
      rows: XIO.REPORT.lastRow + XIO.REPORT_SECOND_OFFSET,
      caption: '總表的版面（顏色、框線在 Excel 裡才看得到）。'
    });

    els.exportBtn.disabled = false;
    var refs = work.refs || [];
    var links = linkRefs(work.links);
    var inputs = read.values;
    var rmaQty = work.rma.qty;
    IEApp.state.withReport = {
      outName: outputName(src),
      mainSheet: work.main.name,
      reportSheet: SHEET_NAME,
      reportRows: work.report.rows,
      monthCol: work.report.monthCol,
      monthCount: work.report.months,
      inputs: inputs,
      build: function () { return buildWorkbook(src, refs, links, rmaQty, inputs); }
    };
  }

  // 各列從哪裡抓、抓到的數值；有找不到的回傳 true
  function renderLinks(work) {
    var months = XIO.REPORT.months;
    var rowLabel = {};
    var group = '';
    XIO.REPORT_ROWS.forEach(function (d) {
      if (d.group) group = d.group;
      if (d.label) group = '';
      rowLabel[d.key] = d.label || ((d.group || group) === 'MPS(PCS)' ? 'MPS(PCS) ' + d.unit : d.unit);
    });
    var warn = false;

    var head = h('tr', null,
      h('th', { class: 'colhead', scope: 'col', text: '總表' }),
      h('th', { class: 'colhead', scope: 'col', text: '來源' }));
    for (var i = 0; i < months; i++) {
      var r = work.refs && work.refs[i];
      head.appendChild(h('th', { class: 'colhead', scope: 'col',
        text: XIO.numToCol(XIO.REPORT.monthCol + i) + (r && r.text ? '　' + r.text : '') }));
    }
    var body = h('tbody');
    XIO.REPORT_LINKS.concat(MP_LINKS).forEach(function (link) {
      var info = work.links[link.key];
      var src;
      var isMp = !!info.sheet;
      if (!info.found && isMp) {
        warn = true;
        src = '找不到';
        els.notes.appendChild(h('li', { class: 'is-warn' }, '現有人力 ' + rowLabel[link.key] + '：' + info.reason + '，這一列先留空白。'));
      } else if (isMp) {
        src = info.sheet + ' 工作表「' + link.label + '」第 ' + info.row + ' 列，' +
          XIO.numToCol(info.refs[0].col) + ' 到 ' + XIO.numToCol(info.refs[info.refs.length - 1].col) + ' 行' +
          (info.from === 'saved' ? '（第 8 步上次確認的位置）' : '');
        var blank = info.refs.filter(function (x) { return x.value === null; }).length;
        if (blank) {
          warn = true;
          els.notes.appendChild(h('li', { class: 'is-warn' }, '現有人力 ' + rowLabel[link.key] + '：有 ' + blank + ' 個月是空白或不是數字。'));
        }
      } else if (!info.found) {
        warn = true;
        src = '找不到「' + link.label + '」';
        els.notes.appendChild(h('li', { class: 'is-warn' }, rowLabel[link.key] + '：' + work.main.name +
          ' 工作表裡找不到「' + link.label + '」這一列（第 ' + link.step + ' 步做的），這一列先留空白。'));
      } else {
        var first = info.refs[0], last = info.refs[info.refs.length - 1];
        src = '「' + link.label + '」第 ' + info.row + ' 列' +
          (first ? '，' + XIO.numToCol(first.col) + ' 到 ' + XIO.numToCol(last.col) + ' 行' : '');
        if (info.refs.length < months) {
          warn = true;
          els.notes.appendChild(h('li', { class: 'is-warn' }, rowLabel[link.key] + '：「' + link.label + '」那一列只找到 ' +
            info.refs.length + ' 個月的數值，其餘留空白。'));
        }
      }
      var tr = h('tr', null,
        h('th', { class: 'rowhead', scope: 'row', text: (isMp ? '現有人力 ' : '') + rowLabel[link.key] }),
        h('td', { text: src }));
      for (var m = 0; m < months; m++) {
        var ref = info.refs[m];
        var shown = !ref ? '' : ref.value === null ? (isMp ? '（空白）' : '（重新計算後才有）')
          : isMp ? String(Math.round(ref.value * 10) / 10) : Math.round(ref.value).toLocaleString('en-US');
        tr.appendChild(h('td', {
          class: ref && !(isMp && ref.value === null) ? 'is-num' : 'is-missing',
          text: shown
        }));
      }
      body.appendChild(tr);
    });

    // MPS(PCS) RMA：第 7 步各項目數量的合計
    (function () {
      var q = work.rma;
      var tr = h('tr', null,
        h('th', { class: 'rowhead', scope: 'row', text: rowLabel.mpsRma }),
        h('td', { text: q.items.length ? 'RMA 項目（' + q.items.join('、') + '）的數量合計' : '找不到 RMA 項目' }));
      for (var m = 0; m < months; m++) {
        var list = q.qty[m] || [];
        tr.appendChild(h('td', { class: list.length ? 'is-num' : 'is-missing',
          text: list.length ? list.reduce(function (a, b) { return a + b; }, 0).toLocaleString('en-US') : '' }));
      }
      body.insertBefore(tr, body.children[2] || null);
      if (!q.items.length) {
        warn = true;
        els.notes.appendChild(h('li', { class: 'is-warn' }, 'MPS(PCS) RMA：MPS 工作表裡找不到 RMA 項目（第 7 步），這一列先留空白。'));
      } else if (q.bad) {
        warn = true;
        els.notes.appendChild(h('li', { class: 'is-warn' }, 'RMA 區塊有 ' + q.bad + ' 格的公式不是「數量*CT.」的樣子，那幾格沒有算進 MPS(PCS) RMA。'));
      }
    })();

    var ex = work.links.mpsMass.found && work.links.mpsMass.refs[0];
    els.result.textContent = '';
    els.result.appendChild(h('p', { class: 'sheet-caption', text: '從其他工作表抓的列（需求工時只顯示整數）：' }));
    els.result.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': '總表已填的列' },
      h('table', { class: 'sheet is-result is-check' }, h('thead', null, head), body)));
    if (ex) {
      els.result.appendChild(h('p', { class: 'hint', text: '每格都是公式，例如總表 D' + XIO.REPORT_ROWS[0].row + ' ＝ =\'' +
        ex.sheet + "'!" + XIO.numToCol(ex.col) + ex.row + '。' }));
    }
    return warn;
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  /* ---------- 產生活頁簿 ---------- */
  function buildWorkbook(src, refs, links, rmaQty, inputs) {
    return buildFrom(src).then(function (wb) {
      XIO.addReportSheet(wb, { title: inputs.title, monthRefs: refs, links: links, rmaQty: rmaQty, inputs: inputs, yoLimit: inputs.yoLimit });
      return wb;
    });
  }

  function outputName(src) {
    var stem = src.name.replace(/\.[^.]+$/, '');
    if (stem.slice(-SUFFIX.length) === SUFFIX) stem = stem.slice(0, -SUFFIX.length);
    return stem + SUFFIX + '.xlsx';
  }

  /* ---------- 匯出 ---------- */
  function onExport() {
    var ready = IEApp.state.withReport;
    if (!ready) return;
    els.exportBtn.disabled = true;
    els.exportNote.textContent = '產生檔案中…';
    ready.build()
      .then(function (wb) { return XIO.toBlob(wb); })
      .then(function (blob) {
        ui.download(blob, ready.outName);
        els.exportNote.textContent = '已匯出 ' + ready.outName + '，在瀏覽器的下載資料夾。';
      })
      .catch(function (err) {
        els.exportNote.textContent = '匯出失敗：' + (err && err.message ? err.message : err);
      })
      .then(function () {
        els.exportBtn.disabled = false;
      });
  }

  IEApp.registerStep({ id: 'report', render: render });
})(window);
