/*
 * 第 2 步：匯總
 *   算出 MPS 各月的預計總產量（Total），再用 Status 分成量產（M）和試產（R）的合計。
 *   表格放在篩選後的 MPS 資料下面：最後一列資料之後空一列開始，共三列，對齊各月的行；
 *   文字標籤在第一個月左邊那一行。每格都是公式。
 *
 *   當月（第一個月）比較特別：數量要先加上 WIP。
 *   以 Status 在 C 行、WIP 在 D 行、當月在 E 行、資料在第 5 到 74 列為例：
 *
 *            當月                                                             其他月份
 *     Total  =SUM(D5:D74)+SUM(E5:E74)                                        =SUM(F5:F74)
 *     量產   =SUMIF($C$5:$C$74,"M",D5:D74)+SUMIF($C$5:$C$74,"M",E5:E74)       =SUMIF($C$5:$C$74,"M",F5:F74)
 *     試產   =SUMIF($C$5:$C$74,"R",D5:D74)+SUMIF($C$5:$C$74,"R",E5:E74)       =SUMIF($C$5:$C$74,"R",F5:F74)
 *
 *   Status 不是 M 也不是 R 的列只算進 Total，畫面上會列出來提醒。
 *   Total 那一列是淺綠色底（#B5E48C）；整塊外框粗線、裡面細格線。
 *
 *   WIP 和六個月各在哪一行由使用者指定，和第 5 步（計算需求工時）共用同一組設定，
 *   在其中一步改了，另一步也會跟著變，避免兩邊對不起來。
 *   Status 在哪一行：和第 1 步、第 6 步共用同一個設定（第 1 步指定，這裡改了另外兩步也跟著變）。
 *
 *   來源有兩種：
 *     a. 剛才第 1 步篩選好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 1 步匯出的檔案
 *
 * 做完之後把結果放在 IEApp.state.mpsTotals，下一步直接拿來用。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STORE_KEY = 'step2';
  var STEP1_KEY = 'step1';
  var REF_SHEET = 'P003';
  var LABELS = { total: 'Total' };
  // Status 等於 key 的列算進這一組（不分大小寫）：M = 量產、R = 試產（和第 6 步共用）
  var GROUPS = XIO.STATUS_GROUPS;
  var STATUS_GUESS = /status|狀態/i;
  var SUFFIX = '_匯總';
  var KEY_COL = 1;              // 料號在第 1 行
  var WIP_DEFAULT = 4;          // 第 4 行是當月 WIP
  var MONTH_FIRST = 5;          // 第 5 行起是月份
  var MONTH_COUNT = 6;          // 當月 + 往後五個月

  var S = {
    upload: null,         // 上傳的第 1 步檔案 { name, buffer }
    preferUpload: false,
    work: null            // 拿來預覽、統計的活頁簿 { src, loading, error, wb, main, mainExt }
  };
  var settings = { startRow: '', wipCol: '', monthCols: '', statusCol: '' };
  var els = {};
  var timer = null;

  function loadSettings() {
    var s = {};
    try { s = IEApp.settings.read(STORE_KEY) || {}; } catch (e) { s = {}; }
    var shared = ui.columns.get();          // WIP、月份的行和第 5 步共用
    settings = {
      startRow: s.startRow === undefined ? '' : s.startRow,
      statusCol: ui.statusCol.get(),
      wipCol: shared.wipCol,
      monthCols: shared.monthCols
    };
  }
  function saveSettings() {
    try {
      IEApp.settings.write(STORE_KEY, { startRow: settings.startRow });
    } catch (e) { /* 存不了就算了 */ }
    ui.columns.set({ wipCol: settings.wipCol, monthCols: settings.monthCols });
    ui.statusCol.set(settings.statusCol);
  }

  function isRefName(name) { return String(name).toLowerCase() === REF_SHEET.toLowerCase(); }

  function activeSource() {
    var prev = IEApp.state.mpsFiltered;
    if (S.upload && (S.preferUpload || !prev)) return { kind: 'file', key: S.upload, name: S.upload.name };
    if (prev) return { kind: 'prev', key: prev, name: prev.outName, prev: prev };
    return null;
  }

  function buildFrom(src) {
    return src.kind === 'prev' ? Promise.resolve(src.prev.build()) : XIO.load(src.key.buffer.slice(0));
  }

  function pickMain(wb, src) {
    var others = wb.worksheets.filter(function (w) { return !isRefName(w.name); });
    var wanted = src.kind === 'prev' ? src.prev.sheetName : null;
    return others.filter(function (w) { return w.name === wanted; })[0] || others[0] || null;
  }

  // MPS 的資料從第幾列開始：接著第 1 步做下來的話沿用第 1 步的設定
  function chainStartRow() {
    var src = activeSource();
    return src && src.kind === 'prev' ? src.prev.headerEnd + 1 : 0;
  }
  function startRow() {
    var chain = chainStartRow();
    if (chain) return chain;
    var n = parseInt(XIO.normalizeText(settings.startRow), 10);
    if (n >= 1) return n;
    try {
      var s1 = IEApp.settings.read(STEP1_KEY) || {};
      n = parseInt(XIO.normalizeText(s1.startRow === undefined ? '' : s1.startRow), 10);
    } catch (e) { n = 0; }
    return n >= 1 ? n : 5;
  }

  function colOf(letter, max) {
    var p = XIO.parseColumns(letter || '');
    return p.cols.length === 1 && !p.bad.length && p.cols[0] <= max ? p.cols[0] : 0;
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};
    loadSettings();     // 每次進來都重讀：WIP、月份、Status 可能在別的步驟改過

    els.srcInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.drop = ui.fileDrop({
      id: 'ft-base',
      title: '上傳第 1 步匯出的 Excel',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openUpload
    });
    els.switchBtn = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 1 步篩選 MPS',
      onclick: function () { IEApp.go('mps-filter'); }
    });
    els.srcMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    els.statusCol = h('select', { id: 'ft-status', class: 'in-select' });
    els.wipCol = h('select', { id: 'ft-wip', class: 'in-select' });
    els.monthCols = h('input', {
      id: 'ft-months', type: 'text', class: 'in in-select', autocomplete: 'off', spellcheck: 'false',
      placeholder: '例如 E:J'
    });
    els.mainPreview = h('div', { class: 'sheet-area' });
    els.startRow = h('input', { id: 'ft-start', type: 'text', class: 'in in-tiny', inputmode: 'numeric' });
    els.startNote = h('span');

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.notes = h('ul', { class: 'notes' });
    els.result = h('div', { class: 'sheet-area' });
    els.others = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：匯入 P003',
      onclick: function () { IEApp.go('p003-import'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '匯總' }),
      h('p', { text: '在篩選後的 MPS 下面加上各月的預計總產量（Total），再用 Status 分成量產（M）和試產（R）的合計。當月會先加上 WIP 再算。每格都是公式。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 1 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg),
      h('li', { class: 'task' },
        h('h2', { text: 'Status、WIP 和各月產量在哪幾行' }),
        h('div', { class: 'field-grid' },
          h('label', { for: 'ft-status', text: 'Status 在' }), els.statusCol,
          h('label', { for: 'ft-wip', text: '當月 WIP 在' }), els.wipCol,
          h('label', { for: 'ft-months', text: '各月產量在' }), els.monthCols),
        h('p', { class: 'hint', text: 'Status 預設是第 1 步指定的那一行，和第 6 步共用：M 算量產、R 算試產。' }),
        h('p', { class: 'hint', text: '各月產量的第一行是當月，會先加上 WIP 再算；連續的行可以寫成 E:J，也可以點下面表格最上方的行號來選。這裡的設定和第 5 步（計算需求工時）共用。' }),
        els.mainPreview,
        h('details', { class: 'advanced' },
          h('summary', { text: '報表版面不一樣時' }),
          h('div', { class: 'field-row' },
            h('label', { for: 'ft-start', text: 'MPS 的資料從第' }), els.startRow, els.startNote))),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.notes, els.result, els.others,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

    els.statusCol.addEventListener('change', function () {
      settings.statusCol = els.statusCol.value;
      saveSettings();
      els.exportNote.textContent = '';
      refresh();
    });
    els.wipCol.addEventListener('change', function () {
      settings.wipCol = els.wipCol.value;
      saveSettings();
      els.exportNote.textContent = '';
      refresh();
    });
    els.monthCols.addEventListener('input', function () {
      settings.monthCols = els.monthCols.value;
      saveSettings();
      els.exportNote.textContent = '';
      clearTimeout(timer);
      timer = setTimeout(function () { refresh(true); }, 250);
    });
    els.startRow.addEventListener('change', function () {
      settings.startRow = els.startRow.value;
      saveSettings();
      refresh();
    });

    refresh();
  }

  /* ---------- 來源 ---------- */
  function openUpload(file) {
    els.srcMsg.hidden = true;
    els.drop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        var others = res.workbook.worksheets.filter(function (w) { return !isRefName(w.name); });
        if (!others.length) throw new Error('這個檔案沒有 MPS 工作表。請上傳第 1 步匯出的檔案。');
        S.upload = { name: file.name, buffer: res.buffer };
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

  function ensureWork(src) {
    if (S.work && S.work.src === src.key) return;
    var work = { src: src.key, loading: true };
    S.work = work;
    buildFrom(src).then(function (wb) {
      if (S.work !== work) return;
      var main = pickMain(wb, src);
      if (!main) throw new Error('找不到 MPS 工作表。');
      work.wb = wb;
      work.main = main;
      work.mainExt = XIO.extent(main);
      work.loading = false;
      refresh();
    }).catch(function (err) {
      if (S.work !== work) return;
      work.loading = false;
      work.error = err && err.message ? err.message : String(err);
      refresh();
    });
  }

  function fillSelect(select, labels, colCount, chosen) {
    select.textContent = '';
    select.appendChild(h('option', { value: '', text: '請選擇' }));
    for (var c = 1; c <= colCount; c++) {
      var letter = XIO.numToCol(c);
      var label = labels[c] ? letter + '　' + labels[c] : letter;
      if (label.length > 42) label = label.slice(0, 41) + '…';
      select.appendChild(h('option', { value: letter, text: label }));
    }
    select.value = chosen ? XIO.numToCol(chosen) : '';
  }

  // 連續的行寫成 E:J，其餘用逗號分開
  function compact(cols) {
    if (!cols.length) return '';
    var contiguous = cols.every(function (c, i) { return i === 0 || c === cols[i - 1] + 1; });
    if (contiguous && cols.length > 1) return XIO.numToCol(cols[0]) + ':' + XIO.numToCol(cols[cols.length - 1]);
    return XIO.formatColumns(cols);
  }

  function toggleMonth(c) {
    var text = XIO.normalizeText(settings.monthCols) ? settings.monthCols : els.monthCols.value;
    var cols = XIO.parseColumns(text).cols;
    var i = cols.indexOf(c);
    if (i >= 0) cols.splice(i, 1); else cols.push(c);
    cols.sort(function (a, b) { return a - b; });
    settings.monthCols = compact(cols);
    saveSettings();
    els.exportNote.textContent = '';
    var scroller = els.mainPreview.querySelector('.sheet-scroll');
    var left = scroller ? scroller.scrollLeft : 0;
    refresh();
    scroller = els.mainPreview.querySelector('.sheet-scroll');
    if (scroller) scroller.scrollLeft = left;
    var btn = els.mainPreview.querySelector('button[data-col="' + c + '"]');
    if (btn) btn.focus({ preventScroll: true });
  }

  /* ---------- 更新畫面 ---------- */
  // keepMonthText：使用者正在打字時不要改寫輸入框的內容
  function refresh(keepMonthText) {
    var src = activeSource();
    var prev = IEApp.state.mpsFiltered;
    IEApp.state.mpsTotals = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;
    els.notes.textContent = '';
    els.result.textContent = '';
    els.others.textContent = '';

    if (!src) {
      els.srcInfo.textContent = '還沒有第 1 步的結果。回到第 1 步篩選 MPS，或上傳之前匯出的檔案。';
      els.drop.setIdle();
    } else if (src.kind === 'prev') {
      els.srcInfo.textContent = '沿用第 1 步剛篩選好的結果：' + prev.outName + '（' +
        prev.matchedRows.toLocaleString('en-US') + ' 列資料）。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.upload) els.drop.setFile(S.upload.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.drop.setIdle();
    } else {
      els.srcInfo.textContent = '使用上傳的檔案。';
      els.drop.setFile(S.upload.name);
    }
    els.backBtn.hidden = !!src;
    els.switchBtn.hidden = !(S.upload && prev);
    if (src) els.switchBtn.textContent = src.kind === 'file' ? '改用第 1 步剛篩選好的結果' : '改用上傳的檔案';

    var chain = chainStartRow();
    els.startRow.value = String(startRow());
    els.startRow.disabled = !!chain;
    els.startNote.textContent = chain ? '列開始（沿用第 1 步的設定）' : '列開始';

    function clear() {
      [els.wipCol, els.statusCol].forEach(function (sel) {
        sel.textContent = '';
        sel.appendChild(h('option', { value: '', text: '請選擇' }));
      });
      els.mainPreview.textContent = '';
    }
    if (!src) { clear(); setSummary('先準備好第 1 步的 Excel。'); return; }

    ensureWork(src);
    var work = S.work;
    if (work.loading) { clear(); setSummary('準備資料中…'); return; }
    if (work.error) { clear(); setSummary(work.error, true); return; }

    var sr = startRow();
    var headerEnd = sr - 1;
    var cols = Math.max(work.mainExt.colCount, 1);
    var labels = XIO.columnLabels(work.main, headerEnd >= 2 ? 2 : 1, Math.max(headerEnd, 1), cols);

    // WIP、月份：先用記住的，沒有就用預設的位置（第 4 行、第 5 到 10 行）
    var wipCol = colOf(settings.wipCol, cols) || (WIP_DEFAULT <= cols ? WIP_DEFAULT : 0);
    var parsed = XIO.parseColumns(settings.monthCols);
    var monthCols = parsed.cols;
    if (!XIO.normalizeText(settings.monthCols)) {
      monthCols = [];
      for (var c = MONTH_FIRST; c < MONTH_FIRST + MONTH_COUNT && c <= cols; c++) monthCols.push(c);
    }

    // Status：共用設定（第 1 步指定的）→ 表頭有「Status」的那一行
    var statusCol = colOf(settings.statusCol, cols);
    if (!statusCol) {
      for (var g = 1; g <= cols; g++) if (STATUS_GUESS.test(labels[g] || '')) { statusCol = g; break; }
    }

    fillSelect(els.statusCol, labels, cols, statusCol);
    fillSelect(els.wipCol, labels, cols, wipCol);
    if (!keepMonthText) els.monthCols.value = XIO.normalizeText(settings.monthCols) ? settings.monthCols : compact(monthCols);

    var keep = {};
    monthCols.forEach(function (mc) { keep[mc] = true; });
    ui.renderSheet(els.mainPreview, work.main, {
      headerEnd: headerEnd, lastRow: work.mainExt.lastRow, colCount: cols, rows: 5, keep: keep,
      onToggle: toggleMonth,
      caption: work.main.name + ' 工作表（前 5 列資料）。標起來的是各月產量。'
    });

    if (parsed.bad.length) {
      setSummary('看不懂「' + parsed.bad.join('、') + '」。行號請用英文字母，例如 E:J。', true);
      return;
    }
    if (!statusCol) { setSummary('選擇 Status 在哪一行。'); return; }
    if (!wipCol) { setSummary('選擇當月 WIP 在哪一行。'); return; }
    if (!monthCols.length) { setSummary('輸入各月產量在哪幾行。'); return; }
    var clash = monthCols.filter(function (mc) { return mc === wipCol || mc === KEY_COL || mc === statusCol || mc > cols; });
    if (clash.length) {
      setSummary(XIO.formatColumns(clash) + ' 行不能當作各月產量（和 Status、WIP 或料號重複，或超出報表範圍）。', true);
      return;
    }

    var opts = {
      dataStartRow: sr, keyCol: KEY_COL, wipCol: wipCol, monthCols: monthCols, statusCol: statusCol,
      groups: GROUPS, labels: LABELS
    };
    var plan = XIO.planMpsTotals(work.main, opts);
    if (!plan.found) {
      setSummary('第 ' + sr + ' 列之後沒有料號。請確認上傳的檔案。', true);
      return;
    }

    var firstL = XIO.numToCol(monthCols[0]);
    var lastL = XIO.numToCol(monthCols[monthCols.length - 1]);
    var hasError = plan.months.some(function (m) { return !!m.error; });
    var endRow = plan.rows[plan.rows.length - 1];
    setSummary('匯總會填在 ' + work.main.name + ' 工作表的第 ' + plan.totalRow + ' 到 ' + endRow +
      ' 列（最後一列資料是第 ' + plan.lastRow + ' 列，空一列），' + firstL + ' 到 ' + lastL + ' 行，共 ' +
      monthCols.length + ' 個月；文字標籤在 ' + XIO.numToCol(plan.labelCol) + ' 行。', false, hasError);

    if (hasError) {
      var bad = plan.months.filter(function (m) { return !!m.error; }).map(function (m) { return XIO.numToCol(m.col); });
      els.notes.appendChild(h('li', { class: 'is-warn' },
        bad.join('、') + ' 行有不是數字的格子，Excel 裡這幾個月的匯總會顯示錯誤。請先檢查 MPS 的數量。'));
    }
    GROUPS.forEach(function (gr, k) {
      if (!plan.counts[k]) {
        els.notes.appendChild(h('li', { class: 'is-warn' },
          XIO.numToCol(statusCol) + ' 行沒有「' + gr.key + '」的列，' + gr.label + '每個月都會是 0。請確認 Status 選對行。'));
      }
    });
    if (plan.others.length) {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '有 ' + plan.others.length + ' 列的 Status 不是 M 也不是 R，只會算進 Total，不算進量產、試產。清單在下面。'));
    }
    if (monthCols.length !== MONTH_COUNT) {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '目前選了 ' + monthCols.length + ' 個月份，通常是 ' + MONTH_COUNT + ' 個（當月加往後五個月）。'));
    }

    renderResult(plan, labels, opts);
    renderOthers(plan);
    els.exportBtn.disabled = false;
    els.nextBtn.disabled = false;

    // 給下一步用（欄位和第 1 步的結果一樣，下一步可以直接接著做）
    var rows = plan.lastRow - plan.firstRow + 1;
    IEApp.state.mpsTotals = {
      sourceName: src.name,
      outName: outputName(src),
      sheetName: work.main.name,
      matchedRows: src.kind === 'prev' ? src.prev.matchedRows : rows,
      headerEnd: headerEnd,
      wipCol: wipCol,
      monthCols: monthCols,
      statusCol: statusCol,
      build: function () { return buildWorkbook(src, opts); }
    };
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  function showNumber(v) {
    return Math.round(v).toLocaleString('en-US');
  }

  function monthHead(m, labels, i) {
    var text = XIO.numToCol(m.col) + (labels[m.col] ? '　' + labels[m.col] : '') + (i === 0 ? '（加 WIP）' : '');
    return h('th', { class: 'colhead', scope: 'col', text: text });
  }

  // 匯總列算出來的數字，以及當月和下一個月填入的公式
  function renderResult(plan, labels, opts) {
    var head = h('tr', null,
      h('th', { class: 'colhead', scope: 'col', text: '列' }),
      h('th', { class: 'colhead', scope: 'col', text: '項目' }));
    plan.months.forEach(function (m, i) { head.appendChild(monthHead(m, labels, i)); });

    var items = [{ label: LABELS.total, row: plan.totalRow, value: function (m) { return m.total; }, cls: ' is-total' }];
    GROUPS.forEach(function (gr, k) {
      items.push({ label: gr.label + '（' + gr.key + '，' + plan.counts[k] + ' 列）', row: plan.groupRows[k],
        value: function (m) { return m.groups[k]; }, cls: '' });
    });

    var body = h('tbody');
    items.forEach(function (it) {
      var tr = h('tr', null,
        h('th', { class: 'rowhead', scope: 'row', text: String(it.row) }),
        h('td', { text: it.label }));
      plan.months.forEach(function (m) {
        tr.appendChild(h('td', { class: 'is-num' + it.cls, text: m.error ? m.error : showNumber(it.value(m)) }));
      });
      body.appendChild(tr);
    });

    els.result.appendChild(h('p', { class: 'sheet-caption', text: '匯總結果。' }));
    els.result.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': '匯總結果' },
      h('table', { class: 'sheet is-result' }, h('thead', null, head), body)));

    var names = [LABELS.total].concat(GROUPS.map(function (gr) { return gr.label; }));
    function explain(i, prefix) {
      var L = XIO.numToCol(plan.months[i].col);
      var f = XIO.mpsTotalsFormulas(plan, opts, i);
      els.result.appendChild(h('p', { class: 'hint', text: prefix + plan.rows.map(function (r, j) {
        return names[j] + ' ' + L + r + ' =' + f[j];
      }).join('　') }));
    }
    explain(0, '當月（' + XIO.numToCol(plan.months[0].col) + ' 行，加 WIP）的公式：');
    if (plan.months.length > 1) explain(1, '其他月份（以 ' + XIO.numToCol(plan.months[1].col) + ' 行為例）：');
  }

  // Status 不是 M 也不是 R 的列
  function renderOthers(plan) {
    if (!plan.others.length) return;
    var head = h('tr', null,
      h('th', { class: 'colhead', scope: 'col', text: '列' }),
      h('th', { class: 'colhead', scope: 'col', text: '料號' }),
      h('th', { class: 'colhead', scope: 'col', text: 'Status' }));
    var body = h('tbody');
    plan.others.forEach(function (o) {
      body.appendChild(h('tr', null,
        h('th', { class: 'rowhead', scope: 'row', text: String(o.row) }),
        h('td', { text: o.label }),
        h('td', { text: o.status === '' ? '（空白）' : o.status })));
    });
    els.others.appendChild(h('p', { class: 'sheet-caption', text: 'Status 不是 M 也不是 R 的列（只算進 Total）。' }));
    els.others.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': 'Status 不是 M 或 R 的列' },
      h('table', { class: 'sheet is-result' }, h('thead', null, head), body)));
  }

  /* ---------- 產生活頁簿 ---------- */
  function buildWorkbook(src, opts) {
    return buildFrom(src).then(function (wb) {
      XIO.addMpsTotals(pickMain(wb, src), opts);
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
    var ready = IEApp.state.mpsTotals;
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

  IEApp.registerStep({ id: 'mps-totals', render: render });
})(window);
