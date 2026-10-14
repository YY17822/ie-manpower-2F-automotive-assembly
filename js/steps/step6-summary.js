/*
 * 第 6 步：需求工時匯總
 *   在需求工時那一塊的下面（最後一列資料之後空一列）加三列，每格都是公式。
 *   以 Status 在 C 行、需求工時在 N 行、資料在第 5 到 74 列為例：
 *
 *     Production 總需求工時   =SUM(N5:N74)                     每月所有料號加總
 *     量產總需求工時          =SUMIF($C$5:$C$74,"M",N5:N74)    Status 是 M 的料號
 *     試產總需求工時          =SUMIF($C$5:$C$74,"R",N5:N74)    Status 是 R 的料號
 *
 *   這三列連同左邊的文字標籤圍成一塊：外框粗線、裡面細格線，標籤置中。
 *   Production 那一列是淺藍色底（#A2D2FF）。不再用顏色標示前三大。
 *   Status 不是 M 也不是 R 的列只算進 Production，畫面上會列出來提醒。
 *
 *   Status 在哪一行：和第 1 步、第 2 步共用同一個設定。
 *
 *   來源有兩種：
 *     a. 剛才第 5 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 5 步匯出的檔案（MPS 工作表裡要有「需求工時」）
 *
 * 做完之後把結果放在 IEApp.state.withSummary，下一步直接拿來用。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STORE_KEY = IEApp.storeKey('step5');       // 鍵名是舊的步驟編號，改掉會讓記住的設定不見，所以沿用
  var STEP1_KEY = IEApp.storeKey('step1');
  var REF_SHEET = 'P003';
  var TITLE = '需求工時';
  var SUFFIX = '_工時匯總';
  var GROUPS = XIO.STATUS_GROUPS;     // M = 量產、R = 試產（和第 2 步共用）
  var LABELS = { total: 'Production 總需求工時', groups: GROUPS.map(function (g) { return g.label + '總需求工時'; }) };
  var STATUS_GUESS = /status|狀態/i;

  var S = {
    upload: null,         // 上傳的第 5 步檔案 { name, buffer }
    preferUpload: false,
    work: null            // 拿來預覽的活頁簿 { src, loading, error, wb, main }
  };
  var settings = loadSettings();
  var els = {};

  function loadSettings() {
    var s = {};
    try { s = JSON.parse(localStorage.getItem(STORE_KEY) || '{}') || {}; } catch (e) { s = {}; }
    return { startRow: s.startRow === undefined ? '' : s.startRow, statusCol: ui.statusCol.get() };
  }
  function saveSettings() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify({ startRow: settings.startRow })); } catch (e) { /* 存不了就算了 */ }
    ui.statusCol.set(settings.statusCol);
  }

  function isRefName(name) { return String(name).toLowerCase() === REF_SHEET.toLowerCase(); }

  function activeSource() {
    var prev = IEApp.state.withDemand;
    if (S.upload && (S.preferUpload || !prev)) return { kind: 'file', key: S.upload, name: S.upload.name };
    if (prev) return { kind: 'prev', key: prev, name: prev.outName, prev: prev };
    return null;
  }

  function buildFrom(src) {
    return src.kind === 'prev' ? src.prev.build() : XIO.load(src.key.buffer.slice(0));
  }

  function pickMain(wb, src) {
    var others = wb.worksheets.filter(function (w) { return !isRefName(w.name); });
    var wanted = src.kind === 'prev' ? src.prev.mainSheet : null;
    return others.filter(function (w) { return w.name === wanted; })[0] || others[0] || null;
  }

  function chainStartRow() {
    var src = activeSource();
    return src && src.kind === 'prev' && src.prev.dataStartRow ? src.prev.dataStartRow : 0;
  }
  function startRow() {
    var chain = chainStartRow();
    if (chain) return chain;
    var n = parseInt(XIO.normalizeText(settings.startRow), 10);
    if (n >= 1) return n;
    try {
      var s1 = JSON.parse(localStorage.getItem(STEP1_KEY) || '{}') || {};
      n = parseInt(XIO.normalizeText(s1.startRow === undefined ? '' : s1.startRow), 10);
    } catch (e) { n = 0; }
    return n >= 1 ? n : 5;
  }

  function colOf(letter, max) {
    var p = XIO.parseColumns(letter || '');
    return p.cols.length === 1 && !p.bad.length && p.cols[0] <= max ? p.cols[0] : 0;
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

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};
    settings = loadSettings();     // 每次進來都重讀：Status 可能在別的步驟改過

    els.srcInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.drop = ui.fileDrop({
      id: 'f5-base',
      title: '上傳第 5 步匯出的 Excel',
      sub: '或把檔案拖進這裡。MPS 工作表裡要有「需求工時」',
      onFile: openUpload
    });
    els.switchBtn = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 5 步計算需求工時',
      onclick: function () { IEApp.go('demand-hours'); }
    });
    els.srcMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });
    els.startRow = h('input', { id: 'f5-start', type: 'text', class: 'in in-tiny', inputmode: 'numeric' });
    els.startNote = h('span');

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.notes = h('ul', { class: 'notes' });
    els.result = h('div', { class: 'sheet-area' });
    els.statusCol = h('select', { id: 'f5-status', class: 'in-select' });
    els.others = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：RMA 需求工時',
      onclick: function () { IEApp.go('rma'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '需求工時匯總' }),
      h('p', { text: '在需求工時下面加上每個月正常生產的總需求工時（Production），再用 Status 分成量產（M）和試產（R）的合計。每格都是公式。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 5 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg,
        h('details', { class: 'advanced' },
          h('summary', { text: '報表版面不一樣時' }),
          h('div', { class: 'field-row' },
            h('label', { for: 'f5-start', text: 'MPS 的資料從第' }), els.startRow, els.startNote))),
      h('li', { class: 'task' },
        h('h2', { text: 'Status 在哪一行' }),
        h('div', { class: 'field-row' },
          h('label', { for: 'f5-status', text: 'Status 在' }), els.statusCol),
        h('p', { class: 'hint', text: 'M 算量產、R 算試產。預設是第 1 步指定的那一行，和第 2 步共用。' })),
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
        if (!others.length) throw new Error('這個檔案沒有 MPS 工作表。請上傳第 5 步匯出的檔案。');
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
    var prev = IEApp.state.withDemand;
    IEApp.state.withSummary = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;
    els.notes.textContent = '';
    els.result.textContent = '';
    els.others.textContent = '';

    if (!src) {
      els.srcInfo.textContent = '還沒有第 5 步的結果。回到第 5 步計算需求工時，或上傳之前匯出的檔案。';
      els.drop.setIdle();
    } else if (src.kind === 'prev') {
      els.srcInfo.textContent = '沿用第 5 步剛做好的結果：' + prev.outName + '。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.upload) els.drop.setFile(S.upload.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.drop.setIdle();
    } else {
      els.srcInfo.textContent = '使用上傳的檔案。';
      els.drop.setFile(S.upload.name);
    }
    els.backBtn.hidden = !!src;
    els.switchBtn.hidden = !(S.upload && prev);
    if (src) els.switchBtn.textContent = src.kind === 'file' ? '改用第 5 步剛做好的結果' : '改用上傳的檔案';

    var chain = chainStartRow();
    els.startRow.value = String(startRow());
    els.startRow.disabled = !!chain;
    els.startNote.textContent = chain ? '列開始（沿用第 1 步的設定）' : '列開始';

    function clearStatus() {
      els.statusCol.textContent = '';
      els.statusCol.appendChild(h('option', { value: '', text: '請選擇' }));
    }
    if (!src) { clearStatus(); setSummary('先準備好第 5 步的 Excel。'); return; }

    ensureWork(src);
    var work = S.work;
    if (work.loading) { clearStatus(); setSummary('準備資料中…'); return; }
    if (work.error) { clearStatus(); setSummary(work.error, true); return; }

    var sr = startRow();
    var keyCol = src.kind === 'prev' && src.prev.keyCol ? src.prev.keyCol : 1;

    // Status：共用設定（第 1 步指定的）→ 表頭有「Status」的那一行
    var ext = XIO.extent(work.main);
    var cols = Math.max(ext.colCount, 1);
    var labels = XIO.columnLabels(work.main, sr - 1 >= 2 ? 2 : 1, Math.max(sr - 1, 1), cols);
    var statusCol = colOf(settings.statusCol, cols);
    if (!statusCol) {
      for (var g = 1; g <= cols; g++) if (STATUS_GUESS.test(labels[g] || '')) { statusCol = g; break; }
    }
    fillSelect(els.statusCol, labels, cols, statusCol);
    if (!statusCol) { setSummary('選擇 Status 在哪一行。'); return; }

    var opts = { dataStartRow: sr, keyCol: keyCol, title: TITLE, statusCol: statusCol, groups: GROUPS, labels: LABELS };
    var plan = XIO.planSummary(work.main, opts);

    if (!plan.found) {
      setSummary('MPS 工作表裡找不到「需求工時」。請先做第 5 步計算需求工時。', true);
      return;
    }
    if (!plan.lastRow) {
      setSummary('第 ' + sr + ' 列之後沒有料號。請確認上傳的檔案。', true);
      return;
    }

    var firstL = XIO.numToCol(plan.startCol);
    var lastL = XIO.numToCol(plan.startCol + plan.width - 1);
    var hasError = plan.months.some(function (m) { return !!m.error; });
    setSummary('匯總會填在 ' + work.main.name + ' 工作表的第 ' + plan.totalRow + ' 到 ' + plan.restRow +
      ' 列（最後一列資料是第 ' + plan.lastRow + ' 列，空一列），' + firstL + ' 到 ' + lastL + ' 行，共 ' +
      plan.width + ' 個月。', false, hasError);

    if (hasError) {
      var bad = plan.months.filter(function (m) { return !!m.error; })
        .map(function (m) { return XIO.numToCol(m.col) + (m.header ? '（' + m.header + '）' : ''); });
      els.notes.appendChild(h('li', { class: 'is-warn' },
        bad.join('、') + ' 行有算不出來的格子（例如 #DIV/0!），Excel 裡這幾個月的匯總也會顯示錯誤。請先回第 4 步處理。'));
    }

    GROUPS.forEach(function (gr, k) {
      if (!plan.counts[k]) {
        els.notes.appendChild(h('li', { class: 'is-warn' },
          XIO.numToCol(statusCol) + ' 行沒有「' + gr.key + '」的列，' + gr.label + '每個月都會是 0。請確認 Status 選對行。'));
      }
    });
    if (plan.others.length) {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '有 ' + plan.others.length + ' 列的 Status 不是 M 也不是 R，只會算進 Production，不算進量產、試產。清單在下面。'));
    }

    renderResult(plan, opts);
    renderOthers(plan);
    els.exportBtn.disabled = false;

    // 給下一步用
    IEApp.state.withSummary = {
      outName: outputName(src),
      mainSheet: work.main.name,
      dataStartRow: sr,
      keyCol: keyCol,
      demandCol: plan.startCol,
      monthCount: plan.width,
      totalRow: plan.totalRow,
      groupRows: plan.groupRows,
      restRow: plan.restRow,
      statusCol: statusCol,
      build: function () { return buildWorkbook(src, opts); }
    };
    els.nextBtn.disabled = false;
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  function showNumber(v) {
    return Math.round(v).toLocaleString('en-US');             // 和 Excel 一樣只顯示整數
  }

  function monthHead(m) {
    return h('th', { class: 'colhead', scope: 'col', text: XIO.numToCol(m.col) + (m.header ? '　' + m.header : '') });
  }

  // 三個匯總列算出來的數字，以及第一個月填入的公式
  function renderResult(plan, opts) {
    var head = h('tr', null,
      h('th', { class: 'colhead', scope: 'col', text: '列' }),
      h('th', { class: 'colhead', scope: 'col', text: '項目' }));
    plan.months.forEach(function (m) { head.appendChild(monthHead(m)); });

    var items = [{ label: LABELS.total, row: plan.totalRow, value: function (m) { return m.total; } }];
    GROUPS.forEach(function (gr, k) {
      items.push({ label: LABELS.groups[k] + '（' + gr.key + '，' + plan.counts[k] + ' 列）', row: plan.groupRows[k],
        value: function (m) { return m.groups[k]; } });
    });

    var body = h('tbody');
    items.forEach(function (it) {
      var tr = h('tr', null,
        h('th', { class: 'rowhead', scope: 'row', text: String(it.row) }),
        h('td', { text: it.label }));
      plan.months.forEach(function (m) {
        tr.appendChild(h('td', { class: 'is-num', text: m.error ? m.error : showNumber(it.value(m)) }));
      });
      body.appendChild(tr);
    });

    els.result.appendChild(h('p', { class: 'sheet-caption', text: '匯總結果。' }));
    els.result.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': '匯總結果' },
      h('table', { class: 'sheet is-result' }, h('thead', null, head), body)));

    var m0 = plan.months[0];
    if (!m0) return;
    var L = XIO.numToCol(m0.col);
    var f = XIO.summaryFormulas(plan, opts, m0);
    var names = [LABELS.total].concat(LABELS.groups);
    els.result.appendChild(h('p', { class: 'hint', text: L + ' 行填入的公式（其他月份一樣）：' +
      plan.rows.map(function (r, j) { return names[j] + ' ' + L + r + ' =' + f[j]; }).join('　') }));
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
    els.others.appendChild(h('p', { class: 'sheet-caption', text: 'Status 不是 M 也不是 R 的列（只算進 Production）。' }));
    els.others.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': 'Status 不是 M 或 R 的列' },
      h('table', { class: 'sheet is-result' }, h('thead', null, head), body)));
  }

  /* ---------- 產生活頁簿 ---------- */
  function buildWorkbook(src, opts) {
    return buildFrom(src).then(function (wb) {
      XIO.addSummaryRows(pickMain(wb, src), opts);
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
    var ready = IEApp.state.withSummary;
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

  IEApp.registerStep({ id: 'summary', render: render });
})(window);
