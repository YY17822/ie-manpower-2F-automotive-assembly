/*
 * 第 1 步：篩選 MPS 報表（好幾份報表合成一份）
 *   1. 上傳 MPS 的 Excel，一次可以好幾份（通常 3、4 份），每份報表版面一樣：
 *      第 1–4 列表頭、第 5 列開始資料、A 行是選取欄
 *   2. 每份報表各自輸入「A 行等於多少」和「要保留哪些行」
 *      ——各報表保留的行位置可以不同，但保留的行數和順序（相對位置）要一樣
 *   3. 指定 Status 在哪一行（區分量產、試產，後面總表會用到）
 *   4. 匯出一份新的 Excel：表頭用報表 1 的（含合併儲存格），
 *      各報表符合條件的列依序往下接；同一個料號出現在不同報表就保留成好幾列
 *
 * 做完之後把結果放在 IEApp.state.mpsFiltered，第 2 步直接拿來用。
 * Status 在新檔的哪一行：IEApp.state.mpsFiltered.statusCol（數字），
 * 也記在設定裡（statusCol，英文字母），之後的步驟上傳檔案時也讀得到。
 *
 * 名詞：行 = 直的（A、B、C…），列 = 橫的（1、2、3…）。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STORE_KEY = IEApp.storeKey('step1');
  var PREVIEW_ROWS = 8;     // 預覽顯示幾列資料
  var MAX_CHIPS = 24;
  var STATUS_GUESS = /status|狀態/i;

  // reports：每份報表各自的條件，依上傳順序（報表 1、報表 2…），下個月照同樣順序上傳就會自動帶入
  // statusPos：Status 是「第幾個保留的行」（1 起算），也就是新檔的第幾行
  var DEFAULTS = { startRow: 5, selCol: 'A', reports: [], statusPos: '' };

  // 這一步的狀態（切到別的步驟再回來時還在）
  // S.reports[i] = { fileName, workbook, ws, info, infoKey, els }
  var S = { reports: [], result: null };
  var settings = loadSettings();
  var els = {};
  var timer = null;
  var seq = 0;

  /* ---------- 記住上次輸入的條件（每個月做的事情差不多） ---------- */
  function loadSettings() {
    var s = {};
    try { s = JSON.parse(localStorage.getItem(STORE_KEY) || '{}') || {}; } catch (e) { s = {}; }
    var out = {};
    Object.keys(DEFAULTS).forEach(function (k) { out[k] = s[k] === undefined ? DEFAULTS[k] : s[k]; });
    if (!Array.isArray(out.reports)) out.reports = [];
    // 只有一份報表的舊版設定（values、cols）當作報表 1
    if (!out.reports.length && (s.values || s.cols)) out.reports.push({ values: s.values || '', cols: s.cols || '' });
    return out;
  }
  function saveSettings() {
    var keep = {
      startRow: settings.startRow, selCol: settings.selCol, reports: settings.reports,
      statusPos: settings.statusPos,
      statusCol: statusPos() ? XIO.numToCol(statusPos()) : ''
    };
    try { localStorage.setItem(STORE_KEY, JSON.stringify(keep)); } catch (e) { /* 存不了就算了 */ }
  }
  function slot(i) {
    while (settings.reports.length <= i) settings.reports.push({ values: '', cols: '' });
    return settings.reports[i];
  }

  /* ---------- 讀取目前的設定值 ---------- */
  function startRow() {
    var n = parseInt(XIO.normalizeText(settings.startRow), 10);
    return n >= 1 ? n : DEFAULTS.startRow;
  }
  function selCol() {
    var p = XIO.parseColumns(settings.selCol);
    return p.cols.length === 1 && !p.bad.length ? p.cols[0] : 1;
  }
  function statusPos() {
    var n = parseInt(settings.statusPos, 10);
    return n >= 1 ? n : 0;
  }
  function wantedValues(i) {
    return XIO.normalizeText(slot(i).values).split(/[,、;]+/)
      .map(function (v) { return v.trim(); })
      .filter(function (v) { return v !== ''; });
  }
  function keptCols(i) {
    return XIO.parseColumns(slot(i).cols);
  }
  function outputName() {
    var base = S.reports[0].fileName.replace(/\.[^.]+$/, '');
    return base + (S.reports.length > 1 ? '_等' + S.reports.length + '份' : '') + '_篩選.xlsx';
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};

    els.drop = ui.fileDrop({
      id: 'f-file',
      title: '選擇 MPS 的 Excel 檔（可以一次選好幾個）',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFiles: addFiles
    });
    els.fileList = h('ol', { class: 'report-list' });
    els.fileMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    els.startRow = h('input', {
      id: 'f-start', type: 'text', class: 'in in-tiny', inputmode: 'numeric', value: String(settings.startRow)
    });
    els.selCol = h('input', {
      id: 'f-selcol', type: 'text', class: 'in in-tiny', autocomplete: 'off', value: String(settings.selCol)
    });

    els.cards = h('div', { class: 'report-cards' });
    els.status = h('select', { id: 'f-status', class: 'in-select' });
    els.statusChecks = h('ul', { class: 'notes' });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.parts = h('ul', { class: 'notes' });
    els.outPreview = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：匯總',
      onclick: function () { IEApp.go('mps-totals'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '篩選 MPS 報表' }),
      h('p', { text: '上傳好幾份 MPS 報表，各自留下 A 行等於指定數值的列和要的行，往下接成一份新的 Excel。表頭（第 1–4 列）用報表 1 的，原樣保留。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '上傳 MPS 報表' }),
        els.drop.el, els.fileList, els.fileMsg,
        h('p', { class: 'hint', text: '依上傳順序叫報表 1、報表 2…，合併時也照這個順序往下接。每份報表輸入的條件會記住，下個月照同樣順序上傳就會自動帶入。' }),
        h('details', { class: 'advanced' },
          h('summary', { text: '報表版面不一樣時' }),
          h('div', { class: 'field-row' },
            h('label', { for: 'f-start', text: '資料從第' }), els.startRow, h('span', { text: '列開始（之前的列都當作表頭保留）' })),
          h('div', { class: 'field-row' },
            h('label', { for: 'f-selcol', text: '選取條件在' }), els.selCol, h('span', { text: '行' })),
          h('p', { class: 'hint', text: '每份報表都用同一組版面設定。' }))),
      h('li', { class: 'task' },
        h('h2', { text: '各報表的選取條件和要保留的行' }),
        h('p', { class: 'hint', text: '每份報表保留的行位置可以不同，但行數和順序要一樣：報表 1 保留的第 1 個行，和報表 2 保留的第 1 個行，會放在新檔的同一行。' }),
        els.cards),
      h('li', { class: 'task' },
        h('h2', { text: 'Status 在哪一行' }),
        h('div', { class: 'field-row' },
          h('label', { for: 'f-status', text: 'Status' }), els.status),
        h('p', { class: 'hint', text: 'Status 用來區分料號是量產還是試產，後面的總表會用到。從報表 1 保留的行裡選；其他報表會用同一個相對位置。' }),
        els.statusChecks),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.parts, els.outPreview,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

    bindSetting(els.startRow, 'startRow');
    bindSetting(els.selCol, 'selCol');
    els.status.addEventListener('change', function () {
      settings.statusPos = els.status.value;
      els.exportNote.textContent = '';
      recompute();
    });

    S.reports.forEach(function (R) { R.infoKey = ''; });
    renderReports();
  }

  function bindSetting(input, key) {
    input.addEventListener('input', function () {
      settings[key] = input.value;
      S.reports.forEach(function (R) { R.infoKey = ''; });
      schedule();
    });
  }
  function schedule() {
    els.exportNote.textContent = '';
    clearTimeout(timer);
    timer = setTimeout(recompute, 200);
  }

  /* ---------- 開檔 ---------- */
  function addFiles(files) {
    els.fileMsg.hidden = true;
    els.drop.setBusy(files.map(function (f) { return f.name; }).join('、'));
    var errors = [];
    // 依選檔順序一份一份讀，報表編號才會跟著選的順序
    files.reduce(function (p, file) {
      return p.then(function () {
        return ui.readWorkbook(file).then(function (res) {
          S.reports.push({ id: ++seq, fileName: file.name, workbook: res.workbook, ws: res.workbook.worksheets[0], info: null, infoKey: '' });
        }).catch(function (err) {
          errors.push(file.name + '：' + (err && err.message ? err.message : String(err)));
        });
      });
    }, Promise.resolve()).then(function () {
      if (errors.length) {
        els.fileMsg.textContent = '有檔案讀不到。' + errors.join('；');
        els.fileMsg.hidden = false;
      }
      els.exportNote.textContent = '';
      renderReports();
    });
  }

  function removeReport(i) {
    S.reports.splice(i, 1);
    settings.reports.splice(i, 1);    // 條件跟著報表走
    els.exportNote.textContent = '';
    renderReports();
  }
  function moveUp(i) {
    if (i < 1) return;
    var r = S.reports[i]; S.reports[i] = S.reports[i - 1]; S.reports[i - 1] = r;
    slot(i);
    var s = settings.reports[i]; settings.reports[i] = settings.reports[i - 1]; settings.reports[i - 1] = s;
    els.exportNote.textContent = '';
    renderReports();
  }

  /* ---------- 各報表的卡片（檔案有增減、換順序時整個重畫） ---------- */
  function renderReports() {
    if (S.reports.length) {
      els.drop.setFile('已上傳 ' + S.reports.length + ' 份報表', '點這裡再加入報表');
    } else {
      els.drop.setIdle();
    }

    els.fileList.textContent = '';
    S.reports.forEach(function (R, i) {
      els.fileList.appendChild(h('li', { class: 'report-item' },
        h('span', { class: 'report-tag', text: '報表 ' + (i + 1) }),
        h('span', { class: 'report-name', text: R.fileName }),
        i > 0 ? h('button', { type: 'button', class: 'btn-link', text: '往上移', onclick: function () { moveUp(i); } }) : null,
        h('button', { type: 'button', class: 'btn-link', text: '移除', onclick: function () { removeReport(i); } })));
    });

    els.cards.textContent = '';
    S.reports.forEach(function (R, i) {
      els.cards.appendChild(buildCard(R, i));
    });
    if (!S.reports.length) els.cards.appendChild(h('p', { class: 'hint', text: '上傳報表之後，這裡會出現每份報表的設定。' }));

    recompute();
  }

  function buildCard(R, i) {
    var st = slot(i);
    var E = R.els = {};
    var pre = 'f-r' + R.id + '-';

    E.sheet = h('select', { id: pre + 'sheet', 'aria-label': '工作表' });
    R.workbook.worksheets.forEach(function (ws, k) {
      E.sheet.appendChild(h('option', { value: String(k), text: ws.name }));
    });
    E.sheet.value = String(R.workbook.worksheets.indexOf(R.ws));
    E.sheet.addEventListener('change', function () {
      R.ws = R.workbook.worksheets[parseInt(E.sheet.value, 10)];
      R.infoKey = '';
      els.exportNote.textContent = '';
      recompute();
    });

    E.selLabel = h('span', { text: 'A' });
    E.values = h('input', {
      id: pre + 'values', type: 'text', class: 'in in-short', autocomplete: 'off',
      inputmode: 'decimal', placeholder: '例如 1', value: st.values
    });
    E.chips = h('div', { class: 'chips' });
    E.cols = h('input', {
      id: pre + 'cols', type: 'text', class: 'in in-wide', autocomplete: 'off', spellcheck: 'false',
      placeholder: '例如 B, C, U, Z, AE', value: st.cols
    });
    E.sameBtn = i > 0
      ? h('button', { type: 'button', class: 'btn-link', text: '和報表 1 保留一樣的行', onclick: function () { copyCols(i); } })
      : null;
    E.msg = h('p', { class: 'summary', 'aria-live': 'polite' });
    E.preview = h('div', { class: 'sheet-area' });

    E.values.addEventListener('input', function () { slot(i).values = E.values.value; schedule(); });
    E.cols.addEventListener('input', function () { slot(i).cols = E.cols.value; schedule(); });

    return h('section', { class: 'report-card', 'aria-labelledby': pre + 'title' },
      h('h3', { id: pre + 'title', class: 'report-title' },
        h('span', { class: 'report-tag', text: '報表 ' + (i + 1) }),
        h('span', { class: 'report-name', text: R.fileName })),
      h('div', { class: 'field-row', hidden: R.workbook.worksheets.length < 2 },
        h('label', { for: pre + 'sheet', text: '工作表' }), E.sheet),
      h('div', { class: 'field-row' },
        h('label', { for: pre + 'values' }, E.selLabel, ' 行等於'), E.values),
      h('p', { class: 'hint', text: '要選多個數值時用逗號分開，例如 1, 3。' }),
      E.chips,
      h('div', { class: 'field-row' },
        h('label', { for: pre + 'cols', text: '要保留的行' }), E.cols),
      h('p', { class: 'hint', text: '連續的行可以寫成 U:Z。也可以直接點下面表格最上方的行號來選。' }),
      E.sameBtn ? h('div', { class: 'link-row' }, E.sameBtn) : null,
      E.msg, E.preview);
  }

  function copyCols(i) {
    var R = S.reports[i];
    slot(i).cols = slot(0).cols;
    if (R.els) R.els.cols.value = slot(i).cols;
    els.exportNote.textContent = '';
    recompute();
  }

  /* ---------- 重新計算並更新畫面 ---------- */
  function recompute() {
    saveSettings();
    S.result = null;
    IEApp.state.mpsFiltered = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;
    els.parts.textContent = '';
    els.outPreview.textContent = '';

    var sr = startRow();
    var sc = selCol();
    // 每份報表各自檢查
    var ready = S.reports.map(function (R, i) { return checkReport(R, i, sr, sc); });

    renderStatus(ready);

    if (!S.reports.length) { setSummary('先上傳 MPS 報表。'); return; }
    var badIdx = [];
    ready.forEach(function (r, i) { if (!r.ok) badIdx.push(i + 1); });
    if (badIdx.length) {
      setSummary('報表 ' + badIdx.join('、') + ' 的設定還沒好，請看上面的說明。', true);
      return;
    }

    // 保留的行數要一樣
    var n0 = ready[0].cols.length;
    var diff = [];
    ready.forEach(function (r, i) { if (r.cols.length !== n0) diff.push('報表 ' + (i + 1) + ' 保留 ' + r.cols.length + ' 行'); });
    if (diff.length) {
      setSummary('每份報表保留的行數要一樣。報表 1 保留 ' + n0 + ' 行，' + diff.join('、') + '。', true);
      return;
    }

    var sp = statusPos();
    if (!sp || sp > n0) { setSummary('選擇 Status 在哪一行。'); return; }
    ui.statusCol.set(XIO.numToCol(sp));       // 第 2 步、第 6 步共用

    var parts = S.reports.map(function (R, i) {
      return { ws: R.ws, values: ready[i].values, keepCols: ready[i].cols, lastRow: R.info.lastRow };
    });
    var common = { dataStartRow: sr, selectorCol: sc };
    var result = XIO.filterMany(parts, common);
    S.result = result;

    var text = '合併後共 ' + result.matchedRows.toLocaleString('en-US') + ' 列資料、' + n0 + ' 行；Status 在 ' +
      XIO.numToCol(sp) + ' 行。';
    setSummary(text);
    result.parts.forEach(function (p, i) {
      els.parts.appendChild(h('li', {
        text: '報表 ' + (i + 1) + '（' + S.reports[i].fileName + '）：' + p.matched.toLocaleString('en-US') + ' 列' +
          (p.matched ? '，在新檔第 ' + p.firstRow + '–' + p.lastRow + ' 列' : '')
      }));
    });

    var mark = {}; mark[sp] = true;
    ui.renderSheet(els.outPreview, result.worksheet, {
      headerEnd: result.headerEnd,
      lastRow: result.lastRow,
      colCount: result.colCount,
      rows: PREVIEW_ROWS,
      keep: mark,
      caption: '匯出後的樣子（前 ' + PREVIEW_ROWS + ' 列資料，Status 那一行用螢光筆標示）。'
    });
    els.exportBtn.disabled = false;
    els.nextBtn.disabled = false;

    // 給第 2 步用。build() 每次都產生一份全新的活頁簿，後面的步驟可以放心往裡面加工作表。
    IEApp.state.mpsFiltered = {
      sourceName: S.reports.map(function (R) { return R.fileName; }).join('、'),
      outName: outputName(),
      sheetName: result.worksheet.name,
      matchedRows: result.matchedRows,
      headerEnd: result.headerEnd,
      statusCol: sp,
      build: function () { return XIO.filterMany(parts, common).workbook; }
    };
  }

  // 檢查一份報表、更新它的卡片；回傳 { ok, values, cols }
  function checkReport(R, i, sr, sc) {
    var E = R.els;
    var key = sr + ':' + sc + ':' + (R.ws ? R.ws.name : '');
    if (R.infoKey !== key) {
      R.info = XIO.inspect(R.ws, { dataStartRow: sr, selectorCol: sc });
      R.infoKey = key;
    }
    var info = R.info;
    var colName = XIO.numToCol(sc);
    var values = wantedValues(i);
    var wanted = {};
    values.forEach(function (v) { wanted[XIO.matchKey(v)] = true; });
    var parsed = keptCols(i);
    var keep = {};
    parsed.cols.forEach(function (c) { keep[c] = true; });
    var out = { ok: false, values: values, cols: parsed.cols };

    E.selLabel.textContent = colName;
    renderChips(R, i, info, wanted);
    ui.renderSheet(E.preview, R.ws, {
      headerEnd: sr - 1,
      lastRow: info.lastRow,
      colCount: Math.max(info.colCount, 1),
      rows: PREVIEW_ROWS,
      keep: keep,
      onToggle: function (c) { toggleColumn(i, c); },
      selectorCol: sc,
      wanted: values.length ? wanted : null,
      caption: '報表 ' + (i + 1) + ' 原始內容（前 ' + PREVIEW_ROWS + ' 列資料）。共 ' + info.dataRows.toLocaleString('en-US') +
        ' 列資料、' + info.colCount + ' 行（到 ' + XIO.numToCol(info.colCount || 1) + ' 行）。'
    });

    function msg(text, isError, isWarn) {
      E.msg.textContent = text;
      E.msg.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
    }

    if (!info.dataRows) { msg('第 ' + sr + ' 列之後沒有資料。請確認上傳的檔案和工作表。', true); return out; }
    if (!values.length) { msg('輸入 ' + colName + ' 行要等於的數值。'); return out; }
    if (parsed.bad.length) {
      msg('看不懂「' + parsed.bad.join('、') + '」。行號請用英文字母，例如 B 或 U:Z。', true);
      return out;
    }
    if (!parsed.cols.length) { msg('輸入要保留的行。'); return out; }

    var matched = 0;
    values.forEach(function (v) {
      var k = XIO.matchKey(v);
      info.values.forEach(function (x) { if (x.key === k) matched += x.count; });
    });
    if (!matched) {
      var seen = info.values.slice(0, 8).map(function (v) { return v.label; }).join('、');
      msg(colName + ' 行沒有等於 ' + values.join('、') + ' 的列。' +
        (seen ? '這份報表的 ' + colName + ' 行有：' + seen + '。' : colName + ' 行從第 ' + sr + ' 列起都是空白。'), true);
      return out;
    }

    var text = '符合 ' + matched.toLocaleString('en-US') + ' 列（共 ' + info.dataRows.toLocaleString('en-US') +
      ' 列資料），保留 ' + parsed.cols.length + ' 行。';
    var beyond = parsed.cols.filter(function (c) { return c > info.colCount; });
    if (beyond.length) text += ' 注意：' + XIO.formatColumns(beyond) + ' 行在報表裡沒有資料。';
    msg(text, false, beyond.length > 0);
    out.ok = true;
    return out;
  }

  /* ---------- Status 在哪一行 ---------- */
  function labelsOf(R, cols) {
    var headerEnd = startRow() - 1;
    if (headerEnd < 1 || !cols.length) return [];
    var all = XIO.columnLabels(R.ws, 1, headerEnd, Math.max.apply(null, cols));
    return cols.map(function (c) { return all[c] || ''; });
  }

  function renderStatus(ready) {
    els.status.textContent = '';
    els.statusChecks.textContent = '';
    var first = ready[0];
    if (!first || !first.cols.length) {
      els.status.appendChild(h('option', { value: '', text: '先設定報表 1 要保留的行' }));
      els.status.disabled = true;
      return;
    }
    els.status.disabled = false;
    var cols = first.cols;
    var labels = labelsOf(S.reports[0], cols);

    // 還沒選過的話，猜表頭有「Status」的那一行
    if (!statusPos()) {
      for (var g = 0; g < labels.length; g++) {
        if (STATUS_GUESS.test(labels[g])) { settings.statusPos = String(g + 1); saveSettings(); break; }
      }
    }
    var sp = statusPos();

    els.status.appendChild(h('option', { value: '', text: '請選擇' }));
    cols.forEach(function (c, k) {
      var label = labels[k] ? '：' + shorten(labels[k]) : '';
      els.status.appendChild(h('option', {
        value: String(k + 1),
        text: '新檔 ' + XIO.numToCol(k + 1) + ' 行（報表 1 的 ' + XIO.numToCol(c) + ' 行' + label + '）'
      }));
    });
    els.status.value = sp && sp <= cols.length ? String(sp) : '';
    if (!sp || sp > cols.length) return;

    // 其他報表在同一個相對位置的行，表頭寫什麼——讓使用者確認有沒有對到
    var base = labels[sp - 1] || '';
    ready.forEach(function (r, i) {
      if (!r.cols.length) return;
      var c = r.cols[sp - 1];
      if (!c) {
        els.statusChecks.appendChild(h('li', { class: 'is-warn', text: '報表 ' + (i + 1) + '：保留的行不到 ' + sp + ' 行，對不到 Status。' }));
        return;
      }
      var lab = labelsOf(S.reports[i], [c])[0] || '';
      var same = !base || labelKey(lab) === labelKey(base) || (STATUS_GUESS.test(base) && STATUS_GUESS.test(lab));
      els.statusChecks.appendChild(h('li', {
        class: same ? null : 'is-warn',
        text: '報表 ' + (i + 1) + '：' + XIO.numToCol(c) + ' 行' + (lab ? '（' + shorten(lab) + '）' : '（表頭空白）') +
          (same ? '' : '——表頭和報表 1 的不一樣，請確認保留的行順序有沒有對到。')
      }));
    });
  }

  function labelKey(t) { return XIO.normalizeText(t || '').replace(/\s+/g, '').toLowerCase(); }
  function shorten(t) { return t.length > 24 ? t.slice(0, 24) + '…' : t; }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  /* ---------- 選取行出現過的值（點一下就填入） ---------- */
  function renderChips(R, i, info, wanted) {
    var box = R.els.chips;
    box.textContent = '';
    if (!info.values.length && !info.blankCount) return;
    var col = XIO.numToCol(selCol());
    box.appendChild(h('span', { class: 'chips-label', text: '報表裡 ' + col + ' 行的值' }));
    info.values.slice(0, MAX_CHIPS).forEach(function (v) {
      var on = !!wanted[v.key];
      box.appendChild(h('button', {
        type: 'button', class: 'chip', 'aria-pressed': on ? 'true' : 'false',
        onclick: function () { toggleValue(i, v); }
      }, h('span', { class: 'chip-value', text: v.label }),
         h('span', { class: 'chip-count', text: v.count.toLocaleString('en-US') + ' 列' })));
    });
    if (info.values.length > MAX_CHIPS) {
      box.appendChild(h('span', { class: 'chips-more', text: '另有 ' + (info.values.length - MAX_CHIPS) + ' 種' }));
    }
    if (info.blankCount) {
      box.appendChild(h('span', { class: 'chips-more', text: '空白 ' + info.blankCount.toLocaleString('en-US') + ' 列' }));
    }
  }

  function toggleValue(i, v) {
    var list = wantedValues(i);
    var kept = list.filter(function (x) { return XIO.matchKey(x) !== v.key; });
    if (kept.length === list.length) kept.push(v.label);
    slot(i).values = kept.join(', ');
    S.reports[i].els.values.value = slot(i).values;
    els.exportNote.textContent = '';
    recompute();
  }

  function toggleColumn(i, c) {
    var E = S.reports[i].els;
    var cols = keptCols(i).cols;
    var k = cols.indexOf(c);
    if (k >= 0) cols.splice(k, 1); else cols.push(c);
    cols.sort(function (a, b) { return a - b; });
    slot(i).cols = XIO.formatColumns(cols);
    E.cols.value = slot(i).cols;
    var scroller = E.preview.querySelector('.sheet-scroll');
    var left = scroller ? scroller.scrollLeft : 0;
    els.exportNote.textContent = '';
    recompute();
    scroller = E.preview.querySelector('.sheet-scroll');
    if (scroller) scroller.scrollLeft = left;
    var btn = E.preview.querySelector('button[data-col="' + c + '"]');
    if (btn) btn.focus({ preventScroll: true });
  }

  /* ---------- 匯出 ---------- */
  function onExport() {
    var ready = IEApp.state.mpsFiltered;
    if (!S.result || !ready) return;
    var rows = ready.matchedRows;
    var name = ready.outName;
    els.exportBtn.disabled = true;
    els.exportNote.textContent = '產生檔案中…';
    XIO.toBlob(ready.build()).then(function (blob) {
      ui.download(blob, name);
      els.exportNote.textContent = '已匯出 ' + name + '（' + rows.toLocaleString('en-US') + ' 列），在瀏覽器的下載資料夾。';
    }).catch(function (err) {
      els.exportNote.textContent = '匯出失敗：' + (err && err.message ? err.message : err);
    }).then(function () {
      els.exportBtn.disabled = false;
    });
  }

  IEApp.registerStep({ id: 'mps-filter', render: render });
})(window);
