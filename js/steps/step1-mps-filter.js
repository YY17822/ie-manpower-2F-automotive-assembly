/*
 * 第 1 步：篩選 MPS 報表（好幾份報表合成一份）
 *   1. 上傳 MPS 的 Excel，一次可以好幾份（通常 3、4 份），每份報表版面一樣：
 *      第 1–4 列表頭、第 5 列開始資料、A 行是選取欄
 *   2. 每份報表各自選篩選方式，並輸入「要保留哪些行」：
 *        a. A 行等於多少（原本的方式）
 *        b. 指定某一行，等於或不等於哪些值；這一行空白的列一律不要
 *        c. 指定某一行，模糊比對（含有某些字）：含有的列選取或排除；這一行空白的列一律不要
 *      兩種方式各報表可以混用
 *      ——各報表保留的行位置可以不同，但保留的行數和順序（相對位置）要一樣
 *   3. 指定 Status 在哪一行（區分量產、試產，後面總表會用到）
 *   4. 每份報表指定類別：取像、車用、光機系統（一定要選，不會自動猜）
 *   5. 匯出一份新的 Excel：表頭用報表 1 的（含合併儲存格），底下依類別分段（XIO.filterSections）：
 *        【取像】 取像的 MPS → 空 3 列 → 【車用】 車用的 MPS → 空 1 列、3 列給第 2 步的匯總表（取像＋車用）
 *        → 空 6 列 → 【光機系統】 光機系統的 MPS → 空 1 列給第 2 步的匯總表（光機系統）
 *      類別標題跨所有保留的行合併、粗體置中、底色 #E9C46A。同一類有好幾份時照上傳順序往下接，中間不畫線；
 *      同一個料號出現在不同報表就保留成好幾列。這個月沒有檔案的類別整段不放。
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

  var STORE_KEY = 'step1';                    // 共用設定的組名（存進 Supabase 的 key 是 step1.xxx）
  var PREVIEW_ROWS = 8;     // 預覽顯示幾列資料
  var OUT_PREVIEW_ROWS = 60; // 匯出後的樣子顯示幾列（要看得到各段）
  var MAX_CHIPS = 24;
  var MAX_FZ_LIST = 300;    // 模糊比對的預覽最多列出幾列
  var STATUS_GUESS = /status|狀態/i;

  // reports：每份報表各自的條件，依上傳順序（報表 1、報表 2…），下個月照同樣順序上傳就會自動帶入
  //   cat：類別（'image' 取像、'auto' 車用、'optics' 光機系統，見 XIO.MPS_SECTIONS）
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
    try { s = IEApp.settings.read(STORE_KEY) || {}; } catch (e) { s = {}; }
    var out = {};
    Object.keys(DEFAULTS).forEach(function (k) { out[k] = s[k] === undefined ? DEFAULTS[k] : s[k]; });
    if (!Array.isArray(out.reports)) out.reports = [];
    // 只有一份報表的舊版設定（values、cols）當作報表 1
    if (!out.reports.length && (s.values || s.cols)) out.reports.push({ values: s.values || '', cols: s.cols || '' });
    return out;
  }
  function saveSettings() {
    // Status 在新檔的哪一行（英文字母）存在 step1-2-6，和第 2、6 步共用（ui.statusCol）
    var keep = { startRow: settings.startRow, selCol: settings.selCol, reports: settings.reports, statusPos: settings.statusPos };
    try { IEApp.settings.write(STORE_KEY, keep); } catch (e) { /* 存不了就算了 */ }
  }
  function catOf(i) {
    var c = slot(i).cat;
    return XIO.MPS_SECTIONS.some(function (s) { return s.key === c; }) ? c : '';
  }
  function catLabel(key) {
    var s = XIO.MPS_SECTIONS.filter(function (x) { return x.key === key; })[0];
    return s ? s.label : '';
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
  // 篩選方式：'a' = A 行等於（原本的方式，用 values），'col' = 指定行等於／不等於（用 fcol、op、fvalues），
  // 'fz' = 指定行模糊比對（用 zCol、zOp：'eq' 含有的列選取／'ne' 含有的列排除、zValues：各輸入框的值）
  // 舊設定沒有 mode，當作 'a'
  function modeOf(i) { var m = slot(i).mode; return m === 'col' || m === 'fz' ? m : 'a'; }
  function opOf(i) { return modeOf(i) === 'col' && slot(i).op === 'neq' ? 'neq' : 'eq'; }
  function zExclude(i) { return modeOf(i) === 'fz' && slot(i).zOp === 'ne'; }
  function oneCol(text) {
    var p = XIO.parseColumns(text || '');
    return p.cols.length === 1 && !p.bad.length ? p.cols[0] : 0;
  }
  function filterCol(i) { return oneCol(slot(i).fcol); }
  function zColOf(i) { return oneCol(slot(i).zCol); }
  // 模糊比對各輸入框的值。沒設定過是 3 個空的；讀的時候不寫回設定，免得打開畫面就被當成修改
  function zVals(i) {
    var z = slot(i).zValues;
    return Array.isArray(z) && z.length ? z.map(function (v) { return String(v == null ? '' : v); }) : ['', '', ''];
  }
  // 這份報表用哪一行篩選（「指定行」「模糊比對」還沒填好時是 0）
  function colFor(i) {
    var m = modeOf(i);
    return m === 'col' ? filterCol(i) : m === 'fz' ? zColOf(i) : selCol();
  }
  function valueText(i) { return (modeOf(i) === 'col' ? slot(i).fvalues : slot(i).values) || ''; }
  function setValueText(i, t) { if (modeOf(i) === 'col') slot(i).fvalues = t; else slot(i).values = t; }
  function wantedValues(i) {
    if (modeOf(i) === 'fz') {
      return zVals(i).map(function (v) { return XIO.normalizeText(v); }).filter(function (v) { return v !== ''; });
    }
    return XIO.normalizeText(valueText(i)).split(/[,、;]+/)
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
      h('p', { text: '上傳好幾份 MPS 報表，指定各自是取像、車用還是光機系統，再依選取條件留下要的列和要的行，依類別分段排成一份新的 Excel。表頭（第 1–4 列）用報表 1 的，原樣保留。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '上傳 MPS 報表' }),
        els.drop.el, els.fileList, els.fileMsg,
        h('p', { class: 'hint', text: '依上傳順序叫報表 1、報表 2…。每份報表都要選類別（取像、車用、光機系統）；新檔裡依 取像 → 車用 → 光機系統 分段，同一類有好幾份時照上傳順序往下接。每份報表的類別和條件會記住，下個月照同樣順序上傳就會自動帶入。' }),
        h('details', { class: 'advanced' },
          h('summary', { text: '報表版面不一樣時' }),
          h('div', { class: 'field-row' },
            h('label', { for: 'f-start', text: '資料從第' }), els.startRow, h('span', { text: '列開始（之前的列都當作表頭保留）' })),
          h('div', { class: 'field-row' },
            h('label', { for: 'f-selcol', text: '選取條件在' }), els.selCol, h('span', { text: '行' })),
          h('p', { class: 'hint', text: '每份報表都用同一組版面設定。「選取條件在哪一行」只用在「A 行等於」這種篩選方式。' }))),
      h('li', { class: 'task' },
        h('h2', { text: '各報表的選取條件和要保留的行' }),
        h('p', { class: 'hint', text: '篩選方式每份報表各自選，可以混用：「A 行等於」是原本的方式；「指定行」可以選任何一行，留下等於或不等於某些值的列，那一行空白的列一律不要。' }),
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
      var cat = h('select', { id: 'f-r' + R.id + '-cat', class: 'in-select cat-select', 'aria-label': '報表 ' + (i + 1) + ' 的類別' });
      cat.appendChild(h('option', { value: '', text: '選擇類別' }));
      XIO.MPS_SECTIONS.forEach(function (s) { cat.appendChild(h('option', { value: s.key, text: s.label })); });
      cat.value = catOf(i);
      if (!cat.value) cat.classList.add('is-empty');
      cat.addEventListener('change', function () {
        slot(i).cat = cat.value;
        els.exportNote.textContent = '';
        renderReports();
      });
      els.fileList.appendChild(h('li', { class: 'report-item' },
        h('span', { class: 'report-tag', text: '報表 ' + (i + 1) }),
        cat,
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

    // 篩選方式：A 行等於（原本的）／指定行 等於、不等於
    var mode = modeOf(i);
    E.modeA = h('input', { id: pre + 'mode-a', type: 'radio', name: pre + 'mode', value: 'a' });
    E.modeC = h('input', { id: pre + 'mode-c', type: 'radio', name: pre + 'mode', value: 'col' });
    E.modeA.checked = mode === 'a';
    E.modeC.checked = mode === 'col';
    E.modeALabel = h('span', { text: 'A' });
    E.fcol = h('input', {
      id: pre + 'fcol', type: 'text', class: 'in in-tiny', autocomplete: 'off', spellcheck: 'false',
      placeholder: '例如 D', value: st.fcol || ''
    });
    E.opEq = h('input', { id: pre + 'op-eq', type: 'radio', name: pre + 'op', value: 'eq' });
    E.opNe = h('input', { id: pre + 'op-ne', type: 'radio', name: pre + 'op', value: 'neq' });
    E.opEq.checked = st.op !== 'neq';
    E.opNe.checked = st.op === 'neq';
    E.fvalues = h('input', {
      id: pre + 'fvalues', type: 'text', class: 'in in-short', autocomplete: 'off',
      placeholder: '例如 M, R', value: st.fvalues || ''
    });
    E.boxA = h('div', { class: 'filter-box', hidden: mode !== 'a' },
      h('div', { class: 'field-row' },
        h('label', { for: pre + 'values' }, E.selLabel, ' 行等於'), E.values),
      h('p', { class: 'hint', text: '要選多個數值時用逗號分開，例如 1, 3。' }));
    E.boxC = h('div', { class: 'filter-box', hidden: mode !== 'col' },
      h('div', { class: 'field-row' },
        h('label', { for: pre + 'fcol', text: '篩選的行' }), E.fcol),
      h('div', { class: 'field-row radio-row', role: 'radiogroup', 'aria-label': '等於或不等於' },
        h('label', { for: pre + 'op-eq' }, E.opEq, '等於'),
        h('label', { for: pre + 'op-ne' }, E.opNe, '不等於'),
        E.fvalues),
      h('p', { class: 'hint', text: '多個值用逗號分開，也可以點下面出現過的值。這一行空白的列一律不要；選「不等於」又沒填值，就是只拿掉空白的列。' }));

    function setMode(m) {
      slot(i).mode = m;
      E.boxA.hidden = m !== 'a';
      E.boxC.hidden = m !== 'col';
      E.boxZ.hidden = m !== 'fz';
      R.infoKey = '';
      els.exportNote.textContent = '';
      recompute();
    }
    function setOp(op) {
      slot(i).op = op;
      els.exportNote.textContent = '';
      recompute();
    }
    E.modeA.addEventListener('change', function () { if (E.modeA.checked) setMode('a'); });
    E.modeC.addEventListener('change', function () { if (E.modeC.checked) setMode('col'); });
    E.opEq.addEventListener('change', function () { if (E.opEq.checked) setOp('eq'); });
    E.opNe.addEventListener('change', function () { if (E.opNe.checked) setOp('neq'); });
    E.fcol.addEventListener('input', function () { slot(i).fcol = E.fcol.value; R.infoKey = ''; schedule(); });
    E.fvalues.addEventListener('input', function () { slot(i).fvalues = E.fvalues.value; schedule(); });

    // 「模糊比對」方式的欄位：哪一行、含有的列選取或排除、值的輸入框（預設 3 個，可以加、可以刪）
    E.modeZ = h('input', { id: pre + 'mode-z', type: 'radio', name: pre + 'mode', value: 'fz' });
    E.modeZ.checked = mode === 'fz';
    E.zCol = h('input', {
      id: pre + 'zcol', type: 'text', class: 'in in-tiny', autocomplete: 'off', spellcheck: 'false',
      placeholder: '例如 C', value: st.zCol || '', 'aria-label': '要模糊比對的行'
    });
    E.zOpSel = h('input', { id: pre + 'zop-eq', type: 'radio', name: pre + 'zop', value: 'eq' });
    E.zOpExc = h('input', { id: pre + 'zop-ne', type: 'radio', name: pre + 'zop', value: 'ne' });
    E.zOpSel.checked = st.zOp !== 'ne';
    E.zOpExc.checked = st.zOp === 'ne';
    E.zBox = h('div', { class: 'fz-values' });
    E.zAdd = h('button', { type: 'button', class: 'link-btn fz-add', text: '＋ 再加一個輸入框', onclick: function () { addZInput(i); } });
    E.zPreview = h('div', { class: 'fz-preview', 'aria-live': 'polite' });
    E.boxZ = h('div', { class: 'filter-box', hidden: mode !== 'fz' },
      h('div', { class: 'field-row' },
        E.zCol, h('span', { text: '行含有下面任一個值的列：' }),
        h('span', { class: 'fz-op', role: 'radiogroup', 'aria-label': '含有的列要選取還是排除' },
          h('label', { class: 'check-label', for: pre + 'zop-eq' }, E.zOpSel, h('span', { text: '選取' })),
          h('label', { class: 'check-label', for: pre + 'zop-ne' }, E.zOpExc, h('span', { text: '排除' })))),
      E.zBox, E.zAdd,
      h('p', { class: 'hint', text: '儲存格的字含有任一個值就算符合，英文大小寫、全形半形不分。這一行空白的列一律不留。' }),
      E.zPreview);
    buildZInputs(i);

    E.modeZ.addEventListener('change', function () { if (E.modeZ.checked) setMode('fz'); });
    E.zCol.addEventListener('input', function () { slot(i).zCol = E.zCol.value; R.infoKey = ''; schedule(); });
    [E.zOpSel, E.zOpExc].forEach(function (r) {
      r.addEventListener('change', function () {
        if (!r.checked) return;
        slot(i).zOp = r.value;
        els.exportNote.textContent = '';
        recompute();
      });
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
        catOf(i) ? h('span', { class: 'cat-tag', text: catLabel(catOf(i)) }) : h('span', { class: 'cat-tag is-empty', text: '還沒選類別' }),
        h('span', { class: 'report-name', text: R.fileName })),
      h('div', { class: 'field-row', hidden: R.workbook.worksheets.length < 2 },
        h('label', { for: pre + 'sheet', text: '工作表' }), E.sheet),
      h('div', { class: 'field-row radio-row', role: 'radiogroup', 'aria-label': '篩選方式' },
        h('span', { class: 'field-name', text: '篩選方式' }),
        h('label', { for: pre + 'mode-a' }, E.modeA, E.modeALabel, ' 行等於'),
        h('label', { for: pre + 'mode-c' }, E.modeC, '指定行 等於／不等於'),
        h('label', { for: pre + 'mode-z' }, E.modeZ, '指定某一行，模糊比對（含有某些字）')),
      E.boxA, E.boxC, E.boxZ,
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

    var noCat = [];
    S.reports.forEach(function (R, i) { if (!catOf(i)) noCat.push(i + 1); });
    if (noCat.length) {
      setSummary('報表 ' + noCat.join('、') + ' 還沒選類別（取像、車用、光機系統）。在上面「上傳 MPS 報表」的清單裡選。', true);
      return;
    }

    var parts = S.reports.map(function (R, i) {
      return {
        ws: R.ws, cat: catOf(i), values: ready[i].values, op: ready[i].op, selectorCol: ready[i].col,
        fuzzy: ready[i].fuzzy, exclude: ready[i].exclude,
        keepCols: ready[i].cols, lastRow: R.info.lastRow
      };
    });
    var common = { dataStartRow: sr, selectorCol: sc };
    var result = XIO.filterSections(parts, common);
    S.result = result;

    var text = '合併後共 ' + result.matchedRows.toLocaleString('en-US') + ' 列資料、' + n0 + ' 行；Status 在 ' +
      XIO.numToCol(sp) + ' 行。';
    setSummary(text);
    result.sections.forEach(function (s) {
      var who = [];
      result.parts.forEach(function (p, i) { if (p.cat === s.key) who.push('報表 ' + (i + 1)); });
      els.parts.appendChild(h('li', {
        text: '【' + s.label + '】標題在第 ' + s.titleRow + ' 列，' + who.join('、') + ' 共 ' + s.matched.toLocaleString('en-US') + ' 列' +
          (s.matched ? '，在第 ' + s.firstRow + '–' + s.lastRow + ' 列' : '')
      }));
    });
    var lastSec = result.sections[result.sections.length - 1];
    var groupsShown = {};
    result.sections.forEach(function (s, k) {
      var nextSec = result.sections[k + 1];
      if (nextSec && nextSec.group === s.group) return;
      if (groupsShown[s.group]) return;
      groupsShown[s.group] = true;
      els.parts.appendChild(h('li', { class: 'hint-item',
        text: XIO.MPS_SECTION_GROUPS[s.group].label + '的匯總表（第 2 步做）會放在第 ' + (s.lastRow + 2) + '–' + (s.lastRow + 4) + ' 列' +
          (s === lastSec ? '。' : '，下面空 6 列才是下一段。') }));
    });

    var mark = {}; mark[sp] = true;
    ui.renderSheet(els.outPreview, result.worksheet, {
      headerEnd: result.headerEnd,
      lastRow: result.lastRow,
      colCount: result.colCount,
      rows: Math.min(result.lastRow - result.headerEnd, OUT_PREVIEW_ROWS),
      keep: mark,
      caption: '匯出後的樣子（前 ' + Math.min(result.lastRow - result.headerEnd, OUT_PREVIEW_ROWS) + ' 列，Status 那一行用螢光筆標示；匯總表第 2 步才會放進去）。'
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
      sections: result.sections,
      build: function () { return XIO.filterSections(parts, common).workbook; }
    };
  }

  // 檢查一份報表、更新它的卡片；回傳 { ok, values, cols }
  function checkReport(R, i, sr, sc) {
    var E = R.els;
    var mode = modeOf(i);
    var op = opOf(i);
    var fc = colFor(i);             // 這份報表用哪一行篩選；「指定行」還沒填好時是 0
    var key = sr + ':' + (fc || 1) + ':' + (R.ws ? R.ws.name : '');
    if (R.infoKey !== key) {
      R.info = XIO.inspect(R.ws, { dataStartRow: sr, selectorCol: fc || 1 });
      R.infoKey = key;
    }
    var info = R.info;
    var colName = XIO.numToCol(fc || 1);
    var values = wantedValues(i);
    var wanted = {};
    values.forEach(function (v) { wanted[XIO.matchKey(v)] = true; });
    var parsed = keptCols(i);
    var keep = {};
    parsed.cols.forEach(function (c) { keep[c] = true; });
    var exc = zExclude(i);
    var terms = mode === 'fz' ? XIO.fuzzyTerms(values) : null;
    var out = { ok: false, values: values, cols: parsed.cols, col: fc, op: op, fuzzy: mode === 'fz', exclude: exc };
    E.zPreview.textContent = '';

    E.selLabel.textContent = XIO.numToCol(sc);
    E.modeALabel.textContent = XIO.numToCol(sc);
    if (fc) renderChips(R, i, info, wanted); else E.chips.textContent = '';
    ui.renderSheet(E.preview, R.ws, {
      headerEnd: sr - 1,
      lastRow: info.lastRow,
      colCount: Math.max(info.colCount, 1),
      rows: PREVIEW_ROWS,
      keep: keep,
      onToggle: function (c) { toggleColumn(i, c); },
      selectorCol: fc || 1,
      // 「不等於」時就算沒填值也要把空白的列變淡
      wanted: !fc || mode === 'fz' ? null : (op === 'neq' || values.length ? wanted : null),
      wantedOp: op,
      fuzzyTerms: fc && terms && terms.length ? terms : null,
      exclude: exc,
      caption: '報表 ' + (i + 1) + ' 原始內容（前 ' + PREVIEW_ROWS + ' 列資料）。共 ' + info.dataRows.toLocaleString('en-US') +
        ' 列資料、' + info.colCount + ' 行（到 ' + XIO.numToCol(info.colCount || 1) + ' 行）。'
    });

    function msg(text, isError, isWarn) {
      E.msg.textContent = text;
      E.msg.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
    }

    if (!info.dataRows) { msg('第 ' + sr + ' 列之後沒有資料。請確認上傳的檔案和工作表。', true); return out; }
    if ((mode === 'col' || mode === 'fz') && !fc) {
      var fp = XIO.parseColumns((mode === 'fz' ? slot(i).zCol : slot(i).fcol) || '');
      var badCol = fp.bad.length > 0 || fp.cols.length > 1;
      var what = mode === 'fz' ? '模糊比對' : '篩選';
      msg(badCol ? what + '的行只能填一個行號（英文字母），例如 D。' : '輸入要' + what + '的行，例如 D。', badCol);
      return out;
    }
    if (mode === 'fz' && !values.length) {
      msg('在輸入框填要比對的字（' + colName + ' 行含有這些字的列' + (exc ? '排除' : '選取') + '）。');
      return out;
    }
    if (!values.length && op === 'eq') { msg('輸入 ' + colName + ' 行要等於的數值。'); return out; }
    if (parsed.bad.length) {
      msg('看不懂「' + parsed.bad.join('、') + '」。行號請用英文字母，例如 B 或 U:Z。', true);
      return out;
    }
    if (!parsed.cols.length) { msg('輸入要保留的行。'); return out; }

    var matched;
    var blankNote = mode === 'col' ? info.blankCount : 0;   // 那一行空白、不留的列數
    if (mode === 'fz') {
      var fz = XIO.fuzzyRows(R.ws, { dataStartRow: sr, selectorCol: fc, lastRow: info.lastRow, values: values, exclude: exc });
      renderFzPreview(E.zPreview, fz, terms, colName, exc);
      matched = fz.rows.length;
      blankNote = fz.blank;
      if (!matched) {
        msg(exc
          ? colName + ' 行每一列都含有 ' + values.join('、') + ' 其中之一（或是空白），一列都不會留下。'
          : colName + ' 行沒有含有 ' + values.join('、') + ' 的列。', true);
        return out;
      }
    } else {
      var hit = 0;                    // 篩選行「等於」輸入值的列數
      info.values.forEach(function (x) { if (wanted[x.key]) hit += x.count; });
      var nonBlank = info.dataRows - info.blankCount;
      matched = op === 'neq' ? nonBlank - hit : hit;
      if (!matched && op === 'neq') {
        msg(nonBlank
          ? colName + ' 行除了 ' + values.join('、') + ' 之外沒有別的值（空白的列不算），一列都不會留下。'
          : colName + ' 行從第 ' + sr + ' 列起都是空白，一列都不會留下。', true);
        return out;
      }
      if (!matched) {
        var seen = info.values.slice(0, 8).map(function (v) { return v.label; }).join('、');
        msg(colName + ' 行沒有等於 ' + values.join('、') + ' 的列。' +
          (seen ? '這份報表的 ' + colName + ' 行有：' + seen + '。' : colName + ' 行從第 ' + sr + ' 列起都是空白。'), true);
        return out;
      }
    }

    var text = '符合 ' + matched.toLocaleString('en-US') + ' 列（共 ' + info.dataRows.toLocaleString('en-US') +
      ' 列資料），保留 ' + parsed.cols.length + ' 行。';
    if (blankNote) {
      text += ' ' + colName + ' 行空白的 ' + blankNote.toLocaleString('en-US') + ' 列不要。';
    }
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
    var col = XIO.numToCol(colFor(i));
    box.appendChild(h('span', { class: 'chips-label',
      text: '報表裡 ' + col + ' 行的值' +
        (modeOf(i) === 'fz' ? '（點一下填進輸入框，再點一下清掉）' : opOf(i) === 'neq' ? '（點選的值會被排除）' : '') }));
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
    if (modeOf(i) === 'fz') {
      // 模糊比對：已經在輸入框裡就清掉，不在就填進第一個空的輸入框（沒有空的就加一個）
      var zs = zVals(i);
      var at = -1, empty = -1;
      zs.forEach(function (x, k) {
        if (at < 0 && XIO.normalizeText(x) !== '' && XIO.matchKey(x) === v.key) at = k;
        if (empty < 0 && XIO.normalizeText(x) === '') empty = k;
      });
      if (at >= 0) zs[at] = '';
      else if (empty >= 0) zs[empty] = v.label;
      else zs.push(v.label);
      slot(i).zValues = zs;
      buildZInputs(i);
      els.exportNote.textContent = '';
      recompute();
      return;
    }
    var list = wantedValues(i);
    var kept = list.filter(function (x) { return XIO.matchKey(x) !== v.key; });
    if (kept.length === list.length) kept.push(v.label);
    setValueText(i, kept.join(', '));
    var E = S.reports[i].els;
    (modeOf(i) === 'col' ? E.fvalues : E.values).value = valueText(i);
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

  /* ---------- 模糊比對的輸入框（預設三個，可以再加、可以刪；至少留一個） ---------- */
  function buildZInputs(i, focusIndex) {
    var R = S.reports[i];
    var E = R.els;
    E.zBox.textContent = '';
    var list = zVals(i);
    var pre = 'f-r' + R.id + '-zval-';
    list.forEach(function (val, k) {
      var id = pre + k;
      var input = h('input', {
        id: id, type: 'text', class: 'in in-short', autocomplete: 'off', spellcheck: 'false',
        placeholder: k === 0 ? '例如 Lens' : '', value: val
      });
      input.addEventListener('input', function () {
        var z = zVals(i);
        z[k] = input.value;
        slot(i).zValues = z;
        schedule();
      });
      var row = h('div', { class: 'fz-value' },
        h('label', { class: 'fz-label', for: id, text: '值 ' + (k + 1) }), input);
      if (list.length > 1) {
        row.appendChild(h('button', {
          type: 'button', class: 'link-btn', text: '刪除', 'aria-label': '刪除值 ' + (k + 1),
          onclick: function () { removeZInput(i, k); }
        }));
      }
      E.zBox.appendChild(row);
    });
    if (focusIndex !== undefined) {
      var el = document.getElementById(pre + focusIndex);
      if (el) el.focus();
    }
  }
  function addZInput(i) {
    var z = zVals(i);
    z.push('');
    slot(i).zValues = z;
    buildZInputs(i, z.length - 1);
    els.exportNote.textContent = '';
    recompute();
  }
  function removeZInput(i, k) {
    var z = zVals(i);
    z.splice(k, 1);
    if (!z.length) z.push('');
    slot(i).zValues = z;
    buildZInputs(i, Math.min(k, z.length - 1));
    els.exportNote.textContent = '';
    recompute();
  }

  /* ---------- 模糊比對的預覽：逐列列出會留下的列 ---------- */
  function renderFzPreview(box, res, terms, colName, exc) {
    box.textContent = '';
    var note = exc
      ? '（' + colName + ' 行含有任一值的 ' + res.hit.toLocaleString('en-US') + ' 列排除；'
      : '（';
    note += colName + ' 行空白的 ' + res.blank.toLocaleString('en-US') + ' 列不留）';
    box.appendChild(h('div', { class: 'blank-head' },
      h('span', { class: 'fz-count', text: '篩選出來的列：' + res.rows.length.toLocaleString('en-US') + ' 列' }),
      h('span', { text: note })));
    if (!res.rows.length) {
      box.appendChild(h('p', { class: 'blank-none', text: '沒有符合的列。' }));
      return;
    }
    var list = h('ul', { class: 'blank-items fz-items' });
    res.rows.slice(0, MAX_FZ_LIST).forEach(function (it) {
      list.appendChild(h('li', { class: 'fz-item' },
        h('span', { class: 'blank-row', text: '第 ' + it.row + ' 列' }),
        markText(it.value, terms),
        h('span', { class: 'blank-text fz-ctx', text: it.text })));
    });
    if (res.rows.length > MAX_FZ_LIST) {
      list.appendChild(h('li', { class: 'blank-more', text: '另有 ' + (res.rows.length - MAX_FZ_LIST).toLocaleString('en-US') + ' 列沒有列出（一樣會留下）' }));
    }
    box.appendChild(list);
  }
  // 把含有的字用螢光筆標出來
  function markText(text, terms) {
    var box = h('span', { class: 'fz-val' });
    var low = text.toLowerCase();
    var pos = 0;
    while (pos < text.length) {
      var best = -1, len = 0;
      terms.forEach(function (t) {
        var j = low.indexOf(t, pos);
        if (j >= 0 && (best < 0 || j < best || (j === best && t.length > len))) { best = j; len = t.length; }
      });
      if (best < 0 || len === 0) { box.appendChild(document.createTextNode(text.slice(pos))); break; }
      if (best > pos) box.appendChild(document.createTextNode(text.slice(pos, best)));
      box.appendChild(h('mark', { text: text.slice(best, best + len) }));
      pos = best + len;
    }
    return box;
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
