/*
 * excel-io.js — 所有 Excel 讀寫與排版都集中在這裡（使用 ExcelJS）。
 * 各步驟只呼叫 ExcelIO，不直接碰 ExcelJS，之後要調整匯出格式只需要改這個檔。
 *
 * 名詞：行 = 直的（A、B、C…），列 = 橫的（1、2、3…）。
 */
(function (global) {
  'use strict';

  /* ---------- 行號（字母）與數字互轉 ---------- */

  function colToNum(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n;
  }

  function numToCol(n) {
    var s = '';
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  // 全形轉半形、去頭尾空白（使用者常用全形輸入 Ａ、１）
  function normalizeText(s) {
    return String(s).normalize('NFKC').trim();
  }

  /*
   * 解析「B, C, U:Z, AE」這類輸入，回傳 { cols: [2,3,21,...], bad: ['看不懂的字'] }
   * 接受逗號、頓號、空白、分號分隔；範圍可用 : 或 - 或 ~
   */
  function parseColumns(text) {
    var cols = {};
    var bad = [];
    var tokens = normalizeText(text || '').toUpperCase().split(/[\s,、;]+/).filter(Boolean);
    tokens.forEach(function (tok) {
      var m = /^([A-Z]{1,3})(?:[:\-~]([A-Z]{1,3}))?$/.exec(tok);
      if (!m) { bad.push(tok); return; }
      var a = colToNum(m[1]);
      var b = m[2] ? colToNum(m[2]) : a;
      if (a > 16384 || b > 16384) { bad.push(tok); return; }
      if (a > b) { var t = a; a = b; b = t; }
      for (var c = a; c <= b; c++) cols[c] = true;
    });
    var list = Object.keys(cols).map(Number).sort(function (x, y) { return x - y; });
    return { cols: list, bad: bad };
  }

  function formatColumns(cols) {
    return cols.map(numToCol).join(', ');
  }

  /* ---------- 儲存格的值 ---------- */

  // 把 ExcelJS 的各種值型態（公式、RichText、超連結…）化成單純的值
  function plainValue(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v;
    if (typeof v !== 'object') return v;
    if (v.formula !== undefined || v.sharedFormula !== undefined) return plainValue(v.result);
    if (Array.isArray(v.richText)) {
      return v.richText.map(function (t) { return t.text; }).join('');
    }
    if (v.error !== undefined) return v.error;
    if (v.text !== undefined) return plainValue(v.text);
    if (v.result !== undefined) return plainValue(v.result);
    return String(v);
  }

  // 合併儲存格裡不是左上角的格子，值要去找左上角那一格
  function cellRawValue(cell) {
    if (!cell) return null;
    var src = cell.isMerged ? cell.master : cell;
    var v = src.value;
    // ExcelJS 的 value 會把「計算結果剛好是 0」的公式結果丟掉（變成沒有結果），用 result 補回來
    if (v && typeof v === 'object' && !(v instanceof Date) && v.result === undefined &&
        (v.formula !== undefined || v.sharedFormula !== undefined)) {
      var res = src.result;
      if (res !== undefined && res !== null) v = Object.assign({}, v, { result: res });
    }
    return v;
  }

  // 比對用的鍵：1、"1"、"１"、"1.0" 都視為相同
  function matchKey(v) {
    v = plainValue(v);
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString();
    if (typeof v === 'number') return String(v);
    var s = normalizeText(v);
    if (s !== '' && !isNaN(Number(s))) return String(Number(s));
    return s;
  }

  /* ---------- 模糊比對（第 1 步第三種選取方式）：儲存格文字含有任一個值就算符合，英文大小寫不分 ---------- */
  function fuzzyText(v) {
    v = plainValue(v);
    if (v === null || v === undefined) return '';
    if (v instanceof Date) v = v.toISOString().slice(0, 10);
    return normalizeText(v).toLowerCase();
  }
  function fuzzyTerms(values) {
    return (values || []).map(fuzzyText).filter(function (s) { return s !== ''; });
  }
  // 回傳儲存格文字含有的第一個值（已轉小寫）；都沒有回傳 ''
  function fuzzyHit(v, terms) {
    var s = fuzzyText(v);
    if (s === '') return '';
    for (var i = 0; i < terms.length; i++) if (s.indexOf(terms[i]) >= 0) return terms[i];
    return '';
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  // 預覽畫面顯示用的文字
  function displayText(cell) {
    var v = plainValue(cellRawValue(cell));
    if (v === null || v === undefined) return '';
    if (v instanceof Date) {
      var d = v.getUTCFullYear() + '/' + (v.getUTCMonth() + 1) + '/' + v.getUTCDate();
      if (v.getUTCHours() || v.getUTCMinutes()) {
        d += ' ' + pad2(v.getUTCHours()) + ':' + pad2(v.getUTCMinutes());
      }
      return d;
    }
    if (typeof v === 'number') {
      var fmt = (cell && cell.numFmt) || '';
      if (fmt.indexOf('%') >= 0) return (Math.round(v * 10000) / 100) + '%';
      var rounded = Math.round(v * 10000) / 10000;
      if (fmt.indexOf('#,##') >= 0) return rounded.toLocaleString('en-US');
      return String(rounded);
    }
    return String(v);
  }

  // 匯出時寫入的值：公式換成計算結果（篩掉列與行之後，原公式的參照會錯位）
  function exportValue(cell) {
    var v = cellRawValue(cell);
    if (v && typeof v === 'object' && !(v instanceof Date) &&
        (v.formula !== undefined || v.sharedFormula !== undefined)) {
      return v.result === undefined ? null : v.result;
    }
    return v === undefined ? null : v;
  }

  /* ---------- 讀檔 ---------- */

  function ioError(code, message) {
    var e = new Error(message);
    e.code = code;
    return e;
  }

  /*
   * 讀檔前的修正：字型旗標。
   *
   * ExcelJS 4.x 讀字型設定時，只要看到 <b>、<i>、<strike> 這些標籤就當成「有」，
   * 不管後面是不是寫著 val="0"（＝沒有）。SAP 這類系統匯出的報表會把「沒有粗體、
   * 沒有刪除線」明寫成 <b val="0"/> <strike val="0"/>，結果整張表被讀成
   * 粗體＋斜體＋刪除線，匯出後就全部被畫上橫線。
   *
   * 做法：.xlsx 其實是一個 zip 壓縮檔。交給 ExcelJS 之前，先把裡面的樣式檔
   * （styles.xml）和文字檔（sharedStrings.xml）解開，拿掉這些「明寫為沒有」的標籤，
   * 再包回去。解壓縮用下面自己寫的 inflateRaw()，不靠瀏覽器的功能，所以 Chrome、
   * Edge、Safari 新舊版本的結果都一樣。任何一步不順利就原封不動交給 ExcelJS。
   */
  var OFF_FLAGS = /<(b|i|strike|outline|shadow|condense|extend)\s+val\s*=\s*(["'])(?:0|false)\2\s*(?:\/>|>\s*<\/\1>)/g;
  var FIX_ENTRIES = { 'xl/styles.xml': true, 'xl/sharedStrings.xml': true };

  /* 解開 zip 裡用 deflate 壓縮的資料。src：壓縮資料；size：解開後的大小。 */
  var LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
  var LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
  var DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
  var DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
  var CODE_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

  function huffman(lengths, start, n) {
    var count = new Uint16Array(16);
    var symbol = new Uint16Array(n);
    var offs = new Uint16Array(16);
    var i;
    for (i = 0; i < n; i++) count[lengths[start + i]]++;
    for (i = 1; i < 15; i++) offs[i + 1] = offs[i] + count[i];
    for (i = 0; i < n; i++) {
      var len = lengths[start + i];
      if (len) symbol[offs[len]++] = i;
    }
    return { count: count, symbol: symbol };
  }

  function inflateRaw(src, size) {
    var out = new Uint8Array(size);
    var op = 0, ip = 0, bitBuf = 0, bitCnt = 0;

    function bits(n) {
      while (bitCnt < n) {
        if (ip >= src.length) throw new Error('inflate: 資料不完整');
        bitBuf |= src[ip++] << bitCnt;
        bitCnt += 8;
      }
      var v = bitBuf & ((1 << n) - 1);
      bitBuf >>>= n;
      bitCnt -= n;
      return v;
    }
    function decode(h) {
      var code = 0, first = 0, index = 0;
      for (var len = 1; len <= 15; len++) {
        code |= bits(1);
        var cnt = h.count[len];
        if (code - cnt < first) return h.symbol[index + (code - first)];
        index += cnt;
        first = (first + cnt) << 1;
        code <<= 1;
      }
      throw new Error('inflate: 編碼錯誤');
    }
    function block(lenCode, distCode) {
      for (;;) {
        var sym = decode(lenCode);
        if (sym < 256) {
          if (op >= size) throw new Error('inflate: 超出大小');
          out[op++] = sym;
        } else if (sym === 256) {
          return;
        } else {
          sym -= 257;
          if (sym >= 29) throw new Error('inflate: 長度錯誤');
          var len = LEN_BASE[sym] + bits(LEN_EXTRA[sym]);
          var ds = decode(distCode);
          if (ds >= 30) throw new Error('inflate: 距離錯誤');
          var dist = DIST_BASE[ds] + bits(DIST_EXTRA[ds]);
          if (dist > op || op + len > size) throw new Error('inflate: 距離錯誤');
          for (var k = 0; k < len; k++, op++) out[op] = out[op - dist];
        }
      }
    }

    var last, lengths, i;
    do {
      last = bits(1);
      var type = bits(2);
      if (type === 0) {                      // 沒壓縮的區塊
        bitBuf = 0; bitCnt = 0;
        if (ip + 4 > src.length) throw new Error('inflate: 資料不完整');
        var n = src[ip] | (src[ip + 1] << 8);
        ip += 4;
        if (ip + n > src.length || op + n > size) throw new Error('inflate: 資料不完整');
        out.set(src.subarray(ip, ip + n), op);
        ip += n; op += n;
      } else if (type === 1) {               // 固定編碼
        lengths = new Uint8Array(288 + 30);
        for (i = 0; i < 144; i++) lengths[i] = 8;
        for (; i < 256; i++) lengths[i] = 9;
        for (; i < 280; i++) lengths[i] = 7;
        for (; i < 288; i++) lengths[i] = 8;
        for (; i < 318; i++) lengths[i] = 5;
        block(huffman(lengths, 0, 288), huffman(lengths, 288, 30));
      } else if (type === 2) {               // 動態編碼
        var nlen = bits(5) + 257, ndist = bits(5) + 1, ncode = bits(4) + 4;
        if (nlen > 286 || ndist > 30) throw new Error('inflate: 編碼表錯誤');
        lengths = new Uint8Array(320);
        for (i = 0; i < ncode; i++) lengths[CODE_ORDER[i]] = bits(3);
        var codeLen = huffman(lengths, 0, 19);
        lengths = new Uint8Array(320);
        i = 0;
        while (i < nlen + ndist) {
          var sym = decode(codeLen);
          if (sym < 16) {
            lengths[i++] = sym;
          } else {
            var prev = 0, rep;
            if (sym === 16) {
              if (i === 0) throw new Error('inflate: 編碼表錯誤');
              prev = lengths[i - 1];
              rep = 3 + bits(2);
            } else if (sym === 17) {
              rep = 3 + bits(3);
            } else {
              rep = 11 + bits(7);
            }
            if (i + rep > nlen + ndist) throw new Error('inflate: 編碼表錯誤');
            while (rep--) lengths[i++] = prev;
          }
        }
        block(huffman(lengths, 0, nlen), huffman(lengths, nlen, ndist));
      } else {
        throw new Error('inflate: 區塊類型錯誤');
      }
    } while (!last);
    if (op !== size) throw new Error('inflate: 大小不符');
    return out;
  }

  var crcTable = null;
  function crc32(bytes) {
    if (!crcTable) {
      crcTable = new Uint32Array(256);
      for (var n = 0; n < 256; n++) {
        var c = n;
        for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
        crcTable[n] = c >>> 0;
      }
    }
    var crc = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) crc = crcTable[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  // 讀出 zip 裡每個檔案的資訊與（仍是壓縮狀態的）內容。看不懂的格式回傳 null。
  function readZip(buffer) {
    var dv = new DataView(buffer);
    var size = buffer.byteLength;
    var eocd = -1;
    for (var i = size - 22; i >= 0 && i >= size - 22 - 65535; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) return null;
    var count = dv.getUint16(eocd + 10, true);
    var p = dv.getUint32(eocd + 16, true);
    if (count === 0xFFFF || p === 0xFFFFFFFF) return null;       // zip64，不處理
    var decoder = new TextDecoder();
    var entries = [];
    for (var k = 0; k < count; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) return null;
      var e = {
        flags: dv.getUint16(p + 8, true),
        method: dv.getUint16(p + 10, true),
        time: dv.getUint16(p + 12, true),
        date: dv.getUint16(p + 14, true),
        crc: dv.getUint32(p + 16, true),
        csize: dv.getUint32(p + 20, true),
        usize: dv.getUint32(p + 24, true),
        intAttr: dv.getUint16(p + 36, true),
        extAttr: dv.getUint32(p + 38, true)
      };
      var nameLen = dv.getUint16(p + 28, true);
      var extraLen = dv.getUint16(p + 30, true);
      var commentLen = dv.getUint16(p + 32, true);
      var local = dv.getUint32(p + 42, true);
      if (e.flags & 1) return null;                              // 有加密，不處理
      if (e.csize === 0xFFFFFFFF || e.usize === 0xFFFFFFFF || local === 0xFFFFFFFF) return null;
      e.nameBytes = new Uint8Array(buffer, p + 46, nameLen);
      e.name = decoder.decode(e.nameBytes);
      if (dv.getUint32(local, true) !== 0x04034b50) return null;
      var dataStart = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
      if (dataStart + e.csize > size) return null;
      e.data = new Uint8Array(buffer, dataStart, e.csize);
      entries.push(e);
      p += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  // 把檔案清單重新包成 zip
  function writeZip(entries) {
    var total = 22;
    entries.forEach(function (e) { total += 30 + 46 + 2 * e.nameBytes.length + e.data.length; });
    var out = new Uint8Array(total);
    var dv = new DataView(out.buffer);
    var p = 0;
    entries.forEach(function (e) {
      e.offset = p;
      e.outFlags = e.flags & (e.method === 8 ? 0x0806 : 0x0800);  // 留下壓縮等級與 UTF-8 檔名旗標
      dv.setUint32(p, 0x04034b50, true);
      dv.setUint16(p + 4, 20, true);
      dv.setUint16(p + 6, e.outFlags, true);
      dv.setUint16(p + 8, e.method, true);
      dv.setUint16(p + 10, e.time, true);
      dv.setUint16(p + 12, e.date, true);
      dv.setUint32(p + 14, e.crc, true);
      dv.setUint32(p + 18, e.data.length, true);
      dv.setUint32(p + 22, e.usize, true);
      dv.setUint16(p + 26, e.nameBytes.length, true);
      out.set(e.nameBytes, p + 30);
      out.set(e.data, p + 30 + e.nameBytes.length);
      p += 30 + e.nameBytes.length + e.data.length;
    });
    var cdStart = p;
    entries.forEach(function (e) {
      dv.setUint32(p, 0x02014b50, true);
      dv.setUint16(p + 4, 20, true);
      dv.setUint16(p + 6, 20, true);
      dv.setUint16(p + 8, e.outFlags, true);
      dv.setUint16(p + 10, e.method, true);
      dv.setUint16(p + 12, e.time, true);
      dv.setUint16(p + 14, e.date, true);
      dv.setUint32(p + 16, e.crc, true);
      dv.setUint32(p + 20, e.data.length, true);
      dv.setUint32(p + 24, e.usize, true);
      dv.setUint16(p + 28, e.nameBytes.length, true);
      dv.setUint16(p + 36, e.intAttr, true);
      dv.setUint32(p + 38, e.extAttr, true);
      dv.setUint32(p + 42, e.offset, true);
      out.set(e.nameBytes, p + 46);
      p += 46 + e.nameBytes.length;
    });
    dv.setUint32(p, 0x06054b50, true);
    dv.setUint16(p + 8, entries.length, true);
    dv.setUint16(p + 10, entries.length, true);
    dv.setUint32(p + 12, p - cdStart, true);
    dv.setUint32(p + 16, cdStart, true);
    return out.buffer;
  }

  // 回傳修正過的檔案內容（ArrayBuffer）；不需要修或修不了就回傳原本的
  function fixFontFlags(buffer) {
    try {
      var entries = readZip(buffer);
      if (!entries) return buffer;
      var changed = false;
      entries.forEach(function (e) {
        if (!FIX_ENTRIES[e.name] || (e.method !== 8 && e.method !== 0)) return;
        var bytes = e.method === 8 ? inflateRaw(e.data, e.usize) : e.data;
        if (crc32(bytes) !== e.crc) return;                       // 解出來的內容不對，不處理
        if (bytes[0] === 0xFF || bytes[0] === 0xFE) return;       // 不是 UTF-8，不處理
        var text = new TextDecoder('utf-8').decode(bytes);
        var fixed = text.replace(OFF_FLAGS, '');
        if (fixed === text) return;
        var plain = new TextEncoder().encode(fixed);
        e.data = plain;                                           // 改過的檔案不壓縮，直接存
        e.method = 0;
        e.crc = crc32(plain);
        e.usize = plain.length;
        changed = true;
      });
      return changed ? writeZip(entries) : buffer;
    } catch (err) {
      return buffer;
    }
  }

  function load(arrayBuffer) {
    var head = new Uint8Array(arrayBuffer.slice(0, 4));
    if (head[0] === 0xD0 && head[1] === 0xCF) {
      return Promise.reject(ioError('OLD_XLS',
        '這個檔案是舊版 .xls 格式，或有密碼保護。請在 Excel 用「另存新檔」存成 .xlsx（並取消密碼）後再上傳。'));
    }
    if (!(head[0] === 0x50 && head[1] === 0x4B)) {
      return Promise.reject(ioError('NOT_XLSX',
        '這不是 Excel 活頁簿（.xlsx）。請確認上傳的是 .xlsx 檔。'));
    }
    var wb = new global.ExcelJS.Workbook();
    return wb.xlsx.load(fixFontFlags(arrayBuffer)).then(function () {
      if (!wb.worksheets.length) throw ioError('NO_SHEET', '這個活頁簿裡沒有工作表。');
      return wb;
    }, function (err) {
      if (err && err.code) throw err;
      throw ioError('READ_FAIL', '讀不了這個 Excel 檔：' + (err && err.message ? err.message : err));
    });
  }

  /* ---------- 合併儲存格 ---------- */

  function listMerges(ws) {
    var out = [];
    var merges = ws._merges || {};
    Object.keys(merges).forEach(function (key) {
      var m = merges[key];
      if (m) out.push({ top: m.top, left: m.left, bottom: m.bottom, right: m.right });
    });
    return out;
  }

  /* ---------- 檢視工作表內容 ---------- */

  /*
   * 回傳：
   *   lastRow     最後一列有資料的列號
   *   colCount    最右邊有資料的行號
   *   dataRows    資料列數（從 dataStartRow 起）
   *   values      選取行出現過的值 [{ key, label, count }]，依出現順序
   *   blankCount  選取行是空白的列數
   */
  function inspect(ws, opts) {
    var dataStartRow = opts.dataStartRow;
    var selectorCol = opts.selectorCol;
    var lastRow = 0;
    var colCount = 0;
    var counts = {};
    var order = [];
    var nonBlank = 0;

    ws.eachRow({ includeEmpty: false }, function (row, r) {
      var rowHasValue = false;
      row.eachCell({ includeEmpty: false }, function (cell, c) {
        var v = cellRawValue(cell);
        if (v === null || v === undefined || v === '') return;
        rowHasValue = true;
        if (c > colCount) colCount = c;
      });
      if (!rowHasValue) return;
      if (r > lastRow) lastRow = r;
      if (r < dataStartRow) return;
      var cell = row.findCell(selectorCol);
      var key = matchKey(cellRawValue(cell));
      if (key === '') return;
      nonBlank++;
      if (!counts[key]) {
        counts[key] = { key: key, label: displayText(cell), count: 0 };
        order.push(key);
      }
      counts[key].count++;
    });

    var dataRows = Math.max(0, lastRow - dataStartRow + 1);
    return {
      lastRow: lastRow,
      colCount: colCount,
      dataRows: dataRows,
      values: order.map(function (k) { return counts[k]; }),
      blankCount: dataRows - nonBlank
    };
  }

  /* ---------- 工作表的範圍 ---------- */

  // 逐列、逐格走訪「實際存在」的列與儲存格（包含只有格式、沒有值的格子）
  function eachStoredRow(ws, fn) {
    var rows = ws._rows;
    if (Array.isArray(rows)) {
      for (var i = 0; i < rows.length; i++) if (rows[i]) fn(rows[i], i + 1);
    } else {
      ws.eachRow({ includeEmpty: false }, fn);
    }
  }
  function eachStoredCell(row, fn) {
    var cells = row._cells;
    if (Array.isArray(cells)) {
      for (var i = 0; i < cells.length; i++) if (cells[i]) fn(cells[i], i + 1);
    } else {
      row.eachCell({ includeEmpty: false }, fn);
    }
  }

  // 回傳 { lastRow, colCount }：最後一列、最右一行「有值」的位置
  function extent(ws) {
    var lastRow = 0;
    var colCount = 0;
    eachStoredRow(ws, function (row, r) {
      var has = false;
      eachStoredCell(row, function (cell, c) {
        var v = cellRawValue(cell);
        if (v === null || v === undefined || v === '') return;
        has = true;
        if (c > colCount) colCount = c;
      });
      if (has && r > lastRow) lastRow = r;
    });
    return { lastRow: lastRow, colCount: colCount };
  }

  /* ---------- 篩選並產生新的活頁簿 ---------- */

  function cleanObject(o) {
    var out = {};
    Object.keys(o).forEach(function (k) { if (o[k] !== undefined) out[k] = o[k]; });
    return out;
  }

  /*
   * opts:
   *   dataStartRow  資料從第幾列開始（之前的列是表頭，全部保留）
   *   selectorCol   選取條件所在的行（數字，A=1）
   *   values        要選取的值（字串陣列）
   *   keepCols      要保留的行（數字陣列，已排序）
   *   lastRow       最後一列資料的列號
   *
   * 表頭列的樣式、合併儲存格、行寬、列高都會照原樣帶到新檔；
   * 合併儲存格如果只保留了其中幾行，會縮成只合併那幾行。
   */
  // 第 1 步篩選時把儲存格的填滿色彩拿掉（變透明）：原報表有些列塗了底色，
  // 留著的話會蓋過後面步驟標的前三大淺橘色，匯出的檔案也會一直帶著那些顏色。
  function withoutFill(style) {
    var out = Object.assign({}, style || {});
    delete out.fill;
    return out;
  }

  /*
   * 模糊比對的結果，給畫面預覽用（和 appendFilteredRows() 同一套規則）
   *   opts  { dataStartRow, selectorCol, lastRow, values, exclude }
   *   回傳  { rows: [{ row, value: 那一格的文字, text: 這一列前幾格 }], hit: 含有任一值的列數, blank: 那一行空白的列數 }
   *   rows 是會留下的列：exclude=false 時是含有任一值的列，exclude=true 時是都不含的列。那一行空白的列一律不留。
   */
  function fuzzyRows(ws, opts) {
    var terms = fuzzyTerms(opts.values);
    var colCount = extent(ws).colCount;
    var out = { rows: [], hit: 0, blank: 0 };
    for (var r = opts.dataStartRow; r <= opts.lastRow; r++) {
      var row = ws.findRow(r);
      if (!row) continue;
      var raw = cellRawValue(row.findCell(opts.selectorCol));
      if (fuzzyText(raw) === '') {
        var any = false;
        for (var c0 = 1; c0 <= colCount && !any; c0++) {
          var v0 = plainValue(cellRawValue(row.findCell(c0)));
          any = v0 !== null && v0 !== undefined && v0 !== '';
        }
        if (any) out.blank++;
        continue;
      }
      var hit = fuzzyHit(raw, terms) !== '';
      if (hit) out.hit++;
      if (opts.exclude ? hit : !hit) continue;
      var pv0 = plainValue(raw);
      var parts = [];
      for (var c = 1; c <= colCount && parts.length < 3; c++) {
        if (c === opts.selectorCol) continue;
        var pv = plainValue(cellRawValue(row.findCell(c)));
        var s = pv === null || pv === undefined ? '' : normalizeText(pv instanceof Date ? pv.toISOString().slice(0, 10) : pv);
        if (s !== '') parts.push(s.length > 30 ? s.slice(0, 30) + '…' : s);
      }
      out.rows.push({
        row: r,
        value: normalizeText(pv0 instanceof Date ? pv0.toISOString().slice(0, 10) : pv0),
        text: parts.join('｜')
      });
    }
    return out;
  }

  function filter(ws, opts) {
    var ExcelJS = global.ExcelJS;
    var dataStartRow = opts.dataStartRow;
    var headerEnd = dataStartRow - 1;
    var keepCols = opts.keepCols;

    var newIndex = {}; // 原行號 -> 新行號
    keepCols.forEach(function (c, i) { newIndex[c] = i + 1; });

    var out = new ExcelJS.Workbook();
    out.creator = 'IE 人力需求小程式';
    out.created = new Date();
    var props = ws.properties || {};
    var ows = out.addWorksheet(ws.name || 'MPS', {
      properties: cleanObject({
        defaultRowHeight: props.defaultRowHeight,
        defaultColWidth: props.defaultColWidth
      })
    });

    // 行寬
    keepCols.forEach(function (c, i) {
      var w = ws.getColumn(c).width;
      if (w) ows.getColumn(i + 1).width = w;
    });

    // 表頭列：每一格先照抄（值只有左上角那格有，其餘合併格只抄樣式）
    var r, i, src, dst, srcRow, dstRow;
    for (r = 1; r <= headerEnd; r++) {
      srcRow = ws.findRow(r);
      if (!srcRow) continue;
      dstRow = ows.getRow(r);
      if (srcRow.height) dstRow.height = srcRow.height;
      for (i = 0; i < keepCols.length; i++) {
        src = srcRow.findCell(keepCols[i]);
        if (!src) continue;
        dst = dstRow.getCell(i + 1);
        dst.style = withoutFill(src.style);
        if (!src.isMerged || src.master === src) dst.value = exportValue(src);
      }
    }

    // 表頭的合併儲存格：對應到新的行號後重新合併
    listMerges(ws).forEach(function (m) {
      if (m.top > headerEnd) return;
      var bottom = Math.min(m.bottom, headerEnd);
      var kept = keepCols.filter(function (c) { return c >= m.left && c <= m.right; });
      if (!kept.length) return;
      var n1 = newIndex[kept[0]];
      var n2 = newIndex[kept[kept.length - 1]];
      var master = ws.getCell(m.top, m.left);
      var target = ows.getCell(m.top, n1);
      target.value = exportValue(master);
      if (kept[0] !== m.left) target.style = withoutFill(master.style);
      if (n2 > n1 || bottom > m.top) {
        ows.mergeCellsWithoutStyle(m.top, n1, bottom, n2);
      }
    });

    // 第 1 列的大標題如果落在沒保留的行，補到新檔的左上角，避免標題不見
    srcRow = headerEnd >= 1 ? ws.findRow(1) : null;
    if (srcRow) {
      var titleCell = null;
      srcRow.eachCell({ includeEmpty: false }, function (cell) {
        if (titleCell || (cell.isMerged && cell.master !== cell)) return;
        var tv = cell.value;
        if (tv !== null && tv !== undefined && tv !== '') titleCell = cell;
      });
      var titleKept = false;
      for (i = 1; i <= keepCols.length; i++) {
        var cv = cellRawValue(ows.getCell(1, i));
        if (cv !== null && cv !== undefined && cv !== '') { titleKept = true; break; }
      }
      if (titleCell && !titleKept) {
        var a1 = ows.getCell(1, 1);
        a1.value = exportValue(titleCell);
        a1.style = withoutFill(titleCell.style);
      }
    }

    // 資料列：選取行的值符合才保留
    var rowsOut = appendFilteredRows(ows, dataStartRow, ws, opts);
    var matched = rowsOut.matched;
    var total = rowsOut.total;
    var outRow = rowsOut.nextRow;

    // 凍結窗格、顯示比例
    var v = (ws.views && ws.views[0]) || {};
    var view = cleanObject({ showGridLines: v.showGridLines, zoomScale: v.zoomScale });
    if (v.state === 'frozen') {
      view.state = 'frozen';
      view.xSplit = keepCols.filter(function (c) { return c <= (v.xSplit || 0); }).length;
      view.ySplit = v.ySplit || 0;
    }
    if (Object.keys(view).length) ows.views = [view];

    // 列印設定：紙張方向照原檔，每頁重複表頭
    var ps = ws.pageSetup || {};
    Object.assign(ows.pageSetup, cleanObject({
      orientation: ps.orientation,
      paperSize: ps.paperSize,
      fitToPage: ps.fitToPage,
      fitToWidth: ps.fitToWidth,
      fitToHeight: ps.fitToHeight,
      printTitlesRow: headerEnd >= 1 ? '1:' + headerEnd : undefined
    }));

    return {
      workbook: out,
      worksheet: ows,
      matchedRows: matched,
      totalRows: total,
      headerEnd: headerEnd,
      lastRow: outRow - 1,
      colCount: keepCols.length
    };
  }

  /*
   * 把 ws 裡「選取行的值符合」的資料列，依序寫到 ows 的第 outRow 列起。
   * o: { dataStartRow, selectorCol, values, op, keepCols, lastRow }
   *   op  'neq' = 留下選取行「不等於」values 的列；其他（沒給）= 留下「等於」的列
   *   fuzzy true = 模糊比對：選取行的字含有任一個 values 就算符合（英文大小寫、全形半形不分）；
   *         exclude=false 留符合的列，exclude=true 留都不符合的列（這時不看 op）
   *   三種都一樣：選取行空白的列一律不要
   * 回傳 { matched, total, nextRow }（nextRow = 下一個空的列號）
   */
  function appendFilteredRows(ows, outRow, ws, o) {
    var sel = selectRows(ws, o);
    var next = copyRows(ows, outRow, ws, sel.rows, o.keepCols);
    return { matched: sel.rows.length, total: sel.total, nextRow: next };
  }

  // 符合選取條件的列（列號，由上往下）；規則見 appendFilteredRows()
  function selectRows(ws, o) {
    var wanted = {};
    o.values.forEach(function (v) { wanted[matchKey(v)] = true; });
    var fzTerms = o.fuzzy ? fuzzyTerms(o.values) : null;   // 模糊比對（第三種選取方式）
    var rows = [];
    var total = 0;
    for (var r = o.dataStartRow; r <= o.lastRow; r++) {
      total++;
      var srcRow = ws.findRow(r);
      if (!srcRow) continue;
      var raw = cellRawValue(srcRow.findCell(o.selectorCol));
      if (fzTerms) {
        // 模糊比對：那一行空白的列一律不要；選取留含有任一值的列，排除留都不含的列
        if (fuzzyText(raw) === '') continue;
        var fzHit = fuzzyHit(raw, fzTerms) !== '';
        if (o.exclude ? fzHit : !fzHit) continue;
      } else {
        var key = matchKey(raw);
        if (key === '') continue;                                // 選取行空白的列一律不要
        if (o.op === 'neq' ? !!wanted[key] : !wanted[key]) continue;
      }
      rows.push(r);
    }
    return { rows: rows, total: total };
  }

  // 把 ws 的這幾列（只要 keepCols 這幾行）依序寫到 ows 的第 outRow 列起；回傳下一個空的列號
  function copyRows(ows, outRow, ws, rows, keepCols) {
    rows.forEach(function (r) {
      var srcRow = ws.findRow(r);
      var dstRow = ows.getRow(outRow++);
      if (srcRow.height) dstRow.height = srcRow.height;
      for (var i = 0; i < keepCols.length; i++) {
        var src = srcRow.findCell(keepCols[i]);
        if (!src) continue;
        var dst = dstRow.getCell(i + 1);
        dst.style = withoutFill(src.style);
        var val = exportValue(src);
        if (val !== null) dst.value = val;
      }
    });
    return outRow;
  }

  /* ---------- 第 1 步：挑出 RMA ---------- */

  /*
   * RMA 的條件（一組，套用到所有報表留下來的列）：
   *   rma = { col, fuzzy, values, skip }
   *     col     新檔的第幾行（1 起算）；每份報表用它自己保留的第 col 個行去比
   *     fuzzy   true = 模糊比對（含有任一個值）；false = 等於任一個值
   *     values  要比對的值（字串陣列，空的不算）
   *     skip    { '報表 id:列號': true }：模糊比對時使用者取消勾選的列（不算 RMA）
   * 那一行空白的列不會是 RMA。col 是 0 或沒有值時，不挑出任何 RMA。
   */
  function rmaMatcher(rma) {
    if (!rma || !rma.col) return null;
    var values = (rma.values || []).filter(function (v) { return normalizeText(v) !== ''; });
    if (!values.length) return null;
    if (rma.fuzzy) {
      var terms = fuzzyTerms(values);
      return function (raw) { return fuzzyText(raw) !== '' && fuzzyHit(raw, terms) !== ''; };
    }
    var wanted = {};
    values.forEach(function (v) { wanted[matchKey(v)] = true; });
    return function (raw) { var k = matchKey(raw); return k !== '' && !!wanted[k]; };
  }
  function rmaRaw(ws, r, keepCols, col) {
    var src = keepCols[col - 1];
    var row = ws.findRow(r);
    return src && row ? cellRawValue(row.findCell(src)) : null;
  }

  /*
   * 給畫面用：各報表留下來的列裡，RMA 那一行出現過哪些值、哪些列符合 RMA 的條件。
   *   parts、common 和 filterSections() 一樣（parts[k].id 是報表的代號，用來對 skip）
   * 回傳 { values: [{ key, label, count }], blank, hits: [{ id, part, row, value, text, skipped }] }
   *   hits 依 取像 → 車用 → 光機系統 的順序（和放進【RMA】的順序一樣）
   */
  function rmaScan(parts, common, rma) {
    var out = { values: [], blank: 0, hits: [] };
    if (!rma || !rma.col) return out;
    var match = rmaMatcher(rma);
    var counts = {}, order = [];
    var skip = rma.skip || {};
    sectionOrder(parts).forEach(function (k) {
      var p = parts[k];
      var o = partOpts(p, common);
      selectRows(p.ws, o).rows.forEach(function (r) {
        var raw = rmaRaw(p.ws, r, p.keepCols, rma.col);
        var key = matchKey(raw);
        if (key === '') { out.blank++; return; }
        if (!counts[key]) {
          var pv = plainValue(raw);
          counts[key] = { key: key, label: pv instanceof Date ? pv.toISOString().slice(0, 10) : normalizeText(pv), count: 0 };
          order.push(key);
        }
        counts[key].count++;
        if (!match || !match(raw)) return;
        var ctx = [];
        for (var i = 0; i < p.keepCols.length && ctx.length < 3; i++) {
          if (i === rma.col - 1) continue;
          var cv = plainValue(rmaRaw(p.ws, r, p.keepCols, i + 1));
          var s = cv === null || cv === undefined ? '' : normalizeText(cv instanceof Date ? cv.toISOString().slice(0, 10) : cv);
          if (s !== '') ctx.push(s.length > 30 ? s.slice(0, 30) + '…' : s);
        }
        out.hits.push({ id: p.id, part: k, row: r, value: counts[key].label, text: ctx.join('｜'), skipped: !!skip[p.id + ':' + r] });
      });
    });
    out.values = order.map(function (k) { return counts[k]; });
    return out;
  }

  // 報表的順序：照 取像 → 車用 → 光機系統，同一類照上傳順序（回傳 parts 的索引）
  function sectionOrder(parts) {
    var out = [];
    MPS_SECTIONS.forEach(function (sec) {
      parts.forEach(function (p, k) { if (p.cat === sec.key) out.push(k); });
    });
    return out;
  }
  function partOpts(p, common) {
    return {
      dataStartRow: common.dataStartRow, selectorCol: p.selectorCol || common.selectorCol,
      values: p.values, op: p.op, fuzzy: !!p.fuzzy, exclude: !!p.exclude,
      keepCols: p.keepCols, lastRow: p.lastRow
    };
  }

  /*
   * 好幾份 MPS 報表篩選後，依類別分段排成一張表（第 1 步）。
   *   parts   [{ id, ws, cat, values, keepCols, lastRow, selectorCol, op, fuzzy, exclude }]，依上傳順序（id 對 rma.skip 用）
   *           cat：'image' 取像、'auto' 車用、'optics' 光機系統（MPS_SECTIONS）
   *   common  { dataStartRow, selectorCol }
   * 版面：表頭（第 1 列到 dataStartRow-1）用第一份報表的。之後照 取像 → 車用 → 光機系統 的順序，
   *   每一段先放一列類別標題（跨所有保留的行合併、粗體置中、底色 #E9C46A），底下是這一類各份報表符合條件的列；
   *   同一類有好幾份時照上傳順序往下接，中間不畫線。
   *   取像最後一列之後空 3 列才是【車用】；取像＋車用這一組最後一列之後，空 1 列留 3 列給第 2 步的匯總表，
   *   再空 6 列才是【光機系統】。這個月沒有檔案的類別整段不放。
   *   rma     RMA 的條件（見 rmaMatcher()）。有給的話（就算沒有條件）最後一定放【RMA】段：
   *           最後一張匯總表下面空 3 列放【RMA】標題（底色 #FFCAD4），底下是挑出來的列（照 取像 → 車用 → 光機系統 的順序），
   *           再空 1 列給第 2 步的 RMA Total。挑去 RMA 的列不會留在原本的段裡；一類的列全部挑走時，那一類整段不放。
   * 回傳和 filter() 一樣，另外多：
   *   parts     [{ matched, rma, total, firstRow, lastRow, cat }]（和 parts 同順序；matched 不含挑去 RMA 的列）
   *   sections  [{ key, label, group, titleRow, firstRow, lastRow, matched }]（最後一個是 RMA）
   */
  function filterSections(parts, common, rma) {
    // 先只做表頭（資料一列都不放）
    var res = filter(parts[0].ws, Object.assign(partOpts(parts[0], common), { lastRow: common.dataStartRow - 1 }));
    var ws = res.worksheet;
    var colCount = res.colCount;
    var match = rmaMatcher(rma);
    var skip = (rma && rma.skip) || {};
    // 每份報表留下來的列，分成一般的和 RMA
    var picked = parts.map(function (p) {
      var sel = selectRows(p.ws, partOpts(p, common));
      var normal = [], rmaRows = [];
      sel.rows.forEach(function (r) {
        var isRma = match && !skip[p.id + ':' + r] && match(rmaRaw(p.ws, r, p.keepCols, rma.col));
        (isRma ? rmaRows : normal).push(r);
      });
      return { total: sel.total, normal: normal, rma: rmaRows };
    });
    var info = parts.map(function (p, k) {
      return { matched: 0, rma: picked[k].rma.length, total: picked[k].total, firstRow: 0, lastRow: 0, cat: p.cat };
    });
    var sections = [];
    var cursor = res.headerEnd + 1;
    var prev = null;
    var matched = 0;
    var headCell = res.headerEnd >= 1 ? ws.findCell(res.headerEnd, 1) : null;
    var font = Object.assign({}, headCell && headCell.style && headCell.style.font, { bold: true });
    MPS_SECTIONS.forEach(function (sec) {
      var idx = [];
      parts.forEach(function (p, k) { if (p.cat === sec.key && picked[k].normal.length) idx.push(k); });
      if (!idx.length) return;                     // 這一類沒有列（沒有檔案，或全部都挑去 RMA）：整段不放
      if (prev) cursor += prev.group === sec.group ? SECTION_GAP : 1 + TOTALS_ROWS + GROUP_GAP;
      var titleRow = cursor;
      writeSectionTitle(ws, titleRow, colCount, sec.label, font, SECTION_FILL);
      var next = titleRow + 1;
      var s = { key: sec.key, label: sec.label, group: sec.group, titleRow: titleRow, firstRow: next, lastRow: titleRow, matched: 0 };
      idx.forEach(function (k) {
        var rows = picked[k].normal;
        var end = copyRows(ws, next, parts[k].ws, rows, parts[k].keepCols);
        info[k].matched = rows.length;
        info[k].firstRow = next;
        info[k].lastRow = end - 1;
        s.matched += rows.length;
        matched += rows.length;
        next = end;
      });
      s.lastRow = next - 1;
      sections.push(s);
      cursor = next;
      prev = sec;
    });
    // 【RMA】：最後一張匯總表下面空 3 列（沒有其他段時接在表頭下面）。沒有挑出 RMA 也一樣放標題
    if (rma) {
      if (prev) cursor += 1 + TOTALS_ROWS + RMA_GAP;
      var rTitle = cursor;
      writeSectionTitle(ws, rTitle, colCount, RMA_SECTION.label, font, RMA_FILL);
      var rNext = rTitle + 1;
      var rs = { key: RMA_SECTION.key, label: RMA_SECTION.label, group: RMA_SECTION.group, titleRow: rTitle, firstRow: rNext, lastRow: rTitle, matched: 0 };
      sectionOrder(parts).forEach(function (k) {
        var rows = picked[k].rma;
        if (!rows.length) return;
        rNext = copyRows(ws, rNext, parts[k].ws, rows, parts[k].keepCols);
        rs.matched += rows.length;
      });
      matched += rs.matched;
      rs.lastRow = rNext - 1;
      sections.push(rs);
      cursor = rNext;
    }
    res.matchedRows = matched;
    res.totalRows = picked.reduce(function (n, x) { return n + x.total; }, 0);
    res.lastRow = cursor - 1;
    res.parts = info;
    res.sections = sections;
    stripInnerThickLines(ws, sections, colCount);
    return res;
  }

  // 類別標題列：跨所有保留的行合併，粗體置中，底色 #E9C46A（RMA 是 #FFCAD4），外框細線
  function writeSectionTitle(ws, r, colCount, text, font, fill) {
    var THIN = { style: 'thin' };
    for (var c = 1; c <= Math.max(colCount, 1); c++) {
      var border = { top: THIN, bottom: THIN };
      if (c === 1) border.left = THIN;
      if (c === colCount) border.right = THIN;
      ws.getCell(r, c).style = {
        font: font,
        fill: fill || SECTION_FILL,
        alignment: { horizontal: 'center', vertical: 'middle' },
        border: border
      };
    }
    ws.getCell(r, 1).value = text;
    if (colCount > 1) ws.mergeCellsWithoutStyle(r, 1, r, colCount);
  }

  // 原報表帶過來的粗框線（例如每份原報表最後一列的黑粗線）拿掉，各段裡面只留細線
  function stripInnerThickLines(ws, sections, colCount) {
    sections.forEach(function (s) {
      for (var r = s.firstRow; r <= s.lastRow; r++) {
        for (var c = 1; c <= colCount; c++) {
          var cell = ws.findCell(r, c);
          if (!cell) continue;
          var border = Object.assign({}, cell.style && cell.style.border);
          var changed = false;
          if (isThickLine(border.bottom)) { delete border.bottom; changed = true; }
          if (isThickLine(border.top)) { delete border.top; changed = true; }
          if (changed) cell.style = Object.assign({}, cell.style, { border: border });
        }
      }
    });
  }

  function isThickLine(b) { return !!(b && (b.style === 'medium' || b.style === 'thick')); }

  /* ---------- 把整張工作表複製到另一個活頁簿 ---------- */

  /*
   * 把 src 工作表原樣複製進 dstWorkbook，取名為 name。
   * 值、格式、合併儲存格、行寬、列高、隱藏的行列、凍結窗格、自動篩選都會帶過去；
   * 公式換成計算結果（跨檔案之後原公式的參照會失效）。
   * dstWorkbook 裡如果已經有同名的工作表，會先移除再放新的。
   * 回傳 { worksheet, lastRow, colCount }
   */
  function copySheet(src, dstWorkbook, name) {
    // Excel 的工作表名稱不分大小寫，p003 和 P003 算同名
    dstWorkbook.worksheets
      .filter(function (w) { return String(w.name).toLowerCase() === String(name).toLowerCase(); })
      .forEach(function (w) { dstWorkbook.removeWorksheet(w.id); });

    var props = src.properties || {};
    var ws = dstWorkbook.addWorksheet(name, {
      properties: cleanObject({
        defaultRowHeight: props.defaultRowHeight,
        defaultColWidth: props.defaultColWidth
      })
    });

    var ext = extent(src);
    var maxCol = ext.colCount;

    eachStoredRow(src, function (row, r) {
      var nrow = null;
      eachStoredCell(row, function (cell, c) {
        if (!nrow) nrow = ws.getRow(r);
        var ncell = nrow.getCell(c);
        ncell.style = cell.style;
        if (c > maxCol) maxCol = c;
        if (cell.isMerged && cell.master !== cell) return;
        var v = exportValue(cell);
        if (v !== null && v !== '') ncell.value = v;
      });
      if (row.height || row.hidden) {
        if (!nrow) nrow = ws.getRow(r);
        if (row.height) nrow.height = row.height;
        if (row.hidden) nrow.hidden = true;
      }
    });

    for (var c = 1; c <= maxCol; c++) {
      var col = src.getColumn(c);
      if (col.width) ws.getColumn(c).width = col.width;
      if (col.hidden) ws.getColumn(c).hidden = true;
    }

    listMerges(src).forEach(function (m) {
      ws.mergeCellsWithoutStyle(m.top, m.left, m.bottom, m.right);
    });

    var v = (src.views && src.views[0]) || {};
    var view = cleanObject({ showGridLines: v.showGridLines, zoomScale: v.zoomScale });
    if (v.state === 'frozen') {
      view.state = 'frozen';
      view.xSplit = v.xSplit || 0;
      view.ySplit = v.ySplit || 0;
    }
    if (Object.keys(view).length) ws.views = [view];

    if (src.autoFilter) ws.autoFilter = src.autoFilter;

    var ps = src.pageSetup || {};
    Object.assign(ws.pageSetup, cleanObject({
      orientation: ps.orientation,
      paperSize: ps.paperSize,
      fitToPage: ps.fitToPage,
      fitToWidth: ps.fitToWidth,
      fitToHeight: ps.fitToHeight,
      printTitlesRow: ps.printTitlesRow
    }));

    return { worksheet: ws, lastRow: ext.lastRow, colCount: ext.colCount };
  }

  /* ---------- 表頭文字 ---------- */

  // 猜表頭在哪一列：前幾列當中，有值的格子最多的那一列（大標題通常只有一兩格有值）
  function guessHeaderRow(ws, maxRows) {
    var best = 1, bestCount = -1;
    for (var r = 1; r <= (maxRows || 10); r++) {
      var row = ws.findRow(r);
      if (!row) continue;
      var n = 0;
      eachStoredCell(row, function (cell) {
        if (cell.isMerged && cell.master !== cell) return;
        var v = cell.value;
        if (v !== null && v !== undefined && v !== '') n++;
      });
      if (n > bestCount) { bestCount = n; best = r; }
    }
    return best;
  }

  // 每一行的表頭文字（fromRow 到 toRow 之間，由上到下串起來；合併儲存格用左上角的字）
  // 回傳陣列，labels[行號] = '2027/01 月合計'
  function columnLabels(ws, fromRow, toRow, colCount) {
    var labels = [];
    for (var c = 1; c <= colCount; c++) {
      var parts = [];
      for (var r = fromRow; r <= toRow; r++) {
        var row = ws.findRow(r);
        var cell = row ? row.findCell(c) : null;
        var t = cell ? normalizeText(displayText(cell)).replace(/\s+/g, ' ') : '';
        if (t && parts.indexOf(t) < 0) parts.push(t);
      }
      labels[c] = parts.join(' ');
    }
    return labels;
  }

  /* ---------- 查表：用料號到另一張工作表找數值 ---------- */

  // 料號比對用的鍵：去頭尾空白、全形轉半形、不分大小寫。
  // 數字和文字的 123 視為相同，但 0012 和 12 不同（料號前面的 0 有意義）。
  function lookupKey(v) {
    v = plainValue(v);
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString();
    return normalizeText(v).toUpperCase();
  }

  // 把參照工作表的某一行做成索引：{ 鍵: [列號, 列號…] }（由上到下）
  function buildIndex(ref, keyCol) {
    var index = {};
    eachStoredRow(ref, function (row, r) {
      var k = lookupKey(cellRawValue(row.findCell(keyCol)));
      if (k === '') return;
      if (!index[k]) index[k] = [];
      index[k].push(r);
    });
    return index;
  }

  function sheetRef(name) {
    return "'" + String(name).replace(/'/g, "''") + "'!";
  }

  // 在表頭列裡找文字是 header 的那一行（重做這一步時，覆蓋原本那一行而不是再加一行）
  function findHeaderColumn(ws, headerEnd, header) {
    var want = lookupKey(header);
    var found = 0;
    for (var r = 1; r <= headerEnd && !found; r++) {
      var row = ws.findRow(r);
      if (!row) continue;
      eachStoredCell(row, function (cell, c) {
        if (found || (cell.isMerged && cell.master !== cell)) return;
        if (lookupKey(cell.value) === want) found = c;
      });
    }
    return found;
  }

  // 把儲存格的值當成數字：數字照用、文字的 "100" 也算；空白回傳 null；不是數字回傳 NaN
  function toNumber(v) {
    v = plainValue(v);
    if (typeof v === 'number') return v;
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return NaN;
    var t = normalizeText(v).replace(/,/g, '');
    if (t === '') return null;
    var n = Number(t);
    return isNaN(n) ? NaN : n;
  }

  // 取出 ref 第 r 列要帶回去的數值。有給 o.refDivCol 時是「取值行 ÷ 除數行」。
  // 回傳 { value, error, numFmt, top, bottom }；error 是 Excel 會顯示的錯誤（例如 #DIV/0!）
  function lookupValue(ref, r, o) {
    var row = ref.findRow(r);
    var topCell = row ? row.findCell(o.refValueCol) : null;
    if (!o.refDivCol) {
      return { value: plainValue(cellRawValue(topCell)), error: null, numFmt: topCell ? topCell.numFmt : undefined };
    }
    var bottomCell = row ? row.findCell(o.refDivCol) : null;
    var top = toNumber(cellRawValue(topCell));
    var bottom = toNumber(cellRawValue(bottomCell));
    var out = {
      value: null, error: null, numFmt: 'General',
      top: topCell ? displayText(topCell) : '', bottom: bottomCell ? displayText(bottomCell) : ''
    };
    if (top !== null && isNaN(top)) out.error = '#VALUE!';
    else if (bottom !== null && isNaN(bottom)) out.error = '#VALUE!';
    else if (bottom === null || bottom === 0) out.error = '#DIV/0!';
    else out.value = (top === null ? 0 : top) / bottom;
    return out;
  }

  function sameResult(a, b) {
    if (a.error || b.error) return a.error === b.error;
    if (typeof a.value === 'number' && typeof b.value === 'number') {
      return Math.abs(a.value - b.value) <= 1e-12 * Math.max(1, Math.abs(a.value), Math.abs(b.value));
    }
    return lookupKey(a.value) === lookupKey(b.value);
  }

  /*
   * 先算出查表結果，不動工作表（給畫面預覽、統計用）。
   *
   * o.dataStartRow  main 的資料從第幾列開始
   * o.keyCol        main 拿來對照的行（放料號的那一行）
   * o.refKeyCol     ref 要在哪一行找料號
   * o.refValueCol   ref 找到之後取哪一行的數值（例如 Labor）
   * o.refDivCol     （可省略）要除以 ref 的哪一行（例如 Base Qty）；有給的話帶回的是 取值行 ÷ 這一行
   * o.header        新增那一行的表頭文字，例如 'CT.'
   * o.index         （可省略）buildIndex(ref, o.refKeyCol) 的結果，重複計算時可以沿用
   *
   * 同一個料號在 ref 出現很多列時，取最上面那一列（和 VLOOKUP 一樣）。
   *
   * 回傳：
   *   targetCol   要填入的行（原本最後一行之後空一行；已經有同名表頭就用那一行）
   *   isNew       是不是新增的行
   *   rows        [{ row, key, label, refRow, value, error, top, bottom }]  refRow = 0 表示找不到
   *   found / missing / blank   找到、找不到、料號空白的列數
   *   errors      找到了、但算不出數值的列數（除數是 0 或空白、或不是數字）
   *   multi       找到的料號裡，有幾個在 ref 出現不只一列
   *   multiDiff   其中有幾個「各列算出來的數值不一樣」
   */
  function planLookup(main, ref, o) {
    var headerEnd = o.dataStartRow - 1;
    var ext = extent(main);
    var existing = findHeaderColumn(main, headerEnd, o.header);
    var lastCol = ext.colCount;
    var targetCol = existing || lastCol + 2;
    var index = o.index || buildIndex(ref, o.refKeyCol);

    var rows = [];
    var found = 0, missing = 0, blank = 0, errors = 0;
    var multiKeys = {}, diffKeys = {};
    // 只看資料列：類別標題、空列、匯總表都跳過（mpsLayout）
    var layout = mpsLayout(main, o);
    var lastDataRow = layout.count ? layout.dataRows[layout.dataRows.length - 1] : 0;
    for (var n = 0; n < layout.dataRows.length; n++) {
      var r = layout.dataRows[n];
      var row = main.findRow(r);
      var cell = row ? row.findCell(o.keyCol) : null;
      var key = lookupKey(cellRawValue(cell));
      if (key === '') { blank++; continue; }
      var hits = index[key];
      var item = { row: r, key: key, label: displayText(cell), refRow: 0, value: null, error: null };
      if (hits && hits.length) {
        found++;
        var first = lookupValue(ref, hits[0], o);
        item.refRow = hits[0];
        item.value = first.value;
        item.error = first.error;
        item.numFmt = first.numFmt;
        item.top = first.top;
        item.bottom = first.bottom;
        if (first.error) errors++;
        if (hits.length > 1 && !multiKeys[key]) {
          multiKeys[key] = true;
          for (var i = 1; i < hits.length; i++) {
            if (!sameResult(first, lookupValue(ref, hits[i], o))) { diffKeys[key] = true; break; }
          }
        }
        item.multi = hits.length;
        item.differs = !!diffKeys[key];
      } else {
        missing++;
      }
      rows.push(item);
    }

    return {
      targetCol: targetCol,
      isNew: !existing,
      lastCol: existing ? existing : lastCol,
      lastDataRow: lastDataRow,
      rows: rows,
      found: found,
      missing: missing,
      blank: blank,
      errors: errors,
      multi: Object.keys(multiKeys).length,
      multiDiff: Object.keys(diffKeys).length
    };
  }

  // 某一列要填入的公式（不含開頭的 =），例如 'P003'!M5 或 'P003'!M5/'P003'!K5
  function lookupFormula(refName, r, o) {
    var prefix = sheetRef(refName);
    var f = prefix + numToCol(o.refValueCol) + r;
    if (o.refDivCol) f += '/' + prefix + numToCol(o.refDivCol) + r;
    return f;
  }

  /*
   * 把查表結果寫進 main 工作表：每一列填入公式（='P003'!M5 或 ='P003'!M5/'P003'!K5），
   * 點儲存格就看得到數值是從 ref 的哪一格來的。找不到的料號留空白並標上底色。
   * 參數和 planLookup 一樣，回傳的也是同一份結果。
   */
  var MISSING_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFD9D9' } };
  function isMissingFill(fill) {
    return !!(fill && fill.fgColor && fill.fgColor.argb === MISSING_FILL.fgColor.argb);
  }

  function addLookupColumn(main, ref, o) {
    var plan = planLookup(main, ref, o);
    var headerEnd = o.dataStartRow - 1;
    var col = plan.targetCol;

    // 表頭：樣式照抄對照那一行（料號那一行）的表頭，大標題那一列不動
    if (headerEnd >= 1) {
      var top = headerEnd >= 2 ? 2 : 1;
      if (plan.isNew) {
        for (var r = top; r <= headerEnd; r++) {
          var src = main.findCell(r, o.keyCol);
          if (src) main.getCell(r, col).style = src.style;
        }
        if (headerEnd > top) main.mergeCellsWithoutStyle(top, col, headerEnd, col);
      }
      var headCell = main.getCell(top, col);
      (headCell.isMerged ? headCell.master : headCell).value = o.header;
    }

    // 資料：樣式照抄原本最後一行（重做時沿用這一行原本的樣式），數字格式改用 ref 那一格的
    var done = {};
    plan.rows.forEach(function (item) {
      var cell = main.getCell(item.row, col);
      var base = main.findCell(item.row, plan.isNew ? plan.lastCol : col);
      var style = Object.assign({}, base ? base.style : {});
      if (isMissingFill(style.fill)) delete style.fill;
      done[item.row] = true;
      if (item.refRow) {
        style.numFmt = item.numFmt || 'General';
        cell.style = style;
        cell.value = {
          formula: lookupFormula(ref.name, item.refRow, o),
          result: item.error ? { error: item.error }
            : (item.value === null || item.value === '' ? 0 : item.value)
        };
      } else {
        style.fill = MISSING_FILL;
        cell.style = style;
        cell.value = null;
      }
    });

    // 重做時，把這一行裡不再對得到的舊內容清掉（需求工時匯總在 CT. 右邊，不會清到）
    if (!plan.isNew) {
      eachStoredRow(main, function (row, r) {
        if (r < o.dataStartRow || r > plan.lastDataRow || done[r]) return;
        var old = row.findCell(col);
        if (old) old.value = null;
      });
    }

    if (plan.isNew) main.getColumn(col).width = 11;

    // 請 Excel 開檔時重新計算一次公式
    var wb = main.workbook;
    if (wb) wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });

    return plan;
  }

  /* ---------- 需求工時：各月數量 × CT. ---------- */

  // 儲存格的值當成數字用；公式結果是錯誤（#DIV/0! 等）時回傳 { error }
  function numberOrError(cell) {
    var raw = cellRawValue(cell);
    if (raw && typeof raw === 'object' && !(raw instanceof Date)) {
      var res = raw.result !== undefined ? raw.result : raw;
      if (res && typeof res === 'object' && res.error) return { error: res.error };
      if (raw.error) return { error: raw.error };
    }
    var n = toNumber(raw);
    if (n !== null && isNaN(n)) return { error: '#VALUE!' };
    // blank：這一格真的什麼都沒有（公式不算空白，就算它還沒有計算結果）
    var empty = raw === null || raw === undefined || raw === '';
    return { value: n === null ? 0 : n, blank: empty };
  }

  // 找出已經存在的區塊（表頭文字是 title 的那一格）：回傳 { col, width }，沒有就回傳 null
  function findBlock(ws, headerEnd, title) {
    var col = findHeaderColumn(ws, headerEnd, title);
    if (!col) return null;
    var width = 1;
    listMerges(ws).forEach(function (m) {
      if (m.left === col && m.top <= headerEnd && m.right > m.left) {
        var cell = ws.findCell(m.top, m.left);
        if (cell && lookupKey(cell.value) === lookupKey(title)) width = m.right - m.left + 1;
      }
    });
    return { col: col, width: width };
  }

  /*
   * 先算出需求工時，不動工作表（給畫面預覽、統計用）。
   *
   * o.dataStartRow  資料從第幾列開始
   * o.keyCol        料號在哪一行（只拿來顯示）
   * o.wipCol        當月 WIP 在哪一行
   * o.monthCols     各月數量在哪幾行（陣列，第一個是當月）
   * o.ctCol         CT. 在哪一行
   * o.title         大表頭文字，例如 '需求工時'
   * o.headers       各月的小表頭文字（使用者自己輸入的月份）
   *
   * 公式：當月 =(WIP+當月數量)*CT.，其餘各月 =數量*CT.
   *
   * 回傳：
   *   startCol    要填入的第一行（原本最後一行之後空一行；已經有同名大表頭就用原本的位置）
   *   isNew       是不是新增
   *   headers     各月的小表頭文字
   *   rows        [{ row, label, ct, noCt, values: [數字或 { error }], formulas: [...] }]
   *   noCt        沒有 CT. 的列數（CT. 是空白）
   *   errors      算不出來的列數（CT. 或數量是錯誤值、不是數字）
   */
  function planDemand(main, o) {
    var headerEnd = o.dataStartRow - 1;
    var ext = extent(main);
    var old = findBlock(main, headerEnd, o.title);
    var startCol = old ? old.col : ext.colCount + 2;

    var headers = o.monthCols.map(function (c, i) {
      return o.headers && o.headers[i] !== undefined ? normalizeText(o.headers[i]) : '';
    });

    var ctL = numToCol(o.ctCol), wipL = numToCol(o.wipCol);
    var rows = [];
    var noCt = 0, errors = 0;
    // 只算資料列：類別標題、空列、匯總表都跳過（mpsLayout）
    var dataRows = mpsLayout(main, o).dataRows;
    for (var n = 0; n < dataRows.length; n++) {
      var r = dataRows[n];
      var row = main.findRow(r);
      if (!row) continue;
      var keyCell = row.findCell(o.keyCol);
      var label = keyCell ? displayText(keyCell) : '';
      if (normalizeText(label) === '') continue;
      var ct = numberOrError(row.findCell(o.ctCol));
      var wip = numberOrError(row.findCell(o.wipCol));
      var item = { row: r, label: label, ct: ct.error ? ct.error : ct.value, noCt: !!ct.blank, values: [], formulas: [] };
      var bad = false;
      o.monthCols.forEach(function (c, i) {
        var qty = numberOrError(row.findCell(c));
        var qL = numToCol(c);
        var err = ct.error || qty.error || (i === 0 ? wip.error : null);
        if (i === 0) {
          item.formulas.push('(' + wipL + r + '+' + qL + r + ')*' + ctL + r);
          item.values.push(err ? { error: err } : (wip.value + qty.value) * ct.value);
        } else {
          item.formulas.push(qL + r + '*' + ctL + r);
          item.values.push(err ? { error: err } : qty.value * ct.value);
        }
        if (err) bad = true;
      });
      if (item.noCt) noCt++;
      if (bad) errors++;
      rows.push(item);
    }

    return {
      startCol: startCol,
      isNew: !old,
      oldWidth: old ? old.width : 0,
      headers: headers,
      rows: rows,
      noCt: noCt,
      errors: errors
    };
  }

  var INTEGER_FORMAT = '#,##0';

  function centered(style) {
    return Object.assign({}, style, {
      alignment: Object.assign({}, style && style.alignment, { horizontal: 'center', vertical: 'middle' })
    });
  }

  /*
   * 把需求工時寫進工作表。參數和 planDemand 一樣，回傳的也是同一份結果。
   *
   * 表頭：大表頭（o.title）放在第 2 列、跨所有月份置中；
   *       各月的小表頭放在第 3 列到表頭最後一列，每一行各自上下合併、置中。
   * 資料：每格都是公式，點儲存格看得到是哪幾格相乘。沒有 CT. 的列一樣寫公式（結果是 0），
   *       並標上和 CT. 一樣的淺紅色底，方便辨認。
   */
  function addDemandColumns(main, o) {
    var plan = planDemand(main, o);
    var headerEnd = o.dataStartRow - 1;
    var n = o.monthCols.length;
    var start = plan.startCol;
    var end = start + n - 1;
    var c, r;

    // 重做時，先把原本那一塊清掉（表頭的合併、內容、樣式）
    if (!plan.isNew) {
      var oldEnd = start + Math.max(plan.oldWidth, n) - 1;
      listMerges(main).forEach(function (m) {
        if (m.left >= start && m.right <= oldEnd) main.unMergeCells(m.top, m.left, m.bottom, m.right);
      });
      eachStoredRow(main, function (row) {
        for (var cc = start; cc <= oldEnd; cc++) {
          var cell = row.findCell(cc);
          if (cell) { cell.value = null; cell.style = {}; }
        }
      });
    }

    // 表頭的位置
    var titleRow = headerEnd >= 3 ? 2 : (headerEnd === 2 ? 1 : 0);
    var subTop = headerEnd >= 3 ? 3 : headerEnd;       // 小表頭從第幾列開始（到 headerEnd）

    // 小表頭：樣式照抄該月數量那一行的表頭
    if (headerEnd >= 1) {
      for (var i = 0; i < n; i++) {
        c = start + i;
        for (r = subTop; r <= headerEnd; r++) {
          var srcHead = main.findCell(r, o.monthCols[i]);
          main.getCell(r, c).style = centered(srcHead ? srcHead.style : {});
        }
        main.getCell(subTop, c).value = plan.headers[i];
        if (headerEnd > subTop) main.mergeCellsWithoutStyle(subTop, c, headerEnd, c);
      }
    }
    // 大表頭：跨所有月份置中
    if (titleRow) {
      for (c = start; c <= end; c++) {
        var srcTitle = main.findCell(titleRow, o.monthCols[c - start]);
        main.getCell(titleRow, c).style = centered(srcTitle ? srcTitle.style : {});
      }
      main.getCell(titleRow, start).value = o.title;
      if (end > start) main.mergeCellsWithoutStyle(titleRow, start, titleRow, end);
    }

    // 資料
    plan.rows.forEach(function (item) {
      for (var i = 0; i < n; i++) {
        var cell = main.getCell(item.row, start + i);
        var base = main.findCell(item.row, o.monthCols[i]);
        var style = Object.assign({}, base ? base.style : {});
        style.numFmt = INTEGER_FORMAT;                 // 只顯示整數（儲存格裡的數值仍保留小數）
        if (item.noCt) style.fill = MISSING_FILL;
        cell.style = style;
        cell.value = { formula: item.formulas[i], result: item.values[i] };
      }
    });

    for (c = start; c <= end; c++) main.getColumn(c).width = 12;

    var wb = main.workbook;
    if (wb) wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });

    return plan;
  }

  /* ---------- MPS 匯總：各月預計總產量、前三大、其他 ---------- */

  var TOTAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB5E48C' } };   // 淺綠色（#B5E48C）：Total 那一列的底色

  /*
   * 先算出 MPS 各月的匯總，不動工作表（給畫面預覽用）。
   *
   * o.dataStartRow  資料從第幾列開始
   * o.keyCol        料號在哪一行（用來判斷資料到第幾列）
   * o.wipCol        當月 WIP 在哪一行
   * o.monthCols     各月預計產量在哪幾行（陣列，第一個是當月）
   * o.statusCol     Status 在哪一行（量產、試產用這一行分）
   * o.groups        [{ key: 'M', label: '量產' }, { key: 'R', label: '試產' }]：Status 等於 key 的列算進這一組
   *
   * 當月（第一個月）的數量一律先加上 WIP。
   *
   * 回傳：
   *   found       有沒有資料
   *   firstRow / lastRow   資料的第一列、最後一列
   *   totalRow    Total 那一列；groupRows[k] 是各組那一列（Total 下面依序往下）
   *   rows        [totalRow, groupRows…]
   *   labelCol    文字標籤放在哪一行（第一個月左邊那一行）
   *   months      [{ col, total, groups: [各組合計], error }]
   *   counts      各組有幾列料號（陣列，和 groups 同順序）
   *   others      Status 不是任何一組的列 [{ row, label, status }]（不算進量產、試產，只算進 Total）
   */
  // Status 分組：M 算量產、R 算試產（不分大小寫）。第 2 步（產量匯總）和第 6 步（需求工時匯總）共用。
  var STATUS_GROUPS = [{ key: 'M', label: '量產' }, { key: 'R', label: '試產' }];

  function statusKey(v) {
    v = plainValue(v);
    if (v === null || v === undefined) return '';
    return normalizeText(v).toUpperCase();
  }

  // 每一列屬於哪一組：回傳 { groupOf: { 列號: 第幾組（-1 = 都不是） }, counts: [各組列數], others: [{ row, label, status }] }
  // rowList：要看的資料列（mpsLayout 的 dataRows）。沒有指定 Status 的行（o.statusCol）時，每一列都當作「都不是」。
  function statusGroupsOf(main, o, rowList, groups) {
    var groupOf = {};
    var counts = groups.map(function () { return 0; });
    var others = [];
    for (var n = 0; n < rowList.length; n++) {
      var r = rowList[n];
      var row = main.findRow(r);
      var keyCell = row ? row.findCell(o.keyCol) : null;
      if (!keyCell || normalizeText(displayText(keyCell)) === '') { groupOf[r] = -1; continue; }
      var statusCell = o.statusCol ? row.findCell(o.statusCol) : null;
      var sk = statusKey(cellRawValue(statusCell));
      var g = -1;
      groups.forEach(function (gr, k) { if (g < 0 && sk === statusKey(gr.key)) g = k; });
      groupOf[r] = g;
      if (g >= 0) counts[g]++;
      else others.push({ row: r, label: displayText(keyCell), status: statusCell ? displayText(statusCell) : '' });
    }
    return { groupOf: groupOf, counts: counts, others: others };
  }

  /* ---------- MPS 工作表的分段：取像、車用、光機系統（第 1 步排好，後面各步驟照這個找資料列） ---------- */

  // 第 1 步依類別分段排列：表頭下面每一段上面有一列跨欄的類別標題（底色 #E9C46A），底下是那一類的 MPS。
  // group：哪幾段一起匯總（取像＋車用一起算一張匯總表，光機系統單獨一張）
  var MPS_SECTIONS = [
    { key: 'image', label: '取像', group: 0 },
    { key: 'auto', label: '車用', group: 0 },
    { key: 'optics', label: '光機系統', group: 1 }
  ];
  var MPS_SECTION_GROUPS = [{ label: '取像＋車用' }, { label: '光機系統' }, { label: 'RMA' }];
  var SECTION_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9C46A' } };
  // RMA：第 1 步從各報表留下來的列裡挑出來的，放在最下面（最後一張匯總表下面空 3 列），標題底色 #FFCAD4。
  // 匯總表只有 Total 一列（不分量產、試產）。這個月沒有挑出 RMA 也一樣放標題和 Total（沒有項目）。
  var RMA_SECTION = { key: 'rma', label: 'RMA', group: 2 };
  var RMA_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCAD4' } };
  var RMA_GAP = 3;                              // 最後一張匯總表下面空幾列才是【RMA】
  var SECTION_GAP = 3;                          // 同一組的兩段之間空幾列（取像最後一列之後空 3 列才是【車用】）
  var GROUP_GAP = 6;                            // 一組的匯總表下面空幾列才是下一組（【光機系統】）
  var TOTALS_ROWS = 1 + STATUS_GROUPS.length;   // 匯總表幾列：Total、量產、試產

  // 這一列是不是類別標題：第 1 行的字剛好是「取像」「車用」「光機系統」「RMA」
  function sectionTitleOf(ws, r) {
    var cell = ws.findCell(r, 1);
    if (!cell) return null;
    var t = normalizeText(displayText(cell));
    for (var i = 0; i < MPS_SECTIONS.length; i++) if (MPS_SECTIONS[i].label === t) return MPS_SECTIONS[i];
    if (t === RMA_SECTION.label) return RMA_SECTION;
    return null;
  }

  /*
   * 找出 MPS 工作表的分段。
   *   o  { dataStartRow, keyCol }
   * 回傳 {
   *   sections  [{ key, label, group, titleRow, firstRow, lastRow, count }]
   *             每一段資料的範圍：類別標題下面、料號那一行有內容的第一列到最後一列
   *   groups    [{ label, isRma, sections, ranges: [[第一列, 最後一列]…], firstRow, lastRow, tableRow }]
   *             要一起匯總的段（取像＋車用、光機系統、RMA）；tableRow = 匯總表第一列（這一組最後一列資料之後空一列）
   *             RMA 沒有項目時 ranges 是空的，tableRow 是標題下面空一列
   *   legacy    true = 舊版的檔案（沒有類別標題）：整塊資料當作一段、一張匯總表
   *   dataRows  料號那一行有內容的資料列（列號，由上往下）；標題列、空列、匯總表都不算
   *   count     資料列數
   * }
   */
  function mpsLayout(ws, o) {
    var keyCol = o.keyCol || 1;
    var ext = extent(ws);
    var sections = [];
    var dataRows = [];
    var cur = null;
    var legacy = false;
    for (var r = o.dataStartRow; r <= ext.lastRow; r++) {
      var sec = sectionTitleOf(ws, r);
      if (sec) {
        cur = { key: sec.key, label: sec.label, group: sec.group, titleRow: r, firstRow: 0, lastRow: 0, count: 0 };
        sections.push(cur);
        continue;
      }
      var cell = ws.findCell(r, keyCol);
      if (!cell || normalizeText(displayText(cell)) === '') continue;
      if (!cur) {
        cur = { key: '', label: '', group: 0, titleRow: 0, firstRow: 0, lastRow: 0, count: 0 };
        sections.push(cur);
        legacy = true;
      }
      if (!cur.firstRow) cur.firstRow = r;
      cur.lastRow = r;
      cur.count++;
      dataRows.push(r);
    }
    var groups = [];
    sections.forEach(function (s) {
      var isRma = s.key === RMA_SECTION.key;
      if (!s.count && !isRma) return;
      var g = null;
      groups.forEach(function (x) { if (x.index === s.group) g = x; });
      if (!g) {
        g = { index: s.group, label: legacy ? '' : MPS_SECTION_GROUPS[s.group].label, isRma: isRma, sections: [], ranges: [], firstRow: s.firstRow, lastRow: 0 };
        groups.push(g);
      }
      g.sections.push(s);
      if (s.count) {
        g.ranges.push([s.firstRow, s.lastRow]);
        g.lastRow = s.lastRow;
      } else {
        g.lastRow = s.titleRow;                    // RMA 沒有項目：Total 在標題下面空一列
      }
    });
    groups.forEach(function (g) { g.tableRow = g.lastRow + 2; });
    return { sections: sections, groups: groups, legacy: legacy, dataRows: dataRows, count: dataRows.length };
  }

  // 某一組（取像＋車用、光機系統）的資料列
  function groupRowsOf(layout, g) {
    return layout.dataRows.filter(function (r) {
      return g.ranges.some(function (x) { return r >= x[0] && r <= x[1]; });
    });
  }
  // 公式裡的範圍：['D6:D40', 'D45:D80']；abs = 加 $（Status 那一行）
  function rangeRefs(col, ranges, abs) {
    var L = numToCol(col);
    return ranges.map(function (x) { return abs ? '$' + L + '$' + x[0] + ':$' + L + '$' + x[1] : L + x[0] + ':' + L + x[1]; });
  }
  function sumOf(list) { return 'SUM(' + list.join(',') + ')'; }
  function sumIfOf(statusList, crit, list) {
    return list.map(function (x, k) { return 'SUMIF(' + statusList[k] + ',' + crit + ',' + x + ')'; }).join('+');
  }
  function critOf(gr) { return '"' + String(gr.key).replace(/"/g, '""') + '"'; }

  // 匯總表旁邊註明是哪一組（文字標籤左邊那一行，跨匯總表的列合併）。那一行是料號的行時不放。
  function writeGroupNote(main, rows, col, text, keyCol, font) {
    if (!text || col < 1 || col === keyCol) return;
    var top = rows[0], bottom = rows[rows.length - 1];
    listMerges(main).forEach(function (m) {
      if (m.left <= col && m.right >= col && m.top <= bottom && m.bottom >= top) main.unMergeCells(m.top, m.left, m.bottom, m.right);
    });
    var THIN = { style: 'thin' }, THICK = { style: 'medium' };
    var f = Object.assign({}, font || {}, { bold: true });
    rows.forEach(function (r) {
      var cell = main.getCell(r, col);
      cell.value = null;
      cell.style = {
        font: f,
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
        border: { top: r === top ? THICK : THIN, bottom: r === bottom ? THICK : THIN, left: THICK, right: THICK }
      };
    });
    main.getCell(top, col).value = text;
    if (bottom > top) main.mergeCellsWithoutStyle(top, col, bottom, col);
    var column = main.getColumn(col);
    if (!column.width || column.width < 12) column.width = 12;   // 放得下「取像＋車用」
  }

  function planMpsTotals(main, o) {
    var layout = mpsLayout(main, o);
    if (!layout.count) return { found: false };
    var groups = o.groups || STATUS_GROUPS;
    var firstMonth = Math.min.apply(null, o.monthCols);
    var labelCol = firstMonth > 1 ? firstMonth - 1 : 0;
    var counts = groups.map(function () { return 0; });
    var others = [];
    var tables = layout.groups.map(function (g) {
      var tg = g.isRma ? [] : groups;                // RMA 只有 Total，不分量產、試產
      var sg = g.isRma ? { groupOf: {}, counts: [], others: [] } : statusGroupsOf(main, o, groupRowsOf(layout, g), groups);
      sg.counts.forEach(function (n, k) { counts[k] += n; });
      others = others.concat(sg.others);
      var months = o.monthCols.map(function (c, i) {
        var error = null;
        var total = 0;
        var sums = tg.map(function () { return 0; });
        g.ranges.forEach(function (x) {
          for (var r = x[0]; r <= x[1]; r++) {
            var row = main.findRow(r);
            var qty = numberOrError(row ? row.findCell(c) : null);
            var wip = i === 0 ? numberOrError(row ? row.findCell(o.wipCol) : null) : { value: 0 };
            if (qty.error || wip.error) { if (!error) error = qty.error || wip.error; continue; }
            var v = qty.value + wip.value;
            total += v;
            if (sg.groupOf[r] >= 0) sums[sg.groupOf[r]] += v;
          }
        });
        return { col: c, total: total, groups: sums, error: error };
      });
      var totalRow = g.tableRow;
      var groupRows = tg.map(function (gr, k) { return totalRow + 1 + k; });
      return {
        label: g.label,
        isRma: !!g.isRma,
        groups: tg,
        sections: g.sections,
        ranges: g.ranges,
        firstRow: g.firstRow,
        lastRow: g.lastRow,
        totalRow: totalRow,
        groupRows: groupRows,
        rows: [totalRow].concat(groupRows),
        labelCol: labelCol,
        months: months,
        counts: sg.counts,
        others: sg.others
      };
    });
    return {
      found: true,
      legacy: layout.legacy,
      sections: layout.sections,
      count: layout.count,
      tables: tables,
      labelCol: labelCol,
      counts: counts,
      others: others
    };
  }

  // 一張匯總表各月的公式（畫面上的說明也用這個，Excel 和畫面才會一致）
  // table = planMpsTotals(...).tables[k]；回傳 [Total 的公式, 各組的公式…]（不含開頭的 =）
  // 取像＋車用這張表有兩段範圍：=SUM(E6:E40,E45:E80)、=SUMIF($C$6:$C$40,"M",E6:E40)+SUMIF($C$45:$C$80,"M",E45:E80)
  function mpsTotalsFormulas(table, o, i) {
    var m = table.months[i];
    var month = rangeRefs(m.col, table.ranges);
    var wip = rangeRefs(o.wipCol, table.ranges);
    var stat = rangeRefs(o.statusCol, table.ranges, true);
    if (!table.ranges.length) return [''].concat(table.groups.map(function () { return ''; }));   // RMA 沒有項目：填 0
    var out = [i === 0 ? sumOf(wip) + '+' + sumOf(month) : sumOf(month)];
    table.groups.forEach(function (gr) {
      var crit = critOf(gr);
      out.push(i === 0 ? sumIfOf(stat, crit, wip) + '+' + sumIfOf(stat, crit, month) : sumIfOf(stat, crit, month));
    });
    return out;
  }

  /*
   * 把 MPS 匯總寫進工作表，每一組一張（取像＋車用、光機系統；舊版沒有類別標題的檔案只有一張），
   * 放在那一組最後一列資料之後空一列，對齊各月的行，每格都是公式。
   * 以 Status 在 C 行、WIP 在 D 行、當月在 E 行、資料在第 6 到 74 列為例（只有一段時）：
   *
   *            當月（要加 WIP）                                                    其他月份
   *   Total    =SUM(D6:D74)+SUM(E6:E74)                                           =SUM(F6:F74)
   *   量產     =SUMIF($C$6:$C$74,"M",D6:D74)+SUMIF($C$6:$C$74,"M",E6:E74)          =SUMIF($C$6:$C$74,"M",F6:F74)
   *   試產     =SUMIF($C$6:$C$74,"R",D6:D74)+SUMIF($C$6:$C$74,"R",E6:E74)          =SUMIF($C$6:$C$74,"R",F6:F74)
   *
   * 取像＋車用有兩段時，每個範圍都寫兩段（中間的標題列、空列不算進去），見 mpsTotalsFormulas()。
   * Status 不是 M 也不是 R 的列只會算進 Total（畫面上會提醒）。
   *
   * 外觀：文字標籤放在第一個月左邊那一行、置中；標籤加上各月圍成一塊，外框粗線、裡面細格線。
   *       Total 那一列整列是淺綠色底（#B5E48C）。標籤再左邊一行跨三列合併，註明是哪一組（粗體）。
   *
   * o.labels = { total }，各組的文字標籤用 o.groups[k].label
   * 其餘參數和 planMpsTotals 一樣，回傳的也是同一份結果。
   */
  function addMpsTotals(main, o) {
    var plan = planMpsTotals(main, o);
    if (!plan.found) return plan;
    var inMonth = {};
    o.monthCols.forEach(function (c) { inMonth[c] = true; });
    plan.tables.forEach(function (t) { writeMpsTotalsTable(main, t, o); });

    // 舊版會在各月加「前三大淺橘色」的條件式格式，重做時拿掉
    main.conditionalFormattings = (main.conditionalFormattings || []).filter(function (cf) {
      var mm = /^([A-Z]+)\d+:([A-Z]+)\d+$/.exec(String(cf.ref || ''));
      return !(mm && mm[1] === mm[2] && inMonth[colToNum(mm[1])]);
    });

    var wb = main.workbook;
    if (wb) wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });

    return plan;
  }

  function writeMpsTotalsTable(main, plan, o) {
    var rows = plan.rows;
    var texts = [o.labels.total].concat(plan.groups.map(function (gr) { return gr.label; }));
    var leftCol = plan.labelCol || Math.min.apply(null, o.monthCols);
    var rightCol = Math.max.apply(null, o.monthCols);
    var noteCol = plan.label && plan.labelCol > 1 && plan.labelCol - 1 !== o.keyCol ? plan.labelCol - 1 : 0;

    // 重做時，先把這幾列在這一塊裡原本的內容清掉
    rows.forEach(function (r) {
      var row = main.findRow(r);
      if (!row) return;
      eachStoredCell(row, function (cell, c) {
        if (c < (noteCol || leftCol) || c > rightCol) return;
        if (cell.isMerged) return;
        cell.value = null; cell.style = {};
      });
    });

    var THIN = { style: 'thin' }, THICK = { style: 'medium' };
    function boxBorder(r, c) {
      return {
        top: r === rows[0] ? THICK : THIN,
        bottom: r === rows[rows.length - 1] ? THICK : THIN,
        left: c === leftCol ? THICK : THIN,
        right: c === rightCol ? THICK : THIN
      };
    }
    var sample = plan.firstRow ? main.findCell(plan.firstRow, o.monthCols[0]) : null;
    var font = sample && sample.style && sample.style.font ? Object.assign({}, sample.style.font) : undefined;
    if (font) delete font.bold;

    plan.months.forEach(function (m, i) {
      var formulas = mpsTotalsFormulas(plan, o, i);
      var results = [m.total].concat(m.groups);
      rows.forEach(function (r, j) {
        var cell = main.getCell(r, m.col);
        var style = { numFmt: INTEGER_FORMAT, border: boxBorder(r, m.col), alignment: { vertical: 'middle' } };
        if (font) style.font = font;
        if (j === 0) style.fill = TOTAL_FILL;
        cell.style = style;
        cell.value = formulas[j] ? { formula: formulas[j], result: m.error ? { error: m.error } : results[j] } : 0;
      });
    });

    if (plan.labelCol) {
      rows.forEach(function (r, j) {
        var cell = main.getCell(r, plan.labelCol);
        var style = { border: boxBorder(r, plan.labelCol), alignment: { horizontal: 'center', vertical: 'middle' } };
        if (font) style.font = font;
        if (j === 0) style.fill = TOTAL_FILL;
        cell.style = style;
        cell.value = texts[j];
      });
      var labelColumn = main.getColumn(plan.labelCol);
      if (!labelColumn.width || labelColumn.width < 12) labelColumn.width = 12;
    }
    if (noteCol) writeGroupNote(main, rows, noteCol, plan.label, o.keyCol, font);
  }

  /* ---------- 需求工時匯總：每月總需求工時、量產、試產 ---------- */

  var HEAD_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFA2D2FF' } };   // 淺藍色（#A2D2FF）：Production 那一列的底色
  var LABEL_WIDTH = 26;               // 文字標籤那一行至少要多寬，才放得下「Production 總需求工時」

  /*
   * 先算出匯總結果，不動工作表（給畫面預覽用）。
   *
   * o.dataStartRow  資料從第幾列開始
   * o.keyCol        料號在哪一行（用來判斷哪些列是資料）
   * o.title         需求工時那一塊的大表頭文字，用來找出它在哪幾行
   * o.statusCol     Status 在哪一行（量產、試產用這一行分）；只是要找位置時可以省略
   * o.groups        分組（預設 M = 量產、R = 試產）
   *
   * 和第 2 步一樣每一組一張（取像＋車用、光機系統），放在同一列：那一組最後一列資料之後空一列，
   * 只是放在需求工時那幾行底下。
   *
   * 回傳：
   *   found       有沒有找到需求工時那一塊
   *   startCol / width   需求工時從哪一行開始、共幾行（幾個月）
   *   labelCol    匯總列的文字標籤放在哪一行（需求工時左邊那一行）
   *   count       資料列數（0 = 沒有資料）
   *   tables      [{ label, ranges, firstRow, lastRow, totalRow, groupRows, rows, months, counts, others }]
   *               months = [{ col, header, total, groups: [各組合計], error }]
   *   counts / others   全部加起來：各組有幾列、Status 都不是的列（和 planMpsTotals 一樣）
   */
  function planSummary(main, o) {
    var headerEnd = o.dataStartRow - 1;
    var block = findBlock(main, headerEnd, o.title);
    if (!block) return { found: false };
    var layout = mpsLayout(main, o);
    var base = { found: true, startCol: block.col, width: block.width, labelCol: block.col > 1 ? block.col - 1 : 0, count: layout.count, tables: [], counts: [], others: [] };
    if (!layout.count) return base;
    var groups = o.groups || STATUS_GROUPS;
    var subRow = headerEnd >= 3 ? 3 : headerEnd;
    var counts = groups.map(function () { return 0; });
    var others = [];
    base.tables = layout.groups.map(function (g) {
      var tg = g.isRma ? [] : groups;                // RMA 只有一列（RMA總需求工時）
      var sg = g.isRma ? { groupOf: {}, counts: [], others: [] } : statusGroupsOf(main, o, groupRowsOf(layout, g), groups);
      sg.counts.forEach(function (n, k) { counts[k] += n; });
      others = others.concat(sg.others);
      var months = [];
      for (var c = block.col; c < block.col + block.width; c++) {
        var error = null;
        var total = 0;
        var sums = tg.map(function () { return 0; });
        g.ranges.forEach(function (x) {
          for (var r = x[0]; r <= x[1]; r++) {
            var row = main.findRow(r);
            var v = numberOrError(row ? row.findCell(c) : null);
            if (v.error) { if (!error) error = v.error; continue; }
            total += v.value;
            if (sg.groupOf[r] >= 0) sums[sg.groupOf[r]] += v.value;
          }
        });
        var headCell = subRow >= 1 ? main.findCell(subRow, c) : null;
        months.push({ col: c, header: headCell ? displayText(headCell) : '', total: total, groups: sums, error: error });
      }
      var totalRow = g.tableRow;
      var groupRows = tg.map(function (gr, k) { return totalRow + 1 + k; });
      return {
        label: g.label,
        isRma: !!g.isRma,
        groups: tg,
        ranges: g.ranges,
        firstRow: g.firstRow,
        lastRow: g.lastRow,
        totalRow: totalRow,
        groupRows: groupRows,
        rows: [totalRow].concat(groupRows),
        months: months,
        counts: sg.counts,
        others: sg.others
      };
    });
    base.counts = counts;
    base.others = others;
    return base;
  }

  // 需求工時匯總一張表、一個月的公式（不含開頭的 =）：[Production, 各組…]
  function summaryFormulas(table, o, m) {
    if (!table.ranges.length) return [''];          // RMA 沒有項目：填 0
    var range = rangeRefs(m.col, table.ranges);
    var stat = rangeRefs(o.statusCol, table.ranges, true);
    return [sumOf(range)].concat(table.groups.map(function (gr) {
      return sumIfOf(stat, critOf(gr), range);
    }));
  }

  /*
   * 把匯總寫進工作表，每一組一張（和第 2 步的匯總表同一列），每格都是公式。
   * 以 Status 在 C 行、需求工時在 N 行、資料在第 6 到 74 列為例：
   *   Production 總需求工時   =SUM(N6:N74)
   *   量產總需求工時          =SUMIF($C$6:$C$74,"M",N6:N74)
   *   試產總需求工時          =SUMIF($C$6:$C$74,"R",N6:N74)
   * 取像＋車用有兩段時，每個範圍都寫兩段（中間的標題列、空列不算進去）。
   *
   * 外觀：文字標籤放在需求工時左邊那一行、置中；標籤加上各月圍成一塊，
   *       外框粗線、裡面細格線。Production 那一列整列是淺藍色底（#A2D2FF）。
   * 不再標示前三大（舊版加在需求工時上的淺橘色條件式格式，重做時會拿掉）。
   *
   * o.labels = { total, groups: [各組的標籤], rma: RMA 那張的標籤（預設「RMA總需求工時」） }
   * 其餘參數和 planSummary 一樣，回傳的也是同一份結果。
   */
  function addSummaryRows(main, o) {
    var plan = planSummary(main, o);
    if (!plan.found || !plan.count) return plan;
    var endCol = plan.startCol + plan.width - 1;
    plan.tables.forEach(function (t) { writeSummaryTable(main, plan, t, o); });
    if (plan.labelCol) {
      var labelColumn = main.getColumn(plan.labelCol);
      if (!labelColumn.width || labelColumn.width < LABEL_WIDTH) labelColumn.width = LABEL_WIDTH;
    }

    // 舊版在需求工時上加的「前三大淺橘色」條件式格式，重做時拿掉
    main.conditionalFormattings = (main.conditionalFormattings || []).filter(function (cf) {
      var mm = /^([A-Z]+)\d+:([A-Z]+)\d+$/.exec(String(cf.ref || ''));
      return !(mm && colToNum(mm[1]) >= plan.startCol && colToNum(mm[2]) <= endCol);
    });

    var wb = main.workbook;
    if (wb) wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });

    return plan;
  }

  function writeSummaryTable(main, plan, table, o) {
    var endCol = plan.startCol + plan.width - 1;
    var leftCol = plan.labelCol || plan.startCol;       // 整塊最左邊那一行
    var rows = table.rows;
    var texts = table.isRma ? [o.labels.rma || 'RMA總需求工時'] : [o.labels.total].concat(o.labels.groups);

    // 重做時，先把這幾列原本的內容清掉。只清這一塊的行，同一列左邊有 MPS 的匯總。
    rows.forEach(function (r) {
      var row = main.findRow(r);
      if (!row) return;
      eachStoredCell(row, function (cell, c) {
        if (c < leftCol || c > endCol) return;
        cell.value = null; cell.style = {};
      });
    });

    // 框線：最外圍粗線，裡面細線
    var THIN = { style: 'thin' }, THICK = { style: 'medium' };
    function boxBorder(r, c) {
      return {
        top: r === rows[0] ? THICK : THIN,
        bottom: r === rows[rows.length - 1] ? THICK : THIN,
        left: c === leftCol ? THICK : THIN,
        right: c === endCol ? THICK : THIN
      };
    }
    // 字型跟資料列一樣（不加粗）
    var sample = main.findCell(table.lastRow, plan.startCol);
    var font = sample && sample.style && sample.style.font ? Object.assign({}, sample.style.font) : undefined;
    if (font) delete font.bold;

    table.months.forEach(function (m) {
      var formulas = summaryFormulas(table, o, m);
      var results = [m.total].concat(m.groups);
      rows.forEach(function (r, i) {
        var cell = main.getCell(r, m.col);
        var style = { numFmt: INTEGER_FORMAT, border: boxBorder(r, m.col), alignment: { vertical: 'middle' } };
        if (font) style.font = font;
        if (i === 0) style.fill = HEAD_FILL;                   // Production（RMA 那張是 RMA總需求工時）那一列
        cell.style = style;
        cell.value = formulas[i] ? { formula: formulas[i], result: m.error ? { error: m.error } : results[i] } : 0;
      });
    });

    // 文字標籤：需求工時左邊那一行，置中
    if (plan.labelCol) {
      rows.forEach(function (r, i) {
        var cell = main.getCell(r, plan.labelCol);
        var style = { border: boxBorder(r, plan.labelCol), alignment: { horizontal: 'center', vertical: 'middle' } };
        if (font) style.font = font;
        if (i === 0) style.fill = HEAD_FILL;
        cell.style = style;
        cell.value = texts[i];
      });
    }
  }

  /* ---------- 人力報表：各項目在哪一列、當月在哪一行 ---------- */

  // 之後總表要抓的六個項目。key 是程式裡用的名字，label 是畫面上顯示的文字。
  var MANPOWER_ITEMS = [
    { key: 'online', label: 'Online' },
    { key: 'onlineProd', label: 'Online（生產）' },
    { key: 'onlineRework', label: 'Online（重工）' },
    { key: 'offDirect', label: 'Off line（直接）' },
    { key: 'offIndirect', label: 'Off line（間接）' },
    { key: 'ttl', label: 'TTL（On + Off直接）' }
  ];

  // 比對列標題用：全形轉半形、轉小寫、去掉空白和連字號（Off line、Off-line、OFFLINE 都一樣）
  function labelKey(text) {
    return normalizeText(text).toLowerCase().replace(/[\s\-_]/g, '');
  }

  /*
   * 列出人力報表每一列的文字，給使用者選「這個項目在第幾列」。
   * 回傳 [{ row, text, key }]，只列有內容的列：
   *   text  這一列的文字（數字不算；合併儲存格用左上角的字），最多取前三段
   *   key   比對用（labelKey 處理過的 text）
   */
  function manpowerRows(ws, ext) {
    var out = [];
    for (var r = 1; r <= ext.lastRow; r++) {
      var row = ws.findRow(r);
      if (!row) continue;
      var parts = [];
      var any = false;
      for (var c = 1; c <= ext.colCount; c++) {
        var cell = row.findCell(c);
        if (!cell) continue;
        var v = plainValue(cellRawValue(cell));
        if (v === null || v === undefined || v === '') continue;
        any = true;
        if (typeof v !== 'string') continue;
        var t = normalizeText(v).replace(/\s+/g, ' ');
        var asNumber = toNumber(t);
        if (t === '' || (asNumber !== null && !isNaN(asNumber))) continue;   // 文字型態的數字不算標題
        if (parts.indexOf(t) < 0 && parts.length < 3) parts.push(t);
      }
      if (!any) continue;
      var text = parts.join(' ');
      out.push({ row: r, text: text, key: labelKey(text) });
    }
    return out;
  }

  // 依列標題猜六個項目各在哪一列。回傳 { online: 5, onlineProd: 6, … }，猜不到的是 0。
  function guessManpowerRows(rows) {
    function has(k, word) { return k.indexOf(word) >= 0; }
    function isTotal(k) { return has(k, 'ttl') || has(k, 'total') || has(k, '+'); }
    var tests = {
      online: [
        function (k) { return k === 'online'; },
        function (k) { return has(k, 'online') && !has(k, '生產') && !has(k, '重工') && !has(k, 'off') && !isTotal(k); }
      ],
      onlineProd: [function (k) { return has(k, 'online') && has(k, '生產') && !has(k, 'off') && !isTotal(k); }],
      onlineRework: [function (k) { return has(k, 'online') && has(k, '重工') && !has(k, 'off') && !isTotal(k); }],
      offDirect: [function (k) { return has(k, 'offline') && has(k, '直接') && !has(k, '間接') && !isTotal(k); }],
      offIndirect: [function (k) { return has(k, 'offline') && has(k, '間接') && !isTotal(k); }],
      ttl: [
        function (k) { return (has(k, 'ttl') || has(k, 'total')) && has(k, 'on') && has(k, 'off'); },
        function (k) { return has(k, 'ttl'); }
      ]
    };
    var found = {};
    MANPOWER_ITEMS.forEach(function (item) {
      found[item.key] = 0;
      tests[item.key].some(function (test) {
        var hit = rows.filter(function (x) { return x.key !== '' && test(x.key); })[0];
        if (hit) found[item.key] = hit.row;
        return !!hit;
      });
    });
    return found;
  }

  // 第 r 列裡是數字的行（由左到右）
  function numericColumns(ws, r, colCount) {
    var cols = [];
    var row = ws.findRow(r);
    if (!row) return cols;
    for (var c = 1; c <= colCount; c++) {
      var n = toNumber(cellRawValue(row.findCell(c)));
      if (n !== null && !isNaN(n)) cols.push(c);
    }
    return cols;
  }

  // 猜當月在哪一行：把選好的各列裡是數字的行合在一起，剛好是 months 個連續的行時取第一個；
  // 其他情況（例如連過去的月份也列出來）不猜，回傳 0，讓使用者自己選。
  function guessFirstMonthCol(ws, rows, colCount, months) {
    var seen = {};
    rows.forEach(function (r) {
      numericColumns(ws, r, colCount).forEach(function (c) { seen[c] = true; });
    });
    var cols = Object.keys(seen).map(Number).sort(function (x, y) { return x - y; });
    if (cols.length !== months) return 0;
    return cols[cols.length - 1] - cols[0] === months - 1 ? cols[0] : 0;
  }

  // 每一行的表頭文字：從 aboveRow 的上一列往上找，第一個有內容的格子。回傳 labels[行號]
  function labelsAbove(ws, aboveRow, colCount) {
    var labels = [];
    for (var c = 1; c <= colCount; c++) {
      labels[c] = '';
      for (var r = aboveRow - 1; r >= 1; r--) {
        var row = ws.findRow(r);
        var cell = row ? row.findCell(c) : null;
        var t = cell ? normalizeText(displayText(cell)).replace(/\s+/g, ' ') : '';
        if (t !== '') { labels[c] = t; break; }
      }
    }
    return labels;
  }

  /*
   * 依使用者選的列和行，把之後總表要抓的數值讀出來給畫面核對。
   *   rows      { online: 5, … }
   *   firstCol  當月在哪一行
   *   months    幾個月
   * 回傳 { cols: [行號…], values: { online: [{ text, number, blank, bad }…], … }, blank, bad }
   */
  function readManpower(ws, rows, firstCol, months) {
    var out = { cols: [], values: {}, blank: 0, bad: 0 };
    for (var i = 0; i < months; i++) out.cols.push(firstCol + i);
    MANPOWER_ITEMS.forEach(function (item) {
      var r = rows[item.key];
      var row = r ? ws.findRow(r) : null;
      out.values[item.key] = out.cols.map(function (c) {
        var cell = row ? row.findCell(c) : null;
        var n = toNumber(cellRawValue(cell));
        var info = { text: cell ? displayText(cell) : '', number: n, blank: n === null, bad: n !== null && isNaN(n) };
        if (info.blank) out.blank++;
        if (info.bad) out.bad++;
        return info;
      });
    });
    return out;
  }

  /* ---------- 總表：先建立表格的樣式和框線，內容之後再填 ---------- */

  var REPORT_SHEET = '總表';
  var REPORT = {
    firstCol: 2,           // B：項目名稱
    unitCol: 3,            // C：單位、子項目
    monthCol: 4,           // D：第一個月
    months: 6,             // D 到 I
    noteCol: 10,           // J：備註
    titleRow: 2,           // 第 2 列：大標題（例如「製二課 (車用/光機)」，B 到 J 合併）
    noteRow: 4,            // 第 4 列：法定加班上限
    headRow: 5,            // 第 5 列：月份表頭
    lastRow: 28
  };
  var REPORT_TITLE_DEFAULT = '製二課 (車用/光機)';
  var REPORT_FONT = '微軟正黑體';
  var OVERTIME_LIMIT = 46;     // 法定加班上限（小時/人/月）
  var WEEKDAY_OT_HOURS = 2;    // 平日每天加班 2 小時
  var DAY_HOURS = 8;           // 一天工作 8 小時
  var OFFLINE_RATE = 20;       // 只用在列名稱「Off line(20%)」
  // 需求人力 Off line 的人員種類（每個月人數一樣，畫面上輸入，加總填入；之後要增減種類改這裡）
  var REPORT_OFFLINE = ['領班', '物料（含材製損）', '文書', '包裝'];
  var YO_LIMIT_DEFAULT = 78;   // 第二張表：YO 加班上限（小時/人/月），畫面上可以改
  var REPORT_SECOND_OFFSET = 28;   // 第二張表在第一張下面空三列（第一張到第 28 列，第二張的紅色格子在第 32 列）
  // 第二張表直接參照第一張表的列（其他列是第二張表自己算）
  var REPORT_SECOND_COPY = ['mpsMass', 'mpsTrial', 'mpsRma', 'hoursMass', 'hoursTrial', 'hoursRework',
    'weekdays', 'workMinutes', 'overtimeMinutes', 'effMass', 'effTrial', 'effRework', 'attendance',
    'needNewRework', 'needOffline', 'haveOnline', 'haveOffline'];
  function isNum(v) { return typeof v === 'number' && !isNaN(v); }
  var REPORT_COLOR = {
    red: 'FFFF0000',       // 法定加班上限
    head: 'FF366092',      // 表頭深藍
    section: 'FFDCE6F1',   // MPS(PCS) 那三列
    subtotal: 'FFEBF1DE',  // TTL
    loading: 'FFFFCC99',   // Loading Rate
    blue: 'FF4472C4',      // 人力相關的字
    white: 'FFFFFFFF'
  };
  /*
   * 每一列：key 給之後填內容用；label 在 B，unit 在 C；merge = B、C 合併；
   * group = 這一組在 B 行的文字（B 行往下合併到這一組的最後一列）；
   * plain = B 行的字不加粗；blue = 藍色字；section / subtotal / loading = 底色；height = 列高
   */
  var REPORT_ROWS = [
    { row: 6, key: 'mpsMass', group: 'MPS(PCS)', unit: '量產', section: true },
    { row: 7, key: 'mpsTrial', unit: '試產', section: true },
    { row: 8, key: 'mpsRma', unit: 'RMA', section: true },
    { row: 9, key: 'hoursMass', label: '需求工時-量產', unit: 'min/月', plain: true },
    { row: 10, key: 'hoursTrial', label: '需求工時-試產', unit: 'min/月', plain: true },
    { row: 11, key: 'hoursRework', label: '需求工時-重工', unit: 'min/月', plain: true },
    { row: 12, key: 'weekdays', label: '平日工作天數', unit: '天/月', plain: true },
    { row: 13, key: 'holidays', label: '假日工作天數', unit: '天/月', plain: true },
    { row: 14, key: 'workMinutes', label: '工作時間', unit: '分/天', plain: true },
    { row: 15, key: 'overtimeMinutes', label: '加班時間', unit: '分/天', plain: true },
    { row: 16, key: 'effMass', label: '綜合效率% - 量產', merge: true },
    { row: 17, key: 'effTrial', label: '綜合效率% - 試產', merge: true },
    { row: 18, key: 'effRework', label: '綜合效率% - 重工', merge: true },
    { row: 19, key: 'attendance', label: '出勤率%', merge: true },
    { row: 20, key: 'needOnline', group: '需求人力', unit: 'On line', blue: true, plain: true },
    { row: 21, key: 'needNewRework', unit: '新機種+重工', blue: true, height: 36 },
    { row: 22, key: 'needOffline', unit: 'Off line(' + OFFLINE_RATE + '%)', blue: true },
    { row: 23, key: 'needTotal', unit: 'TTL', blue: true, subtotal: true },
    { row: 24, key: 'haveOnline', group: '現有人力', unit: 'On line', blue: true, plain: true },
    { row: 25, key: 'haveOffline', unit: 'Off line', blue: true },
    { row: 26, key: 'haveTotal', unit: 'TTL', blue: true, subtotal: true },
    { row: 27, key: 'gap', label: '人力需求差異(現有-需求)', merge: true, blue: true, thickTop: true },
    { row: 28, key: 'loading', label: 'Loading Rate%(含加班)', merge: true, blue: true, loading: true }
  ];

  /*
   * 總表各列要從 MPS 工作表哪一列抓數值：用列的文字標籤找。
   *   量產、試產               第 2 步匯總（當月已含 WIP）
   *   量產總需求工時、試產總需求工時   第 6 步
   *   RMA 的 Total              第 2 步 RMA 那張（只有 Total 一列；用【RMA】段找到是哪一列）
   *   RMA總需求工時             第 6 步 RMA 那張
   * 匯總表有兩張（取像＋車用、光機系統）時，兩張的同一列相加，例如 ='MPS'!E42+'MPS'!E95。
   */
  var REPORT_LINKS = [
    { key: 'mpsMass', label: '量產', step: 2 },
    { key: 'mpsTrial', label: '試產', step: 2 },
    { key: 'mpsRma', label: 'Total', step: 2, rma: true },
    { key: 'hoursMass', label: '量產總需求工時', step: 6 },
    { key: 'hoursTrial', label: '試產總需求工時', step: 6 },
    { key: 'hoursRework', label: 'RMA總需求工時', step: 6 }
  ];

  function labelMatchKey(v) {
    return normalizeText(v).replace(/\s+/g, ' ').toLowerCase();
  }

  /*
   * 在 MPS 工作表找出總表各列的來源。
   *   o.dataStartRow  資料從第幾列開始（標籤只在這一列以下找，避免找到表頭的字）
   *   o.months        幾個月（6）
   *   o.keyCol        料號在哪一行（找【RMA】段用，預設第 1 行）
   * 回傳 { section: { found, row, rows, labelCol, refs: [{ sheet, col, row, value, parts }] }, … }
   * 數值是標籤右邊的格子，由左往右取 months 個（空白略過，碰到文字就停）。
   * 同一個標籤出現好幾次（匯總表有兩張：取像＋車用、光機系統）時全部都抓，同一個月相加：
   *   refs[m].parts = 各張表那一格 [{ sheet, col, row, value }]，value = 合計（有一格還沒有數值就是 null）
   *   rows = 各張表那一列的列號；row = 第一張的列號
   */
  function findReportLinks(main, o) {
    var ext = extent(main);
    var out = {};
    var rmaGroup = mpsLayout(main, { dataStartRow: o.dataStartRow || 1, keyCol: o.keyCol || 1 }).groups
      .filter(function (g) { return g.isRma; })[0];
    REPORT_LINKS.forEach(function (link) {
      var want = labelMatchKey(link.label);
      var hits = [];
      eachStoredRow(main, function (row, r) {
        if (r < (o.dataStartRow || 1)) return;
        if (link.rma && (!rmaGroup || r !== rmaGroup.tableRow)) return;   // RMA 的 Total 只看【RMA】那張
        eachStoredCell(row, function (cell, c) {
          if (cell.isMerged && cell.master !== cell) return;
          var v = plainValue(cellRawValue(cell));
          if (typeof v === 'string' && labelMatchKey(v) === want) hits.push({ row: r, col: c });
        });
      });
      if (!hits.length) { out[link.key] = { found: false, label: link.label, step: link.step, refs: [] }; return; }
      var lists = hits.map(function (hit) { return rowRefs(main, ext, hit, o.months); });
      var n = Math.min.apply(null, lists.map(function (l) { return l.length; }));
      var refs = [];
      for (var m = 0; m < n; m++) {
        var parts = lists.map(function (l) { return l[m]; });
        var sum = 0, ok = true;
        parts.forEach(function (p) { if (p.value === null) ok = false; else sum += p.value; });
        refs.push({ sheet: main.name, col: parts[0].col, row: parts[0].row, value: ok ? sum : null, parts: parts });
      }
      out[link.key] = {
        found: true, label: link.label, step: link.step,
        row: hits[0].row, rows: hits.map(function (x) { return x.row; }), labelCol: hits[0].col, refs: refs
      };
    });
    return out;
  }

  // 標籤右邊的數值格子，由左往右取 months 個（空白略過，碰到文字就停）
  function rowRefs(main, ext, hit, months) {
    var refs = [];
    var row = main.findRow(hit.row);
    for (var c = hit.col + 1; c <= ext.colCount && refs.length < months; c++) {
      var cell = row.findCell(c);
      var raw = cell ? cellRawValue(cell) : null;
      var v = plainValue(raw);
      if (v === null || v === undefined || v === '') {
        var isFormula = raw && typeof raw === 'object' && (raw.formula !== undefined || raw.sharedFormula !== undefined);
        if (!isFormula) { if (refs.length) break; else continue; }
      }
      if (typeof v === 'string' && isNaN(toNumber(v))) break;
      var n = toNumber(v);
      refs.push({ sheet: main.name, col: c, row: hit.row, value: n === null || isNaN(n) ? null : n });
    }
    return refs;
  }

  function solid(argb) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb } }; }

  /*
   * 在活頁簿裡建立「總表」工作表（已經有的話先刪掉重建）。
   * o.title      第 2 列的大標題（預設「製二課 (車用/光機)」）
   * o.monthRefs  各月表頭要參照的儲存格 [{ sheet: 'MPS', col: 14, row: 3, text: 'Oct' }]；
   *              有給的話表頭填公式（例如 ='MPS'!N3），沒給的那幾格留空白
   * o.links      各列要抓的儲存格 { mpsMass: [{ sheet, col, row, value, parts }], hoursMass: […], … }（findReportLinks 的 refs）；
   *              有給的話那一列的 D 到 I 填公式（例如 ='MPS'!E76）；有 parts（兩張匯總表）時相加，例如 ='MPS'!E42+'MPS'!E95
   * o.inputs     使用者在畫面上輸入的數值：
   *                weekdays     [6 個月的平日工作天數]（null = 沒填）
   *                workMinutes  [6 個月的工作時間（分/天）]
   *                effMass / effTrial / effRework / attendance   [6 個月的 %]（例如 80 代表 80%，存成 0.8）
   *                otHours      每日加班幾小時
   *                otRate       加班出勤人數 %（例如 80 代表 80%）
   *                newModels    需求人力 新機種+重工 的新機種 [{ name, qty: [6 個月的人數] }]
   *                reworkPeople 需求人力 新機種+重工 的重工 [6 個月的人數]
   *                offline      需求人力 Off line 各種人員的人數，順序同 REPORT_OFFLINE（每個月一樣）
   * o.yoLimit    第二張表（第 32 列起）的 YO 加班上限，預設 78
   * 回傳 { worksheet, rows: { hoursMass: 9, … }, rows2, monthCol, months }
   */
  function addReportSheet(wb, o) {
    o = o || {};
    wb.worksheets
      .filter(function (w) { return sameSheetName(w.name, REPORT_SHEET); })
      .forEach(function (w) { wb.removeWorksheet(w.id); });
    var ws = wb.addWorksheet(REPORT_SHEET);
    var R = REPORT;
    var C = REPORT_COLOR;
    var lastMonth = R.monthCol + R.months - 1;
    var THIN = { style: 'thin' }, THICK = { style: 'medium' };
    var CENTER = { horizontal: 'center', vertical: 'middle' };
    var yoLimit = isNum(o.yoLimit) ? o.yoLimit : YO_LIMIT_DEFAULT;
    var OFF2 = REPORT_SECOND_OFFSET;                       // 第二張表往下移幾列
    var bottom = R.lastRow + OFF2;

    // 整張總表都用微軟正黑體（charset 136 = 繁體中文）
    function font(extra) { return Object.assign({ name: REPORT_FONT, size: 12, charset: 136 }, extra); }

    // 行寬；表格外的空白格子也設成微軟正黑體，之後在旁邊打字字型才一致
    ws.getColumn(1).width = 6;
    ws.getColumn(R.firstCol).width = 17;
    ws.getColumn(R.unitCol).width = 15;
    for (var mc = R.monthCol; mc <= lastMonth; mc++) ws.getColumn(mc).width = 10;
    ws.getColumn(R.noteCol).width = 48;
    for (var fc = 1; fc <= R.noteCol + 1; fc++) ws.getColumn(fc).font = font();
    for (var fr = 1; fr <= bottom + 1; fr++) {
      for (fc = 1; fc <= R.noteCol + 1; fc++) ws.getCell(fr, fc).font = font();
    }

    // 大標題：B 到 J 合併、置中、粗體，外框粗線
    var title = o.title === undefined || o.title === null || o.title === '' ? REPORT_TITLE_DEFAULT : String(o.title);
    ws.getRow(R.titleRow).height = 27;
    for (var tc = R.firstCol; tc <= R.noteCol; tc++) {
      ws.getCell(R.titleRow, tc).style = {
        font: font({ size: 16, bold: true }), alignment: CENTER,
        border: { top: THICK, bottom: THICK, left: tc === R.firstCol ? THICK : undefined, right: tc === R.noteCol ? THICK : undefined }
      };
    }
    ws.getCell(R.titleRow, R.firstCol).value = title;
    ws.mergeCellsWithoutStyle(R.titleRow, R.firstCol, R.titleRow, R.noteCol);

    /* ----- 畫一張表的外框、底色、標籤（off = 往下移幾列），回傳各列的列號 ----- */
    function drawFrame(off, label) {
      var noteRow = R.noteRow + off, headRow = R.headRow + off, lastRow = R.lastRow + off;
      ws.getRow(noteRow).height = 21;
      ws.getRow(headRow).height = 21;
      for (var hr = headRow + 1; hr <= lastRow; hr++) ws.getRow(hr).height = 18;
      var thickTop = {};
      REPORT_ROWS.forEach(function (d) { if (d.thickTop) thickTop[d.row + off] = true; });

      // 外框粗線、裡面細線；人力需求差異那一列上面也是粗線
      function border(r, c) {
        return {
          top: r === headRow || thickTop[r] ? THICK : THIN,
          bottom: r === headRow || r === lastRow || thickTop[r + 1] ? THICK : THIN,
          left: c === R.firstCol || r === headRow ? THICK : THIN,
          right: c === R.noteCol || r === headRow ? THICK : THIN
        };
      }

      // 加班上限（紅底白字置中，B、C 合併）
      ws.getCell(noteRow, R.firstCol).value = label;
      ws.getCell(noteRow, R.firstCol).style = {
        font: font({ bold: true, color: { argb: C.white } }), fill: solid(C.red), alignment: CENTER
      };
      ws.getCell(noteRow, R.unitCol).style = { font: font(), fill: solid(C.red) };
      ws.mergeCellsWithoutStyle(noteRow, R.firstCol, noteRow, R.unitCol);

      // 表頭（深藍底白字；B、C 合併，畫一條斜線）
      for (var c = R.firstCol; c <= R.noteCol; c++) {
        var b = border(headRow, c);
        if (c <= R.unitCol) b.diagonal = { up: false, down: true, style: 'medium', color: { argb: 'FF000000' } };
        ws.getCell(headRow, c).style = {
          font: font({ size: c === R.noteCol ? 14 : 12, bold: c !== R.noteCol, color: { argb: C.white } }),
          fill: solid(C.head), alignment: CENTER, border: b
        };
      }
      ws.mergeCellsWithoutStyle(headRow, R.firstCol, headRow, R.unitCol);
      ws.getCell(headRow, R.noteCol).value = '備註';

      // 各列
      var rows = { noteRow: noteRow, headRow: headRow };
      var groupTop = 0;
      REPORT_ROWS.forEach(function (d, k) {
        var r = d.row + off;
        rows[d.key] = r;
        if (d.height) ws.getRow(r).height = d.height;
        var color = d.blue ? { argb: C.blue } : undefined;
        for (var c = R.firstCol; c <= R.noteCol; c++) {
          var f = d.section ? solid(C.section) : d.loading ? solid(C.loading) : null;
          if (d.subtotal && c >= R.unitCol) f = solid(C.subtotal);
          var st = { border: border(r, c), alignment: CENTER, font: font({ color: color }) };
          if (c === R.firstCol) st.font = font({ bold: !d.plain, color: color });
          if (c === R.unitCol) st.font = font({ bold: true, color: color });
          if (c === R.noteCol) st.alignment = { horizontal: 'left', vertical: 'middle' };
          if (f) st.fill = f;
          ws.getCell(r, c).style = st;
        }
        if (d.label) ws.getCell(r, R.firstCol).value = d.label;
        if (d.unit) ws.getCell(r, R.unitCol).value = d.unit;
        if (d.merge) ws.mergeCellsWithoutStyle(r, R.firstCol, r, R.unitCol);

        // MPS(PCS)、需求人力、現有人力：B 行往下合併到這一組的最後一列
        if (d.group) {
          groupTop = r;
          var g = ws.getCell(r, R.firstCol);
          g.value = d.group;
          g.style = Object.assign({}, g.style, {
            font: font({ bold: !!d.section, color: color }),
            alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }
          });
        }
        var nextRow = REPORT_ROWS[k + 1];
        if (groupTop && (!nextRow || nextRow.group || !nextRow.unit || nextRow.label)) {
          if (r > groupTop) ws.mergeCellsWithoutStyle(groupTop, R.firstCol, r, R.firstCol);
          groupTop = 0;
        }
      });
      return rows;
    }

    function cellOf(rows, key, m) { return ws.getCell(rows[key], R.monthCol + m); }
    function addr(rows, key, m) { return numToCol(R.monthCol + m) + rows[key]; }
    function num(rows, key, m) {
      var v = plainValue(cellRawValue(cellOf(rows, key, m)));
      return typeof v === 'number' ? v : (v === null || v === undefined ? 0 : NaN);
    }
    function setFormula(cell, formula, result, fmt) {
      cell.value = isNum(result) && isFinite(result) ? { formula: formula, result: result } : { formula: formula };
      cell.numFmt = fmt;
    }
    function setNote(rows, key, text) {
      var c = ws.getCell(rows[key], R.noteCol);
      c.value = text;
    }

    /* ----- 由其他列算出來的列（兩張表都一樣，只有加班上限不同） ----- */
    function fillComputed(rows, limit) {
      for (var m = 0; m < R.months; m++) {
        var a = function (key) { return addr(rows, key, m); };
        var n = function (key) { return num(rows, key, m); };

        // 假日工作天數 =（加班上限 − 平日工作天數 × 每日加班 2 小時）÷ 每天 8 小時
        setFormula(cellOf(rows, 'holidays', m),
          '(' + limit + '-' + a('weekdays') + '*' + WEEKDAY_OT_HOURS + ')/' + DAY_HOURS,
          (limit - n('weekdays') * WEEKDAY_OT_HOURS) / DAY_HOURS, '0.00');

        var days = '(' + a('weekdays') + '+' + a('holidays') + ')';
        var mins = '(' + a('workMinutes') + '+' + a('overtimeMinutes') + ')';
        var d = n('weekdays') + n('holidays');
        var t = n('workMinutes') + n('overtimeMinutes');

        // 需求人力 On line =（量產工時 ÷ 量產效率 + 試產工時 ÷ 試產效率 + 重工工時 ÷ 重工效率）
        //                   ÷（平日+假日天數）÷（工作時間+加班時間）÷ 出勤率
        setFormula(cellOf(rows, 'needOnline', m),
          '(' + a('hoursMass') + '/' + a('effMass') + '+' + a('hoursTrial') + '/' + a('effTrial') + '+' +
          a('hoursRework') + '/' + a('effRework') + ')/' + days + '/' + mins + '/' + a('attendance'),
          (n('hoursMass') / n('effMass') + n('hoursTrial') / n('effTrial') + n('hoursRework') / n('effRework')) /
            d / t / n('attendance'), '0.0');

        // 新機種+重工、Off line 是畫面上輸入的人數（第一張表填好，第二張表參照第一張）

        // TTL
        setFormula(cellOf(rows, 'needTotal', m), a('needOnline') + '+' + a('needNewRework') + '+' + a('needOffline'),
          n('needOnline') + n('needNewRework') + n('needOffline'), '#,##0.0');
        setFormula(cellOf(rows, 'haveTotal', m), a('haveOnline') + '+' + a('haveOffline'),
          n('haveOnline') + n('haveOffline'), 'General');

        // 人力需求差異 = 現有 TTL − 需求 TTL；Loading Rate = 需求 TTL ÷ 現有 TTL
        var hv = n('haveTotal'), nv = n('needTotal');
        setFormula(cellOf(rows, 'gap', m), a('haveTotal') + '-' + a('needTotal'), hv - nv, '#,##0.0');
        setFormula(cellOf(rows, 'loading', m), a('needTotal') + '/' + a('haveTotal'), hv ? nv / hv : NaN, '0%');
      }
    }

    /* ===== 第一張表：法定加班上限 46 小時 ===== */
    var rows = drawFrame(0, '法定加班上限' + OVERTIME_LIMIT + 'hr/人');

    // 月份表頭：參照 MPS 工作表需求工時的月份
    var refs = o.monthRefs || [];
    for (var i = 0; i < R.months; i++) {
      var ref = refs[i];
      if (ref && ref.sheet && ref.col && ref.row) {
        ws.getCell(rows.headRow, R.monthCol + i).value = {
          formula: sheetRef(ref.sheet) + numToCol(ref.col) + ref.row,
          result: ref.text || ''
        };
      }
    }

    // 從其他工作表抓的列（MPS、人力）
    var links = o.links || {};
    Object.keys(links).forEach(function (key) {
      if (!rows[key]) return;
      (links[key] || []).slice(0, R.months).forEach(function (ref, i) {
        var list = ref.parts && ref.parts.length ? ref.parts : [ref];
        var f = { formula: list.map(function (p) { return sheetRef(p.sheet) + numToCol(p.col) + p.row; }).join('+') };
        if (ref.value !== null && ref.value !== undefined) f.result = ref.value;
        var cell = cellOf(rows, key, i);
        cell.value = f;
        cell.numFmt = ref.numFmt || INTEGER_FORMAT;
      });
    });

    // 使用者輸入的列
    var inp = o.inputs || {};
    for (var m = 0; m < R.months; m++) {
      var wd = inp.weekdays ? inp.weekdays[m] : null;
      if (isNum(wd)) {
        cellOf(rows, 'weekdays', m).value = wd;
        cellOf(rows, 'weekdays', m).numFmt = 'General';
      }
      var wm = inp.workMinutes ? inp.workMinutes[m] : null;
      if (isNum(wm)) {
        cellOf(rows, 'workMinutes', m).value = wm;
        cellOf(rows, 'workMinutes', m).numFmt = INTEGER_FORMAT;
      }
      // 綜合效率、出勤率：輸入 80 存成 0.8，格式顯示成 80%
      ['effMass', 'effTrial', 'effRework', 'attendance'].forEach(function (key) {
        var v = inp[key] ? inp[key][m] : null;
        if (!isNum(v)) return;
        var pc = cellOf(rows, key, m);
        pc.value = Math.round(v * 1e6) / 1e8;
        pc.numFmt = Math.round(v) === v ? '0%' : '0.0%';
      });
      // 加班時間（分/天）= 每日加班小時 × 60 × 加班出勤人數 %
      if (isNum(inp.otHours) && isNum(inp.otRate)) {
        setFormula(cellOf(rows, 'overtimeMinutes', m), inp.otHours + '*60*' + inp.otRate + '%',
          inp.otHours * 60 * inp.otRate / 100, 'General');
      }
    }
    if (isNum(inp.otHours) || isNum(inp.otRate)) {
      setNote(rows, 'overtimeMinutes', '每日加班 ' + (isNum(inp.otHours) ? inp.otHours : '?') +
        ' 小時 × 60 × 加班出勤人數 ' + (isNum(inp.otRate) ? inp.otRate : '?') + '%');
    }

    // 需求人力 新機種+重工：各新機種 + 重工的人數相加，例如 =2+3+1；備註寫是哪些機種
    var models = (inp.newModels || []).filter(function (md) { return md && (md.name || (md.qty || []).some(isNum)); });
    var rework = inp.reworkPeople || [];
    for (m = 0; m < R.months; m++) {
      var parts = models.map(function (md) { return md.qty ? md.qty[m] : null; }).concat([rework[m]]);
      if (!parts.some(isNum)) continue;
      var nums = parts.map(function (v) { return isNum(v) ? v : 0; });
      setFormula(cellOf(rows, 'needNewRework', m), nums.join('+'), nums.reduce(function (x, y) { return x + y; }, 0), 'General');
    }
    if (models.length || rework.some(isNum)) {
      var names = models.map(function (md, k) { return md.name || ('新機種' + (k + 1)); });
      setNote(rows, 'needNewRework', '各月人數 = ' + names.concat(['重工']).join(' + '));
    }

    // 需求人力 Off line：各種人員的人數相加（每個月一樣），例如 =1+2+1+3；備註寫各種人數
    var offCounts = inp.offline || [];
    if (offCounts.some(isNum)) {
      var offNums = REPORT_OFFLINE.map(function (c, k) { return isNum(offCounts[k]) ? offCounts[k] : 0; });
      var offTotal = offNums.reduce(function (x, y) { return x + y; }, 0);
      for (m = 0; m < R.months; m++) setFormula(cellOf(rows, 'needOffline', m), offNums.join('+'), offTotal, 'General');
      setNote(rows, 'needOffline', REPORT_OFFLINE.map(function (c, k) { return c + ' ' + offNums[k]; }).join('、'));
    }

    fillComputed(rows, OVERTIME_LIMIT);

    /* ===== 第二張表：YO 加班上限（預設 78 小時，可在畫面上改） =====
     * 和第一張表一樣，只有假日工作天數改用 YO 的加班上限；
     * 其他輸入、抓來的數值都參照第一張表（例如 =D12），上面改了下面會跟著變。 */
    var rows2 = drawFrame(OFF2, 'YO加班上限' + yoLimit + 'hr/人');
    for (i = 0; i < R.months; i++) {
      var head1 = ws.getCell(rows.headRow, R.monthCol + i);
      if (head1.value === null || head1.value === undefined) continue;
      ws.getCell(rows2.headRow, R.monthCol + i).value = {
        formula: numToCol(R.monthCol + i) + rows.headRow,
        result: plainValue(cellRawValue(head1)) || ''
      };
    }
    REPORT_SECOND_COPY.forEach(function (key) {
      for (var m = 0; m < R.months; m++) {
        var src = cellOf(rows, key, m);
        var v = plainValue(cellRawValue(src));
        if (v === null || v === undefined || v === '') continue;    // 上面空白，下面也留空白
        setFormula(cellOf(rows2, key, m), addr(rows, key, m), typeof v === 'number' ? v : NaN, src.numFmt || 'General');
      }
      var note = ws.getCell(rows[key], R.noteCol).value;
      if (note) setNote(rows2, key, note);
    });
    fillComputed(rows2, yoLimit);

    wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });
    return { worksheet: ws, rows: rows, rows2: rows2, monthCol: R.monthCol, months: R.months, yoLimit: yoLimit };
  }

  function sameSheetName(a, b) { return String(a).toLowerCase() === String(b).toLowerCase(); }

  function toBlob(workbook) {
    return workbook.xlsx.writeBuffer().then(function (buf) {
      return new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
    });
  }

  var ExcelIO = {
    VERSION: 14,            // 程式版本，顯示在畫面右上角；每次改版加 1（index.html 的 ?v= 也要改）
    colToNum: colToNum,
    numToCol: numToCol,
    normalizeText: normalizeText,
    parseColumns: parseColumns,
    formatColumns: formatColumns,
    plainValue: plainValue,
    matchKey: matchKey,
    fuzzyText: fuzzyText,
    fuzzyTerms: fuzzyTerms,
    fuzzyHit: fuzzyHit,
    fuzzyRows: fuzzyRows,
    displayText: displayText,
    load: load,
    fixFontFlags: fixFontFlags,
    inflateRaw: inflateRaw,
    listMerges: listMerges,
    inspect: inspect,
    extent: extent,
    filter: filter,
    filterSections: filterSections,
    rmaScan: rmaScan,
    RMA_SECTION: RMA_SECTION,
    MPS_SECTIONS: MPS_SECTIONS,
    MPS_SECTION_GROUPS: MPS_SECTION_GROUPS,
    mpsLayout: mpsLayout,
    copySheet: copySheet,
    guessHeaderRow: guessHeaderRow,
    columnLabels: columnLabels,
    lookupKey: lookupKey,
    buildIndex: buildIndex,
    planLookup: planLookup,
    lookupFormula: lookupFormula,
    addLookupColumn: addLookupColumn,
    findHeaderColumn: findHeaderColumn,
    planDemand: planDemand,
    addDemandColumns: addDemandColumns,
    planMpsTotals: planMpsTotals,
    addMpsTotals: addMpsTotals,
    mpsTotalsFormulas: mpsTotalsFormulas,
    planSummary: planSummary,
    addSummaryRows: addSummaryRows,
    summaryFormulas: summaryFormulas,
    STATUS_GROUPS: STATUS_GROUPS,
    MANPOWER_ITEMS: MANPOWER_ITEMS,
    manpowerRows: manpowerRows,
    guessManpowerRows: guessManpowerRows,
    guessFirstMonthCol: guessFirstMonthCol,
    labelsAbove: labelsAbove,
    readManpower: readManpower,
    REPORT_SHEET: REPORT_SHEET,
    REPORT: REPORT,
    REPORT_ROWS: REPORT_ROWS,
    addReportSheet: addReportSheet,
    REPORT_LINKS: REPORT_LINKS,
    REPORT_TITLE_DEFAULT: REPORT_TITLE_DEFAULT,
    OFFLINE_RATE: OFFLINE_RATE,
    REPORT_OFFLINE: REPORT_OFFLINE,
    YO_LIMIT_DEFAULT: YO_LIMIT_DEFAULT,
    REPORT_SECOND_OFFSET: REPORT_SECOND_OFFSET,
    findReportLinks: findReportLinks,
    toNumber: toNumber,
    toBlob: toBlob
  };

  global.ExcelIO = ExcelIO;
  if (typeof module !== 'undefined' && module.exports) module.exports = ExcelIO;
})(typeof window !== 'undefined' ? window : globalThis);
