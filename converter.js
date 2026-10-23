// ==================== MODULE: converter.js ====================
(function(global) {
  'use strict';

  // Bảng mã TCVN3 (ABC) -> Unicode
  const TCVN3_MAP = {
    0xA1: 'Ă', 0xA2: 'Â', 0xA3: 'Ê', 0xA4: 'Ô', 0xA5: 'Ơ', 0xA6: 'Ư', 0xA7: 'Đ',
    0xA8: 'ă', 0xA9: 'â', 0xAA: 'ê', 0xAB: 'ô', 0xAC: 'ơ', 0xAD: 'ư', 0xAE: 'đ',
    0xB5: 'à', 0xB6: 'ả', 0xB7: 'ã', 0xB8: 'á', 0xB9: 'ạ',
    0xBB: 'ằ', 0xBC: 'ẳ', 0xBD: 'ẵ', 0xBE: 'ắ', 0xC6: 'ặ',
    0xC7: 'ầ', 0xC8: 'ẩ', 0xC9: 'ẫ', 0xCA: 'ấ', 0xCB: 'ậ',
    0xCC: 'è', 0xCE: 'ẻ', 0xCF: 'ẽ', 0xD0: 'é', 0xD1: 'ẹ',
    0xD2: 'ề', 0xD3: 'ể', 0xD4: 'ễ', 0xD5: 'ế', 0xD6: 'ệ',
    0xD7: 'ì', 0xD8: 'ỉ', 0xDC: 'ĩ', 0xDD: 'í', 0xDE: 'ị',
    0xDF: 'ò', 0xE1: 'ỏ', 0xE2: 'õ', 0xE3: 'ó', 0xE4: 'ọ',
    0xE5: 'ồ', 0xE6: 'ổ', 0xE7: 'ỗ', 0xE8: 'ố', 0xE9: 'ộ',
    0xEA: 'ờ', 0xEB: 'ở', 0xEC: 'ỡ', 0xED: 'ớ', 0xEE: 'ợ',
    0xEF: 'ù', 0xF1: 'ủ', 0xF2: 'ũ', 0xF3: 'ú', 0xF4: 'ụ',
    0xF5: 'ừ', 0xF6: 'ử', 0xF7: 'ữ', 0xF8: 'ứ', 0xF9: 'ự',
    0xFA: 'ỳ', 0xFB: 'ỷ', 0xFC: 'ỹ', 0xFD: 'ý', 0xFE: 'ỵ'
  };

  // Bảng mã Windows-1252 đặc biệt (0x80 - 0x9F)
  const WIN1252_EXTRA = {
    0x80: 0x20AC, 0x82: 0x201A, 0x83: 0x0192, 0x84: 0x201E, 0x85: 0x2026,
    0x86: 0x2020, 0x87: 0x2021, 0x88: 0x02C6, 0x89: 0x2030, 0x8A: 0x0160,
    0x8B: 0x2039, 0x8C: 0x0152, 0x8E: 0x017D, 0x91: 0x2018, 0x92: 0x2019,
    0x93: 0x201C, 0x94: 0x201D, 0x95: 0x2022, 0x96: 0x2013, 0x97: 0x2014,
    0x98: 0x02DC, 0x99: 0x2122, 0x9A: 0x0161, 0x9B: 0x203A, 0x9C: 0x0153,
    0x9E: 0x017E, 0x9F: 0x0178
  };

  const UNICODE_TO_WIN1252 = {};
  for (let b in WIN1252_EXTRA) {
    UNICODE_TO_WIN1252[WIN1252_EXTRA[b]] = Number(b);
  }

  // Các chữ cái có dấu thanh (thường) trong TCVN3 để khôi phục hoa
  const LOWER_TONED_SET = new Set(
    'àảãáạằẳẵắặầẩẫấậèẻẽéẹềểễếệìỉĩíịòỏõóọồổỗốộờởỡớợùủũúụừửữứựỳỷỹýỵ'.split('')
  );

  // Kiểm tra chuỗi có chữ tiếng Việt Unicode chuẩn không
  function hasRealVietnameseUnicode(str) {
    return /[\u0102\u0103\u0110\u0111\u01A0\u01A1\u01AF\u01B0\u1EA0-\u1EF9]/.test(str);
  }

  // Kiểm tra chuỗi có dấu hiệu Mojibake của TCVN3 đọc theo Windows-1252 không
  function hasTcvn3MojibakeSignature(str) {
    return /[µ¶·¸¹»¼½¾ÆÇÈÉÊËÌÎÏÐÑÒÓÔÕÖ×ØÜÝÞßáâãäåæçèéêëìíîïñòóôõö÷øùúûüýþ¨©ª«¬­®]/.test(str);
  }

  // Khôi phục chữ hoa cho các từ viết hoa toàn bộ (có ít nhất 2 chữ cái không dấu thanh đều viết hoa)
  function restoreUppercaseWords(text) {
    if (!text) return '';
    return text.replace(/\p{L}+/gu, function(word) {
      let upperUntonedCount = 0;
      let lowerUntonedCount = 0;
      let hasTonedLower = false;
      for (const ch of word) {
        if (LOWER_TONED_SET.has(ch)) {
          hasTonedLower = true;
        } else if (ch === ch.toUpperCase() && ch !== ch.toLowerCase()) {
          upperUntonedCount++;
        } else if (ch === ch.toLowerCase() && ch !== ch.toUpperCase()) {
          lowerUntonedCount++;
        }
      }
      if (hasTonedLower && upperUntonedCount >= 2 && lowerUntonedCount === 0) {
        return word.toUpperCase();
      }
      return word;
    });
  }

  // Giải mã mảng byte từ CSV (UTF-8 hợp lệ thì giữ nguyên, nếu không giải mã TCVN3 + fallback Win-1252)
  function decodeBytesToText(uint8Array) {
    if (uint8Array.length >= 3 && uint8Array[0] === 0xEF && uint8Array[1] === 0xBB && uint8Array[2] === 0xBF) {
      return new TextDecoder('utf-8').decode(uint8Array.subarray(3));
    }
    try {
      const utf8Decoder = new TextDecoder('utf-8', { fatal: true });
      const decoded = utf8Decoder.decode(uint8Array);
      if (hasRealVietnameseUnicode(decoded) || !hasTcvn3MojibakeSignature(decoded)) {
        return decoded;
      }
    } catch (e) {
      // Không phải UTF-8 hợp lệ -> dùng TCVN3
    }

    let out = '';
    for (let i = 0; i < uint8Array.length; i++) {
      const b = uint8Array[i];
      if (TCVN3_MAP[b] !== undefined) {
        out += TCVN3_MAP[b];
      } else if (WIN1252_EXTRA[b] !== undefined) {
        out += String.fromCharCode(WIN1252_EXTRA[b]);
      } else {
        out += String.fromCharCode(b);
      }
    }
    return restoreUppercaseWords(out);
  }

  // Chuyển chuỗi Mojibake trong XLSX (Windows-1252 -> byte -> TCVN3)
  function decodeMojibakeString(str) {
    if (typeof str !== 'string' || !str) return str;
    let out = '';
    for (let i = 0; i < str.length; i++) {
      const code = str.charCodeAt(i);
      let byteVal = -1;
      if (code <= 0xFF) {
        byteVal = code;
      } else if (UNICODE_TO_WIN1252[code] !== undefined) {
        byteVal = UNICODE_TO_WIN1252[code];
      }
      if (byteVal !== -1 && TCVN3_MAP[byteVal] !== undefined) {
        out += TCVN3_MAP[byteVal];
      } else {
        out += str[i];
      }
    }
    return restoreUppercaseWords(out);
  }

  // Kiểm tra cả trang tính (mảng 2 chiều các dòng) xem có bị lỗi font TCVN3 không
  function fixSheetRowsEncoding(rows) {
    let allText = '';
    for (let r = 0; r < Math.min(rows.length, 60); r++) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c++) {
        if (row[c] != null) allText += String(row[c]) + ' ';
      }
    }
    if (hasRealVietnameseUnicode(allText)) {
      return rows;
    }
    if (!hasTcvn3MojibakeSignature(allText)) {
      return rows;
    }
    return rows.map(row => (row || []).map(cell => {
      if (cell == null) return '';
      return decodeMojibakeString(String(cell));
    }));
  }

  // Phân tích CSV hỗ trợ ngoặc kép, dấu phẩy hoặc chấm phẩy
  function parseCSVText(text) {
    const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    const sampleLines = lines.split('\n').slice(0, 5).join('\n');
    let commaCount = (sampleLines.match(/,/g) || []).length;
    let semiCount = (sampleLines.match(/;/g) || []).length;
    const delimiter = semiCount > commaCount ? ';' : ',';

    const rows = [];
    let currentRow = [];
    let currentVal = '';
    let inQuotes = false;

    for (let i = 0; i < lines.length; i++) {
      const ch = lines[i];
      const nextCh = lines[i + 1];
      if (inQuotes) {
        if (ch === '"' && nextCh === '"') {
          currentVal += '"';
          i++;
        } else if (ch === '"') {
          inQuotes = false;
        } else {
          currentVal += ch;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
        } else if (ch === delimiter) {
          currentRow.push(currentVal.trim());
          currentVal = '';
        } else if (ch === '\n') {
          currentRow.push(currentVal.trim());
          if (currentRow.some(c => c !== '')) {
            rows.push(currentRow);
          }
          currentRow = [];
          currentVal = '';
        } else {
          currentVal += ch;
        }
      }
    }
    if (currentVal !== '' || currentRow.length > 0) {
      currentRow.push(currentVal.trim());
      if (currentRow.some(c => c !== '')) {
        rows.push(currentRow);
      }
    }
    return rows;
  }

  // Trích xuất giá trị thuần từ ô ExcelJS (hỗ trợ cả công thức =ROW()-8 hoặc công thức cộng ship)
  function getCellValue(val, rowNumber, preferFormulaAddition = false) {
    if (val == null) return '';
    if (typeof val === 'object') {
      if (val.richText && Array.isArray(val.richText)) {
        return val.richText.map(t => t.text || '').join('').trim();
      }
      if (preferFormulaAddition && val.formula && /^\s*\d{3,4}(\s*\+\s*\d{3,4})+\s*$/.test(val.formula)) {
        return val.formula.trim();
      }
      if (val.result !== undefined && val.result !== null) {
        return String(val.result).trim();
      }
      if (val.formula && rowNumber) {
        const rowFormulaMatch = String(val.formula).match(/^ROW\(\)\s*-\s*(\d+)$/i);
        if (rowFormulaMatch) {
          return String(rowNumber - parseInt(rowFormulaMatch[1], 10));
        }
      }
      if (val.text !== undefined) return String(val.text ?? '').trim();
    }
    return String(val).trim();
  }

  // Chuẩn hóa chuỗi tiêu đề để so khớp
  function normHeader(s) {
    return String(s || '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[đĐ]/g, 'd')
      .toLowerCase()
      .replace(/[^a-z0-9_+]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // Nhận dạng và trích xuất dữ liệu từ một mảng dòng (sheet)
  function classifyAndExtractSheet(rawRows, sheetName) {
    const rows = fixSheetRowsEncoding(rawRows);
    if (!rows || rows.length === 0) return null;

    // 1. Quét 10 dòng đầu tìm tiêu đề cột
    const maxScan = Math.min(rows.length, 10);
    for (let r = 0; r < maxScan; r++) {
      const row = rows[r] || [];
      const normed = row.map(normHeader);

      // Kiểm tra file dg1 đã chuẩn hóa: Code | Tên khách hàng | Điểm giao hàng | Ship
      const idxCodeClean = normed.findIndex(h => h === 'code' || h === 'ma don' || h === 'so');
      const idxCustClean = normed.findIndex(h => h.includes('ten khach') || h === 'khach hang');
      const idxAddrClean = normed.findIndex(h => h.includes('diem giao') || h.includes('dia chi'));
      const idxShipClean = normed.findIndex(h => h === 'ship' || h.includes('ma ship'));

      if (idxCodeClean !== -1 && idxCustClean !== -1 && idxAddrClean !== -1) {
        const records = [];
        for (let i = r + 1; i < rows.length; i++) {
          const d = rows[i] || [];
          const code = String(d[idxCodeClean] || '').trim();
          const customer = String(d[idxCustClean] || '').trim();
          const address = String(d[idxAddrClean] || '').trim();
          const ship = idxShipClean !== -1 ? String(d[idxShipClean] || '').trim() : '';
          if (code) records.push({ code, customer, address, ship });
        }
        return { type: 'dg1', format: 'clean', sheetName, records };
      }

      // Kiểm tra file giao hàng gốc: rDoc_number, rCv_name, rShip_to, rShip_doc_no
      const idxRDoc = normed.findIndex(h => h.includes('rdoc_number') || h === 'rdoc number');
      const idxRCv = normed.findIndex(h => h.includes('rcv_name') || h === 'rcv name');
      const idxRShipTo = normed.findIndex(h => h.includes('rship_to') || h === 'rship to');
      const idxRShipDoc = normed.findIndex(h => h.includes('rship_doc_no') || h === 'rship doc no');

      if (idxRDoc !== -1 && idxRCv !== -1 && idxRShipTo !== -1) {
        const records = [];
        for (let i = r + 1; i < rows.length; i++) {
          const d = rows[i] || [];
          const code = String(d[idxRDoc] || '').trim();
          const customer = String(d[idxRCv] || '').trim();
          const address = String(d[idxRShipTo] || '').trim();
          const ship = idxRShipDoc !== -1 ? String(d[idxRShipDoc] || '').trim() : '';
          if (code) records.push({ code, customer, address, ship });
        }
        return { type: 'dg1', format: 'raw_header', sheetName, records };
      }

      // Kiểm tra file sản phẩm (50/54): chuẩn hóa (Code | Sản phẩm) hoặc gốc (SO | Ten sp)
      const idxSO = normed.findIndex(h => h === 'code' || h === 'so' || h === 'ma so' || h === 'so_no');
      const idxProd = normed.findIndex(h => h === 'san pham' || h === 'ten sp' || h === 'ten san pham' || h === 'tensp');

      if (idxSO !== -1 && idxProd !== -1) {
        const records = [];
        for (let i = r + 1; i < rows.length; i++) {
          const d = rows[i] || [];
          const code = String(d[idxSO] || '').trim();
          const product = String(d[idxProd] || '').trim();
          if (code && product) records.push({ code, product });
        }
        return { type: 'product', format: 'header', sheetName, records };
      }
    }

    // 2. Tự dò bản xuất mới của file 50 KHÔNG có dòng tiêu đề (khoảng 45 cột)
    let startRow = 0;
    if (rows.length > 1) {
      const nonEmptyFirst = (rows[0] || []).filter(c => String(c || '').trim() !== '').length;
      const nonEmptySecond = (rows[1] || []).filter(c => String(c || '').trim() !== '').length;
      if (nonEmptyFirst <= 3 && nonEmptySecond >= 5) {
        startRow = 1;
      }
    }

    const dataRows = rows.slice(startRow).filter(r => (r || []).some(c => String(c || '').trim() !== ''));
    if (dataRows.length === 0) return null;

    let maxCols = 0;
    dataRows.forEach(r => { if (r.length > maxCols) maxCols = r.length; });

    if (maxCols >= 5) {
      const specRegex = /[a-zA-ZÀ-ỹ].*(\b\d+(?:[.,]\d+)?\s*(?:kg|g|gr|gam|ml|l)\b|\b(?:khay|tui|goi|bo|hop|trai|cu|bap)\b)/i;
      let bestProdCol = -1;
      let maxProdScore = 0;

      for (let c = 0; c < maxCols; c++) {
        let score = 0;
        for (let r = 0; r < dataRows.length; r++) {
          const val = String(dataRows[r][c] || '').trim();
          if (specRegex.test(val)) {
            score++;
          }
        }
        if (score > maxProdScore) {
          maxProdScore = score;
          bestProdCol = c;
        }
      }

      const codePattern = /^[A-Z0-9]{7,16}$/i;
      const soPatternStrict = /^\d{4}[A-Z]\d{5,8}$/i;
      let candidateCols = [];

      for (let c = 0; c < maxCols; c++) {
        if (c === bestProdCol || c === bestProdCol - 1) continue;
        let matchGeneral = 0;
        let matchStrict = 0;
        let nonEmpty = 0;
        for (let r = 0; r < dataRows.length; r++) {
          const val = String(dataRows[r][c] || '').trim();
          if (!val) continue;
          nonEmpty++;
          if (codePattern.test(val) && /\d/.test(val) && /[A-Z]/i.test(val)) {
            matchGeneral++;
          }
          if (soPatternStrict.test(val)) {
            matchStrict++;
          }
        }
        if (nonEmpty > 0 && (matchGeneral / nonEmpty) >= 0.9) {
          candidateCols.push({ col: c, strictRatio: matchStrict / nonEmpty, generalRatio: matchGeneral / nonEmpty });
        }
      }

      if (candidateCols.length > 0 && bestProdCol !== -1) {
        candidateCols.sort((a, b) => b.strictRatio - a.strictRatio || a.col - b.col);
        const soCol = candidateCols[0].col;
        const records = [];
        for (let r = 0; r < dataRows.length; r++) {
          const code = String(dataRows[r][soCol] || '').trim();
          const product = String(dataRows[r][bestProdCol] || '').trim();
          if (code && product) {
            records.push({ code, product });
          }
        }
        if (records.length > 0) {
          return { type: 'product', format: 'headerless_45col', sheetName, records };
        }
      }
    }

    return null;
  }

  // Đọc một File (CSV hoặc XLSX) -> trả về kết quả phân tích
  async function parseUploadedFile(file) {
    const buffer = await file.arrayBuffer();
    const uint8 = new Uint8Array(buffer);
    const fileName = file.name || 'data';
    const isZipXlsx = uint8.length >= 4 && uint8[0] === 0x50 && uint8[1] === 0x4B;

    if (!isZipXlsx) {
      const decodedText = decodeBytesToText(uint8);
      const rows = parseCSVText(decodedText);
      const res = classifyAndExtractSheet(rows, 'CSV');
      if (!res) {
        throw new Error(`Không nhận diện được cấu trúc file CSV: ${fileName}`);
      }
      return { fileName, ...res };
    } else {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer);
      let matchedResult = null;

      for (const ws of wb.worksheets) {
        const rows = [];
        ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          const rVals = [];
          for (let c = 1; c <= row.cellCount; c++) {
            rVals.push(getCellValue(row.getCell(c).value, rowNumber));
          }
          rows.push(rVals);
        });
        const res = classifyAndExtractSheet(rows, ws.name);
        if (res && res.records.length > 0) {
          if (res.format === 'header' || res.format === 'clean' || res.format === 'raw_header') {
            matchedResult = res;
            break;
          }
          if (!matchedResult) matchedResult = res;
        }
      }

      if (!matchedResult) {
        throw new Error(`Không tìm thấy trang tính hợp lệ trong file Excel: ${fileName}`);
      }
      return { fileName, ...matchedResult };
    }
  }

  // ==================== ĐỌC & PHÂN TÍCH FILE ĐIỀU XE ====================
  /**
   * Tách chuỗi Shipment (vd: "340", "351+353", "351 + 353", "306+311+362", "5024+ 028")
   * thành mảng các số Ship rời rạc. Nếu ô không có số Ship hợp lệ thì trả về mảng rỗng [].
   */
  function parseShipmentCellTokens(rawShipment) {
    if (rawShipment == null) return [];
    const str = String(rawShipment).trim();
    if (!str || str.toLowerCase() === 'nan') return [];
    const parts = str.split(/[+,;/]+/);
    const tokens = [];
    for (const p of parts) {
      const cleaned = p.trim();
      if (!cleaned) continue;
      // Lấy cụm chữ số (thường 3 hoặc 4 số) hoặc mã ship hợp lệ
      const digitMatch = cleaned.match(/\d+/g);
      if (digitMatch && digitMatch.length > 0) {
        digitMatch.forEach(d => {
          if (d && !tokens.includes(d)) tokens.push(d);
        });
      }
    }
    return tokens;
  }

  /**
   * Trích xuất danh sách chuyến xe có số Shipment từ 1 mảng dòng 2 chiều của Sheet Điều xe
   * (Hỗ trợ tiêu đề STT và Shipment nằm cùng dòng hoặc lệch dòng như dòng 7 & dòng 8)
   */
  function extractDispatchRowsFromMatrix(rawRows, sheetName) {
    const rows = fixSheetRowsEncoding(rawRows);
    if (!rows || rows.length === 0) return null;

    let colStt = -1, rowStt = -1;
    let colShipment = -1, rowShipment = -1;
    let colTruck = -1, colRoute = -1, colCustStart = -1, colCustEnd = -1;
    let dateLabel = '';

    const maxScan = Math.min(rows.length, 20);
    for (let r = 0; r < maxScan; r++) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c++) {
        const rawCell = String(row[c] || '').trim();
        if (!rawCell) continue;
        if (/ng[àa]y\s*:/i.test(rawCell) && !dateLabel) {
          dateLabel = rawCell;
        }
        const h = normHeader(rawCell);
        if (colStt === -1 && (h === 'stt' || h === 'so thu tu')) {
          colStt = c;
          rowStt = r;
        }
        // Chỉ lấy đúng cột "Shipment" hoặc "Ship" (không lấy cột "SHIP VIN+")
        if (colShipment === -1 && (h === 'shipment' || h === 'ship' || h === 'so ship')) {
          colShipment = c;
          rowShipment = r;
        }
        if (colTruck === -1 && (h.includes('bien so xe') || h === 'so xe')) {
          colTruck = c;
        }
        if (colRoute === -1 && (h === 'route' || h === 'tuyen')) {
          colRoute = c;
        }
        if (colCustStart === -1 && h.includes('ten khach hang')) {
          colCustStart = c;
        }
      }
    }

    if (colStt === -1 || colShipment === -1) {
      return null;
    }

    if (colCustStart !== -1) {
      colCustEnd = Math.max(colCustStart, colShipment - 3);
    }

    const dataStartRow = Math.max(rowStt, rowShipment) + 1;
    const records = [];

    for (let r = dataStartRow; r < rows.length; r++) {
      const row = rows[r] || [];

      // Dừng nếu gặp dòng TỔNG cuối bảng
      let isTotalRow = false;
      for (let c = 0; c < Math.min(row.length, 10); c++) {
        const cellNorm = normHeader(row[c]);
        if (cellNorm === 'tong' || cellNorm.startsWith('tong ')) {
          isTotalRow = true;
          break;
        }
      }
      if (isTotalRow) break;

      const sttVal = String(row[colStt] ?? '').trim();
      const shipRaw = String(row[colShipment] ?? '').trim();

      // Nếu ô không có dữ liệu số ship hoặc không có STT thì bỏ qua
      if (!shipRaw || !sttVal || normHeader(sttVal) === 'stt') continue;

      const ships = parseShipmentCellTokens(shipRaw);
      if (ships.length === 0) continue;

      const truck = colTruck !== -1 ? String(row[colTruck] || '').trim() : '';
      const route = colRoute !== -1 ? String(row[colRoute] || '').trim() : '';

      // Gom tóm tắt khách hàng trên dòng Điều xe để hiển thị đối chiếu
      const custParts = [];
      if (colCustStart !== -1 && colCustEnd >= colCustStart) {
        for (let c = colCustStart; c <= Math.min(colCustEnd, row.length - 1); c++) {
          const v = String(row[c] || '').replace(/\s+/g, ' ').trim();
          if (v && v.toLowerCase() !== 'nan') custParts.push(v);
        }
      }

      records.push({
        rowNumber: r + 1,
        stt: sttVal,
        shipmentRaw: shipRaw,
        ships: ships,
        truck: truck,
        route: route,
        dispatchCustomersSummary: custParts.slice(0, 4).join('; ') + (custParts.length > 4 ? ` (+${custParts.length - 4})` : '')
      });
    }

    return {
      sheetName,
      dateLabel,
      records
    };
  }

  /**
   * Đọc file Điều xe (.xlsx hoặc .csv), trả về danh sách tất cả các trang tính (sheets) hợp lệ
   */
  async function parseDispatchFile(file) {
    const buffer = await file.arrayBuffer();
    const uint8 = new Uint8Array(buffer);
    const fileName = file.name || 'Dieu_xe.xlsx';
    const isZipXlsx = uint8.length >= 4 && uint8[0] === 0x50 && uint8[1] === 0x4B;

    if (!isZipXlsx) {
      const decodedText = decodeBytesToText(uint8);
      const rows = parseCSVText(decodedText);
      const extracted = extractDispatchRowsFromMatrix(rows, 'CSV');
      if (!extracted) {
        throw new Error(`Không tìm thấy cột "STT" và "Shipment" trong file Điều xe: ${fileName}`);
      }
      return {
        fileName,
        sheets: [extracted]
      };
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const sheets = [];

    for (const ws of wb.worksheets) {
      const maxRow = ws.rowCount;
      const rows = [];
      for (let r = 1; r <= maxRow; r++) {
        const rowObj = ws.getRow(r);
        const rVals = [];
        const cellCount = Math.min(Math.max(rowObj.cellCount, 26), 40);
        for (let c = 1; c <= cellCount; c++) {
          rVals.push(getCellValue(rowObj.getCell(c).value, r, true));
        }
        rows.push(rVals);
      }
      const extracted = extractDispatchRowsFromMatrix(rows, ws.name);
      if (extracted) {
        sheets.push(extracted);
      }
    }

    if (sheets.length === 0) {
      throw new Error(`Không tìm thấy trang tính nào có cột "STT" và "Shipment" trong file: ${fileName}`);
    }

    return {
      fileName,
      sheets
    };
  }

  /**
   * So khớp danh sách số Ship của 1 dòng Điều xe (vd: ['340'] hoặc ['351', '353'])
   * với cột Ship trong dg1:
   * - Khớp chính xác (không phân biệt hoa/thường), HOẶC
   * - Khớp 3-4 chữ số cuối của mã Ship trong dg1 (vd: dg1 lưu '2650S00340' hoặc 'SHIP340' hoặc '340')
   */
  function matchShipTokensToDg1Codes(shipTokens, dg1Records) {
    const matchedCodesSet = new Set();
    const matchedShipsSet = new Set();

    if (!shipTokens || shipTokens.length === 0 || !dg1Records || dg1Records.length === 0) {
      return { matchedCodesSet, matchedShipsSet };
    }

    const normalizedTokens = shipTokens.map(t => String(t || '').trim()).filter(Boolean);

    for (const rec of dg1Records) {
      const dgShip = String(rec.ship || '').trim();
      if (!dgShip) continue;
      const dgShipLower = dgShip.toLowerCase();
      const dgTrailingDigitsMatch = dgShip.match(/(\d+)$/);
      const dgTrailingDigits = dgTrailingDigitsMatch ? dgTrailingDigitsMatch[1] : '';

      for (const tok of normalizedTokens) {
        const tokLower = tok.toLowerCase();
        const tokDigits = tok.replace(/\D/g, '');
        let isMatch = false;

        if (dgShipLower === tokLower) {
          isMatch = true;
        } else if (tokDigits && dgTrailingDigits) {
          // Khớp chính xác phần số hoặc khớp đuôi 3-4 chữ số cuối (đảm bảo ký tự ngay trước đuôi là số '0' hoặc không phải chữ số khác)
          if (dgTrailingDigits === tokDigits || dgTrailingDigits.replace(/^0+/, '') === tokDigits.replace(/^0+/, '')) {
            isMatch = true;
          } else if (tokDigits.length >= 3 && dgTrailingDigits.endsWith(tokDigits)) {
            isMatch = true;
          }
        } else if (tokDigits.length >= 3 && dgShip.endsWith(tokDigits)) {
          isMatch = true;
        }

        if (isMatch) {
          if (rec.code) matchedCodesSet.add(String(rec.code).trim());
          matchedShipsSet.add(dgShip);
          break;
        }
      }
    }

    return { matchedCodesSet, matchedShipsSet };
  }

  // Tạo Workbook chuẩn hóa (dg1 hoặc file 50) từ kết quả đã phân tích để tải về ở module "Chuyển file gốc"
  async function buildConvertedWorkbookBuffer(parsedResult) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(parsedResult.type === 'dg1' ? 'DG1' : 'SanPham');

    if (parsedResult.type === 'dg1') {
      ws.columns = [
        { header: 'Code', key: 'code', width: 18 },
        { header: 'Tên khách hàng', key: 'customer', width: 32 },
        { header: 'Điểm giao hàng', key: 'address', width: 42 },
        { header: 'Ship', key: 'ship', width: 18 }
      ];
      parsedResult.records.forEach(r => ws.addRow(r));
    } else {
      ws.columns = [
        { header: 'Code', key: 'code', width: 18 },
        { header: 'Sản phẩm', key: 'product', width: 45 }
      ];
      parsedResult.records.forEach(r => ws.addRow(r));
    }

    const headerRow = ws.getRow(1);
    headerRow.font = { name: 'Arial', size: 11, bold: true };
    headerRow.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
    });

    return await wb.xlsx.writeBuffer();
  }

  global.ConverterModule = {
    decodeBytesToText,
    decodeMojibakeString,
    restoreUppercaseWords,
    parseCSVText,
    parseUploadedFile,
    parseShipmentCellTokens,
    extractDispatchRowsFromMatrix,
    parseDispatchFile,
    matchShipTokensToDg1Codes,
    buildConvertedWorkbookBuffer
  };
})(window);
