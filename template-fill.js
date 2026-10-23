// ==================== MODULE: template-fill.js ====================
(function(global) {
  'use strict';

  function getPlainText(val) {
    if (val == null) return '';
    if (typeof val === 'object') {
      if (val.richText && Array.isArray(val.richText)) {
        return val.richText.map(t => t.text || '').join('').trim();
      }
      if (val.result !== undefined) return String(val.result ?? '').trim();
      if (val.text !== undefined) return String(val.text ?? '').trim();
    }
    return String(val).trim();
  }

  /**
   * Làm sạch giá trị ô khi chép từ mẫu sang trang tính mới:
   * - Tuyệt đối không chép nguyên xi đối tượng sharedFormula / formula (vì khi dịch dòng/nhân bản vùng sẽ làm hỏng tham chiếu XML trong Excel)
   * - Tạo bản sao độc lập cho richText hoặc Date
   */
  function sanitizeCellValue(val) {
    if (val == null) return null;
    if (typeof val !== 'object') return val;
    if (val instanceof Date) return new Date(val.getTime());
    if (val.richText && Array.isArray(val.richText)) {
      return { richText: JSON.parse(JSON.stringify(val.richText)) };
    }
    if (val.formula !== undefined || val.sharedFormula !== undefined) {
      return val.result !== undefined ? val.result : null;
    }
    if (val.hyperlink && val.text !== undefined) {
      return String(val.text);
    }
    return getPlainText(val);
  }

  function colLettersToNumber(letters) {
    let n = 0;
    const s = String(letters || '').toUpperCase();
    for (let i = 0; i < s.length; i++) {
      n = n * 26 + (s.charCodeAt(i) - 64);
    }
    return n;
  }

  function parseRangeStrToModel(rangeStr) {
    const m = String(rangeStr || '').trim().match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
    if (!m) return null;
    const c1 = colLettersToNumber(m[1]);
    const r1 = parseInt(m[2], 10);
    const c2 = colLettersToNumber(m[3]);
    const r2 = parseInt(m[4], 10);
    return {
      top: Math.min(r1, r2),
      left: Math.min(c1, c2),
      bottom: Math.max(r1, r2),
      right: Math.max(c1, c2)
    };
  }

  /**
   * Trích xuất toàn bộ danh sách ô gộp của trang tính (hỗ trợ cả ws._merges lẫn ws.model.merges)
   */
  function extractWorksheetMerges(ws) {
    const map = new Map();
    function addMerge(obj) {
      if (!obj) return;
      const top = obj.top;
      const left = obj.left;
      const bottom = obj.bottom;
      const right = obj.right;
      if (top > 0 && left > 0 && bottom >= top && right >= left && (bottom > top || right > left)) {
        const key = `${top}:${left}:${bottom}:${right}`;
        if (!map.has(key)) {
          map.set(key, { top, left, bottom, right });
        }
      }
    }

    if (ws._merges) {
      Object.values(ws._merges).forEach(m => {
        if (!m) return;
        if (typeof m === 'string') {
          addMerge(parseRangeStrToModel(m));
        } else if (m.model && m.model.top !== undefined) {
          addMerge(m.model);
        } else if (m.top !== undefined) {
          addMerge(m);
        }
      });
    }

    if (ws.model && Array.isArray(ws.model.merges)) {
      ws.model.merges.forEach(rStr => {
        addMerge(parseRangeStrToModel(rStr));
      });
    }

    return Array.from(map.values());
  }

  /**
   * Tự động dò các vùng (Zones) trong trang tính Form Báo Cáo:
   * - Tìm dòng có tiêu đề chứa "Khách hàng" (hoặc "Customer") ở cột J (10)
   *   và "Sản phẩm" (hoặc "Product") ở cột K (11).
   * - Hỗ trợ tiêu đề bảng gộp dọc 2-3 dòng (vd J6:J7, K6:K7): tự gom thành 1 cụm tiêu đề và lấy dòng đáy làm headerRow.
   * - Vùng dữ liệu bắt đầu ngay sau dòng đáy tiêu đề bảng và kéo dài đến trước dòng có chữ "Ghi chú" (Notes).
   */
  function detectZonesInWorksheet(ws) {
    const mergesList = extractWorksheetMerges(ws);
    const rawHeaders = [];
    const noteRows = [];

    function findVerticalMergeBounds(r, c) {
      let top = r;
      let bottom = r;
      for (const m of mergesList) {
        if (r >= m.top && r <= m.bottom && c >= m.left && c <= m.right) {
          top = Math.min(top, m.top);
          bottom = Math.max(bottom, m.bottom);
        }
      }
      return { top, bottom };
    }

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const jText = getPlainText(row.getCell(10).value).toLowerCase();
      const kText = getPlainText(row.getCell(11).value).toLowerCase();

      if ((jText.includes('khách hàng') || jText.includes('customer')) &&
          (kText.includes('sản phẩm') || kText.includes('product'))) {
        const bJ = findVerticalMergeBounds(rowNumber, 10);
        const bK = findVerticalMergeBounds(rowNumber, 11);
        const topRow = Math.min(rowNumber, bJ.top, bK.top);
        const bottomRow = Math.max(rowNumber, bJ.bottom, bK.bottom);

        // Nếu dòng này thuộc cùng 1 tiêu đề bảng gộp dọc với dòng ngay trước đó thì gộp chung
        const prev = rawHeaders[rawHeaders.length - 1];
        if (prev && topRow <= prev.bottomRow + 1) {
          prev.topRow = Math.min(prev.topRow, topRow);
          prev.bottomRow = Math.max(prev.bottomRow, bottomRow);
        } else {
          rawHeaders.push({ topRow, bottomRow });
        }
      }

      // Kiểm tra cả dòng xem có ô nào chứa "Ghi chú:" / "Ghi chú" không
      let hasGhiChu = false;
      for (let c = 1; c <= Math.min(row.cellCount, 15); c++) {
        const txt = getPlainText(row.getCell(c).value).toLowerCase();
        if (txt.includes('ghi chú:') || txt.startsWith('ghi chú')) {
          hasGhiChu = true;
          break;
        }
      }
      if (hasGhiChu) {
        const bNote = findVerticalMergeBounds(rowNumber, 1);
        const noteTop = bNote.top;
        if (noteRows[noteRows.length - 1] !== noteTop) {
          noteRows.push(noteTop);
        }
      }
    });

    if (rawHeaders.length === 0) return [];

    // Ghép từng cụm tiêu đề bảng với dòng "Ghi chú:" tương ứng
    const validPairs = [];
    for (let i = 0; i < rawHeaders.length; i++) {
      const hObj = rawHeaders[i];
      const nextHTop = (i + 1 < rawHeaders.length) ? rawHeaders[i + 1].topRow : (ws.rowCount + 100);
      const nRow = noteRows.find(r => r > hObj.bottomRow && r < nextHTop);
      if (!nRow) continue;
      validPairs.push({
        headerTop: hObj.topRow,
        headerRow: hObj.bottomRow,
        noteRow: nRow
      });
    }

    if (validPairs.length === 0) return [];

    const zones = [];
    for (let idx = 0; idx < validPairs.length; idx++) {
      const pair = validPairs[idx];
      const hRow = pair.headerRow;
      const nRow = pair.noteRow;

      let zoneStart;
      if (idx === 0) {
        zoneStart = 1;
      } else {
        zoneStart = zones[idx - 1].zoneEnd + 1;
      }

      let zoneEnd;
      if (idx + 1 < validPairs.length) {
        const nextPair = validPairs[idx + 1];
        // Vùng 1 có 7 dòng phần đầu (1..7, dòng 1 là tên công ty), Vùng 2 có 6 dòng phần đầu (38..43 -> sau nới là 40..45)
        // Do đó phần chân Vùng 1 kết thúc ở (nextPair.headerRow - 6) = 43 - 6 = 37 (tức nRow + 6 = 31 + 6 = 37)
        const expectedNextHeaderSpan = Math.max(5, validPairs[0].headerRow - 1); // 6 dòng (vd 38..43)
        const candidateZoneEnd = nextPair.headerRow - expectedNextHeaderSpan;
        zoneEnd = Math.max(nRow + 4, candidateZoneEnd);
      } else {
        // Vùng cuối của form mẫu giữ cùng số dòng phần chân (Ghi chú + 6 dòng chữ ký = 7 dòng) như Vùng 1
        // Lưu ý: Không dùng Math.min(ws.rowCount, ...) vì các dòng chữ ký cuối trang có thể là dòng trống chỉ chứa Text Box nổi!
        const footerAfterNote = zones.length > 0
          ? (zones[0].zoneEnd - zones[0].noteRow)
          : 6;
        zoneEnd = Math.max(ws.rowCount, nRow + footerAfterNote);
      }

      zones.push({
        index: idx + 1,
        zoneStart,
        headerTop: pair.headerTop,
        headerRow: hRow,
        dataStart: hRow + 1,
        dataEnd: nRow - 1,
        dataRowCount: (nRow - 1) - (hRow + 1) + 1,
        noteRow: nRow,
        zoneEnd
      });
    }

    return zones;
  }

  /**
   * Kiểm tra một file/buffer Form Báo Cáo có hợp lệ không (dùng cho Admin/Người dùng khi tải form mới)
   */
  async function validateFormTemplateBuffer(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    for (const ws of wb.worksheets) {
      const zones = detectZonesInWorksheet(ws);
      if (zones.length > 0) {
        return {
          valid: true,
          sheetName: ws.name,
          zoneCount: zones.length,
          zones
        };
      }
    }
    return { valid: false, zoneCount: 0, zones: [] };
  }

  /**
   * Tạo Form Báo Cáo mẫu chuẩn (nếu người dùng chưa tải form riêng lên):
   * - Vùng 1: Dòng 1..37 (tiêu đề bảng ở dòng 6..7, 23 dòng dữ liệu 8..30, Ghi chú dòng 31, Chữ ký dòng 32..37)
   *   -> Khi điền dữ liệu sẽ tự động nới thêm 2 dòng thành 25 dòng dữ liệu (tổng 39 dòng: 1..39, tên file ở J4).
   * - Vùng 2: Dòng 38..75 (38 dòng: tiêu đề bảng ở dòng 42..43, 25 dòng dữ liệu 44..68, Ghi chú dòng 69, Chữ ký 70..75)
   *   -> Khi điền dữ liệu (sau khi Vùng 1 nới 2 dòng) Vùng 2 nằm ở dòng 40..77 (38 dòng, tên file ở dòng thứ 4 của vùng = J43),
   *      Vùng 3 ở dòng 78..115 (tên file ở J81), Vùng 4 ở dòng 116..153 (tên file ở J119).
   */
  async function createBuiltInDefaultFormBuffer() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('BaoCaoQC');

    const colWidths = [5, 11, 13, 10, 10, 10, 10, 10, 11, 28, 38, 12];
    colWidths.forEach((w, idx) => {
      ws.getColumn(idx + 1).width = w;
    });

    const thinBorder = {
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' }
    };

    function buildTemplateZone(startRow, isZone1, dataRowsCount) {
      let r = startRow;

      // Dòng 1 của vùng: Tên đơn vị / Công ty
      const rTop = ws.getRow(r++);
      rTop.height = 18;
      rTop.getCell(1).value = 'CÔNG TY CỔ PHẦN CHĂN NUÔI C.P. VIỆT NAM - PHÒNG QUẢN LÝ CHẤT LƯỢNG (QC)';
      rTop.getCell(1).font = { name: 'Times New Roman', size: 10, bold: true };

      // Dòng 2 của vùng: Tiêu đề biểu mẫu
      const rTitle = ws.getRow(r++);
      rTitle.height = 24;
      ws.mergeCells(rTitle.number, 1, rTitle.number, 12);
      for (let c = 1; c <= 12; c++) {
        rTitle.getCell(c).style = {
          font: { name: 'Times New Roman', size: 13, bold: true },
          alignment: { horizontal: 'center', vertical: 'middle' }
        };
      }
      rTitle.getCell(1).value = 'BÁO CÁO KIỂM TRA CHẤT LƯỢNG HÀNG HÓA XUẤT KHO / DELIVERY QC REPORT';

      // Dòng 3 của vùng: Ngày kiểm tra
      const rSub = ws.getRow(r++);
      rSub.height = 18;
      rSub.getCell(1).value = 'Ngày kiểm tra / Date: ....../....../2026        Giờ kiểm tra / Time: ......h......';
      rSub.getCell(1).font = { name: 'Times New Roman', size: 10, italic: true };

      // Dòng thứ 4 của vùng (Vùng 1: J4, Vùng 2 sau khi nới Vùng 1: J43, Vùng 3: J81, Vùng 4: J119): nơi ghi TÊN FILE ở ô J
      const rFileRow = ws.getRow(r++);
      rFileRow.height = 20;
      rFileRow.getCell(1).value = 'Ca kiểm tra / Shift: ....................';
      rFileRow.getCell(1).font = { name: 'Times New Roman', size: 10 };
      rFileRow.getCell(9).value = 'Chuyến / File:';
      rFileRow.getCell(9).font = { name: 'Times New Roman', size: 10, bold: true };
      rFileRow.getCell(10).value = '';
      rFileRow.getCell(10).font = { name: 'Times New Roman', size: 11, bold: true };

      // Dòng thứ 5 của vùng: Thông tin nhiệt độ & vệ sinh xe
      const rInfo1 = ws.getRow(r++);
      rInfo1.height = 18;
      rInfo1.getCell(1).value = 'Nhiệt độ xe lạnh / Temp (°C): ............     Tình trạng vệ sinh xe / Vehicle hygiene: Đạt [  ]   Không đạt [  ]';
      rInfo1.getCell(1).font = { name: 'Times New Roman', size: 10 };

      // Riêng Vùng 1 có thêm dòng 6 thông tin phụ để phần đầu Vùng 1 là 7 dòng (1..7) và Vùng 2 là 6 dòng (38..43)
      if (isZone1) {
        const rInfo2 = ws.getRow(r++);
        rInfo2.height = 18;
        rInfo2.getCell(1).value = 'Phương tiện vận chuyển / Truck No.: ............................     Tuyến giao / Route: ............................';
        rInfo2.getCell(1).font = { name: 'Times New Roman', size: 10 };
      }

      // Dòng tiêu đề bảng (Vùng 1: dòng 7; Vùng 2: dòng 43)
      const rHeader = ws.getRow(r++);
      rHeader.height = 28;
      const headers = [
        'STT\nNo.', 'Mã lô\nBatch', 'Quy cách\nSpec', 'SL xuất\nQty',
        'Ngoại quan\nAppearance', 'Độ tươi\nFreshness', 'Bao bì\nPacking',
        'Nhãn\nLabel', 'Kết luận\nResult', 'Khách hàng / Customer',
        'Sản phẩm / Product', 'Đánh giá\nRemark'
      ];
      headers.forEach((hText, idx) => {
        const cell = rHeader.getCell(idx + 1);
        cell.value = hText;
        cell.style = {
          font: { name: 'Times New Roman', size: 10, bold: true },
          alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } },
          border: JSON.parse(JSON.stringify(thinBorder))
        };
      });

      // Các dòng dữ liệu (Vùng 1 có 23 dòng: 8..30; Vùng 2 có 25 dòng: 44..68)
      for (let i = 1; i <= dataRowsCount; i++) {
        const dRow = ws.getRow(r++);
        dRow.height = 20;
        for (let c = 1; c <= 12; c++) {
          const cell = dRow.getCell(c);
          let align = { vertical: 'middle', horizontal: 'center' };
          if (c === 10) {
            align = { vertical: 'middle', horizontal: 'left', wrapText: true };
          } else if (c === 11) {
            align = { vertical: 'middle', horizontal: 'left', shrinkToFit: true };
          }
          cell.style = {
            font: { name: 'Times New Roman', size: 10 },
            border: JSON.parse(JSON.stringify(thinBorder)),
            alignment: align
          };
        }
      }

      // Dòng Ghi chú (Vùng 1: dòng 31; Vùng 2: dòng 69)
      const rNote = ws.getRow(r++);
      rNote.height = 20;
      ws.mergeCells(rNote.number, 1, rNote.number, 12);
      for (let c = 1; c <= 12; c++) {
        rNote.getCell(c).style = {
          font: { name: 'Times New Roman', size: 9.5, italic: true },
          alignment: { vertical: 'middle', horizontal: 'left' }
        };
      }
      rNote.getCell(1).value = 'Ghi chú: Hàng hóa được kiểm tra đạt tiêu chuẩn chất lượng & VSATTP trước khi xuất kho.';

      // 6 dòng cuối vùng dành cho khung chữ ký (Vùng 1: 32..37; Vùng 2: 70..75)
      for (let f = 0; f < 6; f++) {
        const fRow = ws.getRow(r++);
        fRow.height = 18;
        fRow.getCell(1).value = '';
        fRow.getCell(1).font = { name: 'Times New Roman', size: 10 };
      }
    }

    // Dựng Vùng 1 (dòng 1..37, 23 dòng dữ liệu) và Vùng 2 (dòng 38..75, 25 dòng dữ liệu)
    buildTemplateZone(1, true, 23);
    buildTemplateZone(38, false, 25);

    const rawBuf = await wb.xlsx.writeBuffer();
    // Đính kèm sẵn drawing XML chữ ký vào cuối Vùng 1 (dòng 32) và Vùng 2 (dòng 70)
    return await global.FormShapesModule.injectFormSignatureShapes(null, rawBuf, [32, 70]);
  }

  /**
   * Chia danh sách khách hàng & sản phẩm vào các vùng của Form Báo Cáo (tối đa 25 dòng/vùng):
   * - Nếu khách có <= 25 sản phẩm: dùng First-Fit để lấp đầy vùng còn chỗ (không tách khách <= 25).
   * - Nếu khách có > 25 sản phẩm:
   *   + 25 sản phẩm đầu vào 1 vùng, từ sản phẩm thứ 26 trở đi sang vùng kế tiếp, ghi lại tên khách ở dòng đầu.
   *   + Phần dư KHÔNG BAO GIỜ nằm ở vùng trước phần đầu, và phần dư có thể ghép chung với khách khác vào cùng vùng.
   */
  function allocateDataToFormZones(processedData, zoneCapacity = 25) {
    if (!processedData) return [];

    if (processedData.onlyProductMode) {
      const allProducts = [];
      (processedData.zones || []).forEach(z => {
        (z.customers || []).forEach(c => {
          allProducts.push(...(c.products || []));
        });
      });
      const formZones = [];
      for (let i = 0; i < allProducts.length; i += zoneCapacity) {
        formZones.push({
          usedRows: Math.min(zoneCapacity, allProducts.length - i),
          entries: [{
            customerLabel: '',
            products: allProducts.slice(i, i + zoneCapacity)
          }]
        });
      }
      return formZones;
    }

    const formZones = [];
    // Lấy danh sách khách hàng từ customersList hoặc gom từ processedData.zones
    let customersList = processedData.customersList;
    if (!Array.isArray(customersList) || customersList.length === 0) {
      customersList = [];
      (processedData.zones || []).forEach(z => {
        (z.customers || []).forEach(c => customersList.push(c));
      });
    }

    customersList.forEach(cust => {
      const prods = cust.products || [];
      if (prods.length === 0) return;

      if (prods.length <= zoneCapacity) {
        let placed = false;
        for (let z = 0; z < formZones.length; z++) {
          if (formZones[z].usedRows + prods.length <= zoneCapacity) {
            formZones[z].entries.push({
              customerLabel: cust.displayLabel,
              products: prods.slice()
            });
            formZones[z].usedRows += prods.length;
            placed = true;
            break;
          }
        }
        if (!placed) {
          formZones.push({
            usedRows: prods.length,
            entries: [{
              customerLabel: cust.displayLabel,
              products: prods.slice()
            }]
          });
        }
      } else {
        // Khách có > 25 sản phẩm: cắt thành các đoạn 25 dòng, phần dư luôn ở vùng sau phần đầu
        let minZoneIdx = 0;
        for (let offset = 0; offset < prods.length; offset += zoneCapacity) {
          const chunk = prods.slice(offset, offset + zoneCapacity);
          let placedIdx = -1;

          for (let z = minZoneIdx; z < formZones.length; z++) {
            if (formZones[z].usedRows + chunk.length <= zoneCapacity) {
              formZones[z].entries.push({
                customerLabel: cust.displayLabel,
                products: chunk
              });
              formZones[z].usedRows += chunk.length;
              placedIdx = z;
              break;
            }
          }

          if (placedIdx === -1) {
            formZones.push({
              usedRows: chunk.length,
              entries: [{
                customerLabel: cust.displayLabel,
                products: chunk
              }]
            });
            placedIdx = formZones.length - 1;
          }
          minZoneIdx = placedIdx + 1;
        }
      }
    });

    return formZones;
  }

  // Tính cỡ chữ tự động giảm nếu tên khách dài khi gộp N dòng ở cột J
  function computeCustomerFontSize(text, mergedRowCount, baseSize = 10) {
    const len = String(text || '').length;
    const capacityChars = mergedRowCount * 24;
    if (len <= capacityChars) return baseSize;
    if (len <= capacityChars * 1.4) return Math.max(8, baseSize - 1);
    if (len <= capacityChars * 1.9) return Math.max(7.5, baseSize - 2);
    return 7;
  }

  /**
   * Tạo một object style hoàn toàn độc lập từ ô nguồn (không dùng chung tham chiếu với master.style của ExcelJS)
   */
  function buildClonedStyleObject(srcCell) {
    const styleObj = {};
    if (!srcCell) return styleObj;
    if (srcCell.font) styleObj.font = JSON.parse(JSON.stringify(srcCell.font));
    if (srcCell.alignment) styleObj.alignment = JSON.parse(JSON.stringify(srcCell.alignment));
    if (srcCell.border) styleObj.border = JSON.parse(JSON.stringify(srcCell.border));
    if (srcCell.fill) styleObj.fill = JSON.parse(JSON.stringify(srcCell.fill));
    if (srcCell.numFmt) styleObj.numFmt = srcCell.numFmt;
    if (srcCell.protection) styleObj.protection = JSON.parse(JSON.stringify(srcCell.protection));
    return styleObj;
  }

  function cloneCellStyle(srcCell, dstCell) {
    dstCell.style = buildClonedStyleObject(srcCell);
  }

  /**
   * Xác định cột cuối thực sự của biểu mẫu (tránh trường hợp file mẫu có format thừa tới cột XFD làm lệch vùng in)
   */
  function detectEffectiveMaxCol(srcWs, maxScanRow, mergesList) {
    let maxCol = 12;
    mergesList.forEach(m => {
      if (m.top <= maxScanRow && m.right <= 30) {
        if (m.right > maxCol) maxCol = m.right;
      }
    });
    const limitRow = Math.min(maxScanRow, srcWs.rowCount);
    for (let r = 1; r <= limitRow; r++) {
      const row = srcWs.getRow(r);
      const cLimit = Math.min(30, row.cellCount);
      for (let c = 1; c <= cLimit; c++) {
        const cell = row.getCell(c);
        const hasVal = getPlainText(cell.value) !== '';
        const b = cell.border;
        const hasBorder = Boolean(b && (b.top || b.bottom || b.left || b.right));
        if ((hasVal || hasBorder) && c > maxCol) {
          maxCol = c;
        }
      }
    }
    return maxCol;
  }

  /**
   * Điền dữ liệu vào Form Báo Cáo (luôn tạo Workbook MỚI chỉ có 1 trang tính báo cáo,
   * tuyệt đối không dùng spliceRows để tránh lệch ô gộp và lỗi Excel "We found a problem").
   */
  async function fillReportTemplate(options) {
    const {
      templateBuffer,
      processedData,
      reportFileName = 'Bao_cao_giao_hang',
      writeFileNameToJ4 = true,
      enablePrintSetup = true
    } = options;

    const activeTemplateBuffer = templateBuffer || (await createBuiltInDefaultFormBuffer());

    const srcWb = new ExcelJS.Workbook();
    await srcWb.xlsx.load(activeTemplateBuffer);

    let srcWs = null;
    let detectedZones = [];
    for (const ws of srcWb.worksheets) {
      const z = detectZonesInWorksheet(ws);
      if (z.length > 0) {
        srcWs = ws;
        detectedZones = z;
        break;
      }
    }

    if (!srcWs || detectedZones.length === 0) {
      throw new Error('Không dò được vùng "Khách hàng / Customer" (cột J) và "Sản phẩm / Product" (cột K) trong Form Báo Cáo.');
    }

    const z1 = detectedZones[0];
    const z2 = detectedZones.length > 1 ? detectedZones[1] : z1;
    const targetDataCapacity = z2.dataRowCount >= 25 ? z2.dataRowCount : 25;

    const formZonesData = allocateDataToFormZones(processedData, targetDataCapacity);
    const totalNeededZones = Math.max(2, formZonesData.length);

    // Thu thập toàn bộ thông tin ô gộp của srcWs
    const srcMerges = extractWorksheetMerges(srcWs);
    const maxCol = detectEffectiveMaxCol(srcWs, z2.zoneEnd, srcMerges);

    // Tạo Sổ làm việc MỚI hoàn toàn chỉ chứa 1 trang tính báo cáo
    const dstWb = new ExcelJS.Workbook();
    const safeSheetName = String(srcWs.name || 'BaoCao').replace(/[\[\]\*\/\\\?\:]/g, '_').slice(0, 31) || 'BaoCao';
    const dstWs = dstWb.addWorksheet(safeSheetName);

    // Chép thuộc tính trang tính & chế độ hiển thị (gridLines, zoom)
    if (srcWs.properties) {
      dstWs.properties = Object.assign({}, dstWs.properties, {
        defaultRowHeight: srcWs.properties.defaultRowHeight || 18,
        defaultColWidth: srcWs.properties.defaultColWidth
      });
    }
    if (Array.isArray(srcWs.views) && srcWs.views.length > 0) {
      const v0 = srcWs.views[0];
      dstWs.views = [{
        state: 'normal',
        showGridLines: v0.showGridLines !== undefined ? v0.showGridLines : true,
        zoomScale: v0.zoomScale || 100,
        zoomScaleNormal: v0.zoomScaleNormal || 100
      }];
    }

    // Chép độ rộng cột
    for (let c = 1; c <= maxCol; c++) {
      const srcCol = srcWs.getColumn(c);
      const dstCol = dstWs.getColumn(c);
      if (srcCol && srcCol.width) {
        dstCol.width = srcCol.width;
      }
    }

    // Theo dõi các vùng gộp trên dstWs để tuyệt đối không bao giờ gộp trùng hoặc chồng lấn
    const trackedDstMerges = [];
    function safeMergeOnDst(top, left, bottom, right) {
      if (top <= 0 || left <= 0 || bottom < top || right < left) return false;
      if (top === bottom && left === right) return false;
      for (const ex of trackedDstMerges) {
        const overlaps = !(bottom < ex.top || top > ex.bottom || right < ex.left || left > ex.right);
        if (overlaps) return false;
      }
      try {
        dstWs.mergeCells(top, left, bottom, right);
        trackedDstMerges.push({ top, left, bottom, right });
        return true;
      } catch (e) {
        return false;
      }
    }

    // Hàm chép 1 dòng từ srcWs sang dstWs (chưa gộp ô)
    function copySingleRow(srcRowNum, dstRowNum, clearDataCols = false) {
      const sRow = srcWs.getRow(srcRowNum);
      const dRow = dstWs.getRow(dstRowNum);
      dRow.height = sRow.height || 20;

      // Lấy viền chuẩn của cột K (11 - Sản phẩm, không bị gộp dọc ở mẫu) để khôi phục viền ngang cho cột J nếu mẫu cũ đang gộp cột J
      const kSampleCell = sRow.getCell(11);
      const kBorder = kSampleCell && kSampleCell.border ? JSON.parse(JSON.stringify(kSampleCell.border)) : null;

      for (let c = 1; c <= maxCol; c++) {
        const sCell = sRow.getCell(c);
        const dCell = dRow.getCell(c);
        if (!clearDataCols) {
          if (!sCell.isMerged || (sCell.master && sCell.master.address === sCell.address)) {
            dCell.value = sanitizeCellValue(sCell.value);
          } else {
            dCell.value = null;
          }
        } else {
          dCell.value = null;
        }

        const clonedStyle = buildClonedStyleObject(sCell);

        // Nếu đang ở dòng dữ liệu (clearDataCols = true) và cột J (10) bị mất viền trên/dưới do file mẫu từng gộp ô cột J:
        // khôi phục viền trên/dưới đồng bộ với cột K (11) cùng dòng để các dòng trống hoặc khách 1 sản phẩm luôn đủ khung viền
        if (clearDataCols && c === 10 && kBorder) {
          clonedStyle.border = clonedStyle.border || {};
          if (!clonedStyle.border.top && kBorder.top) {
            clonedStyle.border.top = JSON.parse(JSON.stringify(kBorder.top));
          }
          if (!clonedStyle.border.bottom && kBorder.bottom) {
            clonedStyle.border.bottom = JSON.parse(JSON.stringify(kBorder.bottom));
          }
          if (!clonedStyle.border.left) {
            clonedStyle.border.left = kBorder.left ? JSON.parse(JSON.stringify(kBorder.left)) : { style: 'thin' };
          }
          if (!clonedStyle.border.right) {
            clonedStyle.border.right = kBorder.right ? JSON.parse(JSON.stringify(kBorder.right)) : { style: 'thin' };
          }
        }

        dCell.style = clonedStyle;
      }
    }

    /**
     * Áp dụng lại các ô gộp ngang/dọc nằm trọn trong đoạn [srcStart..srcEnd]:
     * Giữ nguyên font/alignment/fill của ô gốc (master) NHƯNG bảo toàn viền ngoài (top/bottom/left/right)
     * của từng ô thành phần trên biên vùng gộp để không bao giờ bị mất viền phải/viền đáy sau khi gộp!
     */
    function copyMergesInRange(srcStart, srcEnd, dstOffset, isSingleDataRow = false) {
      srcMerges.forEach(m => {
        if (m.top >= srcStart && m.bottom <= srcEnd && m.right <= maxCol) {
          // Trên dòng dữ liệu, bỏ qua các ô gộp chạm vào cột J (10) hoặc gộp dọc nhiều dòng
          if (isSingleDataRow) {
            if (m.top !== m.bottom) return;
            if (m.left <= 10 && m.right >= 10) return;
          }

          const dTop = m.top + dstOffset;
          const dBottom = m.bottom + dstOffset;

          // Lưu lại viền riêng của từng ô trong vùng trước khi gọi mergeCells (vì ExcelJS mergeCells sẽ gán chung master.style)
          const savedBorders = [];
          for (let r = dTop; r <= dBottom; r++) {
            const rowBorders = [];
            const srcR = Math.min(srcEnd, Math.max(srcStart, r - dstOffset));
            for (let c = m.left; c <= m.right; c++) {
              const b = srcWs.getRow(srcR).getCell(c).border || dstWs.getRow(r).getCell(c).border;
              rowBorders.push(b ? JSON.parse(JSON.stringify(b)) : null);
            }
            savedBorders.push(rowBorders);
          }

          const masterSrcCell = srcWs.getRow(m.top).getCell(m.left);
          const masterBaseStyle = buildClonedStyleObject(masterSrcCell);
          const masterBorder = masterBaseStyle.border || {};

          if (safeMergeOnDst(dTop, m.left, dBottom, m.right)) {
            for (let r = dTop; r <= dBottom; r++) {
              for (let c = m.left; c <= m.right; c++) {
                const cellBorder = savedBorders[r - dTop][c - m.left] || {};
                const mergedBorder = {};

                if (r === dTop && (cellBorder.top || masterBorder.top)) {
                  mergedBorder.top = JSON.parse(JSON.stringify(cellBorder.top || masterBorder.top));
                }
                if (r === dBottom && (cellBorder.bottom || masterBorder.bottom)) {
                  mergedBorder.bottom = JSON.parse(JSON.stringify(cellBorder.bottom || masterBorder.bottom));
                }
                if (c === m.left && (cellBorder.left || masterBorder.left)) {
                  mergedBorder.left = JSON.parse(JSON.stringify(cellBorder.left || masterBorder.left));
                }
                if (c === m.right && (cellBorder.right || masterBorder.right)) {
                  mergedBorder.right = JSON.parse(JSON.stringify(cellBorder.right || masterBorder.right));
                }

                // Gán một object style hoàn toàn độc lập cho từng ô trong vùng gộp
                const cellStyle = buildClonedStyleObject(masterSrcCell);
                if (Object.keys(mergedBorder).length > 0) {
                  cellStyle.border = mergedBorder;
                }
                dstWs.getRow(r).getCell(c).style = cellStyle;
              }
            }
          }
        }
      });
    }

    let currentDstRow = 1;
    const pageBreakRows = [];
    const zoneFooterRows = [];
    const cleanTitleName = String(reportFileName || '').replace(/\.xlsx$/i, '').trim();

    for (let zIdx = 0; zIdx < totalNeededZones; zIdx++) {
      const isFirstZone = (zIdx === 0);
      const tplZone = isFirstZone ? z1 : z2;
      const zoneStartDst = currentDstRow;

      // 1. Chép phần đầu vùng (từ zoneStart đến headerRow)
      const headerOffset = currentDstRow - tplZone.zoneStart;
      for (let r = tplZone.zoneStart; r <= tplZone.headerRow; r++) {
        copySingleRow(r, currentDstRow++, false);
      }
      copyMergesInRange(tplZone.zoneStart, tplZone.headerRow, headerOffset, false);

      // Ghi TÊN FILE (bỏ .xlsx) vào ô J ở dòng thứ 4 của mỗi vùng có dữ liệu:
      // Vùng 1 (bắt đầu dòng 1) -> dòng 4 (J4)
      // Vùng 2 (bắt đầu dòng 40) -> dòng 43 (J43)
      // Vùng 3 (bắt đầu dòng 78) -> dòng 81 (J81), Vùng 4 -> J119...
      const hasDataInThisZone = (zIdx < formZonesData.length && formZonesData[zIdx].usedRows > 0) ||
                                (zIdx === 0 && formZonesData.length === 0);
      if (writeFileNameToJ4 && hasDataInThisZone && cleanTitleName) {
        const fileTitleRowNum = zoneStartDst + 3; // Dòng thứ 4 của vùng
        const titleRow = dstWs.getRow(fileTitleRowNum);
        const jCell = titleRow.getCell(10);
        const targetCell = (jCell.isMerged && jCell.master) ? jCell.master : jCell;

        // Lấy phông chữ cùng dòng đó (ưu tiên ô J, sau đó đến các ô có chữ trên cùng dòng thứ 4)
        let rowFont = targetCell.font || jCell.font;
        if (!rowFont || !rowFont.name) {
          for (let c = 1; c <= maxCol; c++) {
            const f = titleRow.getCell(c).font;
            if (f && f.name) {
              rowFont = f;
              break;
            }
          }
        }
        const baseFont = rowFont ? JSON.parse(JSON.stringify(rowFont)) : { name: 'Times New Roman', size: 11 };
        baseFont.bold = true;

        const newJStyle = buildClonedStyleObject(targetCell);
        newJStyle.font = baseFont;
        targetCell.value = cleanTitleName;
        targetCell.style = newJStyle;
      }

      // 2. Dựng đúng 25 dòng dữ liệu (nới Vùng 1 từ 23 lên 25 dòng bằng cách lặp dòng mẫu kế cuối, giữ dòng cuối ở cuối)
      const dataStartDst = currentDstRow;
      const tplDataStart = tplZone.dataStart;
      const tplDataEnd = tplZone.dataEnd;
      const tplMiddleSampleRow = Math.max(tplDataStart, tplDataEnd - 1);

      for (let dIdx = 0; dIdx < targetDataCapacity; dIdx++) {
        let srcSampleRow;
        if (dIdx === targetDataCapacity - 1) {
          srcSampleRow = tplDataEnd;
        } else if (tplDataStart + dIdx < tplDataEnd) {
          srcSampleRow = tplDataStart + dIdx;
        } else {
          srcSampleRow = tplMiddleSampleRow;
        }
        copySingleRow(srcSampleRow, currentDstRow, true);
        copyMergesInRange(srcSampleRow, srcSampleRow, currentDstRow - srcSampleRow, true);
        currentDstRow++;
      }

      // 3. Điền dữ liệu Khách hàng (cột J = 10) và Sản phẩm (cột K = 11) vào vùng này
      if (zIdx < formZonesData.length && formZonesData[zIdx].usedRows > 0) {
        const zData = formZonesData[zIdx];
        let writeRow = dataStartDst;

        for (const entry of zData.entries) {
          const prods = entry.products || [];
          if (prods.length === 0) continue;

          const entryStartRow = writeRow;
          for (let p = 0; p < prods.length; p++) {
            const rObj = dstWs.getRow(writeRow);
            const cJ = rObj.getCell(10);
            const cK = rObj.getCell(11);

            if (p === 0 && entry.customerLabel) {
              cJ.value = entry.customerLabel;
            } else {
              cJ.value = null;
            }
            cK.value = prods[p];
            writeRow++;
          }
          const entryEndRow = writeRow - 1;

          // Lưu lại style & viền trên/dưới chuẩn của đoạn [entryStartRow..entryEndRow] trước khi gộp cột J
          const startCellJ = dstWs.getRow(entryStartRow).getCell(10);
          const endCellJ = dstWs.getRow(entryEndRow).getCell(10);
          const startCellK = dstWs.getRow(entryStartRow).getCell(11);
          const endCellK = dstWs.getRow(entryEndRow).getCell(11);

          const baseStyleJ = buildClonedStyleObject(startCellJ);
          const topBorderJ = (startCellJ.border && startCellJ.border.top) || (startCellK.border && startCellK.border.top) || { style: 'thin' };
          const bottomBorderJ = (endCellJ.border && endCellJ.border.bottom) || (endCellK.border && endCellK.border.bottom) || { style: 'thin' };
          const leftBorderJ = (startCellJ.border && startCellJ.border.left) || { style: 'thin' };
          const rightBorderJ = (startCellJ.border && startCellJ.border.right) || { style: 'thin' };

          // Gộp ô dọc cột J (Khách hàng) theo số sản phẩm của khách (không gộp khi tên rỗng)
          if (entry.customerLabel && entryEndRow > entryStartRow) {
            safeMergeOnDst(entryStartRow, 10, entryEndRow, 10);
          }

          // Định dạng lại toàn bộ ô cột J và K của khách hàng bằng các object style ĐỘC LẬP (tránh lỗi mất viền của ExcelJS)
          const tplFontSizeJ = (baseStyleJ.font && baseStyleJ.font.size) ? baseStyleJ.font.size : 10;
          const custFontSize = computeCustomerFontSize(entry.customerLabel, prods.length, tplFontSizeJ);

          for (let rIdx = entryStartRow; rIdx <= entryEndRow; rIdx++) {
            const rObj = dstWs.getRow(rIdx);
            const cellJ = rObj.getCell(10);
            const cellK = rObj.getCell(11);

            const rowBorderJ = {
              left: JSON.parse(JSON.stringify(leftBorderJ)),
              right: JSON.parse(JSON.stringify(rightBorderJ))
            };
            if (rIdx === entryStartRow) {
              rowBorderJ.top = JSON.parse(JSON.stringify(topBorderJ));
            }
            if (rIdx === entryEndRow) {
              rowBorderJ.bottom = JSON.parse(JSON.stringify(bottomBorderJ));
            }

            const newStyleJ = buildClonedStyleObject(baseStyleJ);
            newStyleJ.font = Object.assign({}, newStyleJ.font || { name: 'Times New Roman' }, { size: custFontSize });
            newStyleJ.alignment = Object.assign({}, newStyleJ.alignment || {}, {
              vertical: 'middle',
              horizontal: (newStyleJ.alignment && newStyleJ.alignment.horizontal) || 'left',
              wrapText: true
            });
            newStyleJ.border = rowBorderJ;
            cellJ.style = newStyleJ;

            const newStyleK = buildClonedStyleObject(cellK);
            newStyleK.font = Object.assign({ name: 'Times New Roman', size: 10 }, newStyleK.font || {});
            newStyleK.alignment = Object.assign({}, newStyleK.alignment || {}, {
              vertical: 'middle',
              horizontal: (newStyleK.alignment && newStyleK.alignment.horizontal) || 'left',
              wrapText: false,
              shrinkToFit: true
            });
            cellK.style = newStyleK;
          }
        }
      }

      // 4. Chép phần chân vùng (từ dòng "Ghi chú:" đến hết vùng)
      const footerOffset = currentDstRow - tplZone.noteRow;
      const signatureStartRowDst = currentDstRow + 1; // Dòng ngay dưới "Ghi chú:" để đặt Text Box chữ ký
      zoneFooterRows.push(signatureStartRowDst);

      for (let r = tplZone.noteRow; r <= tplZone.zoneEnd; r++) {
        copySingleRow(r, currentDstRow++, false);
      }
      copyMergesInRange(tplZone.noteRow, tplZone.zoneEnd, footerOffset, false);

      // Ngắt trang sau mỗi vùng (trừ vùng cuối cùng)
      const zoneEndDst = currentDstRow - 1;
      if (zIdx < totalNeededZones - 1) {
        pageBreakRows.push(zoneEndDst);
      }
    }

    const totalRowsDst = currentDstRow - 1;
    const lastColLetter = dstWs.getColumn(maxCol).letter || 'L';

    // Thiết lập vùng in A4 ngang, tỉ lệ cố định 81%
    if (enablePrintSetup) {
      global.PrintSetupModule.applyPrintSetup(dstWs, {
        enabled: true,
        orientation: 'landscape',
        scale: 81,
        printArea: `A1:${lastColLetter}${totalRowsDst}`,
        pageBreakRows: pageBreakRows
      });
    }

    let outBuffer = await dstWb.xlsx.writeBuffer();

    // Hậu xử lý ngắt trang & chuẩn hóa thẻ <pageSetup> theo đúng thứ tự CT_Worksheet
    if (enablePrintSetup) {
      outBuffer = await global.PrintSetupModule.enforceRowBreaksInXlsxBuffer(
        outBuffer,
        pageBreakRows,
        81,
        'landscape'
      );
    }

    // Dùng JSZip chèn lại các hình vẽ / Text Box chữ ký vào cuối mỗi vùng & đồng bộ theme1.xml
    outBuffer = await global.FormShapesModule.injectFormSignatureShapes(
      activeTemplateBuffer,
      outBuffer,
      zoneFooterRows
    );

    return {
      buffer: outBuffer,
      zoneCount: totalNeededZones,
      usedZoneCount: formZonesData.length,
      totalRows: totalRowsDst
    };
  }

  global.TemplateFillModule = {
    extractWorksheetMerges,
    detectZonesInWorksheet,
    validateFormTemplateBuffer,
    createBuiltInDefaultFormBuffer,
    allocateDataToFormZones,
    fillReportTemplate
  };
})(window);
