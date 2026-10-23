// ==================== MODULE: print-setup.js ====================
(function(global) {
  'use strict';

  /**
   * Chèn thẻ <rowBreaks> vào đúng vị trí chuẩn trong schema OOXML CT_Worksheet:
   * Đứng SAU <sheetData>, <mergeCells>, <printOptions>, <pageMargins>, <pageSetup>, <headerFooter>
   * và đứng TRƯỚC <colBreaks>, <customProperties>, <cellWatches>, <ignoredErrors>, <smartTags>,
   * <drawing>, <legacyDrawing>, <tableParts>, <extLst>, </worksheet>.
   */
  function insertRowBreaksTagInWorksheetXml(sheetXml, rowBreaksXml) {
    let xml = sheetXml
      .replace(/<rowBreaks\b[^>]*>[\s\S]*?<\/rowBreaks>/g, '')
      .replace(/<rowBreaks\b[^>]*\/>/g, '');

    if (!rowBreaksXml) return xml;

    const afterTags = [
      'colBreaks',
      'customProperties',
      'cellWatches',
      'ignoredErrors',
      'smartTags',
      'drawing',
      'legacyDrawing',
      'legacyDrawingHF',
      'drawingHF',
      'picture',
      'oleObjects',
      'controls',
      'webPublishItems',
      'tableParts',
      'extLst'
    ];

    let earliestIdx = -1;
    for (const tag of afterTags) {
      const re = new RegExp(`<${tag}\\b`);
      const m = re.exec(xml);
      if (m && (earliestIdx === -1 || m.index < earliestIdx)) {
        earliestIdx = m.index;
      }
    }

    if (earliestIdx !== -1) {
      return xml.slice(0, earliestIdx) + rowBreaksXml + xml.slice(earliestIdx);
    }
    return xml.replace('</worksheet>', `${rowBreaksXml}</worksheet>`);
  }

  /**
   * Thiết lập vùng in A4 chuẩn cho trang tính ExcelJS
   * - Không dùng fitToPage (vì Excel sẽ bỏ qua ngắt trang thủ công rowBreaks)
   * - Đặt scale cố định (vd 95% cho bảng 2 cột dọc, 81% cho form báo cáo ngang)
   * - Lặp dòng tiêu đề mỗi trang, chân trang "Trang &P / &N"
   */
  function applyPrintSetup(ws, options) {
    const {
      enabled = true,
      orientation = 'portrait', // 'portrait' | 'landscape'
      scale = 95,
      printArea = null,
      printTitlesRow = null,
      pageBreakRows = []
    } = options || {};

    if (!enabled) return;

    const setupObj = {
      paperSize: 9, // A4
      orientation: orientation,
      fitToPage: false,
      scale: scale,
      horizontalCentered: true,
      verticalCentered: false,
      margins: {
        left: 0.35, right: 0.35,
        top: 0.4, bottom: 0.45,
        header: 0.2, footer: 0.2
      }
    };

    if (printArea) {
      setupObj.printArea = printArea;
    }
    if (printTitlesRow) {
      setupObj.printTitlesRow = printTitlesRow;
    }

    ws.pageSetup = setupObj;

    ws.headerFooter = Object.assign({}, ws.headerFooter, {
      oddFooter: '&CTrang &P / &N',
      evenFooter: '&CTrang &P / &N'
    });

    ws._customPageBreakRows = (pageBreakRows || []).slice();
  }

  /**
   * Hậu xử lý gói XLSX bằng JSZip để đảm bảo thẻ <rowBreaks> và <pageSetup fitToPage="0" scale="..."/>
   * nằm đúng thứ tự schema OpenXML (CT_Worksheet) mà Microsoft Excel tôn trọng 100%.
   */
  async function enforceRowBreaksInXlsxBuffer(xlsxBuffer, pageBreakRows, scale, orientation) {
    const zip = await JSZip.loadAsync(xlsxBuffer);
    const sheetPath = 'xl/worksheets/sheet1.xml';
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) return xlsxBuffer;

    let xml = await sheetFile.async('string');

    // 1. Đảm bảo sheetPr không bật pageSetUpPr fitToPage="1"
    xml = xml.replace(/<pageSetUpPr\b[^>]*fitToPage="1"[^>]*\/>/g, '<pageSetUpPr fitToPage="0"/>');

    // 2. Chuẩn hóa thẻ <pageSetup .../> (bỏ r:id mồ côi và fitToWidth/fitToHeight, gán scale/orientation/paperSize chuẩn)
    const targetScale = scale || 81;
    const targetOrient = orientation || 'portrait';
    const cleanPageSetupTag = `<pageSetup paperSize="9" scale="${targetScale}" orientation="${targetOrient}" horizontalDpi="4294967295" verticalDpi="4294967295"/>`;

    if (/<pageSetup\b[^>]*\/>/.test(xml)) {
      xml = xml.replace(/<pageSetup\b[^>]*\/>/, cleanPageSetupTag);
    } else if (/<pageSetup\b[^>]*>[\s\S]*?<\/pageSetup>/.test(xml)) {
      xml = xml.replace(/<pageSetup\b[^>]*>[\s\S]*?<\/pageSetup>/, cleanPageSetupTag);
    }

    // 3. Tạo khối <rowBreaks> chuẩn và chèn đúng vị trí CT_Worksheet
    const uniqueBreaks = Array.from(new Set((pageBreakRows || []).filter(r => r > 0))).sort((a, b) => a - b);
    if (uniqueBreaks.length > 0) {
      const brkItems = uniqueBreaks.map(r => `<brk id="${r}" max="16383" man="1"/>`).join('');
      const rowBreaksXml = `<rowBreaks count="${uniqueBreaks.length}" manualBreakCount="${uniqueBreaks.length}">${brkItems}</rowBreaks>`;
      xml = insertRowBreaksTagInWorksheetXml(xml, rowBreaksXml);
    } else {
      xml = insertRowBreaksTagInWorksheetXml(xml, '');
    }

    zip.file(sheetPath, xml);
    return await zip.generateAsync({ type: 'arraybuffer' });
  }

  /**
   * Xuất file Excel kết quả chuẩn (Mục 3):
   * - 2 cột "Khách hàng" và "Sản phẩm" (hoặc 1 cột "Sản phẩm" khi bật chế độ chỉ lấy sản phẩm)
   * - Các vùng cách nhau 1 dòng trống tô màu vàng (#FFD966, cao 10pt)
   * - Gộp ô Khách hàng dọc theo số sản phẩm, wrap text
   */
  async function buildStandardResultWorkbookBuffer(processedData, options) {
    const {
      enablePrintSetup = true
    } = options || {};

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('KetQua');
    const onlyProduct = processedData.onlyProductMode;

    if (onlyProduct) {
      ws.columns = [
        { header: 'Sản phẩm', key: 'product', width: 52 }
      ];
    } else {
      ws.columns = [
        { header: 'Khách hàng', key: 'customer', width: 38 },
        { header: 'Sản phẩm', key: 'product', width: 48 }
      ];
    }

    const thinBorder = {
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' }
    };

    // Định dạng dòng tiêu đề (dòng 1)
    const headerRow = ws.getRow(1);
    headerRow.height = 24;
    const totalCols = onlyProduct ? 1 : 2;
    for (let c = 1; c <= totalCols; c++) {
      const cell = headerRow.getCell(c);
      cell.style = {
        font: { name: 'Arial', size: 11, bold: true },
        alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } },
        border: JSON.parse(JSON.stringify(thinBorder))
      };
    }

    let currentRow = 2;
    const pageBreakRows = [];
    const zones = processedData.zones || [];

    for (let zIdx = 0; zIdx < zones.length; zIdx++) {
      const zone = zones[zIdx];

      for (const cust of zone.customers) {
        const prods = cust.products || [];
        if (prods.length === 0) continue;

        const startCustRow = currentRow;
        for (let pIdx = 0; pIdx < prods.length; pIdx++) {
          const r = ws.getRow(currentRow);
          r.height = 20;

          if (onlyProduct) {
            const cellP = r.getCell(1);
            cellP.value = prods[pIdx];
          } else {
            const cellC = r.getCell(1);
            const cellP = r.getCell(2);
            if (pIdx === 0) {
              cellC.value = cust.displayLabel || '';
            }
            cellP.value = prods[pIdx];
          }
          currentRow++;
        }
        const endCustRow = currentRow - 1;

        // Gộp ô Khách hàng trước rồi mới gán object style độc lập cho từng ô (tránh mất viền ô phụ)
        if (!onlyProduct && cust.displayLabel && endCustRow > startCustRow) {
          ws.mergeCells(startCustRow, 1, endCustRow, 1);
        }

        // Gán style độc lập cho tất cả các dòng của khách hàng
        for (let rIdx = startCustRow; rIdx <= endCustRow; rIdx++) {
          const r = ws.getRow(rIdx);
          for (let c = 1; c <= totalCols; c++) {
            const cell = r.getCell(c);
            cell.style = {
              font: { name: 'Arial', size: 11 },
              border: JSON.parse(JSON.stringify(thinBorder)),
              alignment: (!onlyProduct && c === 1)
                ? { vertical: 'middle', horizontal: 'left', wrapText: true }
                : { vertical: 'middle', horizontal: 'left', wrapText: false, shrinkToFit: true }
            };
          }
        }
      }

      // Chèn dòng trống tô màu vàng (#FFD966, cao 10pt) ngăn cách giữa các vùng
      if (zIdx < zones.length - 1) {
        const sepRow = ws.getRow(currentRow);
        sepRow.height = 10;
        if (!onlyProduct) {
          ws.mergeCells(currentRow, 1, currentRow, 2);
        }
        for (let c = 1; c <= totalCols; c++) {
          const cell = sepRow.getCell(c);
          cell.style = {
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFD966' } },
            border: JSON.parse(JSON.stringify(thinBorder))
          };
        }
        pageBreakRows.push(currentRow);
        currentRow++;
      }
    }

    const lastRow = Math.max(1, currentRow - 1);
    const lastColLetter = onlyProduct ? 'A' : 'B';

    if (enablePrintSetup) {
      applyPrintSetup(ws, {
        enabled: true,
        orientation: 'portrait',
        scale: 95,
        printArea: `A1:${lastColLetter}${lastRow}`,
        printTitlesRow: '1:1',
        pageBreakRows: pageBreakRows
      });
    }

    const rawBuf = await wb.xlsx.writeBuffer();
    if (enablePrintSetup) {
      return await enforceRowBreaksInXlsxBuffer(rawBuf, pageBreakRows, 95, 'portrait');
    }
    return rawBuf;
  }

  global.PrintSetupModule = {
    applyPrintSetup,
    enforceRowBreaksInXlsxBuffer,
    buildStandardResultWorkbookBuffer
  };
})(window);
