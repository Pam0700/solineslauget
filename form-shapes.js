// ==================== MODULE: form-shapes.js ====================
(function(global) {
  'use strict';

  function escapeXml(str) {
    return String(str ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function sanitizeRawXmlAmpersands(xmlStr) {
    return String(xmlStr ?? '').replace(
      /&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/g,
      '&amp;'
    );
  }

  /**
   * Chèn thẻ <drawing r:id="..."/> vào đúng vị trí chuẩn trong schema OOXML CT_Worksheet:
   * Bắt buộc đứng SAU <sheetData>, <mergeCells>, <pageMargins>, <pageSetup>, <headerFooter>, <rowBreaks>, <colBreaks>...
   * và đứng TRƯỚC <legacyDrawing>, <legacyDrawingHF>, <drawingHF>, <picture>, <oleObjects>, <controls>, <webPublishItems>, <tableParts>, <extLst>, </worksheet>.
   */
  function insertDrawingTagInWorksheetXml(sheetXml, drawingTag) {
    let xml = sheetXml
      .replace(/<drawing\b[^>]*\/>/g, '')
      .replace(/<drawing\b[^>]*>[\s\S]*?<\/drawing>/g, '');

    const afterTags = [
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
      return xml.slice(0, earliestIdx) + drawingTag + xml.slice(earliestIdx);
    }
    return xml.replace('</worksheet>', `${drawingTag}</worksheet>`);
  }

  /**
   * Trích xuất và nhân bản các hình vẽ / Text Box chữ ký ("Người báo cáo (QC) / Reported by",
   * "Kho & vận chuyển / Inspected by") từ gói XLSX form gốc sang gói XLSX kết quả.
   *
   * Lý do: ExcelJS khi đọc/ghi sang Workbook mới sẽ làm mất drawing XML (Text Box chữ ký).
   * Giải pháp:
   * 1. Mở gói XLSX gốc bằng JSZip, đồng bộ theme1.xml (nếu có) và tìm drawing1.xml.
   * 2. Lấy các khối <xdr:twoCellAnchor> hoặc <xdr:oneCellAnchor> là khung chữ ký (<xdr:sp>) ở cuối vùng,
   *    dời chỉ số dòng <xdr:row> (0-indexed) theo độ lệch dòng của từng vùng kết quả (Zone 1..N).
   * 3. Loại bỏ các thẻ <a:extLst> / <xdr:extLst> / <a16:creationId> trùng lặp và đảm bảo mọi cNvPr có id duy nhất.
   * 4. Nếu form gốc không có sẵn drawing XML chữ ký hợp lệ, tự sinh drawing1.xml
   *    chứa 2 Text Box chữ ký chuẩn đặt ở cuối mỗi vùng (đã escape &amp; đầy đủ).
   * 5. Đặt <drawing r:id="..."/> vào đúng thứ tự chuẩn CT_Worksheet trong sheet1.xml.
   */
  async function injectFormSignatureShapes(originalFormBuffer, outputXlsxBuffer, zoneFooterRows) {
    const outZip = await JSZip.loadAsync(outputXlsxBuffer);
    let sourceAnchors = [];
    const drawingHeader = '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

    if (originalFormBuffer) {
      try {
        const origZip = await JSZip.loadAsync(originalFormBuffer);

        // Chép xl/theme/theme1.xml từ form gốc (nếu có) để giữ nguyên bảng màu Theme của form người dùng
        const themePath = 'xl/theme/theme1.xml';
        if (origZip.file(themePath)) {
          const themeBytes = await origZip.file(themePath).async('uint8array');
          outZip.file(themePath, themeBytes);
        }

        const drawingFiles = Object.keys(origZip.files).filter(p => /^xl\/drawings\/drawing\d+\.xml$/i.test(p));
        if (drawingFiles.length > 0) {
          const rawDrawingXml = await origZip.file(drawingFiles[0]).async('string');
          // Bỏ các extLst và mc:AlternateContent chứa creationId / thuộc tính mở rộng để không trùng lặp hoặc thiếu namespace
          const cleanedXml = sanitizeRawXmlAmpersands(rawDrawingXml)
            .replace(/<(?:a|xdr):extLst\b[^>]*>[\s\S]*?<\/(?:a|xdr):extLst>/g, '')
            .replace(/<(?:a|xdr):extLst\b[^>]*\/>/g, '')
            .replace(/<mc:AlternateContent\b[^>]*>[\s\S]*?<\/mc:AlternateContent>/g, '');

          const anchorRegex = /<xdr:(?:twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/xdr:(?:twoCellAnchor|oneCellAnchor)>/g;
          const matches = cleanedXml.match(anchorRegex) || [];

          // Chỉ lấy các anchor là Shape/TextBox (<xdr:sp), không lấy <xdr:pic> thiếu media rels
          const parsedAnchors = matches
            .filter(xmlStr => /<xdr:sp\b/.test(xmlStr) && !/<xdr:pic\b/.test(xmlStr))
            .map(xmlStr => {
              // Xóa các thuộc tính r:embed / r:link / r:id nếu có để tránh lỗi thiếu drawing1.xml.rels
              const safeXmlStr = xmlStr.replace(/\s+r:(?:embed|link|id)="[^"]*"/g, '');
              const fromRowMatch = safeXmlStr.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/);
              const fromRow0 = fromRowMatch ? parseInt(fromRowMatch[1], 10) : 0;
              return { xmlStr: safeXmlStr, fromRow0 };
            })
            // Chỉ lấy các khung chữ ký nằm ở phần chân vùng (dòng >= 15, tránh nhầm logo/tiêu đề ở đầu trang)
            .filter(a => a.fromRow0 >= 15);

          if (parsedAnchors.length > 0) {
            parsedAnchors.sort((a, b) => a.fromRow0 - b.fromRow0);
            const firstRow0 = parsedAnchors[0].fromRow0;
            sourceAnchors = parsedAnchors.filter(a => Math.abs(a.fromRow0 - firstRow0) <= 6);
          }
        }
      } catch (e) {
        // Nếu form gốc không đọc được drawing, tự động dùng bộ sinh TextBox chữ ký chuẩn bên dưới
      }
    }

    let allAnchorsXml = '';
    let shapeIdCounter = 100;

    if (sourceAnchors.length > 0) {
      const baseRow0 = sourceAnchors[0].fromRow0;
      (zoneFooterRows || []).forEach((targetRow1Based) => {
        const targetRow0 = Math.max(0, targetRow1Based - 1);
        const delta = targetRow0 - baseRow0;

        sourceAnchors.forEach(anchor => {
          let shifted = anchor.xmlStr
            .replace(/<xdr:from>([\s\S]*?)<\/xdr:from>/g, (m, inner) => {
              return `<xdr:from>${inner.replace(/<xdr:row>(\d+)<\/xdr:row>/, (_, r) => `<xdr:row>${Math.max(0, parseInt(r, 10) + delta)}</xdr:row>`)}</xdr:from>`;
            })
            .replace(/<xdr:to>([\s\S]*?)<\/xdr:to>/g, (m, inner) => {
              return `<xdr:to>${inner.replace(/<xdr:row>(\d+)<\/xdr:row>/, (_, r) => `<xdr:row>${Math.max(0, parseInt(r, 10) + delta)}</xdr:row>`)}</xdr:to>`;
            })
            .replace(/<((?:xdr:)?cNvPr\b[^>]*?\bid=")\d+(")/g, (_, p1, p2) => `${p1}${++shapeIdCounter}${p2}`);
          allAnchorsXml += shifted;
        });
      });
    } else {
      // Sinh 2 Text Box chữ ký chuẩn ở cuối mỗi vùng:
      // 1) "Người báo cáo (QC) / Reported by" (Cột B -> E)
      // 2) "Kho & vận chuyển / Inspected by" (Cột J -> K)
      (zoneFooterRows || []).forEach((targetRow1Based, idx) => {
        const fromRow0 = Math.max(0, targetRow1Based - 1);
        const toRow0 = fromRow0 + 4;
        const id1 = ++shapeIdCounter;
        const id2 = ++shapeIdCounter;

        allAnchorsXml += buildStandardTextBoxAnchor({
          id: id1,
          name: `TextBox_QC_${idx + 1}`,
          fromCol: 1, toCol: 5,
          fromRow: fromRow0, toRow: toRow0,
          line1: 'Người báo cáo (QC) / Reported by',
          line2: '(Ký, ghi rõ họ tên)'
        });

        allAnchorsXml += buildStandardTextBoxAnchor({
          id: id2,
          name: `TextBox_WH_${idx + 1}`,
          fromCol: 9, toCol: 11,
          fromRow: fromRow0, toRow: toRow0,
          line1: 'Kho & vận chuyển / Inspected by',
          line2: '(Ký, ghi rõ họ tên)'
        });
      });
    }

    const finalDrawingXml = '<' + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + `>\n${drawingHeader}${allAnchorsXml}</xdr:wsDr>`;
    outZip.file('xl/drawings/drawing1.xml', finalDrawingXml);

    // 1. Cập nhật [Content_Types].xml
    const ctPath = '[Content_Types].xml';
    if (outZip.file(ctPath)) {
      let ctXml = await outZip.file(ctPath).async('string');
      if (!ctXml.includes('/xl/drawings/drawing1.xml')) {
        ctXml = ctXml.replace(
          '</Types>',
          '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'
        );
        outZip.file(ctPath, ctXml);
      }
    }

    // 2. Cập nhật xl/worksheets/_rels/sheet1.xml.rels
    const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
    let rId = 'rIdDrawing99';
    if (outZip.file(relsPath)) {
      let relsXml = await outZip.file(relsPath).async('string');
      const existingMatch = relsXml.match(/Id="(rId[^"]+)"[^>]*Type="[^"]*\/drawing"/);
      if (existingMatch) {
        rId = existingMatch[1];
      } else {
        relsXml = relsXml.replace(
          '</Relationships>',
          `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`
        );
        outZip.file(relsPath, relsXml);
      }
    } else {
      const newRelsXml = '<' + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + `>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`;
      outZip.file(relsPath, newRelsXml);
    }

    // 3. Đặt thẻ <drawing r:id="..."/> vào xl/worksheets/sheet1.xml đúng chuẩn thứ tự schema CT_Worksheet
    const sheetPath = 'xl/worksheets/sheet1.xml';
    if (outZip.file(sheetPath)) {
      let sheetXml = await outZip.file(sheetPath).async('string');
      if (!sheetXml.includes('xmlns:r=')) {
        sheetXml = sheetXml.replace(
          '<worksheet ',
          '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        );
      }
      const drawingTag = `<drawing r:id="${rId}"/>`;
      sheetXml = insertDrawingTagInWorksheetXml(sheetXml, drawingTag);
      outZip.file(sheetPath, sheetXml);
    }

    return await outZip.generateAsync({ type: 'arraybuffer' });
  }

  function buildStandardTextBoxAnchor({ id, name, fromCol, toCol, fromRow, toRow, line1, line2 }) {
    const safeName = escapeXml(name);
    const safeLine1 = escapeXml(line1);
    const safeLine2 = escapeXml(line2);
    return `<xdr:twoCellAnchor editAs="oneCell">` +
      `<xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>19050</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from>` +
      `<xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
      `<xdr:sp macro="" textlink="">` +
        `<xdr:nvSpPr>` +
          `<xdr:cNvPr id="${id}" name="${safeName}"/>` +
          `<xdr:cNvSpPr txBox="1"/>` +
        `</xdr:nvSpPr>` +
        `<xdr:spPr>` +
          `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
          `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>` +
          `<a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln>` +
        `</xdr:spPr>` +
        `<xdr:txBody>` +
          `<a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" lIns="45720" tIns="22860" rIns="45720" bIns="22860" anchor="t"/>` +
          `<a:lstStyle/>` +
          `<a:p>` +
            `<a:pPr algn="ctr"/>` +
            `<a:r><a:rPr lang="vi-VN" sz="1000" b="1"><a:latin typeface="Times New Roman"/></a:rPr><a:t>${safeLine1}</a:t></a:r>` +
          `</a:p>` +
          `<a:p>` +
            `<a:pPr algn="ctr"/>` +
            `<a:r><a:rPr lang="vi-VN" sz="900" i="1"><a:latin typeface="Times New Roman"/></a:rPr><a:t>${safeLine2}</a:t></a:r>` +
          `</a:p>` +
        `</xdr:txBody>` +
      `</xdr:sp>` +
      `<xdr:clientData/>` +
    `</xdr:twoCellAnchor>`;
  }

  global.FormShapesModule = {
    injectFormSignatureShapes,
    escapeXml
  };
})(window);
