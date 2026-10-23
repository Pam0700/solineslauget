// ==================== MODULE: product-short.js ====================
(function(global) {
  'use strict';

  const DEFAULT_KEYWORDS = ['5kg', '2kg', '1kg', 'TF', '500g', 'Khay', 'Go'];

  // Bỏ chữ VietGap (không phân biệt hoa thường, cả "Viet Gap"), dọn ngoặc rỗng và khoảng trắng thừa
  function cleanProductName(rawName) {
    if (!rawName) return '';
    let s = String(rawName);
    s = s.replace(/viet\s*gap/gi, '');
    // Dọn ngoặc rỗng như (), [], {}, ( - ), ( , )
    s = s.replace(/\(\s*[-,;.]?\s*\)/g, '');
    s = s.replace(/\[\s*[-,;.]?\s*\]/g, '');
    // Dọn dấu gạch ngang thừa ở cuối hoặc đầu
    s = s.replace(/\s+/g, ' ').trim();
    s = s.replace(/[-–—,\s]+$/, '').trim();
    return s;
  }

  // Chuẩn hóa chuỗi để kiểm tra từ khóa quy cách ("2 KG" -> "2kg", "500 gr" -> "500g")
  function normalizeForSortMatch(name) {
    let s = cleanProductName(name).toLowerCase();
    // Chuẩn hóa khoảng trắng giữa số và đơn vị: "2 kg" -> "2kg", "500 gr" -> "500g", "500 gam" -> "500g"
    s = s.replace(/(\d+(?:[.,]\d+)?)\s*(?:kg|kilo|kilogram)\b/gi, '$1kg');
    s = s.replace(/(\d+(?:[.,]\d+)?)\s*(?:gr|gam|gram|g)\b/gi, '$1g');
    return s;
  }

  // Kiểm tra khớp nguyên cụm từ khóa (vd: "5kg" không khớp "15kg" hay "0.5kg")
  function matchesWholeKeyword(normText, keyword) {
    const kw = String(keyword || '').trim().toLowerCase()
      .replace(/(\d+)\s*kg\b/i, '$1kg')
      .replace(/(\d+)\s*(?:gr|gam|g)\b/i, '$1g');
    if (!kw) return false;

    const escaped = kw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    // Nếu từ khóa bắt đầu bằng số (vd 5kg, 500g), phía trước không được là chữ số, dấu chấm hoặc dấu phẩy thập phân
    if (/^\d/.test(kw)) {
      const re = new RegExp(`(?:^|[^0-9.,])${escaped}(?![a-z0-9])`, 'i');
      return re.test(normText);
    } else {
      // Từ khóa chữ (TF, Khay, Go...): biên từ không phải chữ/số
      const re = new RegExp(`(?<![\\p{L}\\p{N}])${escaped}(?![\\p{L}\\p{N}])`, 'iu');
      return re.test(normText);
    }
  }

  // Tính thứ hạng sắp xếp sản phẩm dựa theo tên GỐC (sau khi bỏ VietGap)
  // Riêng sản phẩm có CẢ "khay" và "500g" xếp sau nhóm TF và trước nhóm chỉ-500g / chỉ-khay
  function getProductSortWeight(originalCleanedName, keywordList) {
    const keywords = (keywordList && keywordList.length > 0) ? keywordList : DEFAULT_KEYWORDS;
    const norm = normalizeForSortMatch(originalCleanedName);

    const hasKhay = matchesWholeKeyword(norm, 'khay');
    const has500g = matchesWholeKeyword(norm, '500g');

    // Tìm vị trí của TF, 500g, Khay trong danh sách từ khóa
    const idxTF = keywords.findIndex(k => k.trim().toLowerCase() === 'tf');
    const idx500g = keywords.findIndex(k => k.trim().toLowerCase() === '500g');
    const idxKhay = keywords.findIndex(k => k.trim().toLowerCase() === 'khay');

    // Kiểm tra các từ khóa đứng trước nhóm (TF / 500g / Khay) như 5kg, 2kg, 1kg, TF
    for (let i = 0; i < keywords.length; i++) {
      const kwLower = keywords[i].trim().toLowerCase();
      // Nếu sản phẩm có CẢ khay và 500g, khi duyệt tới 500g hoặc khay thì trả về trọng số đặc biệt ngay sau TF
      if (hasKhay && has500g && (kwLower === '500g' || kwLower === 'khay')) {
        if (idxTF !== -1) {
          return idxTF + 0.5;
        }
        const first500OrKhay = Math.min(
          idx500g !== -1 ? idx500g : 999,
          idxKhay !== -1 ? idxKhay : 999
        );
        return first500OrKhay - 0.5;
      }

      if (matchesWholeKeyword(norm, keywords[i])) {
        return i;
      }
    }

    return 9999;
  }

  // Sắp xếp danh sách đối tượng sản phẩm { rawClean, display }
  // Sắp xếp tính theo rawClean (tên gốc đã bỏ VietGap), nếu trùng sau rút gọn thì gộp
  function sortAndDedupeProducts(rawProductNames, productShortMap, keywordList) {
    const mapByDisplay = new Map();

    for (const raw of rawProductNames) {
      const cleaned = cleanProductName(raw);
      if (!cleaned) continue;

      const shortName = findProductShortName(cleaned, raw, productShortMap);
      const display = shortName || cleaned;
      const weight = getProductSortWeight(cleaned, keywordList);

      if (!mapByDisplay.has(display)) {
        mapByDisplay.set(display, {
          rawClean: cleaned,
          display: display,
          weight: weight
        });
      } else {
        // Nếu gộp trùng tên rút gọn, giữ trọng số nhỏ nhất (ưu tiên cao hơn)
        const existing = mapByDisplay.get(display);
        if (weight < existing.weight) {
          existing.weight = weight;
          existing.rawClean = cleaned;
        }
      }
    }

    const items = Array.from(mapByDisplay.values());
    items.sort((a, b) => {
      if (a.weight !== b.weight) return a.weight - b.weight;
      return a.rawClean.localeCompare(b.rawClean, 'vi', { sensitivity: 'base' });
    });

    return items;
  }

  function normKey(s) {
    return String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  function findCustomerShortName(rawCustomer, rawAddress, customerShortMap) {
    if (!customerShortMap) return '';
    const key = `${normKey(rawCustomer)}|||${normKey(rawAddress)}`;
    return customerShortMap.get(key) || '';
  }

  function findProductShortName(cleanedProduct, rawProduct, productShortMap) {
    if (!productShortMap) return '';
    return productShortMap.get(normKey(cleanedProduct)) || productShortMap.get(normKey(rawProduct)) || '';
  }

  // Xây dựng Map tra cứu nhanh từ mảng quy ước
  function buildShortMaps(customerRules, productRules) {
    const customerMap = new Map();
    const productMap = new Map();

    (customerRules || []).forEach(r => {
      const shortVal = String(r.shortName || '').trim();
      if (shortVal) {
        const key = `${normKey(r.customer)}|||${normKey(r.address)}`;
        customerMap.set(key, shortVal);
      }
    });

    (productRules || []).forEach(r => {
      const shortVal = String(r.shortProduct || '').trim();
      if (shortVal) {
        productMap.set(normKey(r.product), shortVal);
        productMap.set(normKey(cleanProductName(r.product)), shortVal);
      }
    });

    return { customerMap, productMap };
  }

  // Xử lý gom nhóm khách hàng, sản phẩm và chia vùng First-Fit
  function processDeliveryData(options) {
    const {
      dg1Records,
      productRecords,
      selectedCodesSet,
      customerRules,
      productRules,
      keywords,
      maxRowsPerZone = 25,
      onlyProductMode = false
    } = options;

    const { customerMap, productMap } = buildShortMaps(customerRules, productRules);

    // Map Code -> danh sách sản phẩm
    const codeToProducts = new Map();
    (productRecords || []).forEach(r => {
      const c = String(r.code || '').trim();
      if (!c) return;
      if (!codeToProducts.has(c)) codeToProducts.set(c, []);
      codeToProducts.get(c).push(r.product);
    });

    // Map Code -> dòng dg1
    const codeToDg1 = new Map();
    (dg1Records || []).forEach(r => {
      const c = String(r.code || '').trim();
      if (!c) return;
      if (!codeToDg1.has(c)) codeToDg1.set(c, r);
    });

    const warnings = [];
    const groupedCustomers = new Map();
    const allRawProductsForOnlyMode = [];

    selectedCodesSet.forEach(code => {
      const dg1 = codeToDg1.get(code);
      const prods = codeToProducts.get(code) || [];

      if (!dg1) {
        warnings.push(`Code "${code}" không có trong file giao hàng (dg1).`);
      }
      if (prods.length === 0) {
        warnings.push(`Code "${code}" không tìm thấy sản phẩm nào trong file 50/54.`);
        return;
      }

      if (onlyProductMode) {
        allRawProductsForOnlyMode.push(...prods);
        return;
      }

      const rawCustomer = dg1 ? dg1.customer : `Code ${code}`;
      const rawAddress = dg1 ? dg1.address : '';
      const shortCust = dg1 ? findCustomerShortName(rawCustomer, rawAddress, customerMap) : '';

      // Nếu có quy ước rút gọn khách hàng -> chỉ hiện tên rút gọn, các khách cùng tên rút gọn sẽ gộp làm 1
      let groupKey, displayLabel;
      if (shortCust) {
        groupKey = `SHORT:::${normKey(shortCust)}`;
        displayLabel = shortCust;
      } else {
        groupKey = `RAW:::${normKey(rawCustomer)}|||${normKey(rawAddress)}`;
        displayLabel = rawAddress ? `${rawCustomer}\n${rawAddress}` : rawCustomer;
      }

      if (!groupedCustomers.has(groupKey)) {
        groupedCustomers.set(groupKey, {
          groupKey,
          displayLabel,
          rawCustomer,
          rawAddress,
          hasShortRule: Boolean(shortCust),
          codes: [],
          rawProducts: []
        });
      }
      const grp = groupedCustomers.get(groupKey);
      grp.codes.push(code);
      grp.rawProducts.push(...prods);
    });

    // Chế độ chỉ lấy sản phẩm
    if (onlyProductMode) {
      const sortedProds = sortAndDedupeProducts(allRawProductsForOnlyMode, productMap, keywords);
      const zones = [];
      for (let i = 0; i < sortedProds.length; i += maxRowsPerZone) {
        const chunk = sortedProds.slice(i, i + maxRowsPerZone);
        zones.push({
          zoneIndex: zones.length + 1,
          rowCount: chunk.length,
          isOversize: false,
          customers: [{
            displayLabel: '',
            products: chunk.map(p => p.display)
          }]
        });
      }
      return { zones, warnings, customersList: [], onlyProductMode: true, totalProducts: sortedProds.length };
    }

    // Chế độ đầy đủ Khách hàng + Sản phẩm
    const customersList = [];
    groupedCustomers.forEach(grp => {
      const sortedProds = sortAndDedupeProducts(grp.rawProducts, productMap, keywords);
      if (sortedProds.length > 0) {
        customersList.push({
          displayLabel: grp.displayLabel,
          rawCustomer: grp.rawCustomer,
          rawAddress: grp.rawAddress,
          hasShortRule: grp.hasShortRule,
          codes: grp.codes,
          products: sortedProds.map(p => p.display)
        });
      }
    });

    // Chia vùng First-Fit cho Xuất Excel chuẩn (không tách khách giữa 2 vùng; khách > N chiếm riêng 1 vùng và báo)
    const zones = [];
    customersList.forEach(cust => {
      const count = cust.products.length;
      if (count > maxRowsPerZone) {
        warnings.push(`Khách hàng "${cust.displayLabel.replace(/\n/g, ' - ')}" có ${count} sản phẩm (vượt tối đa ${maxRowsPerZone} dòng/vùng) nên được xếp riêng 1 vùng.`);
        zones.push({
          zoneIndex: zones.length + 1,
          rowCount: count,
          isOversize: true,
          customers: [cust]
        });
      } else {
        let placed = false;
        for (let z = 0; z < zones.length; z++) {
          if (!zones[z].isOversize && (zones[z].rowCount + count) <= maxRowsPerZone) {
            zones[z].customers.push(cust);
            zones[z].rowCount += count;
            placed = true;
            break;
          }
        }
        if (!placed) {
          zones.push({
            zoneIndex: zones.length + 1,
            rowCount: count,
            isOversize: false,
            customers: [cust]
          });
        }
      }
    });

    return { zones, warnings, customersList, onlyProductMode: false };
  }

  // Xuất file Excel mẫu Quy ước rút gọn (liệt kê toàn bộ mục trong dữ liệu đang tải, điền sẵn mục đã có, ô trống tô vàng nhạt)
  async function exportShortRuleSampleExcel(type, dg1Records, productRecords, customerRules, productRules) {
    const { customerMap, productMap } = buildShortMaps(customerRules, productRules);
    const wb = new ExcelJS.Workbook();

    if (type === 'customer') {
      const ws = wb.addWorksheet('KhachHang');
      ws.columns = [
        { header: 'Tên khách hàng (gốc)', key: 'customer', width: 32 },
        { header: 'Điểm giao (gốc)', key: 'address', width: 42 },
        { header: 'Khách hàng (rút gọn)', key: 'shortName', width: 30 }
      ];

      const seen = new Set();
      const rowsToWrite = [];

      // Lấy các mục từ dữ liệu đang tải
      (dg1Records || []).forEach(r => {
        const c = String(r.customer || '').trim();
        const a = String(r.address || '').trim();
        if (!c && !a) return;
        const key = `${normKey(c)}|||${normKey(a)}`;
        if (!seen.has(key)) {
          seen.add(key);
          rowsToWrite.push({
            customer: c,
            address: a,
            shortName: customerMap.get(key) || ''
          });
        }
      });

      // Bổ sung các mục đã có trong quy ước nếu chưa nằm trong file đang tải
      (customerRules || []).forEach(r => {
        const c = String(r.customer || '').trim();
        const a = String(r.address || '').trim();
        if (!c && !a) return;
        const key = `${normKey(c)}|||${normKey(a)}`;
        if (!seen.has(key)) {
          seen.add(key);
          rowsToWrite.push({ customer: c, address: a, shortName: String(r.shortName || '').trim() });
        }
      });

      styleSampleSheet(ws, rowsToWrite, 3);
    } else {
      const ws = wb.addWorksheet('SanPham');
      ws.columns = [
        { header: 'Sản phẩm (gốc)', key: 'product', width: 45 },
        { header: 'Sản phẩm rút gọn', key: 'shortProduct', width: 35 }
      ];

      const seen = new Set();
      const rowsToWrite = [];

      (productRecords || []).forEach(r => {
        const p = cleanProductName(r.product);
        if (!p) return;
        const key = normKey(p);
        if (!seen.has(key)) {
          seen.add(key);
          rowsToWrite.push({
            product: p,
            shortProduct: productMap.get(key) || ''
          });
        }
      });

      (productRules || []).forEach(r => {
        const p = cleanProductName(r.product);
        if (!p) return;
        const key = normKey(p);
        if (!seen.has(key)) {
          seen.add(key);
          rowsToWrite.push({ product: p, shortProduct: String(r.shortProduct || '').trim() });
        }
      });

      styleSampleSheet(ws, rowsToWrite, 2);
    }

    return await wb.xlsx.writeBuffer();
  }

  function styleSampleSheet(ws, rows, shortColIndex) {
    const headerRow = ws.getRow(1);
    headerRow.font = { name: 'Arial', size: 11, bold: true };
    headerRow.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
    });

    rows.forEach(r => {
      const row = ws.addRow(r);
      row.font = { name: 'Arial', size: 11 };
      for (let c = 1; c <= shortColIndex; c++) {
        const cell = row.getCell(c);
        cell.numFmt = '@';
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' },
          bottom: { style: 'thin' }, right: { style: 'thin' }
        };
        if (c === shortColIndex && !String(cell.value || '').trim()) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
        }
      }
    });
  }

  // Đọc file Excel quy ước rút gọn do người dùng tải lên
  async function parseShortRuleExcel(file, expectedType) {
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];
    if (!ws) throw new Error('File Excel quy ước không có trang tính.');

    const result = [];
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return; // Bỏ dòng tiêu đề
      const c1 = String(row.getCell(1).text || row.getCell(1).value || '').trim();
      const c2 = String(row.getCell(2).text || row.getCell(2).value || '').trim();
      const c3 = String(row.getCell(3).text || row.getCell(3).value || '').trim();

      if (expectedType === 'customer') {
        if (c1 || c2) {
          result.push({ customer: c1, address: c2, shortName: c3 });
        }
      } else {
        if (c1) {
          result.push({ product: c1, shortProduct: c2 });
        }
      }
    });
    return result;
  }

  global.ProductShortModule = {
    DEFAULT_KEYWORDS,
    cleanProductName,
    getProductSortWeight,
    sortAndDedupeProducts,
    processDeliveryData,
    exportShortRuleSampleExcel,
    parseShortRuleExcel
  };
})(window);
