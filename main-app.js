// ==================== MODULE: main-app.js (Móc nối giao diện) ====================
(function(global) {
  'use strict';

  const state = {
    dg1Records: [],
    productRecords: [],
    selectedShips: new Set(),
    selectedCodes: new Set(),
    processedResult: null,

    // Dữ liệu file Điều xe
    dispatchFileName: '',
    dispatchSheets: [],
    activeDispatchSheetIdx: 0,
    selectedDispatchStts: new Set(),

    // Quy ước & Form hệ thống
    rulesVersion: 1,
    customerRules: [],
    productRules: [],
    systemTemplateBuffer: null,
    customTemplateBuffer: null,
    templateMeta: { fileName: 'Form_Bao_Cao_QC_Mac_Dinh.xlsx', zoneCount: 2 },

    // Chuyển file gốc
    convertedFiles: [],

    // Quyền Admin
    isAdminRoute: false,
    isAdminAuthed: false,
    adminActiveTab: 'customer', // 'customer' | 'product'
    adminPage: 1,
    adminPageSize: 15
  };

  function showToast(msg, duration = 2600) {
    const el = document.getElementById('toastBar');
    if (!el) return;
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(el._timer);
    el._timer = setTimeout(() => el.classList.remove('show'), duration);
  }

  function triggerDownloadBuffer(buffer, fileName, mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
    let cleanName = String(fileName || 'ket_qua').trim();
    if (!/\.(xlsx|zip)$/i.test(cleanName)) cleanName += '.xlsx';
    const blob = new Blob([buffer], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = cleanName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 300);
  }

  // Điều hướng Menu ngăn kéo ẩn
  const VIEW_TITLES = {
    'view-upload': '1. Tải dữ liệu',
    'view-select': '2. Chọn & xử lý',
    'view-result': '3. Kết quả',
    'view-form': '4. Điền form báo cáo',
    'view-dispatch': '5. Xuất theo Điều xe (.zip)',
    'view-converter': 'Công cụ: Chuyển file gốc',
    'view-shortrules': 'Công cụ: Quy ước rút gọn',
    'view-admin-rules': 'Quản trị: Quản lý quy ước',
    'view-admin-form': 'Quản trị: Form báo cáo hệ thống'
  };

  function toggleDrawer(forceOpen) {
    const drawer = document.getElementById('sidebarDrawer');
    const backdrop = document.getElementById('drawerBackdrop');
    const isOpen = forceOpen !== undefined ? forceOpen : !drawer.classList.contains('open');
    if (isOpen) {
      drawer.classList.add('open');
      backdrop.classList.add('open');
    } else {
      drawer.classList.remove('open');
      backdrop.classList.remove('open');
    }
  }

  function switchView(viewId) {
    document.querySelectorAll('.view-panel').forEach(p => {
      p.classList.toggle('active', p.id === viewId);
    });
    document.querySelectorAll('.nav-item').forEach(btn => {
      btn.classList.toggle('active', btn.getAttribute('data-view') === viewId);
    });
    const titleEl = document.getElementById('currentViewTitle');
    if (titleEl && VIEW_TITLES[viewId]) {
      titleEl.textContent = VIEW_TITLES[viewId];
    }
    toggleDrawer(false);
  }

  function updateStepTicks() {
    const hasUpload = state.dg1Records.length > 0 && state.productRecords.length > 0;
    const hasSelect = getEffectiveSelectedCodesSet().size > 0;
    const hasResult = Boolean(state.processedResult && state.processedResult.zones && state.processedResult.zones.length > 0);
    const hasDispatch = state.dispatchSheets.length > 0;

    document.getElementById('tickStep1').classList.toggle('done', hasUpload);
    document.getElementById('tickStep2').classList.toggle('done', hasSelect);
    document.getElementById('tickStep3').classList.toggle('done', hasResult);
    const tick5 = document.getElementById('tickStep5');
    if (tick5) tick5.classList.toggle('done', hasDispatch);
  }

  // Kiểm tra đường dẫn Admin (?page=admin hoặc /admin hoặc nút chuyển nhanh)
  function applyAdminRouteUI() {
    const adminGroup = document.getElementById('adminNavGroup');
    const btnSync = document.getElementById('btnAdminSyncFromLoaded');
    const badge = document.getElementById('btnSwitchAdminMode');

    if (state.isAdminRoute) {
      adminGroup.style.display = 'block';
      btnSync.style.display = state.isAdminAuthed ? 'inline-flex' : 'none';
      badge.textContent = state.isAdminAuthed ? 'Chế độ: Admin (Đã ĐN)' : 'Chế độ: Admin';
      badge.classList.add('active-admin');
    } else {
      adminGroup.style.display = 'none';
      btnSync.style.display = 'none';
      badge.textContent = 'Chế độ: Thường';
      badge.classList.remove('active-admin');
    }

    const loginCard = document.getElementById('adminLoginCard');
    const manageCard = document.getElementById('adminManageRulesCard');
    if (loginCard && manageCard) {
      loginCard.style.display = state.isAdminAuthed ? 'none' : 'block';
      manageCard.style.display = state.isAdminAuthed ? 'block' : 'none';
    }
  }

  function updateRuleCountersUI() {
    document.getElementById('countCustRulesLabel').textContent = state.customerRules.length;
    document.getElementById('countProdRulesLabel').textContent = state.productRules.length;
    document.getElementById('sysRuleStatus').textContent = `Quy ước: ${state.customerRules.length} KH, ${state.productRules.length} SP`;
    document.getElementById('sysRuleStatus').classList.add('ok');
    document.getElementById('adminVersionBadge').textContent = `Phiên bản: v${state.rulesVersion}`;
  }

  // ==================== BƯỚC 1: TẢI DỮ LIỆU ====================
  async function handleDg1Upload(fileList) {
    if (!fileList || fileList.length === 0) return;
    try {
      const parsed = await global.ConverterModule.parseUploadedFile(fileList[0]);
      if (parsed.type !== 'dg1') {
        showToast('Cảnh báo: File vừa chọn có cấu trúc Sản phẩm, không phải dg1!');
      }
      state.dg1Records = parsed.records || [];
      document.getElementById('dg1StatusText').textContent = `Đã nạp ${state.dg1Records.length} dòng từ "${parsed.fileName}" (${parsed.format}).`;
      rebuildShipAndCodeLists();
      renderDispatchTable();
      updateStepTicks();
      showToast(`Đã nạp file dg1 (${state.dg1Records.length} dòng)`);
    } catch (err) {
      showToast(err.message || 'Lỗi đọc file dg1');
    }
  }

  async function handleProductFilesUpload(fileList) {
    if (!fileList || fileList.length === 0) return;
    let totalLoaded = 0;
    const names = [];
    const combined = [];
    for (const file of Array.from(fileList)) {
      try {
        const parsed = await global.ConverterModule.parseUploadedFile(file);
        combined.push(...(parsed.records || []));
        totalLoaded += (parsed.records || []).length;
        names.push(`${parsed.fileName} (${parsed.records.length})`);
      } catch (err) {
        showToast(`Lỗi đọc ${file.name}: ${err.message}`);
      }
    }
    if (combined.length > 0) {
      state.productRecords.push(...combined);
      document.getElementById('prodStatusText').textContent = `Đã nạp tổng cộng ${state.productRecords.length} dòng sản phẩm: ${names.join(', ')}`;
      rebuildShipAndCodeLists();
      renderDispatchTable();
      updateStepTicks();
      showToast(`Đã nạp thêm ${totalLoaded} dòng sản phẩm`);
    }
  }

  // ==================== BƯỚC 2: DANH SÁCH SHIP & CODE ====================
  function getUniqueShips() {
    const map = new Map();
    state.dg1Records.forEach(r => {
      const ship = String(r.ship || '').trim();
      if (!ship) return;
      if (!map.has(ship)) map.set(ship, []);
      map.get(ship).push(r.code);
    });
    return map;
  }

  function getUniqueCodesInfo() {
    const map = new Map();
    state.dg1Records.forEach(r => {
      const code = String(r.code || '').trim();
      if (!code) return;
      if (!map.has(code)) {
        map.set(code, {
          code,
          customer: r.customer || '',
          address: r.address || '',
          ship: r.ship || ''
        });
      }
    });
    state.productRecords.forEach(r => {
      const code = String(r.code || '').trim();
      if (code && !map.has(code)) {
        map.set(code, {
          code,
          customer: '(Không có trong dg1)',
          address: '',
          ship: ''
        });
      }
    });
    return Array.from(map.values());
  }

  function rebuildShipAndCodeLists() {
    renderShipList();
    renderCodeList();
    updateSelectionCounters();
  }

  function renderShipList() {
    const container = document.getElementById('shipListContainer');
    const query = String(document.getElementById('searchShipInput').value || '').trim().toLowerCase();
    const shipMap = getUniqueShips();
    container.innerHTML = '';

    if (shipMap.size === 0) {
      container.innerHTML = '<div style="padding:12px; color:#64748b;">Chưa có dữ liệu Ship.</div>';
      return;
    }

    shipMap.forEach((codes, ship) => {
      if (query && !ship.toLowerCase().includes(query)) return;
      const row = document.createElement('label');
      row.className = 'list-item';
      const chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.checked = state.selectedShips.has(ship);
      chk.addEventListener('change', () => {
        if (chk.checked) state.selectedShips.add(ship);
        else state.selectedShips.delete(ship);
        updateSelectionCounters();
        updateStepTicks();
      });

      const info = document.createElement('div');
      info.innerHTML = `<div class="code-line-main">${escapeHtml(ship)}</div><div class="code-line-sub">${codes.length} mã Code</div>`;
      row.appendChild(chk);
      row.appendChild(info);
      container.appendChild(row);
    });
  }

  function renderCodeList() {
    const container = document.getElementById('codeListContainer');
    const query = String(document.getElementById('searchCodeInput').value || '').trim().toLowerCase();
    const codesInfo = getUniqueCodesInfo();
    container.innerHTML = '';

    if (codesInfo.length === 0) {
      container.innerHTML = '<div style="padding:12px; color:#64748b;">Chưa có dữ liệu Code.</div>';
      return;
    }

    codesInfo.forEach(item => {
      const searchTarget = `${item.code} ${item.customer} ${item.address} ${item.ship}`.toLowerCase();
      if (query && !searchTarget.includes(query)) return;

      const row = document.createElement('label');
      row.className = 'list-item';
      const chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.checked = state.selectedCodes.has(item.code);
      chk.addEventListener('change', () => {
        if (chk.checked) state.selectedCodes.add(item.code);
        else state.selectedCodes.delete(item.code);
        updateSelectionCounters();
        updateStepTicks();
      });

      const info = document.createElement('div');
      const shipTag = item.ship ? ` [Ship: ${escapeHtml(item.ship)}]` : ' [Không có Ship]';
      info.innerHTML = `
        <div class="code-line-main">${escapeHtml(item.code)} — ${escapeHtml(item.customer)}</div>
        <div class="code-line-sub">${escapeHtml(item.address || 'Không có điểm giao')}${shipTag}</div>
      `;
      row.appendChild(chk);
      row.appendChild(info);
      container.appendChild(row);
    });
  }

  function updateSelectionCounters() {
    document.getElementById('countSelectedShips').textContent = state.selectedShips.size;
    document.getElementById('countSelectedCodes').textContent = state.selectedCodes.size;
  }

  function getEffectiveSelectedCodesSet() {
    const finalCodes = new Set();
    const shipMap = getUniqueShips();

    state.selectedShips.forEach(ship => {
      const codes = shipMap.get(ship) || [];
      codes.forEach(c => finalCodes.add(c));
    });

    state.selectedCodes.forEach(c => finalCodes.add(c));

    const pastedText = String(document.getElementById('pasteCodesInput').value || '');
    pastedText.split(/[\s,;]+/).forEach(tok => {
      const clean = tok.trim();
      if (clean) finalCodes.add(clean);
    });

    return finalCodes;
  }

  function getProcessingOptionsFromUI() {
    const rawKw = String(document.getElementById('sortKeywordsInput').value || '');
    const keywords = rawKw.split(',').map(s => s.trim()).filter(Boolean);
    const maxRows = Math.max(5, parseInt(document.getElementById('maxRowsPerZoneInput').value || '25', 10));
    const onlyProductMode = document.getElementById('chkOnlyProductMode').checked;
    return { keywords, maxRows, onlyProductMode };
  }

  // ==================== BƯỚC 3: XỬ LÝ & HIỂN THỊ KẾT QUẢ ====================
  function runProcessAndGoToResult() {
    const selectedCodesSet = getEffectiveSelectedCodesSet();
    if (selectedCodesSet.size === 0) {
      showToast('Vui lòng chọn ít nhất 1 Ship, 1 Code hoặc dán mã Code!');
      return;
    }

    const { keywords, maxRows, onlyProductMode } = getProcessingOptionsFromUI();

    const result = global.ProductShortModule.processDeliveryData({
      dg1Records: state.dg1Records,
      productRecords: state.productRecords,
      selectedCodesSet,
      customerRules: state.customerRules,
      productRules: state.productRules,
      keywords,
      maxRowsPerZone: maxRows,
      onlyProductMode
    });

    state.processedResult = result;
    renderProcessedResultView(result);
    updateStepTicks();
    switchView('view-result');
  }

  function renderProcessedResultView(result) {
    const warnBox = document.getElementById('resultWarningsBox');
    warnBox.innerHTML = '';
    if (result.warnings && result.warnings.length > 0) {
      result.warnings.forEach(w => {
        const d = document.createElement('div');
        d.textContent = '⚠ ' + w;
        warnBox.appendChild(d);
      });
      warnBox.classList.add('show');
    } else {
      warnBox.classList.remove('show');
    }

    const totalZones = (result.zones || []).length;
    const totalRows = (result.zones || []).reduce((acc, z) => acc + z.rowCount, 0);
    document.getElementById('resultSummaryPill').textContent =
      `Đã xếp ${totalRows} dòng sản phẩm vào ${totalZones} vùng (Tối đa ${document.getElementById('maxRowsPerZoneInput').value} dòng/vùng)`;

    const thead = document.getElementById('previewResultThead');
    const tbody = document.getElementById('previewResultTbody');
    tbody.innerHTML = '';

    if (result.onlyProductMode) {
      thead.innerHTML = '<tr><th style="width:60px;">Vùng</th><th>Sản phẩm</th></tr>';
    } else {
      thead.innerHTML = '<tr><th style="width:60px;">Vùng</th><th style="width:36%;">Khách hàng</th><th>Sản phẩm</th></tr>';
    }

    if (totalZones === 0) {
      tbody.innerHTML = `<tr><td colspan="${result.onlyProductMode ? 2 : 3}" style="text-align:center; color:#64748b;">Không có sản phẩm nào khớp với các Code đã chọn.</td></tr>`;
      return;
    }

    result.zones.forEach((zone, zIdx) => {
      zone.customers.forEach(cust => {
        const prods = cust.products || [];
        prods.forEach((prodName, pIdx) => {
          const tr = document.createElement('tr');
          if (result.onlyProductMode) {
            tr.innerHTML = `
              <td style="color:#64748b; font-size:12px;">V${zone.zoneIndex}</td>
              <td>${escapeHtml(prodName)}</td>
            `;
          } else {
            let custCellHtml = '';
            if (pIdx === 0) {
              custCellHtml = `<td class="pre-line" rowspan="${prods.length}" style="font-weight:500;">${escapeHtml(cust.displayLabel)}</td>`;
            }
            tr.innerHTML = `
              <td style="color:#64748b; font-size:12px;">V${zone.zoneIndex}</td>
              ${custCellHtml}
              <td>${escapeHtml(prodName)}</td>
            `;
          }
          tbody.appendChild(tr);
        });
      });

      if (zIdx < result.zones.length - 1) {
        const sepTr = document.createElement('tr');
        sepTr.className = 'zone-sep';
        sepTr.innerHTML = `<td colspan="${result.onlyProductMode ? 2 : 3}"></td>`;
        tbody.appendChild(sepTr);
      }
    });
  }

  // ==================== BƯỚC 5: XỬ LÝ FILE ĐIỀU XE & XUẤT HÀNG LOẠT (.ZIP) ====================
  async function handleDispatchFileUpload(fileList, switchAfterLoad = false) {
    if (!fileList || fileList.length === 0) return;
    const file = fileList[0];
    try {
      const parsed = await global.ConverterModule.parseDispatchFile(file);
      state.dispatchFileName = parsed.fileName;
      state.dispatchSheets = parsed.sheets || [];
      state.activeDispatchSheetIdx = 0;

      // Chọn sheet đầu tiên có ít nhất 1 dòng Shipment
      const firstNonEmptyIdx = state.dispatchSheets.findIndex(s => s.records && s.records.length > 0);
      if (firstNonEmptyIdx !== -1) {
        state.activeDispatchSheetIdx = firstNonEmptyIdx;
      }

      populateDispatchSheetSelect();
      autoSelectDispatchRowsForCurrentSheet();
      renderDispatchTable();
      updateStepTicks();

      const activeSheet = state.dispatchSheets[state.activeDispatchSheetIdx];
      const msg = `Đã nạp "${parsed.fileName}": ${state.dispatchSheets.length} trang tính (Sheet "${activeSheet.sheetName}" có ${activeSheet.records.length} chuyến có số Ship).`;
      document.getElementById('dispatchStep1StatusText').textContent = msg;
      document.getElementById('dispatchFileInfoText').textContent = msg;
      showToast(`Đã nạp file Điều xe (${activeSheet.records.length} chuyến có Shipment)`);

      if (switchAfterLoad) {
        switchView('view-dispatch');
      }
    } catch (err) {
      showToast(err.message || 'Lỗi đọc file Điều xe', 3500);
    }
  }

  function populateDispatchSheetSelect() {
    const sel = document.getElementById('selectDispatchSheet');
    sel.innerHTML = '';
    if (!state.dispatchSheets || state.dispatchSheets.length === 0) {
      sel.innerHTML = '<option value="">-- Chưa có file Điều xe --</option>';
      sel.disabled = true;
      return;
    }

    sel.disabled = false;
    state.dispatchSheets.forEach((sh, idx) => {
      const opt = document.createElement('option');
      opt.value = String(idx);
      const dateStr = sh.dateLabel ? ` (${sh.dateLabel})` : '';
      opt.textContent = `Sheet "${sh.sheetName}"${dateStr} — ${sh.records.length} chuyến có Shipment`;
      if (idx === state.activeDispatchSheetIdx) opt.selected = true;
      sel.appendChild(opt);
    });

    const activeSh = state.dispatchSheets[state.activeDispatchSheetIdx];
    if (activeSh) {
      const cleanSheet = String(activeSh.sheetName || '').replace(/[^\w.-]+/g, '_');
      document.getElementById('dispatchZipFileNameInput').value = `Bao_cao_QC_Dieu_xe_Ngay_${cleanSheet}.zip`;
    }
  }

  function getActiveDispatchSheetRecords() {
    const sh = state.dispatchSheets[state.activeDispatchSheetIdx];
    return sh && sh.records ? sh.records : [];
  }

  function computeDispatchRowMatchInfo(dispatchRow) {
    const { matchedCodesSet, matchedShipsSet } = global.ConverterModule.matchShipTokensToDg1Codes(
      dispatchRow.ships,
      state.dg1Records
    );

    if (matchedCodesSet.size === 0) {
      return {
        matchedCodesSet,
        matchedShipsSet,
        hasMatch: false,
        customerCount: 0,
        productCount: 0,
        processedData: null
      };
    }

    const { keywords, maxRows, onlyProductMode } = getProcessingOptionsFromUI();
    const processedData = global.ProductShortModule.processDeliveryData({
      dg1Records: state.dg1Records,
      productRecords: state.productRecords,
      selectedCodesSet: matchedCodesSet,
      customerRules: state.customerRules,
      productRules: state.productRules,
      keywords,
      maxRowsPerZone: maxRows,
      onlyProductMode
    });

    let customerCount = 0;
    let productCount = 0;
    (processedData.zones || []).forEach(z => {
      customerCount += (z.customers || []).length;
      productCount += z.rowCount || 0;
    });

    return {
      matchedCodesSet,
      matchedShipsSet,
      hasMatch: productCount > 0 || matchedCodesSet.size > 0,
      customerCount,
      productCount,
      processedData
    };
  }

  function autoSelectDispatchRowsForCurrentSheet() {
    state.selectedDispatchStts.clear();
    const records = getActiveDispatchSheetRecords();
    let matchedCount = 0;
    records.forEach(r => {
      const info = computeDispatchRowMatchInfo(r);
      if (info.hasMatch && info.productCount > 0) {
        state.selectedDispatchStts.add(String(r.stt));
        matchedCount++;
      }
    });
    // Nếu chưa nạp dg1 của ngày đó (0 chuyến khớp), mặc định chọn tất cả chuyến có Shipment
    if (matchedCount === 0) {
      records.forEach(r => state.selectedDispatchStts.add(String(r.stt)));
    }
  }

  function renderDispatchTable() {
    const tbody = document.getElementById('dispatchTableTbody');
    const summaryPill = document.getElementById('dispatchSummaryPill');
    const btnZip = document.getElementById('btnExportDispatchZip');
    if (!tbody) return;

    const records = getActiveDispatchSheetRecords();
    if (records.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="7" style="text-align:center; color:#64748b;">
            Chưa có dữ liệu Điều xe. Hãy tải file Điều xe (.xlsx) lên hoặc bấm "Nạp dữ liệu kiểm thử" ở Bước 1.
          </td>
        </tr>
      `;
      summaryPill.textContent = 'Chưa tải file Điều xe';
      btnZip.disabled = true;
      return;
    }

    const query = String(document.getElementById('searchDispatchInput').value || '').trim().toLowerCase();
    const onlyMatched = document.getElementById('chkDispatchOnlyMatched').checked;

    let totalMatchedRows = 0;
    const enriched = records.map(r => {
      const matchInfo = computeDispatchRowMatchInfo(r);
      if (matchInfo.hasMatch && matchInfo.productCount > 0) totalMatchedRows++;
      return { row: r, matchInfo };
    });

    summaryPill.textContent = `Có ${records.length} chuyến có số Ship | Khớp dg1: ${totalMatchedRows} chuyến | Đang chọn: ${state.selectedDispatchStts.size} file`;
    btnZip.disabled = state.selectedDispatchStts.size === 0;

    const visibleItems = enriched.filter(({ row, matchInfo }) => {
      if (onlyMatched && !(matchInfo.hasMatch && matchInfo.productCount > 0)) return false;
      if (query) {
        const hay = `${row.stt} ${row.shipmentRaw} ${row.ships.join(' ')} ${row.truck} ${row.route} ${row.dispatchCustomersSummary}`.toLowerCase();
        if (!hay.includes(query)) return false;
      }
      return true;
    });

    tbody.innerHTML = '';
    if (visibleItems.length === 0) {
      tbody.innerHTML = `<tr><td colspan="7" style="text-align:center; color:#64748b;">Không có chuyến xe nào khớp bộ lọc hiện tại.</td></tr>`;
      return;
    }

    visibleItems.forEach(({ row, matchInfo }) => {
      const sttKey = String(row.stt);
      const tr = document.createElement('tr');
      const isChecked = state.selectedDispatchStts.has(sttKey);

      const shipsBadges = row.ships.map(s => `<span class="ship-chip">${escapeHtml(s)}</span>`).join('');
      const rawNote = row.ships.length > 1 ? `<div class="code-line-sub">Gốc: ${escapeHtml(row.shipmentRaw)} (${row.ships.length} ship)</div>` : '';

      let matchHtml = '';
      if (matchInfo.hasMatch && matchInfo.productCount > 0) {
        const matchedShipList = Array.from(matchInfo.matchedShipsSet).join(', ');
        matchHtml = `
          <span class="status-pill ok">✓ ${matchInfo.matchedCodesSet.size} Code • ${matchInfo.customerCount} KH • ${matchInfo.productCount} SP</span>
          <div class="code-line-sub">Ship dg1: ${escapeHtml(matchedShipList)}</div>
        `;
      } else if (matchInfo.matchedCodesSet.size > 0) {
        matchHtml = `<span class="status-pill warn">Có ${matchInfo.matchedCodesSet.size} Code (chưa có SP)</span>`;
      } else {
        matchHtml = `<span class="status-pill">Chưa có Ship ${escapeHtml(row.ships.join('+'))} trong dg1</span>`;
      }

      tr.innerHTML = `
        <td style="text-align:center;">
          <input type="checkbox" data-dispatch-stt="${escapeAttr(sttKey)}" ${isChecked ? 'checked' : ''}>
        </td>
        <td>
          <strong style="color:#0f172a; font-size:13.5px;">${escapeHtml(row.stt)}.xlsx</strong>
          <div class="code-line-sub">STT: ${escapeHtml(row.stt)}</div>
        </td>
        <td>
          <div>${shipsBadges}</div>
          ${rawNote}
        </td>
        <td>
          <div class="code-line-main">${escapeHtml(row.truck || '—')}</div>
          <div class="code-line-sub">Tuyến: ${escapeHtml(row.route || '—')}</div>
        </td>
        <td>
          <div style="font-size:12px; color:#334155;">${escapeHtml(row.dispatchCustomersSummary || '—')}</div>
        </td>
        <td>${matchHtml}</td>
        <td style="text-align:center;">
          <button type="button" class="btn btn-sm btn-primary" data-dl-single-stt="${escapeAttr(sttKey)}">
            Tải ${escapeHtml(row.stt)}.xlsx
          </button>
        </td>
      `;
      tbody.appendChild(tr);
    });

    tbody.querySelectorAll('input[data-dispatch-stt]').forEach(chk => {
      chk.addEventListener('change', () => {
        const stt = chk.getAttribute('data-dispatch-stt');
        if (chk.checked) state.selectedDispatchStts.add(stt);
        else state.selectedDispatchStts.delete(stt);
        summaryPill.textContent = `Có ${records.length} chuyến có số Ship | Khớp dg1: ${totalMatchedRows} chuyến | Đang chọn: ${state.selectedDispatchStts.size} file`;
        btnZip.disabled = state.selectedDispatchStts.size === 0;
      });
    });

    tbody.querySelectorAll('button[data-dl-single-stt]').forEach(btn => {
      btn.addEventListener('click', async () => {
        const stt = btn.getAttribute('data-dl-single-stt');
        const targetRow = records.find(r => String(r.stt) === String(stt));
        if (!targetRow) return;
        await exportSingleDispatchRowExcel(targetRow);
      });
    });
  }

  async function buildFilledFormBufferForDispatchRow(dispatchRow) {
    const matchInfo = computeDispatchRowMatchInfo(dispatchRow);
    const { onlyProductMode } = getProcessingOptionsFromUI();

    // Nếu chuyến xe có dữ liệu khớp dg1 & sản phẩm thì điền dữ liệu đó;
    // nếu người dùng vẫn muốn xuất file <STT>.xlsx khi chưa có sản phẩm thì tạo 1 vùng trống với tiêu đề <STT>
    const processedData = (matchInfo.processedData && matchInfo.processedData.zones && matchInfo.processedData.zones.length > 0)
      ? matchInfo.processedData
      : {
          onlyProductMode,
          zones: [{ zoneIndex: 1, rowCount: 0, customers: [] }],
          warnings: []
        };

    const writeJ4 = document.getElementById('chkDispatchWriteJ4').checked;
    const enablePrint = document.getElementById('chkDispatchPrintSetup').checked;
    const tplBuf = state.customTemplateBuffer || state.systemTemplateBuffer;
    const sttFileName = String(dispatchRow.stt).trim();

    const filled = await global.TemplateFillModule.fillReportTemplate({
      templateBuffer: tplBuf,
      processedData,
      reportFileName: sttFileName,
      writeFileNameToJ4: writeJ4,
      enablePrintSetup: enablePrint
    });

    return {
      fileName: `${sttFileName}.xlsx`,
      buffer: filled.buffer,
      zoneCount: filled.zoneCount,
      productCount: matchInfo.productCount
    };
  }

  async function exportSingleDispatchRowExcel(dispatchRow) {
    try {
      showToast(`Đang tạo file ${dispatchRow.stt}.xlsx (Ship: ${dispatchRow.shipmentRaw})...`);
      const built = await buildFilledFormBufferForDispatchRow(dispatchRow);
      triggerDownloadBuffer(built.buffer, built.fileName);
      showToast(`Đã tải ${built.fileName} (${built.productCount} sản phẩm, ${built.zoneCount} vùng)!`);
    } catch (err) {
      showToast(`Lỗi tạo file ${dispatchRow.stt}.xlsx: ${err.message}`, 3500);
    }
  }

  async function exportDispatchBatchZip() {
    const records = getActiveDispatchSheetRecords();
    const selectedRows = records.filter(r => state.selectedDispatchStts.has(String(r.stt)));
    if (selectedRows.length === 0) {
      showToast('Vui lòng chọn ít nhất 1 chuyến xe để xuất gói .zip!');
      return;
    }

    const btnZip = document.getElementById('btnExportDispatchZip');
    const progEl = document.getElementById('dispatchProgressText');
    btnZip.disabled = true;

    try {
      const zip = new JSZip();
      let generatedCount = 0;

      for (let i = 0; i < selectedRows.length; i++) {
        const row = selectedRows[i];
        progEl.textContent = `Đang tạo (${i + 1}/${selectedRows.length}): ${row.stt}.xlsx (Ship ${row.shipmentRaw})...`;
        // Nhường luồng 1 nhịp để trình duyệt cập nhật tiến độ
        await new Promise(r => setTimeout(r, 5));
        const built = await buildFilledFormBufferForDispatchRow(row);
        zip.file(built.fileName, built.buffer);
        generatedCount++;
      }

      progEl.textContent = `Đang nén ${generatedCount} file thành gói .zip...`;
      const zipBuf = await zip.generateAsync({ type: 'arraybuffer' });
      let zipName = String(document.getElementById('dispatchZipFileNameInput').value || 'Bao_cao_QC_Dieu_xe.zip').trim();
      if (!/\.zip$/i.test(zipName)) zipName += '.zip';

      triggerDownloadBuffer(zipBuf, zipName, 'application/zip');
      progEl.textContent = `✓ Đã xuất ${generatedCount} file trong ${zipName}`;
      showToast(`Đã đóng gói thành công ${generatedCount} file báo cáo QC vào ${zipName}!`, 3500);
    } catch (err) {
      progEl.textContent = 'Lỗi: ' + err.message;
      showToast('Lỗi xuất gói ZIP: ' + err.message, 4000);
    } finally {
      btnZip.disabled = state.selectedDispatchStts.size === 0;
    }
  }

  // ==================== CHUYỂN FILE GỐC ====================
  async function handleConverterFiles(fileList) {
    if (!fileList || fileList.length === 0) return;
    state.convertedFiles = [];
    const tbody = document.getElementById('converterResultTbody');
    tbody.innerHTML = '';

    for (const file of Array.from(fileList)) {
      try {
        const parsed = await global.ConverterModule.parseUploadedFile(file);
        const outBuf = await global.ConverterModule.buildConvertedWorkbookBuffer(parsed);
        const baseName = file.name.replace(/\.(csv|xlsx)$/i, '');
        const suffix = parsed.type === 'dg1' ? '_dg1_chuan.xlsx' : '_sp_chuan.xlsx';
        const outName = baseName + suffix;

        state.convertedFiles.push({ outName, buffer: outBuf, parsed });

        const tr = document.createElement('tr');
        const typeDesc = parsed.type === 'dg1'
          ? 'Giao hàng dg1 (Code | Tên khách hàng | Điểm giao hàng | Ship)'
          : `Sản phẩm (${parsed.format === 'headerless_45col' ? 'Tự dò 45 cột không tiêu đề' : 'Code | Sản phẩm'})`;
        tr.innerHTML = `
          <td>${escapeHtml(file.name)}</td>
          <td>${escapeHtml(typeDesc)}</td>
          <td>${parsed.records.length} dòng</td>
          <td><button type="button" class="btn btn-sm btn-primary" data-dl-idx="${state.convertedFiles.length - 1}">Tải .xlsx</button></td>
        `;
        tbody.appendChild(tr);
      } catch (err) {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${escapeHtml(file.name)}</td>
          <td colspan="3" style="color:#dc2626;">Lỗi: ${escapeHtml(err.message)}</td>
        `;
        tbody.appendChild(tr);
      }
    }

    tbody.querySelectorAll('[data-dl-idx]').forEach(btn => {
      btn.addEventListener('click', () => {
        const item = state.convertedFiles[parseInt(btn.getAttribute('data-dl-idx'), 10)];
        if (item) triggerDownloadBuffer(item.buffer, item.outName);
      });
    });

    document.getElementById('btnDownloadAllConvertedZip').disabled = state.convertedFiles.length === 0;
  }

  async function downloadAllConvertedAsZip() {
    if (state.convertedFiles.length === 0) return;
    const zip = new JSZip();
    state.convertedFiles.forEach(item => {
      zip.file(item.outName, item.buffer);
    });
    const zipBuf = await zip.generateAsync({ type: 'arraybuffer' });
    triggerDownloadBuffer(zipBuf, 'File_goc_da_chuyen_doi.zip', 'application/zip');
  }

  // ==================== ADMIN: BẢNG QUẢN LÝ QUY ƯỚC ====================
  function getFilteredAdminRules() {
    const query = String(document.getElementById('adminSearchInput').value || '').trim().toLowerCase();
    const onlyUnshortened = document.getElementById('chkAdminFilterUnshortened').checked;
    const sourceList = state.adminActiveTab === 'customer' ? state.customerRules : state.productRules;

    return sourceList.map((item, realIdx) => ({ item, realIdx })).filter(({ item }) => {
      if (state.adminActiveTab === 'customer') {
        const isShortEmpty = !String(item.shortName || '').trim();
        if (onlyUnshortened && !isShortEmpty) return false;
        if (query) {
          const hay = `${item.customer} ${item.address} ${item.shortName}`.toLowerCase();
          if (!hay.includes(query)) return false;
        }
        return true;
      } else {
        const isShortEmpty = !String(item.shortProduct || '').trim();
        if (onlyUnshortened && !isShortEmpty) return false;
        if (query) {
          const hay = `${item.product} ${item.shortProduct}`.toLowerCase();
          if (!hay.includes(query)) return false;
        }
        return true;
      }
    });
  }

  function renderAdminRulesTable() {
    const thead = document.getElementById('adminRulesThead');
    const tbody = document.getElementById('adminRulesTbody');
    const tabCustBtn = document.getElementById('btnAdminTabCust');
    const tabProdBtn = document.getElementById('btnAdminTabProd');

    tabCustBtn.classList.toggle('btn-primary', state.adminActiveTab === 'customer');
    tabProdBtn.classList.toggle('btn-primary', state.adminActiveTab === 'product');

    if (state.adminActiveTab === 'customer') {
      thead.innerHTML = '<tr><th>Tên khách hàng (gốc)</th><th>Điểm giao (gốc)</th><th>Khách hàng (rút gọn)</th><th style="width:70px;">Xóa</th></tr>';
    } else {
      thead.innerHTML = '<tr><th>Sản phẩm (gốc)</th><th>Sản phẩm rút gọn</th><th style="width:70px;">Xóa</th></tr>';
    }

    const filtered = getFilteredAdminRules();
    const totalPages = Math.max(1, Math.ceil(filtered.length / state.adminPageSize));
    if (state.adminPage > totalPages) state.adminPage = totalPages;

    const startIdx = (state.adminPage - 1) * state.adminPageSize;
    const pageItems = filtered.slice(startIdx, startIdx + state.adminPageSize);

    document.getElementById('adminPaginationInfo').textContent =
      `Trang ${state.adminPage} / ${totalPages} (Tổng ${filtered.length} dòng)`;

    tbody.innerHTML = '';
    if (pageItems.length === 0) {
      tbody.innerHTML = `<tr><td colspan="${state.adminActiveTab === 'customer' ? 4 : 3}" style="text-align:center; color:#64748b;">Không có dòng nào.</td></tr>`;
      return;
    }

    pageItems.forEach(({ item, realIdx }) => {
      const tr = document.createElement('tr');
      if (state.adminActiveTab === 'customer') {
        const emptyShort = !String(item.shortName || '').trim();
        tr.innerHTML = `
          <td><input type="text" class="input" data-field="customer" data-idx="${realIdx}" value="${escapeAttr(item.customer)}"></td>
          <td><input type="text" class="input" data-field="address" data-idx="${realIdx}" value="${escapeAttr(item.address)}"></td>
          <td class="${emptyShort ? 'unshortened-cell' : ''}">
            <input type="text" class="input" data-field="shortName" data-idx="${realIdx}" value="${escapeAttr(item.shortName)}" placeholder="Chưa rút gọn...">
          </td>
          <td><button type="button" class="btn btn-sm" data-del-idx="${realIdx}">Xóa</button></td>
        `;
      } else {
        const emptyShort = !String(item.shortProduct || '').trim();
        tr.innerHTML = `
          <td><input type="text" class="input" data-field="product" data-idx="${realIdx}" value="${escapeAttr(item.product)}"></td>
          <td class="${emptyShort ? 'unshortened-cell' : ''}">
            <input type="text" class="input" data-field="shortProduct" data-idx="${realIdx}" value="${escapeAttr(item.shortProduct)}" placeholder="Chưa rút gọn...">
          </td>
          <td><button type="button" class="btn btn-sm" data-del-idx="${realIdx}">Xóa</button></td>
        `;
      }
      tbody.appendChild(tr);
    });

    tbody.querySelectorAll('input[data-field]').forEach(inp => {
      inp.addEventListener('input', () => {
        const idx = parseInt(inp.getAttribute('data-idx'), 10);
        const field = inp.getAttribute('data-field');
        const list = state.adminActiveTab === 'customer' ? state.customerRules : state.productRules;
        if (list[idx]) {
          list[idx][field] = inp.value;
          if (field === 'shortName' || field === 'shortProduct') {
            inp.parentElement.classList.toggle('unshortened-cell', !inp.value.trim());
          }
        }
      });
    });

    tbody.querySelectorAll('button[data-del-idx]').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.getAttribute('data-del-idx'), 10);
        if (state.adminActiveTab === 'customer') {
          state.customerRules.splice(idx, 1);
        } else {
          state.productRules.splice(idx, 1);
        }
        updateRuleCountersUI();
        renderAdminRulesTable();
      });
    });
  }

  // ==================== DỮ LIỆU KIỂM THỬ TỰ ĐỘNG (BAO GỒM CẢ ĐIỀU XE STT 41 -> 340, STT 48 -> 351+353) ====================
  function loadSampleDataForTesting() {
    const sampleMojibake = global.ConverterModule.decodeMojibakeString('C¤NG TY TNHH B¸CH HãA XANH');

    state.dg1Records = [
      { code: '2650D196696', customer: sampleMojibake, address: 'KHO TỔNG BÌNH DƯƠNG - KCN SÓNG THẦN', ship: '2650S00340' },
      { code: '2650D196697', customer: 'SIÊU THỊ MEGA MARKET THANH XUÂN', address: 'ĐƯỜNG PHẠM VĂN ĐỒNG - HÀ NỘI', ship: '340' },
      { code: '2650D196698', customer: 'CÔNG TY TNHH AEON VIỆT NAM', address: 'AEON LINH ĐÀM - HOÀNG MAI - HÀ NỘI', ship: '340' },
      { code: '2650D196699', customer: 'CÔNG TY TNHH KAMEREO VIỆT NAM', address: 'KHO ĐÔNG ANH - HÀ NỘI', ship: '2650S00351' },
      { code: '2650D196700', customer: 'SIÊU THỊ BRG ĐÔNG ANH', address: 'THỊ TRẤN ĐÔNG ANH - HÀ NỘI', ship: '353' },
      { code: '2650D196701', customer: 'NHÀ HÀNG BBQ BÀ TRIỆU', address: 'BÀ TRIỆU - HAI BÀ TRƯNG - HÀ NỘI', ship: '245' },
      { code: '2650D196702', customer: 'SIÊU THỊ CO.OPMART QUẬN 7', address: 'NGUYỄN VĂN LINH - QUẬN 7 - TP.HCM', ship: '' }
    ];

    state.productRecords = [
      { code: '2650D196696', product: 'Cải bẹ xanh VietGap 500g ()' },
      { code: '2650D196696', product: 'Cà chua beef Viet Gap 500g Khay' },
      { code: '2650D196697', product: 'Bắp cải trắng 5KG (VietGap)' },
      { code: '2650D196697', product: 'Khoai tây Đà Lạt 2 KG' },
      { code: '2650D196697', product: 'Xà lách lô lô xanh TF 300g' },
      { code: '2650D196698', product: 'Hành tây Đà Lạt 1kg' },
      { code: '2650D196698', product: 'Nấm kim châm Khay 150g' },
      { code: '2650D196699', product: 'Ớt chuông đỏ 500 gr' },
      { code: '2650D196699', product: 'Bí đỏ hồ lô 2kg VietGap' },
      { code: '2650D196700', product: 'Dưa leo baby Khay 500g' },
      { code: '2650D196700', product: 'Củ cải trắng 5kg' },
      { code: '2650D196701', product: 'Đùi tỏi gà CP 2kg' },
      { code: '2650D196702', product: 'Rau muống nước 500g VietGap' }
    ];

    state.selectedShips = new Set(['2650S00340', '340', '2650S00351', '353']);
    state.selectedCodes = new Set(['2650D196702']);

    // Nạp mẫu bảng Điều xe đúng như ví dụ của người dùng (STT 1 -> Ship 245, STT 41 -> Ship 340, STT 48 -> Ship 351+353)
    state.dispatchFileName = 'ĐIỀU XE GÀ T10.2026 (Mẫu kiểm thử).xlsx';
    state.dispatchSheets = [
      {
        sheetName: '23',
        dateLabel: 'Ngày: 23/10/2026',
        records: [
          {
            rowNumber: 9,
            stt: '1',
            shipmentRaw: '245',
            ships: ['245'],
            truck: '29H-43417',
            route: 'HN',
            dispatchCustomersSummary: 'BBQ BT; KFC LTT; KFC BẮC LĐ; LOTTE TĐN'
          },
          {
            rowNumber: 49,
            stt: '41',
            shipmentRaw: '340',
            ships: ['340'],
            truck: '29H-89301',
            route: 'HN',
            dispatchCustomersSummary: 'TRƯỜNG HÀ GIẢI PHÓNG; MEGA TX; MARFOR K.LŨ; AEON L.ĐÀM'
          },
          {
            rowNumber: 56,
            stt: '48',
            shipmentRaw: '351+353',
            ships: ['351', '353'],
            truck: '29E-64951',
            route: 'HN2',
            dispatchCustomersSummary: 'KAMEREO G+H; BFOOD; BBQ KHO (SHIP 351); BRG ĐA'
          }
        ]
      }
    ];
    state.activeDispatchSheetIdx = 0;

    document.getElementById('dg1StatusText').textContent = `Đã nạp ${state.dg1Records.length} dòng dg1 kiểm thử (Có Ship 245, 340, 351, 353).`;
    document.getElementById('prodStatusText').textContent = `Đã nạp ${state.productRecords.length} dòng sản phẩm kiểm thử (File 50 & File 54).`;
    document.getElementById('dispatchStep1StatusText').textContent = `Đã nạp mẫu Điều xe: STT 1 (Ship 245), STT 41 (Ship 340 → 41.xlsx), STT 48 (Ship 351+353 → 48.xlsx).`;

    populateDispatchSheetSelect();
    autoSelectDispatchRowsForCurrentSheet();
    renderDispatchTable();
    rebuildShipAndCodeLists();
    updateStepTicks();
    showToast('Đã nạp dữ liệu kiểm thử (kể cả mẫu Điều xe STT 41 & 48)!');
  }

  function escapeHtml(s) {
    return String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function escapeAttr(s) {
    return escapeHtml(s);
  }

  // ==================== KHỞI TẠO SỰ KIỆN ====================
  async function initApp() {
    const searchParams = new URLSearchParams(window.location.search);
    if (searchParams.get('page') === 'admin' || window.location.pathname.endsWith('/admin') || window.__IS_ADMIN_ROUTE__) {
      state.isAdminRoute = true;
    }
    if (typeof google !== 'undefined' && google.script && google.script.url) {
      google.script.url.getLocation(function(loc) {
        if (loc && loc.parameter && loc.parameter.page === 'admin') {
          state.isAdminRoute = true;
          applyAdminRouteUI();
          switchView('view-admin-rules');
        }
      });
    }

    document.getElementById('btnToggleDrawer').addEventListener('click', () => toggleDrawer());
    document.getElementById('btnCloseDrawer').addEventListener('click', () => toggleDrawer(false));
    document.getElementById('drawerBackdrop').addEventListener('click', () => toggleDrawer(false));

    document.querySelectorAll('.nav-item[data-view]').forEach(btn => {
      btn.addEventListener('click', () => switchView(btn.getAttribute('data-view')));
    });

    document.getElementById('btnSwitchAdminMode').addEventListener('click', () => {
      state.isAdminRoute = !state.isAdminRoute;
      applyAdminRouteUI();
      if (state.isAdminRoute) {
        switchView('view-admin-rules');
        showToast('Đã bật giao diện trang Admin (?page=admin)');
      } else {
        switchView('view-upload');
        showToast('Đã chuyển về trang Người dùng thường');
      }
    });

    // Bước 1: Sự kiện tải file
    document.getElementById('inputDg1File').addEventListener('change', e => handleDg1Upload(e.target.files));
    document.getElementById('inputProductFiles').addEventListener('change', e => handleProductFilesUpload(e.target.files));
    document.getElementById('inputDispatchFileStep1').addEventListener('change', e => handleDispatchFileUpload(e.target.files, true));
    document.getElementById('btnGoToDispatchFromStep1').addEventListener('click', () => switchView('view-dispatch'));

    document.getElementById('btnLoadQuickTest').addEventListener('click', loadSampleDataForTesting);
    document.getElementById('btnClearLoadedData').addEventListener('click', () => {
      state.dg1Records = [];
      state.productRecords = [];
      state.selectedShips.clear();
      state.selectedCodes.clear();
      state.processedResult = null;
      state.dispatchFileName = '';
      state.dispatchSheets = [];
      state.selectedDispatchStts.clear();
      document.getElementById('dg1StatusText').textContent = 'Chưa tải file dg1.';
      document.getElementById('prodStatusText').textContent = 'Chưa tải file sản phẩm.';
      document.getElementById('dispatchStep1StatusText').textContent = 'Chưa tải file Điều xe.';
      populateDispatchSheetSelect();
      renderDispatchTable();
      rebuildShipAndCodeLists();
      updateStepTicks();
      showToast('Đã xóa dữ liệu tạm');
    });
    document.getElementById('btnGoToSelect').addEventListener('click', () => switchView('view-select'));

    // Bước 2: Tìm kiếm và chọn Ship / Code
    document.getElementById('searchShipInput').addEventListener('input', renderShipList);
    document.getElementById('searchCodeInput').addEventListener('input', renderCodeList);
    document.getElementById('pasteCodesInput').addEventListener('input', updateStepTicks);

    document.getElementById('btnSelectAllShips').addEventListener('click', () => {
      getUniqueShips().forEach((_, ship) => state.selectedShips.add(ship));
      renderShipList();
      updateSelectionCounters();
      updateStepTicks();
    });
    document.getElementById('btnUnselectAllShips').addEventListener('click', () => {
      state.selectedShips.clear();
      renderShipList();
      updateSelectionCounters();
      updateStepTicks();
    });

    document.getElementById('btnSelectAllCodes').addEventListener('click', () => {
      getUniqueCodesInfo().forEach(item => state.selectedCodes.add(item.code));
      renderCodeList();
      updateSelectionCounters();
      updateStepTicks();
    });
    document.getElementById('btnUnselectAllCodes').addEventListener('click', () => {
      state.selectedCodes.clear();
      renderCodeList();
      updateSelectionCounters();
      updateStepTicks();
    });
    document.getElementById('btnSelectNoShipCodes').addEventListener('click', () => {
      let count = 0;
      getUniqueCodesInfo().forEach(item => {
        if (!String(item.ship || '').trim()) {
          state.selectedCodes.add(item.code);
          count++;
        }
      });
      renderCodeList();
      updateSelectionCounters();
      updateStepTicks();
      showToast(`Đã chọn ${count} mã Code không có Ship`);
    });

    document.getElementById('btnRunProcess').addEventListener('click', runProcessAndGoToResult);

    // Bước 3: Xuất Excel kết quả chuẩn
    document.getElementById('btnExportStandardExcel').addEventListener('click', async () => {
      if (!state.processedResult || !state.processedResult.zones || state.processedResult.zones.length === 0) {
        showToast('Chưa có kết quả xử lý để xuất Excel!');
        return;
      }
      const fileName = document.getElementById('exportFileNameInput').value || 'Ket_qua_giao_hang';
      const enablePrint = document.getElementById('chkStandardPrintSetup').checked;
      const buf = await global.PrintSetupModule.buildStandardResultWorkbookBuffer(state.processedResult, {
        enablePrintSetup: enablePrint
      });
      triggerDownloadBuffer(buf, fileName);
      showToast('Đã xuất file Excel kết quả!');
    });

    document.getElementById('btnGoToFormFill').addEventListener('click', () => switchView('view-form'));

    // Bước 4: Điền Form Báo Cáo
    document.getElementById('inputCustomFormFile').addEventListener('change', async e => {
      const f = e.target.files && e.target.files[0];
      if (!f) return;
      const buf = await f.arrayBuffer();
      const check = await global.TemplateFillModule.validateFormTemplateBuffer(buf);
      if (!check.valid) {
        document.getElementById('customFormStatus').textContent = '⚠ Cảnh báo: Không dò được cột J (Khách hàng) & K (Sản phẩm) trong file này!';
        showToast('Cảnh báo: Form không hợp lệ!');
        return;
      }
      state.customTemplateBuffer = buf;
      document.getElementById('activeFormNameLabel').textContent = `${f.name} (${check.zoneCount} vùng mẫu)`;
      document.getElementById('customFormStatus').textContent = `Đã nạp form "${f.name}": Dò thấy ${check.zoneCount} vùng trên sheet "${check.sheetName}".`;
      showToast(`Đã nạp form báo cáo (${check.zoneCount} vùng)`);
    });

    document.getElementById('btnExportFilledForm').addEventListener('click', async () => {
      if (!state.processedResult || !state.processedResult.zones || state.processedResult.zones.length === 0) {
        if (state.dg1Records.length > 0 && state.productRecords.length > 0) {
          runProcessAndGoToResult();
        } else {
          showToast('Vui lòng nạp dữ liệu và bấm Xử lý trước khi điền form!');
          return;
        }
      }
      try {
        const outName = document.getElementById('formOutputFileNameInput').value || 'Bao_cao_QC_Giao_hang';
        const writeJ4 = document.getElementById('chkWriteFileNameJ4').checked;
        const enablePrint = document.getElementById('chkFormPrintSetup').checked;
        const tplBuf = state.customTemplateBuffer || state.systemTemplateBuffer;

        const filled = await global.TemplateFillModule.fillReportTemplate({
          templateBuffer: tplBuf,
          processedData: state.processedResult,
          reportFileName: outName,
          writeFileNameToJ4: writeJ4,
          enablePrintSetup: enablePrint
        });

        document.getElementById('tickStep4').classList.add('done');
        triggerDownloadBuffer(filled.buffer, outName);
        showToast(`Đã điền xong ${filled.zoneCount} vùng báo cáo (kèm chữ ký QC/Kho)!`);
      } catch (err) {
        showToast('Lỗi điền form: ' + err.message);
      }
    });

    // Bước 5: Xuất hàng loạt theo file Điều xe (.zip)
    document.getElementById('inputDispatchFile').addEventListener('change', e => handleDispatchFileUpload(e.target.files, false));
    document.getElementById('selectDispatchSheet').addEventListener('change', e => {
      const idx = parseInt(e.target.value, 10);
      if (!isNaN(idx) && state.dispatchSheets[idx]) {
        state.activeDispatchSheetIdx = idx;
        const activeSh = state.dispatchSheets[idx];
        const cleanSheet = String(activeSh.sheetName || '').replace(/[^\w.-]+/g, '_');
        document.getElementById('dispatchZipFileNameInput').value = `Bao_cao_QC_Dieu_xe_Ngay_${cleanSheet}.zip`;
        autoSelectDispatchRowsForCurrentSheet();
        renderDispatchTable();
      }
    });

    document.getElementById('searchDispatchInput').addEventListener('input', renderDispatchTable);
    document.getElementById('chkDispatchOnlyMatched').addEventListener('change', renderDispatchTable);

    document.getElementById('btnSelectAllDispatch').addEventListener('click', () => {
      getActiveDispatchSheetRecords().forEach(r => state.selectedDispatchStts.add(String(r.stt)));
      renderDispatchTable();
    });
    document.getElementById('btnUnselectAllDispatch').addEventListener('click', () => {
      state.selectedDispatchStts.clear();
      renderDispatchTable();
    });
    document.getElementById('btnSelectMatchedDispatch').addEventListener('click', () => {
      state.selectedDispatchStts.clear();
      let count = 0;
      getActiveDispatchSheetRecords().forEach(r => {
        const info = computeDispatchRowMatchInfo(r);
        if (info.hasMatch && info.productCount > 0) {
          state.selectedDispatchStts.add(String(r.stt));
          count++;
        }
      });
      renderDispatchTable();
      showToast(`Đã chọn ${count} chuyến xe có dữ liệu khớp dg1`);
    });

    document.getElementById('btnExportDispatchZip').addEventListener('click', exportDispatchBatchZip);

    // Công cụ: Chuyển file gốc
    document.getElementById('inputConverterFiles').addEventListener('change', e => handleConverterFiles(e.target.files));
    document.getElementById('btnDownloadAllConvertedZip').addEventListener('click', downloadAllConvertedAsZip);

    // Công cụ: Quy ước rút gọn
    document.getElementById('btnDownloadCustSample').addEventListener('click', async () => {
      const buf = await global.ProductShortModule.exportShortRuleSampleExcel(
        'customer', state.dg1Records, state.productRecords, state.customerRules, state.productRules
      );
      triggerDownloadBuffer(buf, 'Mau_Quy_Uoc_Khach_Hang.xlsx');
    });

    document.getElementById('btnDownloadProdSample').addEventListener('click', async () => {
      const buf = await global.ProductShortModule.exportShortRuleSampleExcel(
        'product', state.dg1Records, state.productRecords, state.customerRules, state.productRules
      );
      triggerDownloadBuffer(buf, 'Mau_Quy_Uoc_San_Pham.xlsx');
    });

    document.getElementById('inputUploadCustRuleFile').addEventListener('change', async e => {
      const f = e.target.files && e.target.files[0];
      if (!f) return;
      const rules = await global.ProductShortModule.parseShortRuleExcel(f, 'customer');
      state.customerRules = rules;
      updateRuleCountersUI();
      renderAdminRulesTable();
      renderDispatchTable();
      showToast(`Đã nạp ${rules.length} quy ước Khách hàng`);
    });

    document.getElementById('inputUploadProdRuleFile').addEventListener('change', async e => {
      const f = e.target.files && e.target.files[0];
      if (!f) return;
      const rules = await global.ProductShortModule.parseShortRuleExcel(f, 'product');
      state.productRules = rules;
      updateRuleCountersUI();
      renderAdminRulesTable();
      renderDispatchTable();
      showToast(`Đã nạp ${rules.length} quy ước Sản phẩm`);
    });

    document.getElementById('btnAdminSyncFromLoaded').addEventListener('click', async () => {
      if (state.dg1Records.length === 0 && state.productRecords.length === 0) {
        showToast('Chưa có file dữ liệu nào đang tải để lấy mục mới!');
        return;
      }
      try {
        const newCust = state.dg1Records.map(r => ({ customer: r.customer, address: r.address }));
        const newProd = state.productRecords.map(r => r.product);
        const res = await global.CPGas('appendMissingRulesFromLoadedData', {
          newCustomerItems: newCust,
          newProductItems: newProd,
          expectedVersion: state.rulesVersion
        });
        state.rulesVersion = res.version;
        state.customerRules = res.customerRules;
        state.productRules = res.productRules;
        updateRuleCountersUI();
        renderAdminRulesTable();
        showToast(`Đã thêm ${res.addedCust} khách hàng và ${res.addedProd} sản phẩm mới vào cuối quy ước!`);
      } catch (err) {
        showToast(err.message);
      }
    });

    // Quản trị Admin
    document.getElementById('btnAdminLoginSubmit').addEventListener('click', async () => {
      const pwd = document.getElementById('adminPasswordInput').value;
      const errEl = document.getElementById('adminLoginError');
      errEl.textContent = '';
      try {
        await global.CPGas('adminLogin', pwd);
        state.isAdminAuthed = true;
        applyAdminRouteUI();
        renderAdminRulesTable();
        showToast('Đăng nhập Quản trị thành công!');
      } catch (err) {
        errEl.textContent = err.message;
      }
    });

    document.getElementById('btnAdminLogout').addEventListener('click', async () => {
      await global.CPGas('adminLogout');
      state.isAdminAuthed = false;
      applyAdminRouteUI();
      showToast('Đã đăng xuất phiên Quản trị');
    });

    document.getElementById('btnAdminTabCust').addEventListener('click', () => {
      state.adminActiveTab = 'customer';
      state.adminPage = 1;
      renderAdminRulesTable();
    });
    document.getElementById('btnAdminTabProd').addEventListener('click', () => {
      state.adminActiveTab = 'product';
      state.adminPage = 1;
      renderAdminRulesTable();
    });

    document.getElementById('adminSearchInput').addEventListener('input', () => {
      state.adminPage = 1;
      renderAdminRulesTable();
    });
    document.getElementById('chkAdminFilterUnshortened').addEventListener('change', () => {
      state.adminPage = 1;
      renderAdminRulesTable();
    });

    document.getElementById('btnAdminPrevPage').addEventListener('click', () => {
      if (state.adminPage > 1) {
        state.adminPage--;
        renderAdminRulesTable();
      }
    });
    document.getElementById('btnAdminNextPage').addEventListener('click', () => {
      const total = Math.max(1, Math.ceil(getFilteredAdminRules().length / state.adminPageSize));
      if (state.adminPage < total) {
        state.adminPage++;
        renderAdminRulesTable();
      }
    });

    document.getElementById('btnAdminAddRow').addEventListener('click', () => {
      if (state.adminActiveTab === 'customer') {
        state.customerRules.push({ customer: '', address: '', shortName: '' });
      } else {
        state.productRules.push({ product: '', shortProduct: '' });
      }
      const total = Math.max(1, Math.ceil(getFilteredAdminRules().length / state.adminPageSize));
      state.adminPage = total;
      updateRuleCountersUI();
      renderAdminRulesTable();
    });

    document.getElementById('btnAdminSaveServer').addEventListener('click', async () => {
      try {
        const res = await global.CPGas('saveShortRules', {
          expectedVersion: state.rulesVersion,
          customerRules: state.customerRules,
          productRules: state.productRules
        });
        state.rulesVersion = res.version;
        updateRuleCountersUI();
        renderDispatchTable();
        showToast(`Đã lưu quy ước lên hệ thống (Phiên bản v${res.version})`);
      } catch (err) {
        showToast(err.message, 4000);
      }
    });

    document.getElementById('inputAdminImportExcel').addEventListener('change', async e => {
      const f = e.target.files && e.target.files[0];
      if (!f) return;
      const rules = await global.ProductShortModule.parseShortRuleExcel(f, state.adminActiveTab);
      if (state.adminActiveTab === 'customer') state.customerRules = rules;
      else state.productRules = rules;
      updateRuleCountersUI();
      renderAdminRulesTable();
      renderDispatchTable();
      showToast(`Đã nhập ${rules.length} dòng từ Excel`);
    });

    document.getElementById('btnAdminExportExcel').addEventListener('click', async () => {
      const buf = await global.ProductShortModule.exportShortRuleSampleExcel(
        state.adminActiveTab, [], [], state.customerRules, state.productRules
      );
      const name = state.adminActiveTab === 'customer' ? 'Quy_uoc_Khach_hang.xlsx' : 'Quy_uoc_San_pham.xlsx';
      triggerDownloadBuffer(buf, name);
    });

    document.getElementById('btnAdminDownloadCurrentForm').addEventListener('click', async () => {
      const buf = state.systemTemplateBuffer || (await global.TemplateFillModule.createBuiltInDefaultFormBuffer());
      triggerDownloadBuffer(buf, state.templateMeta.fileName || 'Form_Bao_Cao_He_Thong.xlsx');
    });

    document.getElementById('inputAdminReplaceForm').addEventListener('change', async e => {
      const f = e.target.files && e.target.files[0];
      if (!f) return;
      const buf = await f.arrayBuffer();
      const check = await global.TemplateFillModule.validateFormTemplateBuffer(buf);
      const statusEl = document.getElementById('adminReplaceFormStatus');
      if (!check.valid) {
        statusEl.textContent = '⚠ Cảnh báo: Không dò được vùng "Khách hàng / Customer" (cột J) và "Sản phẩm / Product" (cột K) trong file này!';
        statusEl.style.color = '#dc2626';
        showToast('Cảnh báo: File form không hợp lệ!');
        return;
      }
      try {
        const b64 = global.AdminModule.arrayBufferToBase64(buf);
        const res = await global.CPGas('uploadSystemFormTemplate', {
          base64: b64,
          fileName: f.name,
          zoneCount: check.zoneCount
        });
        state.systemTemplateBuffer = buf;
        state.templateMeta = res.templateMeta;
        document.getElementById('adminCurrentFormName').textContent = res.templateMeta.fileName;
        document.getElementById('adminCurrentFormMeta').textContent = `Số vùng dò được: ${check.zoneCount} vùng — Cập nhật: ${res.templateMeta.updatedAt}`;
        document.getElementById('activeFormNameLabel').textContent = `${res.templateMeta.fileName} (${check.zoneCount} vùng)`;
        statusEl.textContent = `✓ Đã lưu form "${f.name}" (${check.zoneCount} vùng) lên hệ thống.`;
        statusEl.style.color = '#16a34a';
        showToast('Đã cập nhật Form Báo Cáo hệ thống!');
      } catch (err) {
        statusEl.textContent = 'Lỗi: ' + err.message;
        statusEl.style.color = '#dc2626';
      }
    });

    // Tự động nạp Quy ước và Form từ hệ thống ngay khi mở trang
    try {
      const boot = await global.CPGas('getBootstrapData');
      state.rulesVersion = boot.version || 1;
      state.customerRules = boot.customerRules || [];
      state.productRules = boot.productRules || [];
      state.isAdminAuthed = Boolean(boot.isAdminAuthenticated);
      if (boot.templateBase64) {
        state.systemTemplateBuffer = global.AdminModule.base64ToArrayBuffer(boot.templateBase64);
      } else {
        state.systemTemplateBuffer = await global.TemplateFillModule.createBuiltInDefaultFormBuffer();
      }
      if (boot.templateMeta) {
        state.templateMeta = boot.templateMeta;
        document.getElementById('adminCurrentFormName').textContent = boot.templateMeta.fileName;
        document.getElementById('activeFormNameLabel').textContent = `${boot.templateMeta.fileName} (${boot.templateMeta.zoneCount || 2} vùng)`;
      }
    } catch (e) {
      state.systemTemplateBuffer = await global.TemplateFillModule.createBuiltInDefaultFormBuffer();
    }

    updateRuleCountersUI();
    applyAdminRouteUI();
    renderAdminRulesTable();
    rebuildShipAndCodeLists();
    renderDispatchTable();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
  } else {
    initApp();
  }
})(window);
