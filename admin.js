// ==================== MODULE: admin.js & CPGas RPC Bridge ====================
(function(global) {
  'use strict';

  // Bộ nhớ đệm trong phiên trình duyệt (không phụ thuộc duy nhất vào localStorage vì iframe GAS có thể chặn)
  const memoryStore = {};

  function safeGetStorage(key) {
    if (memoryStore[key] !== undefined) return memoryStore[key];
    try {
      const val = window.localStorage.getItem(key);
      if (val !== null) memoryStore[key] = val;
      return val;
    } catch (e) {
      return memoryStore[key] || null;
    }
  }

  function safeSetStorage(key, value) {
    memoryStore[key] = String(value);
    try {
      window.localStorage.setItem(key, String(value));
    } catch (e) {}
  }

  function safeRemoveStorage(key) {
    delete memoryStore[key];
    try {
      window.localStorage.removeItem(key);
    } catch (e) {}
  }

  /**
   * Lớp nối gọi máy chủ chung CPGas(fnName, ...args) trả về Promise:
   * 1. Google Apps Script: gọi google.script.run
   * 2. Vercel: gọi fetch('/api/rpc')
   * 3. GitHub Pages / Trình duyệt tĩnh: giả lập đầy đủ ở client để mọi nút chức năng hoạt động trơn tru
   */
  function CPGas(fnName, ...args) {
    const adminToken = safeGetStorage('cp_admin_session_token') || '';

    // 1. Môi trường Google Apps Script
    if (typeof google !== 'undefined' && google.script && google.script.run) {
      return new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(resolve)
          .withFailureHandler(err => reject(new Error(err && err.message ? err.message : String(err))))
          [fnName](...args, adminToken);
      });
    }

    // 2. Môi trường Vercel (nếu có endpoint /api/rpc)
    if (window.__USE_VERCEL_RPC__) {
      return fetch('/api/rpc', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'include',
        body: JSON.stringify({ fn: fnName, args, token: adminToken })
      }).then(async res => {
        const data = await res.json();
        if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`);
        return data.result;
      });
    }

    // 3. Môi trường tĩnh (GitHub Pages / Canvas Preview): xử lý cục bộ với đầy đủ kiểm tra phiên bản, khóa chống dò mật khẩu 10 phút sau 8 lần sai
    return localMockRpc(fnName, args, adminToken);
  }

  // Dữ liệu mặc định ban đầu cho môi trường tĩnh
  const localServerState = {
    version: 1,
    customerRules: [
      { customer: 'CÔNG TY TNHH BÁCH HÓA XANH', address: 'KHO TỔNG BÌNH DƯƠNG - KCN SÓNG THẦN', shortName: 'BHX Bình Dương' },
      { customer: 'CÔNG TY CỔ PHẦN DỊCH VỤ THƯƠNG MẠI TỔNG HỢP WINCOMMERCE', address: 'KHO KHÔ MIỀN NAM - LONG AN', shortName: 'WinMart Long An' },
      { customer: 'CÔNG TY TNHH MM MEGA MARKET VIỆT NAM', address: 'TRẠM TRUNG CHUYỂN BÌNH PHƯỚC - HIỆP BÌNH PHƯỚC', shortName: 'MM Hiệp Bình Phước' }
    ],
    productRules: [
      { product: 'Cải bẹ xanh 500g', shortProduct: 'Cải xanh 500g' },
      { product: 'Cà chua beef 500g Khay', shortProduct: 'Cà beef khay 500g' },
      { product: 'Bắp cải trắng 5KG', shortProduct: 'Bắp cải 5kg' }
    ],
    templateBase64: null,
    templateMeta: { fileName: 'Form_Bao_Cao_QC_Mac_Dinh.xlsx', updatedAt: '2026-10-24 08:00', zoneCount: 2 },
    failedAttempts: 0,
    lockoutUntil: 0,
    validTokens: new Set()
  };

  async function localMockRpc(fnName, args, token) {
    const now = Date.now();

    if (fnName === 'getBootstrapData') {
      const isAuthed = Boolean(token && localServerState.validTokens.has(token));
      return {
        version: localServerState.version,
        customerRules: JSON.parse(JSON.stringify(localServerState.customerRules)),
        productRules: JSON.parse(JSON.stringify(localServerState.productRules)),
        templateBase64: localServerState.templateBase64,
        templateMeta: localServerState.templateMeta,
        isAdminAuthenticated: isAuthed
      };
    }

    if (fnName === 'adminLogin') {
      const password = String(args[0] || '');
      if (localServerState.lockoutUntil > now) {
        const remainMin = Math.ceil((localServerState.lockoutUntil - now) / 60000);
        throw new Error(`Tài khoản đang bị tạm khóa chống dò mật khẩu. Vui lòng thử lại sau ${remainMin} phút.`);
      }
      // Mật khẩu mặc định khi chạy bản tĩnh thử nghiệm là admin123 (trên GAS/Vercel đọc từ ADMIN_PASSWORD)
      const expectedPwd = window.__STATIC_ADMIN_PASS__ || 'admin123';
      if (password !== expectedPwd) {
        localServerState.failedAttempts += 1;
        if (localServerState.failedAttempts >= 8) {
          localServerState.lockoutUntil = now + 10 * 60 * 1000;
          localServerState.failedAttempts = 0;
          throw new Error('Nhập sai 8 lần liên tiếp. Đã khóa đăng nhập trong 10 phút.');
        }
        throw new Error(`Mật khẩu không đúng (Sai ${localServerState.failedAttempts}/8 lần). Gợi ý bản thử nghiệm: admin123`);
      }
      localServerState.failedAttempts = 0;
      localServerState.lockoutUntil = 0;
      const newToken = 'tok_' + Math.random().toString(36).slice(2) + '_' + now;
      localServerState.validTokens.add(newToken);
      safeSetStorage('cp_admin_session_token', newToken);
      return { ok: true, token: newToken };
    }

    if (fnName === 'adminLogout') {
      if (token) localServerState.validTokens.delete(token);
      safeRemoveStorage('cp_admin_session_token');
      return { ok: true };
    }

    // Các hàm bên dưới yêu cầu quyền Admin
    if (!token || !localServerState.validTokens.has(token)) {
      throw new Error('Phiên quản trị hết hạn hoặc chưa đăng nhập.');
    }

    if (fnName === 'saveShortRules') {
      const payload = args[0] || {};
      if (payload.expectedVersion !== undefined && payload.expectedVersion !== localServerState.version) {
        throw new Error(`Xung đột phiên bản! Dữ liệu trên hệ thống đã được người khác cập nhật lên bản v${localServerState.version} (bạn đang ở v${payload.expectedVersion}). Hãy tải lại quy ước trước khi lưu.`);
      }
      localServerState.customerRules = JSON.parse(JSON.stringify(payload.customerRules || []));
      localServerState.productRules = JSON.parse(JSON.stringify(payload.productRules || []));
      localServerState.version += 1;
      return {
        ok: true,
        version: localServerState.version,
        customerRules: localServerState.customerRules,
        productRules: localServerState.productRules
      };
    }

    if (fnName === 'appendMissingRulesFromLoadedData') {
      const { newCustomerItems = [], newProductItems = [], expectedVersion } = args[0] || {};
      if (expectedVersion !== undefined && expectedVersion !== localServerState.version) {
        throw new Error(`Xung đột phiên bản (hệ thống đang ở v${localServerState.version}). Vui lòng nạp lại quy ước.`);
      }

      const norm = s => String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
      const custSet = new Set(localServerState.customerRules.map(r => `${norm(r.customer)}|||${norm(r.address)}`));
      const prodSet = new Set(localServerState.productRules.map(r => norm(r.product)));

      let addedCust = 0;
      let addedProd = 0;

      newCustomerItems.forEach(item => {
        const k = `${norm(item.customer)}|||${norm(item.address)}`;
        if ((item.customer || item.address) && !custSet.has(k)) {
          custSet.add(k);
          localServerState.customerRules.push({
            customer: item.customer,
            address: item.address,
            shortName: ''
          });
          addedCust++;
        }
      });

      newProductItems.forEach(prodName => {
        const cleaned = global.ProductShortModule.cleanProductName(prodName);
        const k = norm(cleaned);
        if (cleaned && !prodSet.has(k)) {
          prodSet.add(k);
          localServerState.productRules.push({
            product: cleaned,
            shortProduct: ''
          });
          addedProd++;
        }
      });

      if (addedCust > 0 || addedProd > 0) {
        localServerState.version += 1;
      }

      return {
        ok: true,
        addedCust,
        addedProd,
        version: localServerState.version,
        customerRules: JSON.parse(JSON.stringify(localServerState.customerRules)),
        productRules: JSON.parse(JSON.stringify(localServerState.productRules))
      };
    }

    if (fnName === 'uploadSystemFormTemplate') {
      const { base64, fileName, zoneCount } = args[0] || {};
      localServerState.templateBase64 = base64;
      localServerState.templateMeta = {
        fileName: fileName || 'Form_Bao_Cao.xlsx',
        updatedAt: new Date().toLocaleString('vi-VN'),
        zoneCount: zoneCount || 2
      };
      return { ok: true, templateMeta: localServerState.templateMeta };
    }

    throw new Error(`Hàm RPC không tồn tại: ${fnName}`);
  }

  function arrayBufferToBase64(buffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
    }
    return btoa(binary);
  }

  function base64ToArrayBuffer(base64) {
    const binary = atob(base64);
    const len = binary.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes.buffer;
  }

  global.CPGas = CPGas;
  global.AdminModule = {
    CPGas,
    safeGetStorage,
    safeSetStorage,
    safeRemoveStorage,
    arrayBufferToBase64,
    base64ToArrayBuffer
  };
})(window);
