// ==================== Google Apps Script: Admin.gs ====================
function verifyAdminSession_(clientToken) {
  var props = PropertiesService.getScriptProperties();
  var userKey = Session.getTemporaryActiveUserKey() || 'anon';
  var savedToken = props.getProperty('SESSION_' + userKey);
  if (savedToken && clientToken && savedToken === clientToken) {
    return true;
  }
  return false;
}

function requireAdmin_(clientToken) {
  if (!verifyAdminSession_(clientToken)) {
    throw new Error('Chưa đăng nhập quyền Quản trị hoặc phiên đã hết hạn.');
  }
}

function adminLogin(password, _clientToken) {
  var props = PropertiesService.getScriptProperties();
  var now = Date.now();
  var lockUntil = parseInt(props.getProperty('ADMIN_LOCK_UNTIL') || '0', 10);

  if (lockUntil > now) {
    var mins = Math.ceil((lockUntil - now) / 60000);
    throw new Error('Đã khóa chống dò mật khẩu. Thử lại sau ' + mins + ' phút.');
  }

  var adminPwd = props.getProperty('ADMIN_PASSWORD') || 'admin123';
  if (String(password) !== String(adminPwd)) {
    var fails = parseInt(props.getProperty('ADMIN_FAIL_COUNT') || '0', 10) + 1;
    if (fails >= 8) {
      props.setProperty('ADMIN_LOCK_UNTIL', String(now + 10 * 60 * 1000));
      props.setProperty('ADMIN_FAIL_COUNT', '0');
      throw new Error('Sai mật khẩu 8 lần. Khóa đăng nhập 10 phút.');
    }
    props.setProperty('ADMIN_FAIL_COUNT', String(fails));
    throw new Error('Mật khẩu không đúng (' + fails + '/8).');
  }

  props.setProperty('ADMIN_FAIL_COUNT', '0');
  props.setProperty('ADMIN_LOCK_UNTIL', '0');

  var userKey = Session.getTemporaryActiveUserKey() || 'anon';
  var newToken = Utilities.getUuid();
  props.setProperty('SESSION_' + userKey, newToken);
  return { ok: true, token: newToken };
}

function adminLogout(_clientToken) {
  var props = PropertiesService.getScriptProperties();
  var userKey = Session.getTemporaryActiveUserKey() || 'anon';
  props.deleteProperty('SESSION_' + userKey);
  return { ok: true };
}

function saveShortRules(payload, clientToken) {
  requireAdmin_(clientToken);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var props = PropertiesService.getScriptProperties();
    var currentVer = parseInt(props.getProperty('RULES_VERSION') || '1', 10);
    if (payload.expectedVersion && payload.expectedVersion !== currentVer) {
      throw new Error('Phiên bản dữ liệu đã thay đổi (v' + currentVer + '). Hãy tải lại trước khi lưu.');
    }

    var db = getOrCreateStorageSheet_();
    // Ghi tab KhachHang định dạng văn bản '@' để "1:" không bị đổi thành giờ
    db.shCust.clearContents();
    var custRows = [['Tên khách hàng (gốc)', 'Điểm giao (gốc)', 'Khách hàng (rút gọn)']];
    (payload.customerRules || []).forEach(function(r) {
      custRows.push([String(r.customer || ''), String(r.address || ''), String(r.shortName || '')]);
    });
    var custRange = db.shCust.getRange(1, 1, custRows.length, 3);
    custRange.setNumberFormat('@');
    custRange.setValues(custRows);

    // Ghi tab SanPham định dạng văn bản '@'
    db.shProd.clearContents();
    var prodRows = [['Sản phẩm (gốc)', 'Sản phẩm rút gọn']];
    (payload.productRules || []).forEach(function(r) {
      prodRows.push([String(r.product || ''), String(r.shortProduct || '')]);
    });
    var prodRange = db.shProd.getRange(1, 1, prodRows.length, 2);
    prodRange.setNumberFormat('@');
    prodRange.setValues(prodRows);

    var nextVer = currentVer + 1;
    props.setProperty('RULES_VERSION', String(nextVer));
    return {
      ok: true,
      version: nextVer,
      customerRules: payload.customerRules || [],
      productRules: payload.productRules || []
    };
  } finally {
    lock.releaseLock();
  }
}

function appendMissingRulesFromLoadedData(payload, clientToken) {
  requireAdmin_(clientToken);
  var current = getBootstrapData(clientToken);
  if (payload.expectedVersion && payload.expectedVersion !== current.version) {
    throw new Error('Xung đột phiên bản v' + current.version + '. Hãy tải lại trang.');
  }

  var norm = function(s) { return String(s || '').replace(/\s+/g, ' ').trim().toLowerCase(); };
  var custMap = {};
  current.customerRules.forEach(function(r) {
    custMap[norm(r.customer) + '|||' + norm(r.address)] = true;
  });
  var prodMap = {};
  current.productRules.forEach(function(r) {
    prodMap[norm(r.product)] = true;
  });

  var addedCust = 0, addedProd = 0;
  (payload.newCustomerItems || []).forEach(function(item) {
    var k = norm(item.customer) + '|||' + norm(item.address);
    if ((item.customer || item.address) && !custMap[k]) {
      custMap[k] = true;
      current.customerRules.push({ customer: item.customer, address: item.address, shortName: '' });
      addedCust++;
    }
  });

  (payload.newProductItems || []).forEach(function(p) {
    var k = norm(p);
    if (p && !prodMap[k]) {
      prodMap[k] = true;
      current.productRules.push({ product: p, shortProduct: '' });
      addedProd++;
    }
  });

  var saved = saveShortRules({
    expectedVersion: current.version,
    customerRules: current.customerRules,
    productRules: current.productRules
  }, clientToken);

  return {
    ok: true,
    addedCust: addedCust,
    addedProd: addedProd,
    version: saved.version,
    customerRules: saved.customerRules,
    productRules: saved.productRules
  };
}

function uploadSystemFormTemplate(payload, clientToken) {
  requireAdmin_(clientToken);
  var props = PropertiesService.getScriptProperties();
  var bytes = Utilities.base64Decode(payload.base64);
  var blob = Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', payload.fileName || 'Form_Bao_Cao.xlsx');
  var file = DriveApp.createFile(blob);
  props.setProperty('FORM_TEMPLATE_DRIVE_ID', file.getId());
  props.setProperty('FORM_TEMPLATE_ZONES', String(payload.zoneCount || 2));
  return {
    ok: true,
    templateMeta: {
      fileName: file.getName(),
      updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      zoneCount: payload.zoneCount || 2
    }
  };
}
