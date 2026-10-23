// ==================== Google Apps Script: Code.gs ====================
function doGet(e) {
  var isAdminPage = (e && e.parameter && e.parameter.page === 'admin');
  var template = HtmlService.createTemplateFromFile('Index');
  template.isAdminRoute = isAdminPage;
  return template.evaluate()
    .setTitle('Công cụ Xử lý Giao hàng & Báo cáo Excel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getOrCreateStorageSheet_() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('RULE_SPREADSHEET_ID');
  var ss = null;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('DB_QuyUoc_GiaoHang');
    props.setProperty('RULE_SPREADSHEET_ID', ss.getId());
  }

  var shCust = ss.getSheetByName('KhachHang');
  if (!shCust) {
    shCust = ss.insertSheet('KhachHang');
    shCust.getRange(1, 1, 1, 3).setNumberFormat('@').setValues([['Tên khách hàng (gốc)', 'Điểm giao (gốc)', 'Khách hàng (rút gọn)']]);
  }
  var shProd = ss.getSheetByName('SanPham');
  if (!shProd) {
    shProd = ss.insertSheet('SanPham');
    shProd.getRange(1, 1, 1, 2).setNumberFormat('@').setValues([['Sản phẩm (gốc)', 'Sản phẩm rút gọn']]);
  }
  return { ss: ss, shCust: shCust, shProd: shProd };
}

function getBootstrapData(token) {
  var props = PropertiesService.getScriptProperties();
  var version = parseInt(props.getProperty('RULES_VERSION') || '1', 10);
  var db = getOrCreateStorageSheet_();

  var custValues = db.shCust.getDataRange().getDisplayValues();
  var customerRules = [];
  for (var i = 1; i < custValues.length; i++) {
    if (custValues[i][0] || custValues[i][1]) {
      customerRules.push({
        customer: custValues[i][0] || '',
        address: custValues[i][1] || '',
        shortName: custValues[i][2] || ''
      });
    }
  }

  var prodValues = db.shProd.getDataRange().getDisplayValues();
  var productRules = [];
  for (var j = 1; j < prodValues.length; j++) {
    if (prodValues[j][0]) {
      productRules.push({
        product: prodValues[j][0] || '',
        shortProduct: prodValues[j][1] || ''
      });
    }
  }

  var formFileId = props.getProperty('FORM_TEMPLATE_DRIVE_ID');
  var templateBase64 = null;
  var templateMeta = null;
  if (formFileId) {
    try {
      var file = DriveApp.getFileById(formFileId);
      templateBase64 = Utilities.base64Encode(file.getBlob().getBytes());
      templateMeta = {
        fileName: file.getName(),
        updatedAt: Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
        zoneCount: parseInt(props.getProperty('FORM_TEMPLATE_ZONES') || '2', 10)
      };
    } catch (e) {}
  }

  return {
    version: version,
    customerRules: customerRules,
    productRules: productRules,
    templateBase64: templateBase64,
    templateMeta: templateMeta,
    isAdminAuthenticated: verifyAdminSession_(token)
  };
}
