/* global SpreadsheetApp, DriveApp, ScriptApp, Utilities, UrlFetchApp, Logger */

var TITLE_INVOICE = '開立發票';
var RE_TRACK_LIST = /trackinfo=(..),(\d{8}),(\d{8})/;
var CSV_HEADER = '"InvoiceNumber","InvoiceDate","InvoiceTime","BuyerIdentifier","BuyerName","BuyerAddress","BuyerTelephoneNumber","BuyerEmailAddress","SalesAmount","FreeTaxSalesAmount","ZeroTaxSalesAmount","TaxType","TaxRate","TaxAmount","TotalAmount","PrintMark","RandomNumber","MainRemark","CarrierType","CarrierId1","CarrierId2","NPOBAN","Description","Quantity","UnitPrice","Amount","Remark"';
var baseUrl, taxId, user, password, spreadsheet, sheet;

function onOpen() {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Load pinkoi.txt', 'loadPinkoi')
      .addItem(TITLE_INVOICE, 'createInvoice')
      .addToUi();
}

function getSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsArr = ss.getSheets();
  var sheets = {};
  sheetsArr.forEach(function(s) {
    sheets[s.getName()] = s;
  });
  return sheets;
}

function fillSheets(sheets, data) {
  var target;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(data).forEach(function(prop) {
    if (!sheets[prop]) {
      target = ss.insertSheet(prop);
      sheets[prop] = target;
    } else {
      sheets[prop].clear();
      target = sheets[prop];
    }
    target.appendRow(['訂單編號', '姓名', '地址', '電話', '發票抬頭', '統編', '產品', '數量', '金額', '小計金額', '總金額', '運費', '金流手續費', '折扣', '備註']);
    data[prop].forEach(function(order) {
      target.appendRow([
        order.oid, order.name, order.address, order.tel, order.taxtitle, order.taxid,
        order.title, order.quantity, order.price, order.subtotal, order.payment, order.handling,
        order.payment_fee, order.reward_deduct, order.message
      ]);
    });
  });
}

function getJSON(filename) {
  var file = DriveApp.getFilesByName(filename).next();
  var content = file.getAs('text/plain').getDataAsString('utf-8');
  var obj = JSON.parse(content);
  return obj;
}

function showResult(res, csv) {
  var ui = SpreadsheetApp.getUi();
  var content = '匯入完成！';
  ui.alert(content);
}

function loadPinkoi() {
  var content = getJSON('pinkoi.txt');
  var sheets = getSheets();
  fillSheets(sheets, content);
  showResult();
}



function onInstall(e) {
  onOpen(e);
}

function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}

///////////////////////////////////////

function getItems() {
  return sheet.getRange('2:100').getValues().filter(function(row) {
    return row[0];
  });
}

function paddy(n, p, c) {
  var padChar = typeof c !== 'undefined' ? c : '0';
  var pad = new Array(1 + p).join(padChar);
  return (pad + n).slice(-pad.length);
}

function getYearPeriod(date) {
  var year = date.getFullYear();
  var period = parseInt(date.getMonth() / 2);
  return [year, period];
}

function push(line, content) {
  line.push('"' + content + '"');
}

function getSalesAmount(row) {
  // 總金額減掉金流手續費
  var total = row[10] - row[12];
  return parseInt(row[5] ? total * 0.95 : total);
}

function getTaxAmount(row) {
  // 總金額減掉金流手續費
  var total = row[10] - row[12];
  return parseInt(row[5] ? total * 0.05 : 0);
}

function buildCSV(rows, invoices, date) {
  var csv = [];
  csv.push(CSV_HEADER);
  var invoice;
  rows.forEach(function(row, index) {
    var line = [];
    if (index > 0 && rows[index - 1][0] === row[0]) {
      push(line, invoice);
      line.push('"","","","","","","","","","","","","","","","","","","","",""');
    } else {
      invoice = invoices.shift();
      push(line, invoice);
      var invoiceDate = date.getFullYear() + paddy(date.getMonth() + 1, 2) + paddy(date.getDate(), 2);
      push(line, invoiceDate);
      push(line, paddy(date.getHours(), 2) + ':' + paddy(date.getMinutes(), 2) + ':00'); // InvoiceTime
      push(line, row[5]); //買方統編
      push(line, row[4] ? row[4] : row[1]); //買方名稱
      push(line, row[2].replace('\n', ' ')); //買方地址
      push(line, row[3]); //買方電話
      push(line, ''); //email
      push(line, getSalesAmount(row)); //應稅銷售額
      push(line, 0); //免稅銷售額
      push(line, 0); //零稅率銷售額
      push(line, 1); //課稅別
      push(line, 0.05); //稅率，預設是 0.05
      push(line, getTaxAmount(row)); //營業稅額
      push(line, row[10]); //總計
      push(line, 'N'); //列印註記
      push(line, ''); //隨機碼
      push(line, ''); //總備註
      push(line, ''); //載具類別
      push(line, ''); //載具明碼
      push(line, ''); //載具隱碼
      push(line, ''); //愛心碼
    }
    push(line, row[6]); //品名
    push(line, row[7]); //數量
    push(line, row[8]); //單價
    push(line, row[9]); //小計金額
    push(line, ''); //備註
    csv.push(line.join(','));
  });
  csv.push('"Finish"');
  return csv.join('\n');
}

function parseResponse(raw) {
  var res = {};
  raw.split('&').forEach(function(part) {
    var arr = part.split('=');
    if (arr[0] === 'rtmessage') {
      res[arr[0]] = Utilities.newBlob(Utilities.base64Decode(arr[1])).getDataAsString();
    } else {
      res[arr[0]] = arr[1];
    }
  });
  return res;
}

function uploadToCXN(csv) {
  var url = baseUrl + '/c0401.php';
  var payload = {
    csv: Utilities.base64Encode(csv, Utilities.Charset.UTF_8),
    id: taxId,
    user: user,
    passwd: password
  };
  var params = {
    method: 'post',
    payload: payload
  };
  var response = UrlFetchApp.fetch(url, params);
  var res = parseResponse(response.getContentText());
  Logger.log(res);
  return res;
}

function getOrderCount(rows) {
  var count = 0;
  rows.forEach(function(row, index) {
    if (index === 0) {
      count++;
    }
    else if (rows[index - 1][0] !== row[0]) {
      count++;
    }
  });
  return count;
}

function getInvoiceNumbers(len, year, period) {
  var url = baseUrl + '/get_track_list.php';
  var invoices = [];
  var payload = {
    year: year,
    period: period,
    size: len,
    id: taxId,
    user: user,
    passwd: password
  };
  var params = {
    method: 'post',
    payload: payload
  };
  var response = UrlFetchApp.fetch(url, params);
  Logger.log(parseResponse(response.getContentText()));
  var matched = response.getContentText().match(RE_TRACK_LIST);
  if (matched) {
    var start = parseInt(matched[2], 10);
    for (var i = 0; i < len; i++) {
      invoices.push(matched[1] + paddy(start + i, 8));
    }
  }
  Logger.log(invoices);
  return invoices;
}

function init() {
  spreadsheet = SpreadsheetApp.getActive();
  sheet = spreadsheet.getSheetByName('cxn');
  baseUrl = sheet.getRange(1, 1).getValue();
  taxId = sheet.getRange(1, 2).getValue();
  user = sheet.getRange(1, 3).getValue();
  password = sheet.getRange(1, 4).getValue();
}

function showInvoiceResult(res, csv) {
  var ui = SpreadsheetApp.getUi();
  var content = JSON.stringify(res, null, 2);
  if (res.rtcode !== '0000') {
    content += '\n' + csv;
  }
  ui.alert(content);
}

function createInvoice() {
  init();
  var dateField = sheet.getRange(1, 5).getValue();
  var date = dateField === '' ? new Date() : new Date(dateField);
  Logger.log(dateField);
  Logger.log(date);
  var rows = getItems();
  var year = date.getFullYear();
  var period = parseInt(date.getMonth() / 2);
  var invoices = getInvoiceNumbers(rows.length, year, period);
  var csv = buildCSV(rows, invoices, date);
  Logger.log(csv);
  var res = uploadToCXN(csv);
  showInvoiceResult(res, csv);
}

/*
 * test functions
 */

 function test4() {
   init();
   var rows = getItems();
   var date = new Date();
   var year = date.getFullYear();
   var period = parseInt(date.getMonth() / 2);
   var orderCount = getOrderCount(rows);
   var invoices = getInvoiceNumbers(orderCount, year, period);
   var csv = buildCSV(rows, invoices, date);
   Logger.log(csv);

   Logger.log(Utilities.base64Encode(csv, Utilities.Charset.UTF_8));
   //Logger.log(csv);
   //var res = uploadToCXN(csv);
 }

 function test3() {
   init();
   var items = getItems();
   items.forEach(function(item) {
     Logger.log(item[0]);
   });
 }

 function test2() {
   var yp = getYearPeriod(new Date());
   Logger.log(yp);
 }

 function test() {
   init();
   getInvoiceNumbers(1, 2015, 1);
 }
