/**
 * PP AUTO ERP - Core System (Updated: 2026-03-16)
 * ระบบจัดการภาษีซื้อ-ขาย, Dashboard VAT และระบบนำทางอัจฉริยะแบบ Dynamic
 */

// --- 1. System Core (ส่วนแกนหลักระบบ) ---

function doGet(e) {
  var page = e.parameter.page || 'index';
  var template = HtmlService.createTemplateFromFile(page);
  
  // เตรียมข้อมูลเมนูส่งให้หน้าเพจ (เพื่อไฮไลต์ปุ่มปัจจุบัน)
  template.menuConfig = getMenuConfig(page); 
  
  return template.evaluate()
      .setTitle('PP AUTO ERP')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชัน include แบบโปร (แก้ไขปัญหาโค้ดหลุดและส่งตัวแปรได้)
function include(filename, menuData) {
  var template = HtmlService.createTemplateFromFile(filename);
  if (menuData) {
    template.menuConfig = menuData;
  }
  return template.evaluate().getContent();
}

function getMenuConfig(currentPage) {
  const menuItems = [
    { id: 'dashboard', label: '📊 Dashboard' },
    { id: 'index', label: '📝 ออกใบกำกับภาษี' },
    { id: 'billing', label: '📅 ทำใบวางบิล' },
    { id: 'purchase', label: '📥 ระบบภาษีซื้อ' },
    // { id: 'sales_tax', label: '📤 4. ระบบภาษีขาย' },         // ✨ เมนูใหม่
    { id: 'dashboard_purchase', label: '📈 รายงานภาษีซื้อ' },
    { id: 'dashboard_sales', label: '📈 รายงานภาษีขาย' },    // ✨ เมนูใหม่
    { id: 'dashboard_vat', label: '💡 สรุปภาษี ภ.พ.30' }
  ];
  const baseUrl = ScriptApp.getService().getUrl();
  return menuItems.map(item => ({
    ...item,
    url: `${baseUrl}?page=${item.id}`,
    isActive: item.id === currentPage
  }));
}

// --- 2. Sales Tax System (ระบบภาษีขาย) ---

// บันทึกภาษีขาย (คีย์เพิ่มเติมนอกระบบ)
function saveSalesTax(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sales_Tax") || ss.insertSheet("Sales_Tax");
  try {
    sheet.appendRow([
      new Date(), payload.date, payload.taxNo, payload.customer, payload.taxID || "N/A",
      payload.branch || "00000", payload.base, payload.vat, payload.total, payload.user, "ออกแล้ว"
    ]);
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; }
}

// ดึงข้อมูลรายงานภาษีขายรวม (รวมจากทั้ง Invoice และ Sales_Tax)
function getSalesTaxReportData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var salesTaxSheet = ss.getSheetByName("Sales_Tax");
  var invoiceSheet = ss.getSheetByName("Invoice");
  var history = [];
  var stats = { total: 0, vat: 0 };

  // 1. ดึงจาก Sales_Tax (ที่คีย์มือ)
  if (salesTaxSheet) {
    var data = salesTaxSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var total = parseFloat(data[i][8]) || 0;
      var vat = parseFloat(data[i][7]) || 0;
      stats.total += total; stats.vat += vat;
      history.push({
        date: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], "GMT+7", "dd/MM/yyyy") : data[i][1],
        taxNo: data[i][2], customer: data[i][3], total: total, vat: vat, source: "Manual"
      });
    }
  }

  // 2. ดึงจากหน้าออกใบกำกับภาษี (Autoดึงจาก Invoice Sheet)
  if (invoiceSheet) {
    var invData = invoiceSheet.getDataRange().getValues();
    for (var i = 1; i < invData.length; i++) {
      var total = parseFloat(invData[i][3]) || 0;
      var vat = total - (total * 100 / 107); // คำนวณ VAT 7%
      stats.total += total; stats.vat += vat;
      history.push({
        date: invData[i][0] instanceof Date ? Utilities.formatDate(invData[i][0], "GMT+7", "dd/MM/yyyy") : invData[i][0],
        taxNo: invData[i][1], customer: invData[i][2], total: total, vat: vat, source: "System"
      });
    }
  }
  return { stats: stats, history: history.reverse() };
}

// --- 3. Purchase Tax System (ระบบภาษีซื้อ) ---

function savePurchase(payload, base64Image) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Purchase_Tax") || ss.insertSheet("Purchase_Tax");
  var data = sheet.getDataRange().getValues();
  var targetRow = -1;

  if (payload.oldTaxNo) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] == payload.oldTaxNo) { targetRow = i + 1; break; }
    }
  }

  if (!payload.oldTaxNo || (payload.oldTaxNo && payload.oldTaxNo !== payload.taxNo)) {
    if (checkDuplicateTaxNo(payload.taxNo)) return { status: "duplicate" };
  }

  try {
    if (targetRow != -1) {
      sheet.getRange(targetRow, 2, 1, 8).setValues([[payload.date, payload.taxNo, payload.vendor, payload.taxID, payload.branch, payload.base, payload.vat, payload.total]]);
      sheet.getRange(targetRow, 12).setValue(payload.paymentStatus);
    } else {
      sheet.appendRow([new Date(), payload.date, payload.taxNo, payload.vendor, payload.taxID, payload.branch, payload.base, payload.vat, payload.total, "", payload.user, payload.paymentStatus]);
    }
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; }
}

// --- 4. VAT Matching & Reporting (สรุปภาษี ภ.พ.30) ---

function getMonthlyVatReport(month, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var salesTaxSheet = ss.getSheetByName("Sales_Tax");
  var purchaseSheet = ss.getSheetByName("Purchase_Tax");
  var invoiceSheet = ss.getSheetByName("Invoice");
  
  var report = { salesVat: 0, salesBase: 0, purchaseVat: 0, purchaseBase: 0, netVat: 0 };

  // คำนวณภาษีขายรวม (Auto + Manual)
  const calculateSales = (sheet, isInvoiceSheet) => {
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var d = new Date(data[i][isInvoiceSheet ? 0 : 1]);
      if (d.getMonth() + 1 == month && d.getFullYear() == year) {
        var total = parseFloat(data[i][isInvoiceSheet ? 3 : 8]) || 0;
        var base = isInvoiceSheet ? (total * 100 / 107) : (parseFloat(data[i][6]) || 0);
        var vat = isInvoiceSheet ? (total - base) : (parseFloat(data[i][7]) || 0);
        report.salesBase += base; report.salesVat += vat;
      }
    }
  };

  calculateSales(salesTaxSheet, false);
  calculateSales(invoiceSheet, true);

  // คำนวณภาษีซื้อ
  if (purchaseSheet) {
    var pData = purchaseSheet.getDataRange().getValues();
    for (var i = 1; i < pData.length; i++) {
      var d = new Date(pData[i][1]);
      if (d.getMonth() + 1 == month && d.getFullYear() == year) {
        report.purchaseBase += parseFloat(pData[i][6]) || 0;
        report.purchaseVat += parseFloat(pData[i][7]) || 0;
      }
    }
  }

  report.netVat = report.salesVat - report.purchaseVat;
  return report;
}

// --- 5. Utilities & Auth (ฟังก์ชันเสริม) ---

function checkDuplicateTaxNo(taxNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Purchase_Tax");
  if (!sheet) return false;
  var data = sheet.getRange("C:C").getValues().flat();
  return data.some(item => item.toString().toLowerCase().trim() === taxNo.toLowerCase().trim());
}

function scanInvoiceWithAI(base64Image) {
  var apiKey = "AIzaSyA1AhcQCPIn3MJDmiO9HENALpyeYXVcoe0";
  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=" + apiKey;
  var base64Data = base64Image.split(',')[1];
  var payload = { "contents": [{"parts": [{"text": "Extract Thai Tax Invoice JSON: {taxNo, taxID, vendor, branch, total, date(YYYY-MM-DD)}"}, {"inline_data": {"mime_type": "image/jpeg", "data": base64Data}}]}] };
  try {
    var res = UrlFetchApp.fetch(url, {method:"post", contentType:"application/json", payload:JSON.stringify(payload)});
    var resultText = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
    return { status: "success", data: JSON.parse(resultText.replace(/```json|```/g, "")) };
  } catch (e) { return { status: "error", message: e.message }; }
}

function getVendorHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Purchase_Tax");
  if (!sheet) return { names: [], dict: {} };
  var data = sheet.getDataRange().getValues();
  var dict = {}, names = [];
  for (var i = 1; i < data.length; i++) {
    var vName = data[i][3], vTax = data[i][4];
    if (vName && vTax) {
      vTax = vTax.toString().trim();
      dict[vTax] = vName; dict[vName] = vTax;
      if (names.indexOf(vName) === -1) names.push(vName);
    }
  }
  return { names: names, dict: dict };
}
function getPurchaseReportData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Purchase_Tax");
  
  // 1. ถ้าไม่มีชีท ให้คืนค่าว่างกลับไป ไม่ให้ระบบค้าง
  if (!sheet) return { stats: { total: 0, vat: 0, unpaid: 0 }, history: [] };
  
  var data = sheet.getDataRange().getValues();
  var stats = { total: 0, vat: 0, unpaid: 0 };
  var history = [];

  // เริ่มวนลูปจากแถวที่ 2 (i=1)
  for (var i = 1; i < data.length; i++) {
    try {
      if (!data[i][2]) continue; // ถ้าไม่มีเลขที่บิล ให้ข้ามแถวนี้ไป

      var total = parseFloat(data[i][8]) || 0;
      var vat = parseFloat(data[i][7]) || 0;
      var status = data[i][11] || "ยังไม่ชำระ";
      
      stats.total += total;
      stats.vat += vat;
      if (status.includes("ยังไม่ชำระ") || status.includes("ยังไม่จ่าย")) {
        stats.unpaid += total;
      }

      // ดักจับเรื่องวันที่ ถ้าไม่ใช่ Date Object ให้แสดงเป็นข้อความธรรมดา
      var dateStr = "";
      if (data[i][1] instanceof Date) {
        dateStr = Utilities.formatDate(data[i][1], "GMT+7", "dd/MM/yyyy");
      } else {
        dateStr = data[i][1] ? data[i][1].toString() : "-";
      }

      history.unshift({
        date: dateStr,
        taxNo: data[i][2],
        vendor: data[i][3],
        total: total,
        vat: vat,
        status: status,
        img: data[i][9]
      });
    } catch(err) {
      console.log("Error at row " + (i+1) + ": " + err.message);
    }
  }
  
  return { stats: stats, history: history };
}

// ฟังก์ชันดึงข้อมูล Invoice ทั้งหมดเพื่อไปทำ Dashboard
function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Invoice");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var results = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][5]) { // คอลัมน์ F ที่เก็บ JSON rawData
      results.push({
        rawData: data[i][5]
      });
    }
  }
  return results;
}
