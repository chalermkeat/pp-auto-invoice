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
  
  // 🚀 ดึงข้อมูลโปรไฟล์บริษัทเจ้าของ (เรียกใช้ฟังก์ชัน getCompanyProfile ที่เราทำไว้)
  var company = getCompanyProfile(); 
  
  var history = [];
  var stats = { total: 0, vat: 0 };

  // 1. ดึงจาก Sales_Tax (ที่คีย์มือ)
  if (salesTaxSheet) {
    var data = salesTaxSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (!data[i][2]) continue; // ข้ามถ้าไม่มีเลขที่บิล
      
      var total = parseFloat(data[i][8]) || 0;
      var vat = parseFloat(data[i][7]) || 0;
      stats.total += total; 
      stats.vat += vat;
      
      history.push({
        date: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], "GMT+7", "dd/MM/yyyy") : (data[i][1] || "-"),
        taxNo: data[i][2], 
        customer: data[i][3], 
        taxID: data[i][4] ? data[i][4].toString().replace(/'/g, "").trim() : "-", // เลขผู้เสียภาษีลูกค้า
        branch: data[i][5] ? data[i][5].toString().replace(/'/g, "").trim() : "00000", // สาขาลูกค้า
        total: total, 
        vat: vat, 
        source: "Manual"
      });
    }
  }

  // 2. ดึงจากหน้าออกใบกำกับภาษี (Auto ดึงจาก Invoice Sheet)
  if (invoiceSheet) {
    var invData = invoiceSheet.getDataRange().getValues();
    for (var i = 1; i < invData.length; i++) {
      if (!invData[i][1]) continue; // ข้ามถ้าไม่มีเลขที่บิล
      
      var total = parseFloat(invData[i][3]) || 0;
      var base = (total * 100 / 107);
      var vat = total - base;
      stats.total += total; 
      stats.vat += vat;
      
      // ดึงข้อมูล TaxID และ Branch จาก JSON ในคอลัมน์ F (Index 5)
      var raw = {};
      try { 
        raw = JSON.parse(invData[i][5] || "{}"); 
      } catch(e) {
        console.log("Error parsing JSON at row " + (i+1));
      }

      history.push({
        date: invData[i][0] instanceof Date ? Utilities.formatDate(invData[i][0], "GMT+7", "dd/MM/yyyy") : (invData[i][0] || "-"),
        taxNo: invData[i][1], 
        customer: invData[i][2], 
        taxID: raw.taxID ? raw.taxID.toString().replace(/'/g, "").trim() : "-",
        branch: raw.branch ? raw.branch.toString().replace(/'/g, "").trim() : "00000",
        total: total, 
        vat: vat, 
        source: "System"
      });
    }
  }

  // 🚀 ส่งกลับทั้งสถิติ, ประวัติ และข้อมูลบริษัทเจ้าของ
  return { 
    stats: stats, 
    history: history.reverse(), 
    company: company 
  };
}

// --- 3. Purchase Tax System (ระบบภาษีซื้อ) ---

// --- แก้ไขใน Code.gs ตรงส่วน savePurchase ---

function savePurchase(payload, base64Image) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Purchase_Tax") || ss.insertSheet("Purchase_Tax");
  
  // 🚀 ซ่อมจุดนี้: เช็กก่อนว่ามีสาขาไหม ถ้าไม่มีให้ใส่ 00000 และเติม ' เพื่อให้เป็นข้อความ
  var branchText = "'" + (payload.branch || "00000"); 

  try {
    var targetRow = -1;
    var data = sheet.getDataRange().getValues();
    if (payload.oldTaxNo) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][2] == payload.oldTaxNo) { targetRow = i + 1; break; }
      }
    }

    if (targetRow != -1) {
      sheet.getRange(targetRow, 2, 1, 8).setValues([[
        payload.date, payload.taxNo, payload.vendor, payload.taxID, 
        branchText, payload.base, payload.vat, payload.total
      ]]);
      sheet.getRange(targetRow, 12).setValue(payload.paymentStatus);
    } else {
      sheet.appendRow([
        new Date(), payload.date, payload.taxNo, payload.vendor, payload.taxID,
        branchText, payload.base, payload.vat, payload.total, "",
        payload.user, payload.paymentStatus
      ]);
    }
    return { status: "success" };
  } catch (e) {
    return { status: "error", message: e.message };
  }
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
        taxID: data[i][4],    // 🚀 เพิ่มบรรทัดนี้: ดึงคอลัมน์ E (Index 4)
        branch: data[i][5],   // 🚀 เพิ่มบรรทัดนี้: ดึงคอลัมน์ F (Index 5)
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

// ฟังก์ชันค้นหาข้อมูลบิลซื้อตามเลขที่ (สำหรับโหมดแก้ไข)
function getPurchaseByTaxNo(taxNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Purchase_Tax");
  if (!sheet) return { status: "error", message: "ไม่พบชีท Purchase_Tax" };
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    // เทียบเลขที่บิล (คอลัมน์ C)
    if (data[i][2].toString().trim() === taxNo.toString().trim()) {
      return {
        status: "success",
        data: {
          date: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], "GMT+7", "yyyy-MM-dd") : data[i][1],
          taxNo: data[i][2],
          vendor: data[i][3],
          taxID: data[i][4],
          branch: data[i][5].toString().replace(/'/g, ""), // ลบเครื่องหมาย ' ออกก่อนแสดงผล
          total: data[i][8],
          paymentStatus: data[i][11]
        }
      };
    }
  }
  return { status: "error", message: "ไม่พบข้อมูลบิลเลขที่ " + taxNo };
}

// 🚀 ฟังก์ชันเทพ: ดึงข้อมูลชื่อบริษัทจากเลขผู้เสียภาษี (นิติบุคคลใหม่)
function fetchCompanyFromCreden(taxID) {
  if (!taxID || taxID.length !== 13) return { status: "error", message: "เลขไม่ครบ 13 หลัก" };
  
  // ท่อทางเลือก 1: ดึงจากฐานข้อมูลสาธารณะ (ผ่าน Creden Data API หรือคล้ายคลึง)
  var url = "https://data.creden.co/api/v1/company/" + taxID; 
  
  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var resCode = response.getResponseCode();
    
    if (resCode === 200) {
      var json = JSON.parse(response.getContentText());
      // ปรับตามโครงสร้าง JSON ของผู้ให้บริการ
      var companyName = json.name_th || json.company_name || ""; 
      if (companyName) {
        return { status: "success", name: companyName };
      }
    }
    
    // ท่อทางเลือก 2: ถ้า API แรกไม่ตอบสนอง หรือหาไม่เจอ
    // ระบบจะพยายามค้นหาจากประวัติใน Invoice เดิม (เผื่อเคยออกบิลขายให้เขา)
    var invoiceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoice");
    if (invoiceSheet) {
      var invData = invoiceSheet.getDataRange().getValues();
      for (var i = 1; i < invData.length; i++) {
        var raw = JSON.parse(invData[i][5] || "{}");
        if (raw.taxID === taxID) return { status: "success", name: raw.custName || "" };
      }
    }

    return { status: "not_found", message: "ไม่พบข้อมูลในฐานข้อมูลกลาง" };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

// 🚀 ฟังก์ชันลบรายการภาษีซื้อจากเลขที่ใบกำกับภาษี
function deletePurchase(taxNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Purchase_Tax");
  if (!sheet) return { status: "error", message: "ไม่พบชีทข้อมูล" };
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][2].toString().trim() === taxNo.toString().trim()) {
      sheet.deleteRow(i + 1); // ลบแถวที่เจอ (+1 เพราะ Array เริ่มที่ 0)
      return { status: "success", message: "ลบข้อมูลเรียบร้อยแล้ว" };
    }
  }
  return { status: "error", message: "ไม่พบเลขที่บิลที่ต้องการลบ" };
}

// 1. ฟังก์ชันดึงข้อมูลโปรไฟล์บริษัทเจ้าของ (จากชีท Company_Profile)
function getCompanyProfile() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Company_Profile");
    if (!sheet) return null;
    
    // ดึงข้อมูลแถวที่ 2 (คอลัมน์ A, B, C, D)
    var data = sheet.getRange(2, 1, 1, 4).getValues()[0];
    
    return {
      name: data[0] || "",       // คอลัมน์ A: ชื่อผู้ประกอบการ
      storeName: data[1] || "",  // คอลัมน์ B: ชื่อสถานประกอบการ
      taxID: data[2] ? data[2].toString().replace(/'/g, "").trim() : "", // คอลัมน์ C: เลขประจำตัวผู้เสียภาษี
      branch: data[3] !== "" ? data[3].toString().replace(/'/g, "").trim() : "00000" // คอลัมน์ D: รหัสสาขา
    };
  } catch (e) {
    console.log("Error getting company profile: " + e.message);
    return null;
  }
}

// 2. ฟังก์ชันหลักสำหรับดึงข้อมูลรายงานภาษีซื้อ
function getPurchaseReportData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Purchase_Tax");
  
  // ดึงข้อมูลบริษัทเจ้าของมาก่อน
  var company = getCompanyProfile();
  
  // ถ้าไม่มีชีทภาษีซื้อ ให้ส่งค่าว่างและข้อมูลบริษัทกลับไป
  if (!sheet) return { stats: { total: 0, vat: 0, unpaid: 0 }, history: [], company: company };
  
  var data = sheet.getDataRange().getValues();
  var stats = { total: 0, vat: 0, unpaid: 0 };
  var history = [];

  // เริ่มวนลูปจากแถวที่ 2 (i=1)
  for (var i = 1; i < data.length; i++) {
    try {
      // ตรวจสอบว่ามีเลขที่บิลหรือไม่ (คอลัมน์ C)
      if (!data[i][2]) continue; 

      var total = parseFloat(data[i][8]) || 0; // คอลัมน์ I: ยอดรวม
      var vat = parseFloat(data[i][7]) || 0;   // คอลัมน์ H: VAT
      var status = data[i][11] || "ยังไม่ชำระ"; // คอลัมน์ L: สถานะ
      
      stats.total += total;
      stats.vat += vat;
      if (status.includes("ยังไม่ชำระ") || status.includes("ยังไม่จ่าย")) {
        stats.unpaid += total;
      }

      // จัดการรูปแบบวันที่
      var dateStr = "";
      if (data[i][1] instanceof Date) { // คอลัมน์ B: วันที่
        dateStr = Utilities.formatDate(data[i][1], "GMT+7", "dd/MM/yyyy");
      } else {
        dateStr = data[i][1] ? data[i][1].toString() : "-";
      }

      // รวบรวมข้อมูลรายการ
      history.unshift({
        date: dateStr,
        taxNo: data[i][2], // คอลัมน์ C
        vendor: data[i][3], // คอลัมน์ D
        taxID: data[i][4] ? data[i][4].toString().replace(/'/g, "").trim() : "-", // คอลัมน์ E
        branch: data[i][5] ? data[i][5].toString().replace(/'/g, "").trim() : "", // คอลัมน์ F
        total: total,
        vat: vat,
        status: status,
        img: data[i][9] // คอลัมน์ J
      });
    } catch(err) {
      console.log("Error at row " + (i+1) + ": " + err.message);
    }
  }
  
  // ส่งข้อมูลทั้งหมดกลับไปที่หน้า HTML
  return { 
    stats: stats, 
    history: history, 
    company: company 
  };
}

// สคริปต์สำหรับรันครั้งเดียว เพื่อดึง Tax ID จาก JSON (คอลัมน์ F) ไปใส่คอลัมน์ G
function extractTaxIdToColumnG() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Invoice");
  if (!sheet) return;
  
  var data = sheet.getDataRange().getValues();
  
  // วนลูปตั้งแต่แถวที่ 2 เป็นต้นไป
  for (var i = 1; i < data.length; i++) {
    var jsonString = data[i][5]; // คอลัมน์ F (Index 5)
    var currentTaxId = data[i][6]; // คอลัมน์ G (Index 6)
    
    // ถ้ามี JSON และคอลัมน์ G ยังว่างเปล่า
    if (jsonString && !currentTaxId) {
      try {
        var parsedData = JSON.parse(jsonString);
        if (parsedData.taxID && parsedData.taxID !== "-") {
          // นำ taxID ไปหยอดลงคอลัมน์ G (แถวที่ i+1, คอลัมน์ 7)
          sheet.getRange(i + 1, 7).setValue("'" + parsedData.taxID);
        }
      } catch (e) {
        console.log("Error parsing JSON in row " + (i + 1));
      }
    }
  }
}

// ฟังก์ชันสำหรับลบข้อมูลบิลภาษีซื้อใน Code.gs
function deletePurchaseRecord(taxNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Purchase_Tax"); // เช็คชื่อชีทให้ตรงกับของคุณเบิ้มด้วยนะครับ
    if (!sheet) return {status: 'error', message: 'ไม่พบฐานข้อมูลภาษีซื้อ'};
    
    var data = sheet.getDataRange().getValues();
    
    // วนลูปหาเลขที่บิลที่ตรงกัน (สมมติว่าเลขที่บิลอยู่คอลัมน์ B หรือ Index 1)
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == taxNo || data[i][2] == taxNo) { 
        sheet.deleteRow(i + 1); // ลบแถวทิ้ง
        return {status: 'success'};
      }
    }
    return {status: 'error', message: 'ไม่พบเลขที่บิลนี้ในระบบ'};
  } catch(e) {
    return {status: 'error', message: e.toString()};
  }
}
