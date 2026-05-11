// ส่วนสำหรับรับข้อมูลจากหน้าเว็บ (GitHub)
function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const action = params.action;
  let result;

  if (action === 'saveData') {
    result = saveData(params.type, params.record); 
  } else if (action === 'deleteData') {
    result = deleteData(params.type, params.backendId); 
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getAllData') {
    const data = getAllData(); 
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ฟังก์ชันเดิมที่คุณมีอยู่แล้ว (ไม่ต้องแก้ข้างใน)
function getAllData() { /* โค้ดเดิมของคุณ [cite: 412-413] */ }
function saveData(type, record) { /* โค้ดเดิมของคุณ [cite: 416-421] */ }
function deleteData(type, backendId) { /* โค้ดเดิมของคุณ [cite: 421-424] */ }


function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('CostPro-FP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemSheet = getOrCreateSheet(ss, "Items");
  const ingSheet = getOrCreateSheet(ss, "Ingredients");
  return { 
    items: getSheetData(itemSheet), 
    ingredients: getSheetData(ingSheet) 
  };
}

function getSheetData(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  return values.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function saveData(type, record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = type === 'ingredient' ? "Ingredients" : "Items";
  const sheet = getOrCreateSheet(ss, sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!record.__backendId) record.__backendId = Date.now().toString();
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  const idIdx = headers.indexOf("__backendId");
  for(let i=1; i<data.length; i++) { if(data[i][idIdx] == record.__backendId) { rowIndex = i + 1; break; } }
  const row = headers.map(h => record[h] || "");
  if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
  return { isOk: true };
}

function deleteData(type, backendId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = type === 'ingredient' ? "Ingredients" : "Items";
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf("__backendId");
  for(let i=1; i<data.length; i++) { if(data[i][idIdx] == backendId) { sheet.deleteRow(i + 1); return { isOk: true }; } }
  return { isOk: false };
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = name === "Ingredients" 
      ? ["__backendId", "name", "category", "quantity", "unit", "price", "created_at"]
      : ["__backendId", "category", "item_name", "ingredients", "total_cost", "selling_price", "profit", "margin_percent", "created_at"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}
