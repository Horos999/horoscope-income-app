// ==== ตั้งแต่ตรงนี้: ทำหน้าที่เป็น JSON API เท่านั้น ====
// หน้าเว็บ (index.html) ย้ายไปโฮสต์บน GitHub Pages แล้ว จึงเรียกเข้ามาที่นี่ผ่าน fetch()
// แทนที่จะใช้ google.script.run เหมือนเดิม ตัวฐานข้อมูล (ชีต Entries) ยังเป็นตัวเดิมทุกอย่าง

function doGet(e) {
  return jsonResponse_(getData());
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'getData':
        result = getData();
        break;
      case 'upsertEntry':
        result = upsertEntry(body.date, body.cumulative);
        break;
      case 'deleteEntry':
        result = deleteEntry(body.id);
        break;
      case 'setRate':
        result = setRate(body.rate);
        break;
      default:
        throw new Error('ไม่รู้จัก action: ' + body.action);
    }
    return jsonResponse_(result);
  } catch (err) {
    return jsonResponse_({ error: String(err) });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Entries');
  if (!sheet) {
    sheet = ss.insertSheet('Entries');
    sheet.appendRow(['ID', 'Date', 'Cumulative']);
  }
  return sheet;
}

function formatDateISO_(d) {
  if (d instanceof Date) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(d);
}

function getData() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const entries = [];
  for (let i = 1; i < values.length; i++) {
    const id = values[i][0];
    const date = values[i][1];
    const cumulative = values[i][2];
    if (id === '' || id === undefined || id === null) continue;
    entries.push({
      id: String(id),
      date: formatDateISO_(date),
      cumulative: Number(cumulative)
    });
  }
  const rate = Number(PropertiesService.getDocumentProperties().getProperty('rate') || 69);
  return { rate: rate, entries: entries };
}

function upsertEntry(date, cumulative) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (formatDateISO_(values[i][1]) === date) {
      sheet.getRange(i + 1, 3).setValue(cumulative);
      return getData();
    }
  }
  const id = new Date().getTime();
  sheet.appendRow([id, date, cumulative]);
  return getData();
}

function deleteEntry(id) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getData();
}

function setRate(rate) {
  PropertiesService.getDocumentProperties().setProperty('rate', String(rate));
  return getData();
}
