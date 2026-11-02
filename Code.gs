// Google Apps Script — dán vào Extensions > Apps Script của Google Sheet
const FOLDER_NAME = 'Snap Chi Tieu';
const SHEET_NAME = 'ChiTieu';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Thời gian', 'Số tiền', 'Danh mục', 'Ghi chú', 'Ảnh', 'Link ảnh']);
      sheet.setFrozenRows(1);
    }

    // Lưu ảnh vào Google Drive
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
    const base64 = data.image.split(',')[1];
    const time = new Date(data.time);
    const name = Utilities.formatDate(time, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.jpg';
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', name));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const imgUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400';

    sheet.appendRow([time, data.amount, data.category, data.note, '', file.getUrl()]);
    const row = sheet.getLastRow();
    sheet.getRange(row, 5).setFormula('=IMAGE("' + imgUrl + '")');
    sheet.getRange(row, 2).setNumberFormat('#,##0"đ"');
    sheet.setRowHeight(row, 80);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// Script gắn với Sheet thì dùng Sheet đó; script độc lập thì tự tạo Sheet "Snap Chi Tieu" và nhớ ID
function getSpreadsheet() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create(FOLDER_NAME);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
