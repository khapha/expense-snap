// Google Apps Script — backend cho Snap Chi Tiêu
// Cài đặt lần đầu: chọn hàm `setup` ở thanh công cụ rồi bấm Run, xem token trong Execution log.
const FOLDER_NAME = 'Snap Chi Tieu';
const SHEET_NAME = 'GiaoDich';
const HEADERS = ['ID', 'Thời gian', 'Loại', 'Số tiền', 'Danh mục', 'Ghi chú', 'Người', 'Ảnh', 'Link ảnh', 'FileId'];
const COL = { id: 1, time: 2, type: 3, amount: 4, category: 5, note: 6, who: 7, image: 8, link: 9, fileId: 10 };
const CLAUDE_MODEL = 'claude-opus-5-5';

const props = () => PropertiesService.getScriptProperties();

// ===== Chạy 1 lần từ editor =====
function setup() {
  let token = props().getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid().slice(0, 8);
    props().setProperty('TOKEN', token);
  }
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailySummary')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailySummary').timeBased().everyDays(1).atHour(21).create();
  getSheet();
  Logger.log('Mã bí mật (nhập vào app): ' + token);
  Logger.log('Sheet: ' + getSpreadsheet().getUrl());
}

// ===== API =====
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const token = props().getProperty('TOKEN');
    if (token && data.token !== token) throw new Error('Sai mã bí mật');
    const handler = ACTIONS[data.action || 'add'];
    if (!handler) throw new Error('Hành động không hợp lệ');
    return json({ ok: true, ...handler(data) });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}

const ACTIONS = {
  add(d) {
    const sheet = getSheet();
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      // Bỏ qua nếu ID đã có (gửi lại từ hàng đợi offline)
      if (findRow(sheet, d.id)) return { duplicate: true };
      const time = new Date(d.time);
      let link = '', fileId = '', formula = '';
      if (d.image) {
        const file = saveImage(d.image, time);
        fileId = file.getId();
        link = file.getUrl();
        formula = '=IMAGE("' + thumbUrl(fileId) + '")';
      }
      sheet.appendRow([d.id || Utilities.getUuid(), time, d.type || 'chi', d.amount, d.category, d.note || '', d.who || '', formula, link, fileId]);
      const row = sheet.getLastRow();
      sheet.getRange(row, COL.amount).setNumberFormat('#,##0"đ"');
      if (fileId) sheet.setRowHeight(row, 80);
    } finally {
      lock.releaseLock();
    }
    return {};
  },

  list(d) {
    const sheet = getSheet();
    const n = sheet.getLastRow() - 1;
    const rows = n > 0 ? sheet.getRange(2, 1, n, HEADERS.length).getValues() : [];
    const since = d.since ? new Date(d.since) : null;
    const items = rows
      .filter(r => r[0] && (!since || new Date(r[1]) >= since))
      .map(r => ({
        id: r[0], time: new Date(r[1]).toISOString(), type: r[2], amount: Number(r[3]),
        category: r[4], note: r[5], who: r[6], thumb: r[9] ? thumbUrl(r[9]) : ''
      }))
      .reverse();
    return { items, budget: Number(props().getProperty('BUDGET') || 0) };
  },

  update(d) {
    const sheet = getSheet();
    const row = findRow(sheet, d.id);
    if (!row) throw new Error('Không tìm thấy giao dịch');
    if (d.amount != null) sheet.getRange(row, COL.amount).setValue(d.amount);
    if (d.category != null) sheet.getRange(row, COL.category).setValue(d.category);
    if (d.note != null) sheet.getRange(row, COL.note).setValue(d.note);
    if (d.type != null) sheet.getRange(row, COL.type).setValue(d.type);
    return {};
  },

  delete(d) {
    const sheet = getSheet();
    const row = findRow(sheet, d.id);
    if (!row) throw new Error('Không tìm thấy giao dịch');
    const fileId = sheet.getRange(row, COL.fileId).getValue();
    if (fileId) {
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
    }
    sheet.deleteRow(row);
    return {};
  },

  setBudget(d) {
    props().setProperty('BUDGET', String(Number(d.budget) || 0));
    return {};
  },

  // Đọc hoá đơn bằng Claude: trả về số tiền, danh mục, ghi chú gợi ý
  ocr(d) {
    const key = props().getProperty('ANTHROPIC_API_KEY');
    if (!key) throw new Error('Chưa cài ANTHROPIC_API_KEY trong Script Properties');
    const [header, base64] = d.image.split(',');
    const mediaType = header.match(/data:(.*?);/)[1];
    const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      headers: {
        'x-api-key': key,
        'anthropic-version': '2023-06-01',
        'anthropic-beta': 'server-side-fallback-2026-07-01'
      },
      payload: JSON.stringify({
        model: CLAUDE_MODEL,
        max_tokens: 1024,
        fallbacks: 'default',
        output_config: {
          effort: 'low',
          format: {
            type: 'json_schema',
            schema: {
              type: 'object',
              properties: {
                amount: { type: 'integer', description: 'Tổng tiền phải trả, đơn vị đồng; 0 nếu không thấy' },
                category: { type: 'string', enum: d.categories },
                note: { type: 'string', description: 'Mô tả ngắn, vd tên quán hoặc món' }
              },
              required: ['amount', 'category', 'note'],
              additionalProperties: false
            }
          }
        },
        messages: [{
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: mediaType, data: base64 } },
            { type: 'text', text: 'Đây là ảnh một khoản chi tiêu (hoá đơn, bill, hoặc món đồ). Hãy đọc tổng số tiền và chọn danh mục phù hợp nhất.' }
          ]
        }]
      })
    });
    const body = JSON.parse(res.getContentText());
    if (res.getResponseCode() !== 200) throw new Error('Claude API: ' + (body.error && body.error.message));
    if (body.stop_reason === 'refusal') throw new Error('Claude từ chối xử lý ảnh này');
    const text = body.content.filter(b => b.type === 'text').map(b => b.text).join('');
    return { result: JSON.parse(text) };
  }
};

// ===== Tóm tắt cuối ngày (trigger 21h) =====
function dailySummary() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const items = ACTIONS.list({ since: today.toISOString() }).items;
  if (!items.length) return;
  const spent = items.filter(i => i.type !== 'thu');
  const total = spent.reduce((a, i) => a + i.amount, 0);
  const fmt = n => Number(n).toLocaleString('vi-VN') + 'đ';
  const lines = spent.map(i => `• ${fmt(i.amount)} — ${i.category}${i.note ? ' (' + i.note + ')' : ''}${i.who ? ' · ' + i.who : ''}`);
  const to = props().getProperty('SUMMARY_EMAIL') || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail(to, `💸 Hôm nay chi ${fmt(total)}`, `Tổng chi hôm nay: ${fmt(total)}\n\n${lines.join('\n')}\n\nSheet: ${getSpreadsheet().getUrl()}`);
}

// ===== Helpers =====
// Script gắn với Sheet thì dùng Sheet đó; script độc lập thì tự tạo Sheet "Snap Chi Tieu" và nhớ ID
function getSpreadsheet() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = props().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create(FOLDER_NAME);
  props().setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function getSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.hideColumns(COL.fileId);
  }
  return sheet;
}

function findRow(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return 0;
  const found = sheet.getRange(2, COL.id, sheet.getLastRow() - 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return found ? found.getRow() : 0;
}

function saveImage(dataUrl, time) {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
  const name = Utilities.formatDate(time, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.jpg';
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(dataUrl.split(',')[1]), 'image/jpeg', name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file;
}

const thumbUrl = id => 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400';

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
