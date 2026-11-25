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
    if (data.update_id !== undefined) return handleTelegram(e, data); // webhook của Telegram
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
  if (!items.length) { tgBroadcast('📝 Hôm nay bạn chưa ghi khoản chi nào — nhớ ghi nhé!'); return; }
  const spent = items.filter(i => i.type !== 'thu');
  const total = spent.reduce((a, i) => a + i.amount, 0);
  const fmt = n => Number(n).toLocaleString('vi-VN') + 'đ';
  const lines = spent.map(i => `• ${fmt(i.amount)} — ${i.category}${i.note ? ' (' + i.note + ')' : ''}${i.who ? ' · ' + i.who : ''}`);
  const to = props().getProperty('SUMMARY_EMAIL') || Session.getEffectiveUser().getEmail();
  tgSummary();
  MailApp.sendEmail(to, `💸 Hôm nay chi ${fmt(total)}`, `Tổng chi hôm nay: ${fmt(total)}\n\n${lines.join('\n')}\n\nSheet: ${getSpreadsheet().getUrl()}`);
}

// ===== Telegram bot =====
// Cài đặt: xem README. Cần Script Properties: TELEGRAM_BOT_TOKEN, WEBAPP_URL, TELEGRAM_CHAT_IDS.
const TG_CATS = {
  chi: ['🍜 Ăn uống', '☕ Cà phê', '🛵 Đi lại', '🛍️ Mua sắm', '🎮 Giải trí', '🏠 Nhà cửa', '💊 Sức khỏe', '📦 Khác'],
  thu: ['💰 Lương', '🎁 Thưởng', '💸 Được cho', '📈 Đầu tư', '📦 Khác']
};
const TG_KEYWORDS = {
  '☕ Cà phê': ['cà phê', 'cafe', 'cafê', 'trà sữa', 'trà'],
  '🍜 Ăn uống': ['ăn', 'cơm', 'phở', 'bún', 'bánh', 'lẩu', 'nhậu', 'bia'],
  '🛵 Đi lại': ['xăng', 'grab', 'taxi', 'xe', 'gửi xe', 'vé'],
  '🛍️ Mua sắm': ['mua', 'áo', 'quần', 'giày', 'shopee', 'lazada'],
  '🎮 Giải trí': ['phim', 'game', 'karaoke', 'du lịch'],
  '🏠 Nhà cửa': ['điện', 'nước nhà', 'tiền nhà', 'internet', 'wifi', 'gas'],
  '💊 Sức khỏe': ['thuốc', 'khám', 'bệnh viện', 'gym'],
  '💰 Lương': ['lương'],
  '🎁 Thưởng': ['thưởng']
};
const TG_HELP = '<b>💸 Bot chi tiêu</b>\n' +
  'Gõ thẳng để ghi: <code>cà phê 35k</code>, <code>ăn trưa 45</code>, <code>lương 12tr</code>\n\n' +
  '/today — hôm nay\n/week — 7 ngày qua\n/month — tháng này\n/top — 5 khoản chi lớn nhất tháng\n' +
  '/budget 5tr — đặt ngân sách tháng\n/undo — xoá khoản vừa ghi';

const vnd = n => String(Math.round(Number(n) || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, '.') + 'đ';
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const dayKey = d => Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd');
const monthKey = d => Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM');
const sumOf = list => list.reduce((a, i) => a + i.amount, 0);
const isChi = i => i.type !== 'thu';
const bar = pct => '█'.repeat(Math.min(10, Math.round(pct / 10))) + '░'.repeat(Math.max(0, 10 - Math.round(pct / 10)));

function tgChatIds() {
  return (props().getProperty('TELEGRAM_CHAT_IDS') || '').split(',').map(s => s.trim()).filter(Boolean);
}

// Chạy 1 lần: đăng ký webhook và menu lệnh cho bot
function setupTelegram() {
  const token = props().getProperty('TELEGRAM_BOT_TOKEN');
  const base = props().getProperty('WEBAPP_URL');
  if (!token || !base) throw new Error('Cần đặt TELEGRAM_BOT_TOKEN và WEBAPP_URL trong Script Properties');
  let secret = props().getProperty('TELEGRAM_SECRET');
  if (!secret) {
    secret = Utilities.getUuid().replace(/-/g, '');
    props().setProperty('TELEGRAM_SECRET', secret);
  }
  const call = (method, payload) => UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(payload)
  }).getContentText();
  Logger.log(call('setWebhook', { url: base + '?tg=' + secret, allowed_updates: ['message'], drop_pending_updates: true }));
  Logger.log(call('setMyCommands', { commands: [
    { command: 'today', description: 'Chi tiêu hôm nay' },
    { command: 'week', description: '7 ngày qua' },
    { command: 'month', description: 'Tháng này' },
    { command: 'top', description: '5 khoản chi lớn nhất tháng' },
    { command: 'budget', description: 'Đặt/xem ngân sách tháng' },
    { command: 'undo', description: 'Xoá khoản vừa ghi' },
    { command: 'help', description: 'Hướng dẫn' }
  ] }));
  if (!tgChatIds().length) Logger.log('Chưa có TELEGRAM_CHAT_IDS: nhắn /start cho bot để lấy Chat ID rồi thêm vào Script Properties.');
}

function handleTelegram(e, update) {
  try {
    const secret = props().getProperty('TELEGRAM_SECRET');
    if (!secret || !e.parameter || e.parameter.tg !== secret) return json({ ok: false });
    const msg = update.message;
    if (!msg || !msg.text) return json({ ok: true });
    // Telegram gửi lại nếu phản hồi chậm — bỏ qua update đã xử lý
    const cache = CacheService.getScriptCache();
    if (cache.get('tg_' + update.update_id)) return json({ ok: true });
    cache.put('tg_' + update.update_id, '1', 600);

    const chatId = String(msg.chat.id);
    if (!tgChatIds().includes(chatId)) {
      tgSend(chatId, '🔒 Chat này chưa được cấp quyền.\nChat ID: <code>' + chatId + '</code>\nThêm ID này vào TELEGRAM_CHAT_IDS trong Script Properties.');
      return json({ ok: true });
    }
    let reply;
    try { reply = tgReply(msg.text.trim(), msg.from || {}); }
    catch (err) { reply = '⚠️ Lỗi: ' + esc(err.message || err); }
    tgSend(chatId, reply);
  } catch (err) {
    console.error(err);
  }
  return json({ ok: true });
}

function tgReply(text, from) {
  if (!text.startsWith('/')) return addFromText(text, from);
  const parts = text.split(/\s+/);
  const cmd = parts[0].toLowerCase().replace(/@.*$/, '');
  const arg = parts.slice(1).join(' ');
  switch (cmd) {
    case '/start':
    case '/help': return TG_HELP;
    case '/today': return reportDay(tgData().items, dayKey(new Date()), '📅 Hôm nay');
    case '/week': return reportWeek(tgData().items);
    case '/month': { const d = tgData(); return reportMonth(d.items, d.budget); }
    case '/top': return reportTop(tgData().items);
    case '/budget': {
      if (!arg) { const b = tgData().budget; return b ? 'Ngân sách tháng: <b>' + vnd(b) + '</b>' : 'Chưa đặt ngân sách. Dùng <code>/budget 5tr</code>'; }
      const amount = parseAmount(arg.toLowerCase()).amount;
      if (!amount) return 'Không hiểu số tiền. Ví dụ: <code>/budget 5tr</code>';
      ACTIONS.setBudget({ budget: amount });
      return '🎯 Đã đặt ngân sách tháng: <b>' + vnd(amount) + '</b>';
    }
    case '/undo': {
      const last = tgData().items[0];
      if (!last) return 'Không có khoản nào để xoá.';
      ACTIONS.delete({ id: last.id });
      return '↩️ Đã xoá: ' + vnd(last.amount) + ' — ' + esc(last.category) + (last.note ? ' (' + esc(last.note) + ')' : '');
    }
    default: return 'Lệnh không có. Gõ /help để xem hướng dẫn.';
  }
}

function addFromText(text, from) {
  const p = parseExpense(text);
  if (!p.amount) return 'Mình chưa thấy số tiền. Thử: <code>cà phê 35k</code> hoặc <code>ăn trưa 45</code>';
  ACTIONS.add({
    id: Utilities.getUuid(), time: new Date().toISOString(), type: p.type, amount: p.amount,
    category: p.category, note: p.note, who: from.first_name || 'Telegram'
  });
  const today = tgData().items.filter(i => i.day === dayKey(new Date()) && isChi(i));
  return (p.type === 'thu' ? '💰 Đã ghi thu ' : '✅ Đã ghi ') + '<b>' + vnd(p.amount) + '</b> — ' + esc(p.category) + (p.note ? ' (' + esc(p.note) + ')' : '') +
    '\nHôm nay đã chi: <b>' + vnd(sumOf(today)) + '</b>';
}

// "35k" "35 nghìn" "1.500.000" "12tr" "45" (số nhỏ hơn 1000 hiểu là nghìn đồng)
function parseAmount(t) {
  let m = t.match(/\d{1,3}(?:[.,]\d{3})+(?!\d)(?:\s*(?:đồng|đ))?/);
  if (m) return { amount: +m[0].replace(/\D/g, ''), raw: m[0] };
  m = t.match(/(\d+(?:[.,]\d+)?)\s*(triệu|tr|nghìn|ngàn|ngh|k|đồng|đ)?(?![\p{L}])/u);
  if (!m) return { amount: 0, raw: '' };
  const n = parseFloat(m[1].replace(',', '.')), u = m[2] || '';
  const amount = /^(triệu|tr)$/.test(u) ? n * 1e6 : /^(nghìn|ngàn|ngh|k)$/.test(u) ? n * 1e3 : (n < 1000 ? n * 1e3 : n);
  return { amount: Math.round(amount), raw: m[0] };
}

function parseExpense(text) {
  const t = text.toLowerCase();
  const { amount, raw } = parseAmount(t);
  const padded = ' ' + t.replace(/[.,!?]/g, ' ') + ' '; // khớp nguyên từ, tránh "xăng" chứa "ăn"
  let category = '📦 Khác';
  for (const c of Object.keys(TG_KEYWORDS)) {
    if (TG_KEYWORDS[c].some(w => padded.includes(' ' + w + ' '))) { category = c; break; }
  }
  const idx = raw ? t.indexOf(raw) : -1;
  const note = (idx < 0 ? text : text.slice(0, idx) + text.slice(idx + raw.length)).replace(/\s+/g, ' ').trim();
  return { amount, category, note, type: TG_CATS.thu.includes(category) && category !== '📦 Khác' ? 'thu' : 'chi' };
}

function tgData() {
  const r = ACTIONS.list({ since: new Date(Date.now() - 70 * 86400000).toISOString() });
  return { budget: r.budget, items: r.items.map(i => Object.assign({}, i, { day: dayKey(i.time), month: monthKey(i.time) })) };
}

function reportDay(all, key, title) {
  const list = all.filter(i => i.day === key);
  if (!list.length) return title + ': chưa ghi khoản nào.';
  const thu = list.filter(i => !isChi(i));
  const lines = list.slice(0, 25).reverse().map(i =>
    (isChi(i) ? '• ' : '+ ') + vnd(i.amount) + ' — ' + esc(i.category) + (i.note ? ' (' + esc(i.note) + ')' : '') + (i.who ? ' · ' + esc(i.who) : ''));
  return '<b>' + title + '</b>\nChi: <b>' + vnd(sumOf(list.filter(isChi))) + '</b>' + (thu.length ? ' · Thu: <b>' + vnd(sumOf(thu)) + '</b>' : '') + '\n' + lines.join('\n');
}

function reportWeek(all) {
  const days = [];
  for (let k = 6; k >= 0; k--) days.push(dayKey(Date.now() - k * 86400000));
  const per = days.map(d => sumOf(all.filter(i => i.day === d && isChi(i))));
  const max = Math.max.apply(null, per.concat(1));
  const total = per.reduce((a, b) => a + b, 0);
  const rows = days.map((d, k) => d.slice(8) + '/' + d.slice(5, 7) + ' ' + ('█'.repeat(Math.round(per[k] / max * 10)) || '▏') + ' ' + vnd(per[k]));
  return '<b>📆 7 ngày qua</b>\nTổng chi: <b>' + vnd(total) + '</b> (TB ' + vnd(total / 7) + '/ngày)\n<pre>' + rows.join('\n') + '</pre>';
}

function reportMonth(all, budget) {
  const now = new Date(), mk = monthKey(now);
  const p = new Date(); p.setDate(1); p.setMonth(p.getMonth() - 1);
  const cur = all.filter(i => i.month === mk);
  if (!cur.length) return 'Tháng này chưa có giao dịch nào.';
  const chi = cur.filter(isChi), total = sumOf(chi), income = sumOf(cur.filter(i => !isChi(i)));
  const prev = sumOf(all.filter(i => i.month === monthKey(p) && isChi(i)));
  const byCat = {};
  chi.forEach(i => { byCat[i.category] = (byCat[i.category] || 0) + i.amount; });
  let out = '<b>🗓 Tháng ' + mk.slice(5) + '/' + mk.slice(0, 4) + '</b>\nChi: <b>' + vnd(total) + '</b>' + (income ? ' · Thu: <b>' + vnd(income) + '</b>' : '');
  if (prev) out += '\n' + (total >= prev ? '▲' : '▼') + ' ' + Math.abs(Math.round((total - prev) / prev * 100)) + '% so với tháng trước';
  if (budget) {
    const pct = Math.round(total / budget * 100);
    const left = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate() - now.getDate() + 1;
    out += '\n\n🎯 ' + bar(pct) + ' ' + pct + '% (' + vnd(total) + '/' + vnd(budget) + ')\n' +
      (total > budget ? 'Đã vượt ' + vnd(total - budget) + ' 😱' : 'Còn ' + vnd(budget - total) + ' ≈ ' + vnd((budget - total) / left) + '/ngày');
  }
  out += '\n\n' + Object.keys(byCat).sort((a, b) => byCat[b] - byCat[a]).map(c => esc(c) + ' ' + vnd(byCat[c]) + ' (' + Math.round(byCat[c] / total * 100) + '%)').join('\n');
  return out;
}

function reportTop(all) {
  const mk = monthKey(new Date());
  const top = all.filter(i => i.month === mk && isChi(i)).sort((a, b) => b.amount - a.amount).slice(0, 5);
  if (!top.length) return 'Tháng này chưa có khoản chi nào.';
  return '<b>🏆 Top chi tháng này</b>\n' + top.map((i, k) => (k + 1) + '. ' + vnd(i.amount) + ' — ' + esc(i.category) + (i.note ? ' (' + esc(i.note) + ')' : '') + ' · ' + i.day.slice(8) + '/' + i.day.slice(5, 7)).join('\n');
}

function tgSend(chatId, text) {
  const token = props().getProperty('TELEGRAM_BOT_TOKEN');
  if (!token) return;
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ chat_id: chatId, text: text.slice(0, 4000), parse_mode: 'HTML', disable_web_page_preview: true })
  });
}

function tgBroadcast(text) {
  tgChatIds().forEach(id => tgSend(id, text));
}

// Tóm tắt cuối ngày gửi qua Telegram (gọi từ dailySummary lúc 21h)
function tgSummary() {
  if (!tgChatIds().length) return;
  const d = tgData();
  tgBroadcast(reportDay(d.items, dayKey(new Date()), '📅 Tổng kết hôm nay') + '\n\n' + reportMonth(d.items, d.budget));
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
