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

  // Đọc hoá đơn bằng AI (Gemini miễn phí hoặc Claude): trả về số tiền, danh mục, ghi chú gợi ý
  ocr(d) {
    return { result: readReceipt(d.image, d.categories) };
  }
};

// ===== Đọc hoá đơn bằng AI =====
// AI_PROVIDER = 'gemini' | 'claude'. Mặc định: gemini nếu có GEMINI_API_KEY, không thì claude.
function readReceipt(image, categories) {
  const provider = props().getProperty('AI_PROVIDER') || (props().getProperty('GEMINI_API_KEY') ? 'gemini' : 'claude');
  const r = provider === 'claude' ? ocrClaude(image, categories) : ocrGemini(image, categories);
  return {
    amount: Math.max(0, Math.round(Number(r.amount) || 0)),
    category: categories.includes(r.category) ? r.category : '📦 Khác',
    note: String(r.note || '').slice(0, 80)
  };
}

const receiptPrompt = categories =>
  'Đây là ảnh một khoản chi tiêu (hoá đơn, bill, hoặc món đồ). Hãy đọc tổng số tiền phải trả (đơn vị đồng, số nguyên, 0 nếu không thấy) ' +
  'và chọn danh mục phù hợp nhất trong: ' + categories.join(', ') + '. Ghi chú là mô tả ngắn như tên quán hoặc món. ' +
  'Chỉ trả về JSON dạng {"amount": 45000, "category": "...", "note": "..."}.';

function ocrGemini(image, categories) {
  const key = props().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('Chưa cài GEMINI_API_KEY trong Script Properties');
  const model = props().getProperty('GEMINI_MODEL') || 'gemini-3.5-flash-lite';
  const [header, base64] = image.split(',');
  const mime = header.match(/data:(.*?);/)[1];
  const call = generationConfig => UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({
      contents: [{ parts: [{ inline_data: { mime_type: mime, data: base64 } }, { text: receiptPrompt(categories) }] }],
      generationConfig
    })
  });
  // 503 = Gemini quá tải tạm thời: thử lại tối đa 3 lần, chờ lâu dần
  const callWithRetry = cfg => {
    let r;
    for (let i = 0; i < 3; i++) {
      r = call(cfg);
      if (r.getResponseCode() !== 503) break;
      Utilities.sleep(2000 * (i + 1));
    }
    return r;
  };
  const schema = {
    type: 'OBJECT',
    properties: { amount: { type: 'INTEGER' }, category: { type: 'STRING', enum: categories }, note: { type: 'STRING' } },
    required: ['amount', 'category', 'note']
  };
  let res = callWithRetry({ responseMimeType: 'application/json', responseSchema: schema });
  // API từ chối lược đồ (khác phiên bản) thì thử lại chỉ với yêu cầu JSON trong prompt
  if (res.getResponseCode() === 400) res = callWithRetry({ responseMimeType: 'application/json' });
  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText());
  if (code === 429) throw new Error('Gemini đã hết hạn mức miễn phí, thử lại sau ít phút');
  if (code === 503) throw new Error('Gemini đang quá tải, hãy gửi lại ảnh sau ít phút');
  if (code !== 200) throw new Error('Gemini API (' + code + '): ' + (body.error && body.error.message));
  const parts = body.candidates && body.candidates[0] && body.candidates[0].content && body.candidates[0].content.parts;
  if (!parts) throw new Error('Gemini không trả kết quả' + (body.promptFeedback ? ': ' + body.promptFeedback.blockReason : ''));
  const text = parts.map(p => p.text || '').join('').replace(/^```(?:json)?|```$/g, '').trim();
  return JSON.parse(text);
}

// Chạy để xem tên model hợp lệ cho GEMINI_MODEL (nếu gặp lỗi "model not found")
function listGeminiModels() {
  const key = props().getProperty('GEMINI_API_KEY');
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=100', { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
  const body = JSON.parse(res.getContentText());
  if (!body.models) return Logger.log(res.getContentText());
  body.models.filter(m => (m.supportedGenerationMethods || []).includes('generateContent')).forEach(m => Logger.log(m.name.replace('models/', '')));
}

function ocrClaude(image, categories) {
  const key = props().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('Chưa cài ANTHROPIC_API_KEY trong Script Properties');
  const [header, base64] = image.split(',');
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
              category: { type: 'string', enum: categories },
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
  return JSON.parse(body.content.filter(b => b.type === 'text').map(b => b.text).join(''));
}

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
  chi: ['🍜 Ăn uống', '☕ Cà phê', '🛒 Đi chợ', '🛵 Đi lại', '🅿️ Gửi xe', '⚡ Điện', '💧 Nước', '🏢 Phí quản lý', '🌐 Internet', '📱 Điện thoại', '🔥 Gas', '🏠 Tiền nhà', '🪑 Đồ dùng nhà cửa', '🧺 Giặt ủi', '🧴 Skincare', '💇 Làm đẹp', '🩹 Thuốc', '💊 Sức khỏe', '🛍️ Mua sắm', '🎮 Giải trí', '📚 Học tập', '🐾 Thú cưng', '🎁 Quà tặng', '🧾 Phát sinh', '📦 Khác'],
  thu: ['💰 Lương', '🎁 Thưởng', '💸 Được cho', '📈 Đầu tư', '📦 Khác']
};
const TG_KEYWORDS = {
  '🍜 Ăn uống': ['ăn', 'cơm', 'phở', 'bún', 'bánh', 'lẩu', 'nhậu', 'bia', 'mì', 'cháo', 'xôi', 'đồ ăn', 'bữa', 'pizza', 'nướng'],
  '☕ Cà phê': ['cà phê', 'cafe', 'cafê', 'trà sữa', 'trà', 'nước ngọt', 'nước suối'],
  '🛒 Đi chợ': ['đi chợ', 'chợ', 'siêu thị', 'bách hóa', 'bách hóa xanh', 'winmart', 'coopmart', 'tạp hóa', 'rau', 'thịt', 'cá', 'trứng', 'trái cây', 'hoa quả'],
  '🛵 Đi lại': ['xăng', 'đổ xăng', 'grab', 'xanh sm', 'taxi', 'xe', 'vé', 'xe buýt', 'buýt', 'rửa xe', 'sửa xe', 'tàu', 'máy bay'],
  '🅿️ Gửi xe': ['gửi xe', 'giữ xe', 'đậu xe', 'bãi xe', 'phí gửi xe', 'vé xe tháng'],
  '⚡ Điện': ['điện', 'tiền điện', 'hóa đơn điện', 'evn'],
  '💧 Nước': ['nước', 'tiền nước', 'nước sinh hoạt', 'hóa đơn nước', 'nước máy'],
  '🏢 Phí quản lý': ['phí quản lý', 'quản lý', 'ban quản lý', 'phí chung cư', 'chung cư', 'phí dịch vụ', 'phí vệ sinh'],
  '🌐 Internet': ['internet', 'wifi', 'mạng', 'cáp quang', 'fpt'],
  '📱 Điện thoại': ['điện thoại', 'sim', '4g', '5g', 'data', 'nạp thẻ', 'nạp tiền', 'cước', 'viettel', 'mobifone', 'vinaphone'],
  '🔥 Gas': ['gas', 'bình gas', 'bình ga', 'tiền gas'],
  '🏠 Tiền nhà': ['tiền nhà', 'tiền thuê', 'thuê nhà', 'thuê phòng', 'tiền phòng', 'tiền trọ', 'trọ'],
  '🪑 Đồ dùng nhà cửa': ['đồ dùng', 'gia dụng', 'đồ gia dụng', 'nhà cửa', 'bột giặt', 'nước giặt', 'nước rửa chén', 'giấy vệ sinh', 'nước lau sàn', 'túi rác', 'xà phòng', 'chén', 'bát', 'nồi', 'chảo', 'khăn', 'gối', 'chăn', 'bóng đèn', 'sửa nhà'],
  '🧺 Giặt ủi': ['giặt', 'giặt ủi', 'giặt đồ', 'ủi', 'tiệm giặt'],
  '🧴 Skincare': ['skincare', 'kem dưỡng', 'sữa rửa mặt', 'serum', 'toner', 'kem chống nắng', 'mặt nạ', 'dưỡng da', 'tẩy trang', 'mỹ phẩm', 'son'],
  '💇 Làm đẹp': ['cắt tóc', 'uốn tóc', 'nhuộm tóc', 'gội đầu', 'làm tóc', 'làm nail', 'nail', 'spa', 'massage', 'salon', 'barber'],
  '🩹 Thuốc': ['thuốc', 'vitamin', 'khẩu trang', 'nhà thuốc', 'pharmacity', 'long châu', 'băng gạc'],
  '💊 Sức khỏe': ['khám', 'bệnh viện', 'phòng khám', 'nha khoa', 'răng', 'xét nghiệm', 'bảo hiểm', 'gym', 'yoga', 'thể thao'],
  '🛍️ Mua sắm': ['mua', 'áo', 'quần', 'giày', 'váy', 'dép', 'túi', 'balo', 'phụ kiện', 'shopee', 'lazada', 'tiki'],
  '🎮 Giải trí': ['phim', 'xem phim', 'cgv', 'game', 'karaoke', 'du lịch', 'đi chơi', 'netflix', 'spotify', 'youtube', 'concert'],
  '📚 Học tập': ['học', 'học phí', 'khóa học', 'sách', 'vở', 'bút', 'văn phòng phẩm', 'tiếng anh', 'udemy'],
  '🐾 Thú cưng': ['thú cưng', 'thú y', 'chó', 'mèo', 'cát mèo', 'pet', 'thức ăn cho'],
  '🎁 Quà tặng': ['quà', 'quà tặng', 'tặng', 'biếu', 'lì xì', 'mừng cưới', 'đám cưới', 'phúng điếu', 'sinh nhật', 'từ thiện'],
  '🧾 Phát sinh': ['phát sinh', 'đột xuất', 'sửa chữa', 'sửa', 'hỏng', 'khẩn cấp', 'phạt', 'lệ phí'],
  '💰 Lương': ['lương'],
  '🎁 Thưởng': ['thưởng']
};
const TG_HELP = '<b>💸 Bot chi tiêu</b>\n' +
  'Gõ thẳng để ghi: <code>cà phê 35k</code>, <code>ăn trưa 45</code>, <code>lương 12tr</code>\n\n' +
  '/today — hôm nay\n/week — 7 ngày qua\n/month — tháng này\n/top — 5 khoản chi lớn nhất tháng\n' +
  '/budget 5tr — đặt ngân sách tháng\n/undo — xoá khoản vừa ghi\n\n' +
  '📸 Gửi ảnh hoá đơn (có thể kèm chú thích như <code>45k</code>) để bot tự đọc và ghi.';

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
  // WEBHOOK_URL = địa chỉ worker chuyển tiếp (xem worker.js); không có thì gọi thẳng Web App
  const base = props().getProperty('WEBHOOK_URL') || props().getProperty('WEBAPP_URL');
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

// Chạy để kiểm tra webhook: xem last_error_message và pending_update_count trong Execution log
function webhookInfo() {
  const token = props().getProperty('TELEGRAM_BOT_TOKEN');
  Logger.log(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getWebhookInfo', { muteHttpExceptions: true }).getContentText());
}

function handleTelegram(e, update) {
  try {
    const secret = props().getProperty('TELEGRAM_SECRET');
    if (!secret || !e.parameter || e.parameter.tg !== secret) return json({ ok: false });
    const msg = update.message;
    if (!msg || !(msg.text || msg.photo)) return json({ ok: true });
    // Telegram gửi lại nếu phản hồi chậm. Đã xong thì bỏ qua; đang xử lý dở quá 60s (lỗi/treo) thì xử lý lại
    const cache = CacheService.getScriptCache();
    const ukey = 'tg_' + update.update_id;
    const state = cache.get(ukey);
    if (state === 'done' || (state && Date.now() - Number(state) < 60000)) return json({ ok: true });
    cache.put(ukey, String(Date.now()), 600);

    const chatId = String(msg.chat.id);
    if (!tgChatIds().includes(chatId)) {
      tgSend(chatId, '🔒 Chat này chưa được cấp quyền.\nChat ID: <code>' + chatId + '</code>\nThêm ID này vào TELEGRAM_CHAT_IDS trong Script Properties.');
      return json({ ok: true });
    }
    let reply;
    const from = Object.assign({}, msg.from, { uid: update.update_id });
    if (msg.photo) tgSend(chatId, '⏳ Đang đọc ảnh, đợi mình một chút...');
    try { reply = msg.photo ? addFromPhoto(msg, from) : tgReply(msg.text.trim(), from); }
    catch (err) { reply = '⚠️ Lỗi: ' + esc(err.message || err); }
    tgSend(chatId, reply);
    cache.put(ukey, 'done', 600);
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
    id: from.uid ? 'tg' + from.uid : Utilities.getUuid(), // id cố định theo update: gửi lại không bị ghi trùng
    time: new Date().toISOString(), type: p.type, amount: p.amount,
    category: p.category, note: p.note, who: from.first_name || 'Telegram'
  });
  const today = tgData().items.filter(i => i.day === dayKey(new Date()) && isChi(i));
  return (p.type === 'thu' ? '💰 Đã ghi thu ' : '✅ Đã ghi ') + '<b>' + vnd(p.amount) + '</b> — ' + esc(p.category) + (p.note ? ' (' + esc(p.note) + ')' : '') +
    '\nHôm nay đã chi: <b>' + vnd(sumOf(today)) + '</b>';
}

// Ảnh hoá đơn: Claude đọc số tiền/danh mục/ghi chú, chú thích của người gửi (nếu có) được ưu tiên
function addFromPhoto(msg, from) {
  const t0 = Date.now();
  const lap = step => console.log('photo ' + step + ': ' + (Date.now() - t0) + 'ms');
  const token = props().getProperty('TELEGRAM_BOT_TOKEN');
  const photo = msg.photo[msg.photo.length - 1]; // bản lớn nhất
  const meta = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getFile?file_id=' + photo.file_id, { muteHttpExceptions: true }).getContentText());
  if (!meta.ok) throw new Error('Không tải được ảnh từ Telegram');
  const bytes = UrlFetchApp.fetch('https://api.telegram.org/file/bot' + token + '/' + meta.result.file_path).getBlob().getBytes();
  const image = 'data:image/jpeg;base64,' + Utilities.base64Encode(bytes);

  lap('downloaded');
  const ai = ACTIONS.ocr({ image, categories: TG_CATS.chi }).result;
  lap('ai done');
  const cap = msg.caption ? parseExpense(msg.caption) : null;
  const amount = (cap && cap.amount) || ai.amount;
  if (!amount) return 'Mình chưa đọc được số tiền trong ảnh. Gửi lại ảnh kèm chú thích, ví dụ: <code>45k</code>';
  const category = cap && TG_CATS.chi.includes(cap.category) && cap.category !== '📦 Khác' ? cap.category : ai.category;
  const note = (cap && cap.note) || ai.note;

  ACTIONS.add({
    id: 'tg' + from.uid, time: new Date().toISOString(), type: 'chi', amount, category, note,
    who: from.first_name || 'Telegram', image
  });
  lap('saved');
  const today = tgData().items.filter(i => i.day === dayKey(new Date()) && isChi(i));
  return '📸 Đã ghi<b>' + vnd(amount) + '</b> — ' + esc(category) + (note ? ' (' + esc(note) + ')' : '') +
    '\nHôm nay đã chi: <b>' + vnd(sumOf(today)) + '</b>\nSai số? Gõ /undo rồi gửi lại kèm chú thích.';
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
  // Từ khóa dài nhất thắng: "gửi xe" (Gửi xe) thắng "xe" (Đi lại), "bột giặt" thắng "giặt"
  let best = 0;
  for (const c of Object.keys(TG_KEYWORDS)) {
    for (const w of TG_KEYWORDS[c]) {
      if (w.length > best && padded.includes(' ' + w + ' ')) { category = c; best = w.length; }
    }
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
