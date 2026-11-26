// Cloudflare Worker: nhận webhook Telegram, trả 200 ngay, rồi chuyển tin sang Apps Script.
// Apps Script trả 302 nên Telegram coi là lỗi và gửi lại; worker này tránh việc đó.
//
// Biến môi trường (Settings → Variables and Secrets):
//   APPS_SCRIPT_URL     (bắt buộc) URL Web App (.../exec) của deployment đang chạy code mới
//   TELEGRAM_SECRET     (nên đặt)  khoá bí mật của webhook = Script Property TELEGRAM_SECRET.
//                       Có đặt thì worker từ chối mọi yêu cầu không mang đúng ?tg=<khoá>.
//   TELEGRAM_BOT_TOKEN  (tuỳ chọn, loại Secret) + TELEGRAM_CHAT_IDS (tuỳ chọn, ngăn cách bằng dấu phẩy):
//                       có đủ cả hai thì khi chuyển tiếp thất bại, worker nhắn lại cho chat đó
//                       "chưa xử lý được, gửi lại nhé" thay vì để tin biến mất.

const TRIES = 3;                 // số lần chuyển tiếp tối đa khi gặp lỗi chắc chắn
const RETRY_DELAYS = [0, 800, 2000];
const ATTEMPT_TIMEOUT_MS = 20000; // quá hạn thì coi là "chưa rõ" (ví dụ AI đọc ảnh chậm): không thử lại, không báo lỗi

export default {
  async fetch(request, env, ctx) {
    if (request.method !== 'POST') return new Response('ok');
    const url = new URL(request.url);

    // Chỉ Telegram (biết khoá trong URL webhook) mới được đi tiếp
    if (env.TELEGRAM_SECRET && !safeEqual(url.searchParams.get('tg') || '', env.TELEGRAM_SECRET)) {
      return new Response('forbidden', { status: 403 });
    }

    const body = await request.text();
    let update;
    try { update = JSON.parse(body); } catch (e) { return new Response('bad request', { status: 400 }); }

    const chatId = update && update.message && update.message.chat ? String(update.message.chat.id) : '';
    // Giữ nguyên query (?tg=...) để Apps Script kiểm tra khoá bí mật
    ctx.waitUntil(forward(env, url.search, body, chatId));
    return new Response('ok', { status: 200 });
  }
};

// Chuyển tiếp sang Apps Script. Thử lại khi gặp lỗi chắc chắn (an toàn vì Apps Script bỏ qua update_id đã xử lý).
async function forward(env, search, body, chatId) {
  let reason = '';
  for (let i = 0; i < TRIES; i++) {
    if (RETRY_DELAYS[i]) await new Promise(r => setTimeout(r, RETRY_DELAYS[i]));
    try {
      const res = await fetch(env.APPS_SCRIPT_URL + search, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body,
        redirect: 'follow',
        signal: AbortSignal.timeout(ATTEMPT_TIMEOUT_MS)
      });
      const verdict = judge(res.status, await res.text());
      if (verdict.ok) return;
      reason = verdict.reason;
      if (!verdict.retry) break;
    } catch (e) {
      if (e && (e.name === 'TimeoutError' || e.name === 'AbortError')) {
        console.warn('Apps Script chưa trả lời kịp, bỏ qua (có thể vẫn đang xử lý)');
        return;
      }
      reason = 'network: ' + String(e && e.message || e);
    }
  }
  console.error('Chuyển tiếp thất bại: ' + reason);
  await notify(env, chatId);
}

// ok: Apps Script nhận được; retry: lỗi tạm thời đáng thử lại
function judge(status, text) {
  if (status === 429 || status >= 500) return { ok: false, retry: true, reason: 'http ' + status };
  if (status >= 400) return { ok: false, retry: false, reason: 'http ' + status + ' (sai URL hoặc quyền deployment?)' };
  let json;
  try { json = JSON.parse(text); } catch (e) {
    return { ok: false, retry: true, reason: 'phản hồi không phải JSON (trang lỗi của Apps Script?)' };
  }
  if (json && json.ok === false) return { ok: false, retry: false, reason: 'Apps Script từ chối (sai khoá bí mật?)' };
  return { ok: true };
}

// Báo cho người dùng khi tin không được xử lý. Chỉ nhắn các chat nằm trong TELEGRAM_CHAT_IDS.
async function notify(env, chatId) {
  if (!env.TELEGRAM_BOT_TOKEN || !env.TELEGRAM_CHAT_IDS || !chatId) return;
  const allowed = String(env.TELEGRAM_CHAT_IDS).split(',').map(s => s.trim()).filter(Boolean);
  if (!allowed.includes(chatId)) return;
  try {
    await fetch('https://api.telegram.org/bot' + env.TELEGRAM_BOT_TOKEN + '/sendMessage', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ chat_id: chatId, text: '⚠️ Mình chưa xử lý được tin vừa rồi (lỗi tạm thời). Bạn gửi lại giúp mình nhé.' })
    });
  } catch (e) {
    console.error('Không gửi được thông báo lỗi: ' + String(e && e.message || e));
  }
}

// So sánh chuỗi không phụ thuộc thời gian dừng sớm
function safeEqual(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}
