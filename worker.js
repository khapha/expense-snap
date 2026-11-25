// Cloudflare Worker: nhận webhook Telegram, trả 200 ngay, rồi chuyển tin sang Apps Script.
// Apps Script trả 302 nên Telegram coi là lỗi và gửi lại; worker này tránh việc đó.
// Biến môi trường cần đặt: APPS_SCRIPT_URL = URL Web App (.../exec)
export default {
  async fetch(request, env, ctx) {
    if (request.method !== 'POST') return new Response('ok');
    const url = new URL(request.url);
    const body = await request.text();
    // Giữ nguyên query (?tg=...) để Apps Script kiểm tra khoá bí mật
    ctx.waitUntil(fetch(env.APPS_SCRIPT_URL + url.search, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body,
      redirect: 'follow'
    }));
    return new Response('ok', { status: 200 });
  }
};
