# Snap Chi Tiêu

Chụp ảnh khoản chi kiểu Locket → nhập số tiền, danh mục → tự lưu vào Google Sheet (ảnh lưu ở Google Drive).

App: https://khapha.github.io/expense-snap/

## Tính năng
- 📸 Chụp ảnh, hoặc ghi **không ảnh**; ghi cả **khoản thu**
- 🧾 **Đọc bill bằng AI** (Claude) — tự điền số tiền, danh mục, ghi chú
- 🎤 **Nhập bằng giọng nói**: "cà phê 35 nghìn", "đổ xăng 50k", "lương 12 triệu"
- 🖼️ **Lịch sử** đọc từ Sheet (đổi máy không mất), chạm để **sửa / xoá**
- 📊 **Thống kê** theo tháng: thu/chi/còn lại, biểu đồ danh mục, theo tuần, so với tháng trước
- 🎯 **Ngân sách** tháng có thanh tiến độ, cảnh báo khi gần vượt
- 📶 **Offline**: không có mạng thì xếp hàng, có mạng tự gửi
- 👫 **Dùng chung**: nhiều người cùng URL + mã bí mật ghi vào 1 Sheet, thấy ảnh của nhau
- 📧 **Email tóm tắt** chi tiêu lúc 21h mỗi ngày
- 🔒 **Mã bí mật**: chỉ ai có mã mới ghi/đọc được
- 📱 Cài như app (PWA): Chia sẻ → Thêm vào màn hình chính

> Widget màn hình chính như Locket thật cần app native iOS/Android — web app không làm được.

## Cài đặt backend (Google Apps Script)
1. Vào https://script.google.com → mở project (hoặc Sheet → Extensions → Apps Script), dán toàn bộ `Code.gs`, `Ctrl+S`.
2. Ở thanh công cụ chọn hàm **`setup`** → **Run** → cấp quyền. Mở **Execution log**, copy **mã bí mật**.
   `setup` cũng tạo trigger gửi email tóm tắt lúc 21h.
3. **Deploy → Manage deployments → ✏️ → Version: New version → Deploy** (giữ nguyên URL).
   Lần đầu thì **Deploy → New deployment → Web app**, Execute as **Me**, Who has access **Anyone**.
4. Trên app: ⚙️ Cài đặt → nhập tên, mã bí mật, ngân sách.

### Bật đọc bill bằng AI (tuỳ chọn)
1. Lấy API key tại https://console.anthropic.com (có tính phí theo lượt dùng).
2. Apps Script → ⚙️ **Project Settings → Script Properties → Add property**:
   `ANTHROPIC_API_KEY` = key của bạn.

### Script Properties khác (tuỳ chọn)
- `SUMMARY_EMAIL`: email nhận tóm tắt (mặc định là email của bạn)
- `TOKEN`: mã bí mật (do `setup` tạo, có thể đổi)

## Bot Telegram

Bot dùng chung Sheet với app: ghi chi tiêu bằng cách nhắn tin, xem thống kê, và nhận tổng kết lúc 21h mỗi ngày.

**Lệnh:** `/today` `/week` `/month` `/top` `/budget 5tr` `/undo` `/help`. Nhắn thẳng `cà phê 35k`, `ăn trưa 45`, `lương 12tr` để ghi.

**Cài đặt:**
1. Trên Telegram nhắn **@BotFather** → `/newbot` → đặt tên → copy **bot token**.
2. Apps Script → **Project Settings → Script Properties**, thêm:
   - `TELEGRAM_BOT_TOKEN` = token từ BotFather
   - `WEBAPP_URL` = URL Web App (`.../exec`)
3. Project Settings → **Time zone** chọn `(GMT+07:00) Asia/Ho_Chi_Minh` (để "hôm nay" tính đúng giờ Việt Nam).
4. Dán `Code.gs` mới, **Deploy → Manage deployments → ✏️ → New version → Deploy**.
5. Chọn hàm **`setupTelegram`** → **Run** (đăng ký webhook). Xem Execution log phải có `"ok":true`.
6. Nhắn `/start` cho bot. Bot trả về **Chat ID** của bạn — thêm Script Property `TELEGRAM_CHAT_IDS` = ID đó (nhiều người thì ngăn cách bằng dấu phẩy). Nhắn `/help` lại để dùng.

Chỉ các Chat ID trong `TELEGRAM_CHAT_IDS` mới dùng được bot; người khác chỉ nhận được ID của họ.

### Gửi ảnh hoá đơn cho bot
Gửi ảnh bill (có thể kèm chú thích như `45k` hoặc `cà phê`) → Claude đọc số tiền, danh mục, ghi chú rồi ghi vào Sheet cùng ảnh. Chú thích của bạn được ưu tiên hơn kết quả AI. Cần Script Property `ANTHROPIC_API_KEY` (xem phần đọc bill ở trên). Đọc sai thì `/undo` rồi gửi lại kèm chú thích.

### Bot lúc có lúc không phản hồi (lỗi 302)
Apps Script trả mã 302 nên Telegram coi là lỗi, gửi lại và trì hoãn tin. Chạy hàm `webhookInfo`: nếu `last_error_message` là `302 Found` thì dùng bộ chuyển tiếp `worker.js` trên Cloudflare Workers (miễn phí):
1. Đăng ký https://dash.cloudflare.com → **Workers & Pages → Create → Create Worker** → đặt tên (vd `expense-bot`) → **Deploy** → **Edit code** → dán `worker.js` → **Deploy**.
2. Worker → **Settings → Variables and Secrets** → thêm `APPS_SCRIPT_URL` = URL Web App (`.../exec`).
3. Copy địa chỉ worker (`https://expense-bot.<tên>.workers.dev`), thêm Script Property `WEBHOOK_URL` = địa chỉ đó.
4. Dán `Code.gs` mới, deploy New version, chạy lại `setupTelegram`, rồi `webhookInfo` để kiểm tra `url` đã là địa chỉ worker.

Dữ liệu nằm ở tab **GiaoDich** của Sheet "Snap Chi Tieu". Tab `ChiTieu` cũ (nếu có) giữ nguyên, không hiện trong app.
