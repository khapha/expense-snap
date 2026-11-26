# Sổ tay dự án — Snap Chi Tiêu

Nhật ký tóm tắt toàn bộ quá trình làm app chi tiêu kiểu Locket: những gì đã làm, lỗi gặp phải, cách sửa, và việc còn dở.
Đây là bản **tóm tắt theo thứ tự thời gian**, không phải bản ghi nguyên văn cuộc trò chuyện. Cố ý **không ghi** token, khoá bí mật, Chat ID, URL deployment, email, số tài khoản vì repo này public.

- App: https://khapha.github.io/expense-snap/
- Repo: https://github.com/khapha/expense-snap

## 1. Mục tiêu
Web app chụp ảnh khoản chi hằng ngày (giống Locket), lưu vào Google Sheet; sau đó mở rộng: thống kê, ngân sách, bot Telegram, AI đọc hoá đơn. Ưu tiên chạy **miễn phí**.

## 2. Kiến trúc
```
Điện thoại (PWA trên GitHub Pages) ──POST──▶ Google Apps Script (Web App) ──▶ Google Sheet + Drive (ảnh)
Telegram ──▶ Cloudflare Worker (trả 200 ngay) ──▶ Apps Script (doPost, nhận diện update_id) ──▶ cùng Sheet
Apps Script ──▶ Gemini (miễn phí) hoặc Claude (trả phí) để đọc hoá đơn
Trigger 21h mỗi ngày ──▶ email + tin Telegram tổng kết
```
File chính: `index.html` (giao diện), `Code.gs` (backend), `worker.js` (bộ chuyển tiếp Cloudflare), `sw.js` + `manifest.json` (PWA).

## 3. Dòng thời gian

### 2026-11-02 — Bản đầu và đưa lên GitHub
- Viết web app: camera vuông bo góc, nhập số tiền, chọn danh mục, gửi lên Apps Script → ghi dòng vào Sheet, ảnh lưu Drive.
- Cài GitHub CLI bằng `winget`, đăng nhập `gh auth login`, tạo repo public, bật GitHub Pages. Mạng hay ngắt kết nối (Connection reset) nên phải thử lại vài lần.
- **Lỗi:** Apps Script báo "Sorry, unable to open the file" → do trình duyệt đăng nhập **nhiều tài khoản Google**. Sửa: dùng cửa sổ ẩn danh, chỉ đăng nhập đúng tài khoản chủ.
- Hardcode URL Web App làm mặc định trong `index.html`.
- **Lỗi:** `Cannot read properties of null (reading 'getSheetByName')` → project Apps Script tạo ở script.google.com (không gắn với Sheet). Sửa: `getSpreadsheet()` tự tạo Sheet "Snap Chi Tieu" và nhớ ID trong Script Properties.

### 2026-11-02 — Nâng cấp lớn
Viết lại backend theo kiểu `action` (`add`, `list`, `update`, `delete`, `setBudget`, `ocr`) và giao diện 3 tab.
- Lịch sử đọc từ Sheet (đổi máy không mất), sửa/xoá khoản chi.
- Thống kê tháng: thu/chi/còn lại, biểu đồ theo danh mục, theo tuần, so với tháng trước.
- Ngân sách tháng + cảnh báo gần vượt.
- Hàng đợi offline (tự gửi khi có mạng), khoản thu, ghi không ảnh, nhập bằng giọng nói tiếng Việt.
- Đọc hoá đơn bằng AI, dùng chung với nhiều người (tên người ghi), email tóm tắt 21h, mã bí mật (`TOKEN`), PWA (icon, offline).
- Lỗi nhỏ: "đổ xăng" bị nhận thành "Ăn uống" vì chuỗi `xăng` chứa `ăn` → sửa bằng **khớp nguyên từ**.
- Không làm được: widget màn hình chính như Locket (cần app native).

### 2026-11-25 — Bot Telegram
- Thêm bot ngay trong Apps Script: ghi bằng tin nhắn (`cà phê 35k`), lệnh `/today /week /month /top /budget /undo /help`, tổng kết 21h.
- Cấu hình: chỉ các Chat ID trong `TELEGRAM_CHAT_IDS` được dùng; webhook có khoá bí mật trong query `?tg=`.
- Hàm test trong trình duyệt (máy không có Node/Python): tách số tiền ("35k", "12tr", "1.500.000", "85.000đ") và danh mục.
- **Lỗi nhập nhầm:** bấm **Debug** thay vì **Run** → hàm dừng ở breakpoint. Sửa: Stop, xoá breakpoint, bấm Run.
- **Lỗi cấp quyền:** popup cấp quyền báo "Sorry, unable to open the file" (nhiều tài khoản) → dùng cửa sổ ẩn danh.
- **Lỗi bot lúc có lúc không:** `getWebhookInfo` báo `Wrong response from the webhook: 302 Found`. Apps Script trả 302, Telegram coi là lỗi, gửi lại và trì hoãn. Sửa: **Cloudflare Worker** (`worker.js`) nhận webhook, trả 200 ngay, chuyển tiếp sang Apps Script. Chọn "Start with Hello World!" khi tạo worker, rồi dán code.
- **Lỗi cấu hình:** `APPS_SCRIPT_URL` ban đầu đặt nhầm vào Script Properties của Google (phải đặt ở Cloudflare). Sau đó phát hiện nó trỏ tới **deployment cũ** chưa có code Telegram.
- **Lỗi DNS:** `Bad webhook: Failed to resolve host` ngay sau khi tạo worker → đợi vài phút rồi chạy lại `setupTelegram` là được.
- **Sự cố dữ liệu:** deployment cũ coi tin Telegram là lệnh "thêm khoản chi" rỗng → 14 dòng rác (thời gian 1970, số tiền 0) trong Sheet. Một dòng do **lệnh test của chính mình** tạo ra (gửi POST thử vào endpoint chưa có mã bí mật). Đã xoá đúng các dòng rác (sau khi được cho phép), giữ lại khoản thật. Về sau chỉ test bằng lệnh chỉ đọc.
- Gom lại: mỗi lần sửa code phải **New version cho đúng deployment mà worker đang dùng**.

### 2026-11-25 — Gửi ảnh hoá đơn cho bot (AI)
- Thêm nhánh ảnh: bot tải ảnh từ Telegram → AI đọc số tiền/danh mục/ghi chú → ghi vào Sheet + lưu ảnh Drive. Chú thích của người gửi được ưu tiên. Bot nhắn "⏳ Đang đọc ảnh…" ngay khi nhận.
- Chọn **Gemini gói miễn phí** làm bộ đọc mặc định, Claude là lựa chọn trả phí (`AI_PROVIDER`).
- **Lỗi 404:** `gemini-2.5-flash-lite` không còn cho tài khoản mới → đổi mặc định sang `gemini-3.5-flash-lite` (có thể đặt qua `GEMINI_MODEL` không cần deploy lại).
- **Lỗi 503** "high demand" → tự thử lại tối đa 3 lần; vẫn quá tải thì báo gọn "hãy gửi lại sau".
- Tài liệu Gemini đọc được mô tả một API mới (`interactions`) với tên model lạ, nên code dùng `generateContent` có phương án dự phòng (bỏ lược đồ JSON nếu bị từ chối 400).

## 4. Cấu hình cần có (chỉ tên, không ghi giá trị)

**Apps Script → Script Properties**
| Tên | Dùng để |
|---|---|
| `TOKEN` | mã bí mật của app (do hàm `setup` tạo) |
| `BUDGET` | ngân sách tháng (do app/bot đặt) |
| `SPREADSHEET_ID` | Sheet tự tạo nếu project không gắn với Sheet |
| `TELEGRAM_BOT_TOKEN` | token từ BotFather |
| `TELEGRAM_SECRET` | khoá bí mật trong URL webhook (do `setupTelegram` tạo) |
| `TELEGRAM_CHAT_IDS` | danh sách Chat ID được phép, ngăn cách bằng dấu phẩy |
| `WEBAPP_URL` | URL Web App `/exec` |
| `WEBHOOK_URL` | địa chỉ worker Cloudflare |
| `GEMINI_API_KEY`, `GEMINI_MODEL` | đọc hoá đơn bằng Gemini |
| `ANTHROPIC_API_KEY`, `AI_PROVIDER` | tuỳ chọn, dùng Claude (trả phí) |
| `SUMMARY_EMAIL` | email nhận tóm tắt (mặc định là email chủ script) |

**Cloudflare Worker → Variables**: `APPS_SCRIPT_URL` = URL Web App của deployment đang chạy code mới; `TELEGRAM_SECRET` = cùng giá trị Script Property `TELEGRAM_SECRET` (worker trả 403 nếu `?tg=` sai); tuỳ chọn `TELEGRAM_BOT_TOKEN` (Secret) + `TELEGRAM_CHAT_IDS` để worker nhắn báo khi chuyển tiếp thất bại. Đổi `TELEGRAM_SECRET` ở Google thì phải đổi cả ở đây rồi chạy lại `setupTelegram`.

**Hàm chạy tay trong Apps Script**: `setup` (tạo token + trigger 21h), `setupTelegram` (đăng ký webhook + menu lệnh), `webhookInfo` (chẩn đoán webhook), `listGeminiModels` (liệt kê model hợp lệ).

## 5. Quy trình cập nhật code (nhớ kỹ)
1. Dán `Code.gs` mới → `Ctrl+S`.
2. **Deploy → Manage deployments → ✏️ → New version → Deploy**, **chọn đúng deployment worker đang trỏ tới** (không bấm "New deployment" vì URL sẽ đổi).
3. Bấm **Run**, không bấm **Debug**.
4. Việc phía Google nên làm trong cửa sổ ẩn danh khi có nhiều tài khoản.

## 6. Bài học / bảng tra lỗi
| Triệu chứng | Nguyên nhân | Cách xử lý |
|---|---|---|
| "Sorry, unable to open the file" | nhiều tài khoản Google | cửa sổ ẩn danh, 1 tài khoản |
| `getSheetByName` of null | script không gắn Sheet | đã xử lý bằng `getSpreadsheet()` |
| Bot im lặng, `last_error_message` = 302 | Apps Script trả 302 | dùng Cloudflare Worker |
| Bot im lặng sau khi sửa code | worker trỏ deployment cũ | đổi `APPS_SCRIPT_URL`, hoặc New version đúng deployment |
| Bot im lặng, `webhookInfo` báo 403 | `TELEGRAM_SECRET` ở worker khác Script Property | sửa cho khớp rồi chạy lại `setupTelegram` |
| `Bad webhook: Failed to resolve host` | DNS `workers.dev` chưa lan | đợi 5–10 phút, chạy lại `setupTelegram` |
| Sheet có dòng 1970, số tiền 0 | tin Telegram vào bản cũ/endpoint không token | xoá dòng, trỏ đúng deployment |
| Gemini 404 model | model bị ngừng cho tài khoản mới | đặt `GEMINI_MODEL`, chạy `listGeminiModels` |
| Gemini 503 | quá tải tạm thời | tự thử lại; gửi lại sau |
| Gemini 429 | hết hạn mức free | đợi hoặc đổi model |

## 7. Chi phí & riêng tư
- Miễn phí: GitHub Pages, Apps Script/Sheets/Drive, Telegram, Cloudflare Workers (gói free), Gemini (gói free, có giới hạn).
- Claude API: trả phí theo lượt dùng.
- Gói free của Gemini: Google có thể dùng nội dung gửi lên để cải thiện sản phẩm, người thật có thể đọc → **không gửi ảnh có thông tin nhạy cảm** (số thẻ, số tài khoản, ảnh chuyển khoản ngân hàng). Ảnh lưu Drive ở chế độ "ai có link đều xem được".

## 8. Việc còn dở
- [ ] Chạy hàm `setup` để tạo mã bí mật `TOKEN` và nhập vào ⚙️ của app. **Chưa làm → ai có URL Web App cũ (đang nằm trong `index.html` public) đều đọc/ghi được Sheet.**
- [ ] Gom về **một URL Web App** duy nhất: đặt URL deployment mới vào ⚙️ của app, rồi bỏ URL cứng trong `index.html`.
- [ ] Đặt lại ngân sách (đang đọc ra 0): `/budget 5tr`.
- [ ] Xác nhận đọc ảnh bằng Gemini chạy thành công với hoá đơn thường (chưa thấy kết quả thành công cuối cùng).
- [ ] Đổi bot token qua BotFather (`/revoke`) vì token từng hiện trong ảnh chụp màn hình; xoá `TELEGRAM_SECRET` rồi chạy lại `setupTelegram` để đổi khoá webhook.
- [ ] Cân nhắc: hỗ trợ ảnh gửi dạng tệp (document) cho bot; widget native nếu muốn giống Locket hơn.
