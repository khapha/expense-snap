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

Dữ liệu nằm ở tab **GiaoDich** của Sheet "Snap Chi Tieu". Tab `ChiTieu` cũ (nếu có) giữ nguyên, không hiện trong app.
