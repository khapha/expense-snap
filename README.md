# Snap Chi Tiêu

Chụp ảnh khoản chi kiểu Locket → nhập số tiền, danh mục → tự lưu vào Google Sheet (ảnh lưu ở Google Drive).

## 1. Tạo backend Google Sheet
1. Tạo một Google Sheet mới.
2. **Extensions → Apps Script**, xoá code mẫu, dán nội dung `Code.gs`, bấm Save.
3. **Deploy → New deployment** → loại **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Cấp quyền (Sheets + Drive), copy **Web app URL** (`.../exec`).

> Khi sửa `Code.gs`, phải **Deploy → Manage deployments → Edit → New version** thì mới có hiệu lực.

## 2. Chạy web app
Camera chỉ hoạt động trên HTTPS hoặc localhost.
- Thử trên máy: `npx serve .` rồi mở `http://localhost:3000`
- Dùng trên điện thoại: đẩy thư mục lên **GitHub Pages / Netlify / Vercel** (miễn phí, có HTTPS), mở link trên điện thoại → "Thêm vào màn hình chính".

Lần đầu mở, dán Web app URL vào ô cài đặt (nút ⚙️).

## Cách dùng
Bấm nút chụp → nhập số tiền → chọn danh mục → bấm chụp lần nữa (hoặc ➤) để gửi. ✕ để chụp lại, 🔄 đổi camera.

Sheet `ChiTieu` gồm: Thời gian · Số tiền · Danh mục · Ghi chú · Ảnh (hiển thị trực tiếp) · Link ảnh. Có thể tạo Pivot Table / biểu đồ theo tháng, danh mục ngay trong Sheet.

**Lưu ý:** URL Web App ai có cũng ghi được vào Sheet — đừng chia sẻ công khai.
