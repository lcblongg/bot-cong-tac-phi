# BOT Công tác phí — bản web

Tự tạo file Excel công tác phí hàng tháng từ hóa đơn xăng. Thiết kế bởi lcblongg.

## Dùng
1. Điền **Thông tin kỳ** (tháng, kỳ từ/đến ngày, địa điểm, tạm ứng…).
2. Upload **hóa đơn** (.pdf + .eml của kỳ).
3. **Số KM**: upload file Numbers (bảng *Route Plan*) — hoặc gõ tay tổng KM.
4. Upload **file Excel mẫu** = file công tác phí tháng trước của bạn.
5. Bấm **Tạo file công tác phí** → tải file `.xlsx` + file zip hóa đơn đã đổi tên.
6. Mở file Excel bằng Excel/Numbers 1 lần cho công thức tự tính lại.

## BOT làm gì
- Đọc từng PDF: số hóa đơn, số tiền ("Tổng cộng tiền thanh toán"), VAT, mã tra cứu,
  website, loại hàng (Xăng→HĐX / Dầu→HĐD), địa chỉ người mua.
- Đổi tên file → `HĐX_<số>` / `HĐD_<số>`; địa chỉ sai → gắn nhãn `[SAI ĐỊA CHỈ]`.
- **KM**: cộng cột "Km di chuyển" bảng *Route Plan* trong kỳ Từ–Đến ngày.
- **Tiền đi lại** = `KM<1000 ? KM×1.500 : 1.000×1.500 + (KM−1000)×1.900`.
- **Nội dung DNTT** = chia kỳ thành N đoạn ngày liền nhau (N = số hóa đơn).
- **Hóa đơn cuối** (ngày trễ nhất / dòng cuối) được chỉnh XUỐNG cho tổng khớp tiền
  đi lại theo KM. Không đủ → báo THIẾU, không bịa tăng.
- Điền: **DNTT**, **Bảng kê chi tiết** (KM, địa điểm), **Tra cứu HĐ** (từ PDF).

## File
- `app.py` — giao diện Streamlit
- `core.py` — toàn bộ logic (dùng chung, làm việc trên bytes)
- `requirements.txt` · `.streamlit/config.toml`
- `HƯỚNG DẪN DEPLOY.md` — cách đưa lên web

## Chạy tại máy
```
pip3 install --user -r requirements.txt
streamlit run app.py
```
