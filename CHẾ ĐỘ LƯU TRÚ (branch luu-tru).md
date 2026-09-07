# Branch `luu-tru` — chế độ có lưu trú / nhiều chuyến

Branch này thêm **chế độ thứ 2** cho app: kỳ có ngủ lại, nhiều chuyến, hóa đơn
khách sạn + per-diem. Chế độ đơn giản (cũ) giữ nguyên.

## Deploy 1 app Streamlit RIÊNG để test (không đụng link chính)
1. https://share.streamlit.io → **Create app**
2. Repository: `lcblongg/bot-cong-tac-phi`
3. **Branch: `luu-tru`**  ← khác branch main
4. Main file path: `app.py`
5. App URL: đặt tên khác, ví dụ `bot-ctp-luutru`
6. Deploy → link test `bot-ctp-luutru.streamlit.app`

Link chính (`main`) vẫn chạy bình thường, không bị ảnh hưởng.

## Test gì
- Chọn "Có lưu trú / nhiều chuyến".
- Nhập bảng chuyến (5-6 dòng: địa điểm, từ–đến ngày, số đêm, KM).
- Upload tất cả hóa đơn (xăng .pdf/.eml + khách sạn .pdf/.eml).
- Bảng "Số tiền hóa đơn xăng" giờ BOT **tự đọc bằng OCR** (số tiền + ngày + đoán chuyến),
  kể cả hóa đơn Petrolimex dạng ảnh. Chỉ cần soát lại, sửa ô nào sai.
- Upload file Excel mẫu tháng trước của người đó.
- Tạo → kiểm tra file Excel: DNTT (xăng theo chuyến + khách sạn + per-diem),
  Bảng kê chi tiết (chuyến, KM, khách sạn), Tra cứu HĐ, tổng tiền, đọc số thành chữ.

## Đã test với file thật của Lại Thanh Sơn (tháng 8/2026)
6 HĐ xăng + 5 HĐKS + per-diem → tổng **9.058.400đ** khớp đúng file làm tay.
Tự dò được layout template (Sơn có DNTT 12 dòng / 2 dòng diễn giải, khác Lê Công Bảo Long).
OCR đọc đúng 6/6 số tiền + ngày hóa đơn xăng Petrolimex ảnh; ghép PDF↔EML qua "mã tra cứu".

## OCR trên Streamlit Cloud
`packages.txt` cài `tesseract-ocr` + `tesseract-ocr-vie`; `requirements.txt` thêm
`pytesseract`, `Pillow`. Nếu Cloud không cài được tesseract → OCR trả rỗng, bảng xăng để
trống, người dùng gõ tay như cũ (không vỡ app).

## Khi OK → merge vào main
`git checkout main && git merge luu-tru && git push` → link chính tự cập nhật.
