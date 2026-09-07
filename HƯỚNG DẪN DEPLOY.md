# Đưa BOT Công tác phí lên web (Streamlit Community Cloud)

Kết quả: 1 đường link `https://…streamlit.app` — bạn bè chỉ cần mở, không cài gì.

---

## Bước 1 — Tài khoản (làm 1 lần)
1. Có tài khoản **GitHub**: https://github.com/signup
2. Đăng nhập **Streamlit Community Cloud** bằng GitHub: https://share.streamlit.io

## Bước 2 — Đưa code lên GitHub

**Cách A — GitHub Desktop (dễ, không cần gõ lệnh)**
1. Tải https://desktop.github.com → đăng nhập.
2. File → *Add Local Repository…* → chọn thư mục `Bot-Streamlit`.
3. Nó hỏi tạo repo → *create a repository* → *Publish repository*
   (bỏ tick "Keep this code private" nếu muốn public, hoặc để private cũng deploy được).

**Cách B — Terminal**
```
cd "/Users/lcblongg/Travel Cost Claim/Bot-Streamlit"
git init && git add -A && git commit -m "BOT cong tac phi - streamlit"
```
Rồi tạo repo rỗng trên github.com, copy URL của nó và:
```
git branch -M main
git remote add origin https://github.com/<tên-bạn>/<tên-repo>.git
git push -u origin main
```

## Bước 3 — Deploy
1. Vào https://share.streamlit.io → **Create app** → **Deploy a public app from GitHub**.
2. Chọn:
   - Repository: repo vừa đẩy lên
   - Branch: `main`
   - Main file path: `app.py`
3. Bấm **Deploy**. Chờ ~2 phút → được link `https://…streamlit.app`.

## Bước 4 — Gửi bạn bè
Copy link đó gửi qua Zalo / email. Xong.

---

## Cập nhật app sau này
Sửa file (`app.py` / `core.py`) → đẩy lên GitHub (GitHub Desktop: Commit → Push /
Terminal: `git add -A && git commit -m "sua" && git push`) → Streamlit **tự deploy lại**.

## Giới hạn người xem (tùy chọn)
App public thì ai có link cũng mở được — nhưng **mỗi phiên độc lập, app không lưu
dữ liệu của ai**. Muốn khóa theo email: trong Streamlit Cloud → app → **Settings →
Sharing** → thêm email được phép.

## Lưu ý riêng tư
Hóa đơn / file Numbers khi dùng app sẽ được **tải lên máy chủ Streamlit** để xử lý
(rồi bỏ, không lưu). Nếu công ty không cho phép dữ liệu qua bên thứ ba → dùng bản
`.command` chạy tại máy thay vì bản web này.

## Công ty / đơn vị khác
Không cần sửa code — trong app có ô **"Đơn vị người mua"** (thanh bên trái) để đổi
tên / MST / địa chỉ chuẩn. Định mức km 1.500 / 1.900 đ nằm trong `core.py` hàm
`mileage_amount()` nếu cần chỉnh.

---

## Chạy thử tại máy trước khi deploy (tùy chọn)
```
pip3 install --user -r requirements.txt
streamlit run app.py
```
