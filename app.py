# -*- coding: utf-8 -*-
"""
BOT Công tác phí — bản web (Streamlit).  Thiết kế bởi lcblongg.

Chạy local:   streamlit run app.py
Deploy:       Streamlit Community Cloud  (xem HƯỚNG DẪN DEPLOY.md)
"""
import datetime as dt
import streamlit as st

import core

st.set_page_config(page_title="BOT Công tác phí", page_icon="✨", layout="centered")

# ----------------------------- style nhẹ -----------------------------
st.markdown("""
<style>
  .stApp { }
  .big-title { font-size: 1.7rem; font-weight: 800; letter-spacing:-.01em; }
  .by { color: #8a93a3; font-size: .85rem; margin-top:-4px; }
  div[data-testid="stMetricValue"] { font-size: 1.25rem; }
  .stDownloadButton button { width: 100%; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="big-title">✨ BOT Công tác phí</div>', unsafe_allow_html=True)
st.markdown('<div class="by">Tự tạo file Excel công tác phí từ hóa đơn xăng · thiết kế bởi lcblongg</div>',
            unsafe_allow_html=True)

# ----------------------------- hướng dẫn sử dụng -----------------------------
with st.expander("📖  HƯỚNG DẪN SỬ DỤNG  —  bấm để mở (lần đầu nên đọc)", expanded=False):
    st.markdown("""
Công cụ này thay bạn dựng file Excel công tác phí hàng tháng. Làm đúng **4 phần** bên dưới rồi bấm nút xanh.

**Phần 1 — Thông tin kỳ**
- *Tháng công tác phí*: gõ `9/2026` (tháng/năm).
- *Kỳ — Từ ngày / Đến ngày*: khoảng thời gian claim (thường 21 tháng trước → 20 tháng này).
- *Địa điểm công tác*, *Tạm ứng*… điền nếu có. Không có thì để 0.

**Phần 2 — Hóa đơn xăng của kỳ**
- Kéo-thả **tất cả** file `.pdf` và `.eml` hóa đơn xăng của kỳ vào ô upload.
- Tên file lộn xộn cũng được — công cụ tự đọc số hóa đơn và đổi tên `HĐX_<số>`.

**Phần 3 — Số KM**
- *Tự tính từ file Numbers*: upload file Numbers của bạn (có bảng **Route Plan**) → tự cộng KM theo kỳ.
- *Gõ tay tổng KM*: nếu không có file Numbers, tự xem rồi nhập tổng KM vào ô.

**Phần 4 — File Excel mẫu (tháng trước)**
- Upload đúng **file công tác phí tháng trước của bạn** (.xlsx). Công cụ sao y cấu trúc, tên, số tài khoản, mã dự án từ file này.

**Bấm ▶ Tạo file công tác phí** → xem kết quả:
- Bảng tóm tắt: Tổng KM · Tiền đi lại · Tổng chi phí · **KHỚP / THIẾU**.
- Bảng hóa đơn (hóa đơn cuối được chỉnh cho tổng khớp tiền đi lại theo KM).
- Nút **TẢI TẤT CẢ (.zip)** = 1 file zip gồm **file Excel + toàn bộ hóa đơn đã đổi tên**.

Cuối cùng: mở file Excel bằng Excel/Numbers **một lần** cho công thức tự tính lại, rồi lưu.
""")
    c1, c2 = st.columns(2)
    try:
        c1.image("guide/01-form.jpg", caption="Điền form + kéo-thả hóa đơn")
        c2.image("guide/03-ketqua.jpg", caption="Kết quả + nút tải tất cả")
    except Exception:
        pass

st.divider()

# ----------------------------- sidebar: đơn vị người mua -----------------------------
with st.sidebar:
    st.subheader("Đơn vị người mua")
    st.caption("BOT kiểm tra địa chỉ người mua trên hóa đơn có khớp không. "
               "Mặc định là CPM Việt Nam — công ty khác thì sửa 3 ô dưới.")
    b_ten = st.text_input("Tên đơn vị", core.TEN_NGUOI_MUA)
    b_mst = st.text_input("Mã số thuế", core.MST_NGUOI_MUA)
    b_dc = st.text_area("Địa chỉ chuẩn", core.DIA_CHI_CHUAN, height=90)
    st.divider()
    st.caption("📖 Hướng dẫn: mở mục **HƯỚNG DẪN SỬ DỤNG** ở đầu trang.")

buyer = {"ten": b_ten.strip(), "mst": b_mst.strip(), "dia_chi": b_dc.strip()}

# ----------------------------- 1. Thông tin kỳ -----------------------------
st.subheader("1 · Thông tin kỳ")
c1, c2, c3 = st.columns(3)
thang_txt = c1.text_input("Tháng công tác phí", placeholder="9/2026")
today = dt.date.today()
tu_ngay = c2.date_input("Kỳ — Từ ngày", value=today.replace(day=21) - dt.timedelta(days=31),
                        format="DD/MM/YYYY")
den_ngay = c3.date_input("Kỳ — Đến ngày", value=today.replace(day=20), format="DD/MM/YYYY")

dia_diem = st.text_input("Địa điểm công tác")
muc_dich = st.text_input("Mục đích công tác (tùy chọn)")

c4, c5, c6 = st.columns(3)
so_dem = c4.number_input("Số đêm khách sạn", min_value=0, value=0, step=1)
so_pd = c5.number_input("Số ngày Per-diem", min_value=0, value=0, step=1)
tam_ung = c6.number_input("Tạm ứng (VND)", min_value=0, value=0, step=100_000)

che_do = st.radio("Kiểm tra địa chỉ hóa đơn", ["Linh hoạt", "Chính xác tuyệt đối"],
                  horizontal=True,
                  help="Linh hoạt: chấp nhận khác cách viết (TP / Thành phố, dấu gạch…). "
                       "Chính xác tuyệt đối: sai địa chỉ thì DỪNG.")

# ----------------------------- 2. Hóa đơn -----------------------------
st.subheader("2 · Hóa đơn xăng của kỳ")
inv_files = st.file_uploader("Kéo-thả tất cả file .pdf và .eml của kỳ vào đây",
                             type=["pdf", "eml"], accept_multiple_files=True)

# ----------------------------- 3. Số KM -----------------------------
st.subheader("3 · Số KM")
km_mode = st.radio("Lấy số KM từ đâu?",
                   ["Tự tính từ file Numbers", "Gõ tay tổng KM"], horizontal=True)
numbers_file = None
km_manual = None
if km_mode.startswith("Tự tính"):
    numbers_file = st.file_uploader("File Numbers (có bảng “Route Plan”)", type=["numbers"])
else:
    km_manual = st.number_input("Tổng KM công tác của kỳ", min_value=0.0, value=0.0, step=1.0)

# ----------------------------- 4. File Excel mẫu -----------------------------
st.subheader("4 · File Excel mẫu (tháng trước)")
tmpl_file = st.file_uploader("File công tác phí THÁNG TRƯỚC của bạn (.xlsx) — BOT sao cấu trúc từ đây",
                             type=["xlsx"])

# ----------------------------- 5. (tùy chọn) danh sách hóa đơn -----------------------------
with st.expander("Tùy chọn — tự đặt thứ tự / mô tả hóa đơn"):
    st.caption("Để trống = BOT tự lập (sắp theo ngày hóa đơn). "
               "Điền vào đây nếu muốn tự quyết thứ tự, mô tả, hoặc ép số tiền.")
    manual_df = st.data_editor(
        [{"Số hóa đơn": "", "Mô tả": "", "Số tiền (ghi đè)": None}],
        num_rows="dynamic", use_container_width=True, key="manual",
    )

# ----------------------------- CHẠY -----------------------------
st.divider()
go = st.button("▶  Tạo file công tác phí", type="primary", use_container_width=True)

if go:
    month, year = core.parse_month_year(thang_txt)
    problems = []
    if not month:
        problems.append("Ô 'Tháng công tác phí' phải dạng M/YYYY, ví dụ 9/2026.")
    if not inv_files:
        problems.append("Chưa upload hóa đơn.")
    if not tmpl_file:
        problems.append("Chưa upload file Excel mẫu tháng trước.")
    if km_mode.startswith("Tự tính") and not numbers_file:
        problems.append("Chọn 'Tự tính' thì phải upload file Numbers (hoặc đổi sang 'Gõ tay').")
    if km_mode.startswith("Gõ tay") and not km_manual:
        problems.append("Nhập tổng KM (> 0).")
    if problems:
        for p in problems:
            st.error(p)
        st.stop()

    cfg = core.Config(
        month=month, year=year, tu_ngay=tu_ngay, den_ngay=den_ngay,
        dia_diem=dia_diem.strip(), muc_dich=muc_dich.strip(),
        so_dem=int(so_dem), so_ngay_pd=int(so_pd), tam_ung=int(tam_ung),
        km_override=float(km_manual) if km_manual else None,
        strict_addr=che_do.startswith("Chính xác"),
        buyer=buyer,
    )

    files = [(f.name, f.getvalue()) for f in inv_files]
    manual = None
    rows = [r for r in manual_df if (r.get("Số hóa đơn") or r.get("Mô tả"))]
    if rows:
        manual = [{"so_hd": r.get("Số hóa đơn"), "mota": r.get("Mô tả"),
                   "override": r.get("Số tiền (ghi đè)")} for r in rows]

    with st.spinner("Đang đọc hóa đơn, tính KM, dựng Excel…"):
        R = core.generate(
            template_bytes=tmpl_file.getvalue(),
            invoice_files=files, cfg=cfg, manual_invoices=manual,
            numbers_bytes=numbers_file.getvalue() if numbers_file else None,
        )

    if R.errors:
        # lỗi -> hiện ngay, không lưu kết quả
        st.session_state.pop("R", None)
        for e in R.errors:
            st.error(e)
        st.stop()

    # lưu vào session_state để KHÔNG bị mất khi bấm nút tải (Streamlit rerun sau mỗi lần tải)
    st.session_state["R"] = R
    st.session_state["files"] = files
    st.session_state["mon_yr"] = (cfg.month, cfg.year)
    st.session_state["tam_ung"] = cfg.tam_ung


# ------- HIỂN THỊ KẾT QUẢ (chạy mỗi lần, đọc từ session_state) -------
R = st.session_state.get("R")
if R:
    files = st.session_state["files"]
    mon, yr = st.session_state["mon_yr"]
    tam_ung = st.session_state["tam_ung"]

    st.divider()
    for w in R.warnings:
        st.warning(w)

    st.success(f"Xong — **{R.out_filename}**")
    cols = st.columns(4)
    cols[0].metric("Tổng KM", f"{R.km:,.0f}", R.km_source)
    cols[1].metric("Tiền đi lại (KM)", f"{R.tien_di_lai:,.0f}")
    cols[2].metric("Tổng chi phí", f"{R.tong_cp:,.0f}")
    if tam_ung:
        cols[3].metric("Thu / chi thêm", f"{R.thu_chi_them:,.0f}")

    if R.khop:
        st.success(f"✅ Tổng hóa đơn {R.tong_hd:,.0f} = tiền đi lại theo KM {R.tien_di_lai:,.0f} — KHỚP.")
    else:
        st.warning(f"⚠️ Tổng hóa đơn {R.tong_hd:,.0f} ≠ tiền đi lại {R.tien_di_lai:,.0f} "
                   f"(lệch {R.tong_hd - R.tien_di_lai:+,.0f}). File vẫn tạo được.")

    st.dataframe(
        [{"HĐ": r["so_hd"], "Loại": r["loai_hd"],
          "Chưa VAT": r["chua_vat"], "VAT": r["vat"], "Thanh toán": r["thanh_toan"],
          "Nội dung": r["mota"], "": "← hóa đơn cuối" if r["is_last"] else ""}
         for r in R.invoices],
        use_container_width=True, hide_index=True,
    )

    if R.km_detail:
        with st.expander(f"Chi tiết KM theo ngày ({len(R.km_detail)} ngày)"):
            st.dataframe(
                [{"Ngày": d.strftime("%d/%m/%Y"), "Km": v} for d, v in sorted(R.km_detail.items())],
                use_container_width=True, hide_index=True,
            )
    for l in R.log:
        st.caption("• " + l)

    st.divider()
    full_zip = core.build_full_zip(R, files)
    st.download_button(
        "⬇️  TẢI TẤT CẢ — file Excel + hóa đơn đã đổi tên (.zip)",
        full_zip, file_name=f"Cong tac phi - {core.MON[mon]} {yr}.zip",
        mime="application/zip", type="primary", use_container_width=True,
    )
    with st.expander("Tải riêng từng phần"):
        c1, c2 = st.columns(2)
        c1.download_button("⬇️  Chỉ file Excel", R.xlsx_bytes, file_name=R.out_filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
        c2.download_button("⬇️  Chỉ hóa đơn đã đổi tên (.zip)",
                           core.build_renamed_zip(files, R.renamed),
                           file_name=f"Hoa don - {core.MON[mon]} {yr}.zip",
                           mime="application/zip", use_container_width=True)

    st.info("Mở file Excel bằng Excel/Numbers **một lần** để công thức (tổng tiền, đọc thành chữ) "
            "tự tính lại, rồi lưu.")
