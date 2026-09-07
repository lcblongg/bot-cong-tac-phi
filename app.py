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

st.markdown("""
<style>
  .big-title { font-size: 1.7rem; font-weight: 800; letter-spacing:-.01em; }
  .by { color: #8a93a3; font-size: .85rem; margin-top:-4px; }
  div[data-testid="stMetricValue"] { font-size: 1.2rem; }
  .stDownloadButton button { width: 100%; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="big-title">✨ BOT Công tác phí</div>', unsafe_allow_html=True)
st.markdown('<div class="by">Tự tạo file Excel công tác phí từ hóa đơn · thiết kế bởi lcblongg</div>',
            unsafe_allow_html=True)

# ----------------------------- hướng dẫn -----------------------------
with st.expander("📖  HƯỚNG DẪN SỬ DỤNG  —  bấm để mở", expanded=False):
    st.markdown("""
**Chế độ đơn giản** (không lưu trú, chỉ xăng): điền tháng/kỳ → upload hóa đơn xăng
→ nhập tổng KM (hoặc upload file Numbers) → upload file Excel mẫu tháng trước → bấm tạo.

**Chế độ có lưu trú / nhiều chuyến**: nhập **bảng chuyến** (địa điểm · từ–đến ngày ·
số đêm · KM mỗi chuyến) → upload tất cả hóa đơn (xăng + khách sạn) → BOT **tự đọc số tiền
hóa đơn xăng** (kể cả hóa đơn Petrolimex dạng ảnh), chỉ cần soát lại → upload file Excel
mẫu → bấm tạo.

BOT lo phần còn lại: phân loại HĐX/HĐKS, đọc tiền khách sạn từ PDF, ghép hóa đơn ↔ chuyến,
chia tỷ lệ mileage, dòng Per-diem, điền DNTT + Bảng kê + Tra cứu HĐ, đọc số thành chữ,
đổi tên file, kiểm địa chỉ người mua. Kết quả: **1 file zip** = Excel + hóa đơn đã đổi tên.

Cuối cùng: mở file Excel bằng Excel/Numbers **một lần** cho công thức tự tính lại, rồi lưu.
""")

st.divider()

# ----------------------------- sidebar -----------------------------
with st.sidebar:
    st.subheader("Đơn vị người mua")
    st.caption("Kiểm tra địa chỉ người mua trên hóa đơn. Mặc định CPM Việt Nam — công ty khác thì sửa.")
    b_ten = st.text_input("Tên đơn vị", core.TEN_NGUOI_MUA)
    b_mst = st.text_input("Mã số thuế", core.MST_NGUOI_MUA)
    b_dc = st.text_area("Địa chỉ chuẩn", core.DIA_CHI_CHUAN, height=90)
buyer = {"ten": b_ten.strip(), "mst": b_mst.strip(), "dia_chi": b_dc.strip()}

# ----------------------------- chế độ -----------------------------
mode = st.radio(
    "Chế độ",
    ["Đơn giản — 1 chuyến, chỉ xăng (dùng KM)",
     "Có lưu trú / nhiều chuyến (khách sạn + per-diem)"],
    help="Đa số đồng nghiệp có ngủ lại → chọn chế độ thứ 2.",
)
MULTI = mode.startswith("Có lưu trú")

# ----------------------------- Thông tin kỳ (chung) -----------------------------
st.subheader("1 · Thông tin kỳ")
c1, c2, c3 = st.columns(3)
thang_txt = c1.text_input("Tháng công tác phí", placeholder="9/2026")
today = dt.date.today()
tu_ngay = c2.date_input("Kỳ — Từ ngày", value=today.replace(day=1) - dt.timedelta(days=11),
                        format="DD/MM/YYYY")
den_ngay = c3.date_input("Kỳ — Đến ngày", value=today.replace(day=20), format="DD/MM/YYYY")
c4, c5 = st.columns(2)
tam_ung = c4.number_input("Tạm ứng (VND)", min_value=0, value=0, step=100_000)
che_do_dc = c5.radio("Kiểm tra địa chỉ hóa đơn", ["Linh hoạt", "Chính xác tuyệt đối"],
                     horizontal=True)
muc_dich = st.text_input("Mục đích công tác (tùy chọn)")

trips_df = gas_df = None
dia_diem = khu_vuc = ""
so_dem = so_pd = 0
numbers_file = None
km_manual = None

if MULTI:
    khu_vuc = st.text_input("Khu vực chung (dòng tiêu đề DNTT)",
                            placeholder="Hưng Yên - Hà Nam - Ninh Bình")

    st.subheader("2 · Bảng chuyến công tác")
    st.caption("Mỗi dòng = 1 chuyến. Chuyến đi lại nội vùng cả tháng (không ngủ) → số đêm = 0.")
    trips_df = st.data_editor(
        [{"Địa điểm": "", "Từ ngày": None, "Đến ngày": None, "Số đêm": 0, "Tổng KM": 0.0}],
        num_rows="dynamic", width='stretch', key="trips",
        column_config={
            "Từ ngày": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Đến ngày": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Số đêm": st.column_config.NumberColumn(min_value=0, step=1),
            "Tổng KM": st.column_config.NumberColumn(min_value=0.0, step=1.0),
        },
    )

    # trips tạm (để gợi ý chuyến cho bảng xăng)
    trips_now = []
    for row in (trips_df or []):
        _dd = str(row.get("Địa điểm") or "").strip()
        _tu = core.parse_date(row.get("Từ ngày"))
        _den = core.parse_date(row.get("Đến ngày"))
        if not _dd and not _tu:
            continue
        trips_now.append(core.Trip(_dd, _tu, _den,
                                   int(row.get("Số đêm") or 0),
                                   float(row.get("Tổng KM") or 0)))

    st.subheader("3 · Hóa đơn (xăng + khách sạn)")
    inv_files = st.file_uploader("Kéo-thả tất cả .pdf và .eml của kỳ (cả xăng lẫn khách sạn)",
                                 type=["pdf", "eml"], accept_multiple_files=True)

    st.subheader("4 · Số tiền hóa đơn xăng — BOT tự đọc, soát lại nếu sai")
    st.caption("BOT tự đọc số tiền + ngày từ hóa đơn xăng (kể cả Petrolimex dạng ảnh) và "
               "đoán 'Chuyến'. Sai thì sửa thẳng trong bảng. "
               "Cột 'Chuyến' = số thứ tự dòng ở bảng chuyến (1, 2, 3…).")
    _gas_default = [{"Số hóa đơn": "", "Số tiền": None, "Chuyến": 1}]
    if inv_files:
        sig = (tuple(sorted((f.name, f.size) for f in inv_files))
               + tuple((str(t.tu), str(t.den)) for t in trips_now))
        if st.session_state.get("gas_seed_sig") != sig:
            with st.spinner("Đang đọc hóa đơn xăng (OCR)…"):
                _pf = core.prefill_gas_rows(
                    [(f.name, f.getvalue()) for f in inv_files], trips_now)
            st.session_state["gas_seed"] = [
                {"Số hóa đơn": r["so_hd"], "Số tiền": r["so_tien"], "Chuyến": r["chuyen"]}
                for r in _pf] or _gas_default
            st.session_state["gas_seed_sig"] = sig
        _gas_default = st.session_state.get("gas_seed", _gas_default)
    gas_df = st.data_editor(
        _gas_default,
        num_rows="dynamic", width='stretch',
        key=f"gas_{hash(st.session_state.get('gas_seed_sig'))}",
        column_config={
            "Số tiền": st.column_config.NumberColumn(min_value=0, step=1000),
            "Chuyến": st.column_config.NumberColumn(min_value=1, step=1),
        },
    )

    st.subheader("5 · File Excel mẫu (tháng trước của bạn)")
    tmpl_file = st.file_uploader("File công tác phí THÁNG TRƯỚC (.xlsx) — BOT sao cấu trúc từ đây",
                                 type=["xlsx"])

else:
    dia_diem = st.text_input("Địa điểm công tác")
    cc1, cc2 = st.columns(2)
    so_dem = cc1.number_input("Số đêm khách sạn", min_value=0, value=0, step=1)
    so_pd = cc2.number_input("Số ngày Per-diem", min_value=0, value=0, step=1)

    st.subheader("2 · Hóa đơn xăng của kỳ")
    inv_files = st.file_uploader("Kéo-thả tất cả file .pdf và .eml của kỳ vào đây",
                                 type=["pdf", "eml"], accept_multiple_files=True)

    st.subheader("3 · Số KM")
    km_mode = st.radio("Lấy số KM từ đâu?",
                       ["Tự tính từ file Numbers", "Gõ tay tổng KM"], horizontal=True)
    if km_mode.startswith("Tự tính"):
        numbers_file = st.file_uploader("File Numbers (có bảng “Route Plan”)", type=["numbers"])
    else:
        km_manual = st.number_input("Tổng KM công tác của kỳ", min_value=0.0, value=0.0, step=1.0)

    st.subheader("4 · File Excel mẫu (tháng trước)")
    tmpl_file = st.file_uploader("File công tác phí THÁNG TRƯỚC của bạn (.xlsx)", type=["xlsx"])

    with st.expander("Tùy chọn — tự đặt thứ tự / mô tả hóa đơn"):
        manual_df = st.data_editor(
            [{"Số hóa đơn": "", "Mô tả": "", "Số tiền (ghi đè)": None}],
            num_rows="dynamic", width='stretch', key="manual",
        )

# ----------------------------- CHẠY -----------------------------
st.divider()
go = st.button("▶  Tạo file công tác phí", type="primary", width='stretch')

if go:
    month, year = core.parse_month_year(thang_txt)
    problems = []
    if not month:
        problems.append("Ô 'Tháng công tác phí' phải dạng M/YYYY, ví dụ 9/2026.")
    if not inv_files:
        problems.append("Chưa upload hóa đơn.")
    if not tmpl_file:
        problems.append("Chưa upload file Excel mẫu tháng trước.")
    if not MULTI:
        if numbers_file is None and not km_manual:
            problems.append("Nhập tổng KM, hoặc upload file Numbers.")
    if problems:
        for p in problems:
            st.error(p)
        st.stop()

    cfg = core.Config(
        month=month, year=year, tu_ngay=tu_ngay, den_ngay=den_ngay,
        dia_diem=dia_diem.strip(), muc_dich=muc_dich.strip(),
        so_dem=int(so_dem), so_ngay_pd=int(so_pd), tam_ung=int(tam_ung),
        km_override=float(km_manual) if km_manual else None,
        strict_addr=che_do_dc.startswith("Chính xác"), buyer=buyer,
    )
    files = [(f.name, f.getvalue()) for f in inv_files]

    with st.spinner("Đang xử lý…"):
        if MULTI:
            trips = []
            for row in (trips_df or []):
                dd = str(row.get("Địa điểm") or "").strip()
                tu = core.parse_date(row.get("Từ ngày"))
                den = core.parse_date(row.get("Đến ngày"))
                if not dd and not tu:
                    continue
                trips.append(core.Trip(dd, tu, den,
                                       int(row.get("Số đêm") or 0),
                                       float(row.get("Tổng KM") or 0)))
            gas_rows = [{"so_hd": r.get("Số hóa đơn"), "so_tien": r.get("Số tiền"),
                         "chuyen": r.get("Chuyến")}
                        for r in (gas_df or []) if (r.get("Số hóa đơn") or r.get("Số tiền"))]
            R = core.generate_multi(tmpl_file.getvalue(), files, cfg, trips, gas_rows, khu_vuc.strip())
        else:
            rows = [r for r in manual_df if (r.get("Số hóa đơn") or r.get("Mô tả"))]
            manual = [{"so_hd": r.get("Số hóa đơn"), "mota": r.get("Mô tả"),
                       "override": r.get("Số tiền (ghi đè)")} for r in rows] or None
            R = core.generate(tmpl_file.getvalue(), files, cfg, manual,
                              numbers_file.getvalue() if numbers_file else None)

    if R.errors:
        st.session_state.pop("R", None)
        for e in R.errors:
            st.error(e)
        st.stop()

    st.session_state["R"] = R
    st.session_state["files"] = files
    st.session_state["mon_yr"] = (cfg.month, cfg.year)
    st.session_state["tam_ung"] = cfg.tam_ung
    st.session_state["multi"] = MULTI


# ------- KẾT QUẢ -------
R = st.session_state.get("R")
if R:
    files = st.session_state["files"]
    mon, yr = st.session_state["mon_yr"]
    tam_ung = st.session_state["tam_ung"]
    is_multi = st.session_state.get("multi", False)

    st.divider()
    for w in R.warnings:
        st.warning(w)
    st.success(f"Xong — **{R.out_filename}**")

    if is_multi:
        m = st.columns(3)
        m[0].metric("Tổng KM", f"{R.km:,.0f}")
        m[1].metric("Tiền đi lại (KM)", f"{R.tien_di_lai:,.0f}")
        m[2].metric("Khách sạn", f"{R.tien_ks:,.0f}")
        m = st.columns(3)
        m[0].metric("Per-diem", f"{R.tien_pd:,.0f}")
        m[1].metric("TỔNG CHI PHÍ", f"{R.tong_cp:,.0f}")
        if tam_ung:
            m[2].metric("Thu / chi thêm", f"{R.thu_chi_them:,.0f}")
    else:
        m = st.columns(4)
        m[0].metric("Tổng KM", f"{R.km:,.0f}", R.km_source)
        m[1].metric("Tiền đi lại (KM)", f"{R.tien_di_lai:,.0f}")
        m[2].metric("Tổng chi phí", f"{R.tong_cp:,.0f}")
        if tam_ung:
            m[3].metric("Thu / chi thêm", f"{R.thu_chi_them:,.0f}")

    if R.khop:
        st.success(f"✅ Tổng hóa đơn xăng {R.tong_hd:,.0f} = tiền đi lại theo KM {R.tien_di_lai:,.0f} — KHỚP.")
    elif R.thieu:
        st.warning(f"⚠️ Thiếu {R.thieu:,.0f}đ hóa đơn xăng so với tiền đi lại theo KM. File vẫn tạo được.")

    st.dataframe(
        [{"HĐ": r["so_hd"], "Loại": r.get("loai_hd") or "per-diem",
          "Chưa VAT": r["chua_vat"], "VAT": r["vat"], "Thanh toán": r["thanh_toan"],
          "Nội dung": r["mota"], "": "← chỉnh" if r.get("is_last") else ""}
         for r in R.invoices],
        width='stretch', hide_index=True,
    )

    if R.km_detail:
        with st.expander(f"Chi tiết KM theo ngày ({len(R.km_detail)} ngày)"):
            st.dataframe([{"Ngày": d.strftime("%d/%m/%Y"), "Km": v}
                          for d, v in sorted(R.km_detail.items())],
                         width='stretch', hide_index=True)
    for l in R.log:
        st.caption("• " + l)

    st.divider()
    st.download_button(
        "⬇️  TẢI TẤT CẢ — file Excel + hóa đơn đã đổi tên (.zip)",
        core.build_full_zip(R, files),
        file_name=f"Cong tac phi - {core.MON[mon]} {yr}.zip",
        mime="application/zip", type="primary", width='stretch',
    )
    with st.expander("Tải riêng từng phần"):
        c1, c2 = st.columns(2)
        c1.download_button("⬇️  Chỉ file Excel", R.xlsx_bytes, file_name=R.out_filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width='stretch')
        c2.download_button("⬇️  Chỉ hóa đơn (.zip)", core.build_renamed_zip(files, R.renamed),
                           file_name=f"Hoa don - {core.MON[mon]} {yr}.zip",
                           mime="application/zip", width='stretch')

    st.info("Mở file Excel bằng Excel/Numbers **một lần** cho công thức (tổng tiền, đọc thành chữ) "
            "tự tính lại, rồi lưu.")
