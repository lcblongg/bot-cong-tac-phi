# -*- coding: utf-8 -*-
"""
core.py — Logic dùng chung cho BOT Công tác phí (CLI + Streamlit).

Tất cả hàm ở đây làm việc trên BYTES / đối tượng file, không đụng tới đường dẫn
cố định, không print, không sys.exit. Kết quả trả về trong dataclass Result.
"""
from __future__ import annotations

import io
import re
import html
import email
import email.header
import zipfile
import tempfile
import unicodedata
import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

# ================== ĐỊA CHỈ / ĐƠN VỊ NGƯỜI MUA (mặc định: CPM VIỆT NAM) ==================
DIA_CHI_CHUAN = "Tầng 5, Số 231 – 233 Lê Thánh Tôn, Phường Bến Thành, TP Hồ Chí Minh, Việt Nam"
TEN_NGUOI_MUA = "CÔNG TY CỔ PHẦN CPM VIỆT NAM"
MST_NGUOI_MUA = "0305908974"

MON = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

DNTT_INV_ROW0, DNTT_INV_ROW_MAX, TRACUU_ROW0 = 27, 33, 3


# ------------------------------ helpers thuần -----------------------------------
def collapse(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def parse_money(s):
    if s is None:
        return None
    s = re.sub(r"[^\d]", "", str(s).strip())
    return int(s) if s else None


def parse_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if v in (None, ""):
        return None
    s = str(v).strip().replace("-", "/").replace(".", "/")
    for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%d/%m/%y", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def to_number(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(re.sub(r"[.,\s]", "", str(v)))
    except ValueError:
        return None


def parse_month_year(v):
    """'9/2026' | '2026-09' | datetime | '9' -> (month, year) hoặc (None, None)."""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.month, v.year
    s = str(v or "").strip()
    mo = re.search(r"\b(\d{1,2})\s*[/\-]\s*(\d{4})\b", s)
    if mo:
        return int(mo.group(1)), int(mo.group(2))
    mo = re.search(r"\b(\d{4})\s*[/\-]\s*(\d{1,2})\b", s)
    if mo:
        return int(mo.group(2)), int(mo.group(1))
    mo = re.match(r"\s*(\d{1,2})\s*$", s)
    if mo:
        return int(mo.group(1)), dt.date.today().year
    return None, None


def chia_ky_cong_tac(tu, den, n):
    """Chia kỳ [tu..den] (cả 2 đầu) thành n đoạn ngày liền nhau, đoạn dài dồn về cuối,
    đoạn cuối kết thúc đúng 'den'.  21/07–20/08 (31 ngày) / 4 hóa đơn -> 7,8,8,8."""
    n = max(1, n)
    tong = (den - tu).days + 1
    base, rem = divmod(tong, n)
    doan, cur = [], tu
    for i in range(n):
        dai = base + (1 if i >= n - rem else 0)
        end = den if i == n - 1 else cur + dt.timedelta(days=max(1, dai) - 1)
        doan.append((cur, end))
        cur = end + dt.timedelta(days=1)
    return doan


def mileage_amount(km):
    """Công thức bậc thang: 1000 km đầu × 1.500đ, phần vượt × 1.900đ."""
    if km < 1000:
        return round(km * 1500)
    return round(1000 * 1500 + (km - 1000) * 1900)


def tidy_url(u):
    u = (u or "").strip().rstrip("/#")
    if not u:
        return ""
    if not u.startswith("http"):
        u = "https://" + u
    if "pvoil.vn" in u:
        return "https://hoadon.pvoil.vn/Invoice/search#"
    if "vnpt-invoice.com.vn" in u:
        return u + "/"
    return u


def norm_addr(s):
    s = unicodedata.normalize("NFC", s or "").lower().strip()
    s = s.replace("–", "-").replace("—", "-")
    s = re.sub(r"\s*-\s*", "-", s)
    s = re.sub(r"tp\.?\s*hồ chí minh|thành phố hồ chí minh|tp\s*hcm|tphcm|hcmc", "hcm", s)
    s = re.sub(r"\bsố\b", " ", s)
    s = s.replace(".", " ").replace(",", " ")
    return re.sub(r"\s+", " ", s).strip()


# ------------------------------ đọc hóa đơn -----------------------------------
def _pdf_text(pdf_bytes: bytes) -> str:
    try:
        from pypdf import PdfReader
        r = PdfReader(io.BytesIO(pdf_bytes))
        return "\n".join((p.extract_text() or "") for p in r.pages)
    except Exception:
        return ""


def extract_invoice(pdf_bytes: bytes, buyer=None) -> dict:
    buyer = buyer or {}
    ten_mua = buyer.get("ten", TEN_NGUOI_MUA)
    mst_mua = buyer.get("mst", MST_NGUOI_MUA)
    dc_chuan = buyer.get("dia_chi", DIA_CHI_CHUAN)

    t = collapse(_pdf_text(pdf_bytes))
    d = {"so_hd": None, "so_tien": None, "vat": 0, "web": "", "ma_tra_cuu": "",
         "mst_ban": "", "dia_chi_mua": "", "dia_chi_ok": False, "ngay": None, "loai": "X"}
    if not t:
        return d

    m = (re.search(r"Số\s*\(No\)?\s*:\s*(\d{5,10})", t)
         or re.search(r"Ký hiệu[^0-9]{0,20}?Số:\s*(\d{5,10})", t)
         or re.search(r"\bSố:\s*(\d{5,10})\b", t))
    if m:
        d["so_hd"] = m.group(1)

    m = re.search(r"[Tt]ra cứu hóa đơn[^:]*tại:\s*(https?://[^\s,)]+|[a-z0-9.\-]+\.vn)", t)
    if m:
        d["web"] = tidy_url(m.group(1))
    m = re.search(r"[Mm]ã tra cứu[^:]*:\s*([A-Za-z0-9\-]+)", t)
    if m:
        d["ma_tra_cuu"] = m.group(1)

    m = re.search(r"Đơn vị bán hàng[^:]*:\s*(.+?)\s*Mã số thuế[^:]*:\s*(\d{10,13})", t)
    if m:
        d["mst_ban"] = m.group(2)
    else:
        m = re.search(r"(SƠN HẢI|XĂNG DẦU[^0-9]{0,40})\s*(\d{10})", t)
        if m:
            d["mst_ban"] = m.group(2)

    kkknt = "KKKNT" in t
    m = (re.search(r"Tổng cộng tiền thanh toán[^:]*:\s*([\d.]+)", t)
         or re.search(r"Cộng tiền hàng[^:]*:\s*([\d.]+)", t))
    if m:
        d["so_tien"] = parse_money(m.group(1))
    if not kkknt:
        m = re.search(r"Tiền thuế GTGT[^:]*:\s*([\d.]+)", t)
        d["vat"] = (parse_money(m.group(1)) or 0) if m else 0

    m = re.search(r"Ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})", t)
    if m:
        try:
            d["ngay"] = dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass

    m = re.search(r"(Xăng|Dầu)\s+[^\n]{0,40}?(Lít|Kg|lít|kg)\b", t, re.I)
    if m:
        d["loai"] = "D" if m.group(1)[0].lower() == "d" else "X"

    # địa chỉ người mua: khối "Tầng N ... Việt Nam" (bên bán là cây xăng, không mở đầu "Tầng")
    m = re.search(r"(Tầng\s*\d+[^\n]{5,120}?(?:Việt\s*Nam|Vietnam)\.?)", t, re.I)
    if m:
        d["dia_chi_mua"] = collapse(m.group(1))
    d["dia_chi_ok"] = (
        ten_mua.lower() in t.lower()
        and mst_mua in t
        and bool(d["dia_chi_mua"])
        and norm_addr(d["dia_chi_mua"]) == norm_addr(dc_chuan)
    )
    return d


def eml_invoice_no(eml_bytes: bytes):
    try:
        m = email.message_from_bytes(eml_bytes)
    except Exception:
        return None
    subj = str(email.header.make_header(email.header.decode_header(m.get("Subject", ""))))
    body = ""
    for p in m.walk():
        if p.get_content_type() in ("text/plain", "text/html"):
            try:
                body += p.get_payload(decode=True).decode(
                    p.get_content_charset() or "utf-8", errors="replace")
            except Exception:
                pass
    text = html.unescape(re.sub(r"<[^>]+>", " ", body))
    for pat in (r"[Ss]ố hóa đơn[:\s]*0*(\d{5,10})",
                r"[Hh]óa đơn điện tử số[:\s]*0*(\d{5,10})",
                r"Số[:\s]*0*(\d{6,10})"):
        mm = re.search(pat, subj + " " + text)
        if mm:
            return mm.group(1)
    return None


# ------------------------------ KM từ Numbers -----------------------------------
def tinh_km(numbers_bytes: bytes, tu, den):
    """(tong_km, {ngày: km}, [sheet thiếu]).  Đọc bảng 'Route Plan' cột Date + Km."""
    from numbers_parser import Document
    with tempfile.NamedTemporaryFile(suffix=".numbers", delete=False) as tmp:
        tmp.write(numbers_bytes)
        tmp_path = tmp.name
    try:
        doc = Document(tmp_path)
        routes = {}
        for sh in doc.sheets:
            for tb in sh.tables:
                if tb.name.strip() == "Route Plan":
                    routes[sh.name] = tb

        thang = set()
        d = tu
        while d <= den:
            thang.add((d.month, d.year))
            d += dt.timedelta(days=1)

        tong, chi_tiet, thieu = 0.0, {}, []
        for (mm, yy) in sorted(thang):
            ten = f"Lịch đi shop {mm}/{yy % 100}"
            tb = routes.get(ten) or next((routes[k] for k in routes if k.strip() == ten), None)
            if tb is None:
                thieu.append(ten)
                continue
            for r, cells in enumerate(tb.rows(values_only=True)):
                if r == 0 or len(cells) < 7:
                    continue
                ngay, km = cells[2], cells[6]
                if isinstance(ngay, dt.datetime):
                    ngay = ngay.date()
                if isinstance(ngay, dt.date) and isinstance(km, (int, float)) and tu <= ngay <= den:
                    tong += km
                    chi_tiet[ngay] = chi_tiet.get(ngay, 0.0) + km
        return tong, chi_tiet, thieu
    finally:
        try:
            Path(tmp_path).unlink()
        except OSError:
            pass


# ------------------------------ dataclasses -----------------------------------
@dataclass
class Config:
    month: int
    year: int
    tu_ngay: dt.date
    den_ngay: dt.date
    dia_diem: str = ""
    muc_dich: str = ""
    so_dem: int = 0
    so_ngay_pd: int = 0
    tam_ung: int = 0
    km_override: float | None = None
    strict_addr: bool = False
    buyer: dict = field(default_factory=dict)   # {ten, mst, dia_chi} — để trống = CPM

    @property
    def thang_label(self):
        return f"{self.month}/{self.year}"


@dataclass
class Result:
    ok: bool = False
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    log: list = field(default_factory=list)
    km: float = 0.0
    km_source: str = ""
    km_detail: dict = field(default_factory=dict)
    tien_di_lai: int = 0
    tien_ks: int = 0
    tien_pd: int = 0
    tong_cp: int = 0
    thu_chi_them: int = 0
    invoices: list = field(default_factory=list)      # dict/hàng đưa vào Excel
    tong_hd: int = 0
    khop: bool = False
    thieu: int = 0
    xlsx_bytes: bytes | None = None
    out_filename: str = ""
    renamed: list = field(default_factory=list)       # (tên cũ, tên mới)


# ------------------------------ điền Excel -----------------------------------
def fill_workbook(template_bytes: bytes, cfg: Config, invoices, km) -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))

    dn = wb["DNTT"]
    dn["G3"] = f"Thời gian/ Time: {cfg.thang_label}"
    c10 = str(dn["C10"].value or "Công tác phí tháng")
    c10 = re.sub(r"th[áa]ng\s*\S*\s*$", f"tháng {cfg.thang_label}", c10.strip())
    if "tháng" not in c10:
        c10 += f" tháng {cfg.thang_label}"
    dn["C10"] = c10
    for r in range(DNTT_INV_ROW0, DNTT_INV_ROW_MAX + 1):
        for col in ("B", "F", "G", "H", "I", "J"):
            dn[f"{col}{r}"].value = None
            dn[f"{col}{r}"].hyperlink = None
    for i, inv in enumerate(invoices):
        r = DNTT_INV_ROW0 + i
        if r > DNTT_INV_ROW_MAX:
            raise ValueError("Quá 7 hóa đơn — mẫu DNTT không đủ dòng.")
        dn[f"A{r}"] = i + 1
        dn[f"B{r}"] = inv["mota"]
        dn[f"F{r}"] = cfg.thang_label
        dn[f"G{r}"] = inv["chua_vat"]
        dn[f"H{r}"] = inv["vat"]
        dn[f"I{r}"] = inv["thanh_toan"]
        dn[f"J{r}"] = inv["so_hd"]

    bk = wb[next(s for s in wb.sheetnames if s.strip() == "Bảng kê chi tiết")]
    bk["H3"] = f"Thời gian/ Time: {cfg.month:02d}/{cfg.year}"
    if cfg.muc_dich:
        bk["C4"] = cfg.muc_dich
    bk["C7"] = cfg.dia_diem
    bk["D7"] = (f"{cfg.tu_ngay.month}/{cfg.tu_ngay.day}/{cfg.tu_ngay.year} - "
                f"{cfg.den_ngay.month}/{cfg.den_ngay.day}/{cfg.den_ngay.year}")
    bk["E7"] = cfg.so_dem
    bk["G7"] = cfg.so_ngay_pd
    bk["D16"] = round(km)
    bk["H23"] = cfg.tam_ung if cfg.tam_ung else None

    tc = wb[next(s for s in wb.sheetnames if s.strip() == "Tra cứu HĐ")]
    for r in range(TRACUU_ROW0, TRACUU_ROW0 + 30):
        for col in ("B", "C", "D", "E"):
            tc[f"{col}{r}"].value = None
            tc[f"{col}{r}"].hyperlink = None
    for i, inv in enumerate(invoices):
        r = TRACUU_ROW0 + i
        tc[f"B{r}"] = inv["loai_hd"]
        tc[f"C{r}"] = inv["web"]
        tc[f"D{r}"] = inv["ma_tra_cuu"]
        tc[f"E{r}"] = inv["so_hd"]

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================== ORCHESTRATOR ==================================
def generate(template_bytes: bytes,
             invoice_files: list,          # [(filename, bytes), ...]  .pdf / .eml
             cfg: Config,
             manual_invoices: list | None = None,   # [{so_hd, mota, override}]
             numbers_bytes: bytes | None = None) -> Result:
    """Chạy toàn bộ quy trình. Không raise — mọi lỗi vào Result.errors."""
    R = Result()

    def err(m): R.errors.append(m)
    def warn(m): R.warnings.append(m)
    def log(m): R.log.append(m)

    # --- validate cơ bản ---
    if not template_bytes:
        err("Chưa có file Excel mẫu (tháng trước).")
    if cfg.den_ngay < cfg.tu_ngay:
        err(f"'Đến ngày' ({cfg.den_ngay:%d/%m/%Y}) sớm hơn 'Từ ngày' ({cfg.tu_ngay:%d/%m/%Y}).")
    pdfs = [(n, b) for n, b in invoice_files if n.lower().endswith(".pdf")]
    emls = [(n, b) for n, b in invoice_files if n.lower().endswith(".eml")]
    if not pdfs:
        err("Chưa upload file PDF hóa đơn nào.")
    if R.errors:
        return R

    # --- 1) đọc từng PDF ---
    log(f"Đọc {len(pdfs)} PDF + {len(emls)} EML.")
    data = {}            # so_hd -> info
    for name, b in pdfs:
        inv = extract_invoice(b, cfg.buyer)
        so = inv["so_hd"]
        if not so:
            warn(f"Không đọc được số hóa đơn trong: {name}")
            continue
        pre = "HĐD" if inv["loai"] == "D" else "HĐX"
        mark = "" if inv["dia_chi_ok"] else " [SAI ĐỊA CHỈ]"
        inv["_pre"] = pre
        new_name = f"{pre}_{so}{mark}.pdf"
        R.renamed.append((name, new_name))
        data[so] = inv
        if not inv["dia_chi_ok"]:
            warn(f"HĐ {so}: SAI/THIẾU ĐỊA CHỈ người mua — "
                 f"thấy “{inv['dia_chi_mua'] or 'không tìm thấy'}”.")

    if not data:
        err("Không PDF nào đọc được số hóa đơn.")
        return R

    for name, b in emls:
        so = (eml_invoice_no(b) or "").lstrip("0")
        match = next((k for k in data if k.lstrip("0") == so and so), None)
        if match:
            pre = data[match]["_pre"]
            mark = "" if data[match]["dia_chi_ok"] else " [SAI ĐỊA CHỈ]"
            R.renamed.append((name, f"{pre}_{match}{mark}.eml"))
        else:
            warn(f"EML không khớp hóa đơn nào: {name}")

    # --- 2) danh sách hóa đơn ---
    rows = []
    if manual_invoices:
        for m in manual_invoices:
            so = re.sub(r"\D", "", str(m.get("so_hd") or "")) or None
            mota = str(m.get("mota") or "").strip()
            if not so and not mota:
                continue
            rows.append({"so_hd": so, "mota": mota, "override": parse_money(m.get("override"))})
        # khớp với PDF
        miss = []
        for row in rows:
            pd_ = data.get(row["so_hd"]) or next(
                (v for k, v in data.items() if k.lstrip("0") == (row["so_hd"] or "").lstrip("0")), None)
            if pd_ is None:
                miss.append(row["so_hd"] or row["mota"])
            else:
                row["_pdf"] = pd_
        if miss:
            err("Hóa đơn trong danh sách KHÔNG thấy PDF: " + ", ".join(map(str, miss))
                + f"\nPDF đang có: {', '.join(sorted(data))}")
            return R
    else:
        # tự lập từ TẤT CẢ pdf: sắp theo ngày; trùng ngày -> tiền nhỏ trước
        items = list(data.values())
        for v in items:
            if v["ngay"] and not (cfg.tu_ngay <= v["ngay"] <= cfg.den_ngay):
                warn(f"HĐ {v['so_hd']} ghi ngày {v['ngay']:%d/%m/%Y} — ngoài kỳ "
                     f"{cfg.tu_ngay:%d/%m}–{cfg.den_ngay:%d/%m} (vẫn dùng).")
        order = sorted(items, key=lambda v: (v["ngay"] or dt.date(2100, 1, 1),
                                             v["so_tien"] or 0, v["so_hd"]))
        rows = [{"so_hd": v["so_hd"], "mota": "", "override": None, "_pdf": v} for v in order]
        log(f"Tự lập {len(rows)} hóa đơn (giữ hết), sắp theo ngày.")

    if not rows:
        err("Không có hóa đơn nào.")
        return R
    if len(rows) > 7:
        err(f"{len(rows)} hóa đơn — mẫu DNTT chỉ 7 dòng. Bớt bớt hóa đơn.")
        return R

    if cfg.strict_addr:
        bad = [r["so_hd"] for r in rows if not r["_pdf"]["dia_chi_ok"]]
        if bad:
            err("Chế độ 'Chính xác tuyệt đối': các HĐ sai địa chỉ — " + ", ".join(bad))
            return R

    # --- 2b) nội dung DNTT: chia kỳ thành N đoạn ngày ---
    doan = chia_ky_cong_tac(cfg.tu_ngay, cfg.den_ngay, len(rows))
    for row, (s, e) in zip(rows, doan):
        loai_hd = "HĐD" if row["_pdf"]["loai"] == "D" else "HĐX"
        loai_txt = "dầu" if loai_hd == "HĐD" else "xăng"
        if not row["mota"]:
            row["mota"] = f"CP {loai_txt} xe: {s:%d/%m/%Y} - {e:%d/%m/%Y}"
        row["loai_hd"] = loai_hd
        row["web"] = row["_pdf"]["web"]
        row["ma_tra_cuu"] = row["_pdf"]["ma_tra_cuu"]
        row["vat"] = int(round(row["_pdf"]["vat"] or 0))
        real = row["override"] if row["override"] is not None else row["_pdf"]["so_tien"]
        if real is None:
            err(f"Không đọc được số tiền HĐ {row['so_hd']}. Nhập tay 'Số tiền (ghi đè)'.")
            return R
        row["thanh_toan"] = int(round(real))

    # --- 3) KM + tiền ---
    if cfg.km_override:
        R.km = float(cfg.km_override)
        R.km_source = "nhập tay"
    elif numbers_bytes:
        try:
            tong, ct, missing = tinh_km(numbers_bytes, cfg.tu_ngay, cfg.den_ngay)
        except Exception as ex:
            err(f"Không đọc được file Numbers: {ex}")
            return R
        if missing:
            err("File Numbers thiếu sheet Route Plan cho: " + ", ".join(missing)
                + "\n-> Chọn 'Gõ tay số KM' và nhập tổng KM.")
            return R
        R.km, R.km_detail, R.km_source = tong, ct, "file Numbers"
    else:
        err("Chưa có số KM: upload file Numbers, hoặc chọn 'Gõ tay' và nhập tổng KM.")
        return R

    R.tien_di_lai = mileage_amount(R.km)
    R.tien_ks = cfg.so_dem * 500_000
    R.tien_pd = cfg.so_ngay_pd * 200_000
    R.tong_cp = R.tien_di_lai + R.tien_ks + R.tien_pd
    R.thu_chi_them = R.tong_cp - cfg.tam_ung

    # --- 4) hóa đơn cuối = số bù (chỉnh XUỐNG, không bịa tăng) ---
    for r in rows:
        r["chua_vat"] = r["thanh_toan"] - r["vat"]
    last = rows[-1]
    goc_last = last["thanh_toan"]
    tong_khac = sum(r["thanh_toan"] for r in rows[:-1])
    can_co = int(round(R.tien_di_lai - tong_khac))

    if can_co < 0:
        err(f"Các hóa đơn TRƯỚC hóa đơn cuối ({tong_khac:,.0f}) đã VƯỢT tiền đi lại theo KM "
            f"({R.tien_di_lai:,.0f}). Bớt hóa đơn hoặc kiểm tra lại KM / kỳ ngày.")
        return R
    if can_co < goc_last:
        last["thanh_toan"] = can_co
        last["chua_vat"] = can_co - last["vat"]
        log(f"Hóa đơn cuối HĐ {last['so_hd']}: {goc_last:,.0f} → {can_co:,.0f} "
            f"(chỉnh XUỐNG cho khớp KM).")
    elif can_co == goc_last:
        log("Tổng hóa đơn khớp sẵn tiền đi lại theo KM — không chỉnh gì.")
    else:
        R.thieu = can_co - goc_last
        warn(f"THIẾU HÓA ĐƠN: còn thiếu {R.thieu:,.0f} đồng so với tiền đi lại theo KM. "
             f"Bổ sung thêm hóa đơn rồi tạo lại. (BOT KHÔNG bịa tăng hóa đơn cuối.)")

    for j, r in enumerate(rows):
        r["is_last"] = (j == len(rows) - 1)
    R.invoices = rows
    R.tong_hd = sum(r["thanh_toan"] for r in rows)
    R.khop = abs(R.tong_hd - R.tien_di_lai) < 1

    # --- 5) xuất Excel ---
    try:
        R.xlsx_bytes = fill_workbook(template_bytes, cfg, rows, R.km)
    except Exception as ex:
        err(f"Lỗi khi điền Excel: {ex}")
        return R
    R.out_filename = f"CS Le Cong Bao Long - {MON[cfg.month]} {cfg.year}.xlsx"
    R.ok = True
    return R


def build_renamed_zip(invoice_files: list, renamed: list) -> bytes:
    """Đóng gói lại các file PDF/EML với tên mới (HĐX_<số>...) thành 1 zip."""
    by_name = {n: b for n, b in invoice_files}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for old, new in renamed:
            if old in by_name:
                z.writestr(new, by_name[old])
    return buf.getvalue()


def build_full_zip(result: "Result", invoice_files: list) -> bytes:
    """1 file zip chứa TẤT CẢ: file Excel công tác phí + các hóa đơn đã đổi tên."""
    by_name = {n: b for n, b in invoice_files}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        if result.xlsx_bytes:
            z.writestr(result.out_filename, result.xlsx_bytes)
        for old, new in result.renamed:
            if old in by_name:
                z.writestr(f"Hóa đơn/{new}", by_name[old])
    return buf.getvalue()
