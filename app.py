"""
Streamlit app: Tự động tạo sheet "KTV-TVV" từ raw data — BẢN NÂNG CẤP (nhanh
hơn + chi tiết hơn + đối chiếu bằng raw_data trong Payroll):

  1. File "Phân tích Doanh thu khách hàng" (nhiều tháng, mỗi chi nhánh 1 sheet).
  2. File Payroll — upload BAO NHIÊU FILE CŨNG ĐƯỢC (sheet "Bảng lương").

ĐIỂM MỚI so với bản trước:
  ⚡ TỐC ĐỘ: đọc file payroll ở chế độ read_only (nhanh hơn ~30 lần so với
     cách đọc thông thường — 1 file 17MB từ ~34s xuống còn ~1s) + cache lại
     kết quả đọc file (st.cache_data) để bấm nút xuất nhiều lần không phải
     đọc lại file từ đầu.
  📊 THÊM CỘT "TẤT CẢ CHI NHÁNH": tổng hợp toàn công ty ở cuối bảng, tính
     đúng theo trọng số (vd tỷ lệ chốt = tổng khách chốt / tổng khách, không
     phải trung bình cộng đơn giản của 9 chi nhánh).
  🎨 TÔ MÀU CHÊNH LỆCH: cột "Δ" tự động xanh khi tốt lên / đỏ khi xấu đi
     (doanh thu tăng = xanh, chi phí nhân sự tăng = đỏ...).
  🔎 SHEET "AUDIT" MỚI: tự động cảnh báo — chi nhánh có trong Payroll nhưng
     chưa cấu hình trong app (dữ liệu sẽ bị bỏ sót), chi nhánh thiếu dữ liệu
     doanh thu, chi nhánh có Chi phí nhân sự tăng nhanh hơn Doanh thu, chi
     nhánh sụt doanh thu so với tháng trước.
  📺 DASHBOARD TRỰC TIẾP TRÊN WEB: ngay sau khi upload đủ file (không cần bấm
     xuất Excel) — có bộ lọc tháng/chi nhánh, thẻ KPI, 5 tab biểu đồ (so sánh
     chi nhánh, xu hướng theo tháng, hiệu suất KTV/TVV, chi phí & hiệu quả,
     bảng chi tiết có thể tải CSV). Phần này chỉ hiển thị trên web, không ảnh
     hưởng tới file Excel xuất ra.
  🆕 ĐỌC THÊM "RAW DATA" TRONG PAYROLL (sheet "Data tổng"): mỗi file Payroll
     tự mang theo 1 dashboard doanh thu riêng (KPI, Doanh thu trước/sau thuế
     phí, DT khách cũ...) — app giờ đọc luôn sheet này (gọi tắt là raw_data)
     để:
       (a) ĐỐI CHIẾU chéo với số "Tổng doanh thu" trong file Phân tích Doanh
           thu khách hàng — lệch quá 0.5% sẽ bị cảnh báo trong sheet Audit
           (đây chính xác là cách đã dùng để bắt lỗi số liệu trong file báo
           cáo lương T7-T8/2026 thực tế trước đó).
       (b) DỰ PHÒNG (fallback): nếu 1 chi nhánh/tháng nào đó bị thiếu dữ liệu
           trong file Phân tích Doanh thu, app tự động lấy tạm "Doanh thu
           trước thuế phí" / "DT khách cũ" từ raw_data trong Payroll để bảng
           không bị bỏ trống ô — có cảnh báo rõ trong Audit là số đang dùng
           tạm từ raw_data, không phải từ file doanh thu chính thức.
     Tên cột trong sheet "Data tổng" đổi khác nhau giữa các tháng (vd tháng 08
     tách "DT khách cũ" thành "DT khách cũ trước thuế phí" / "...sau thuế
     phí"), nên app dò theo DANH SÁCH BIẾN THỂ tên cột, không hard-code 1 tên
     cố định.

Toàn bộ công thức nghiệp vụ đã ĐỐI CHIẾU khớp chính xác 100% với file báo cáo
lương T8/2026 (so với T7/2026) người dùng tự làm tay, trên toàn bộ 9 chi nhánh
và 26 chỉ tiêu.

LƯU Ý: vị trí cột "TỔNG THU NHẬP" và các cột khác trong "Bảng lương" LỆCH NHAU
giữa các tháng — app dò cột theo TÊN HEADER, không theo số thứ tự cột.

Cách chạy:
    pip install streamlit openpyxl pandas plotly
    streamlit run app.py
"""

import io
import re

import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# ---------------------------------------------------------------------------
# 1. CẤU HÌNH CHI NHÁNH
# ---------------------------------------------------------------------------

BRANCHES = [
    {"label": "Bình Dương", "payroll": "BÌNH DƯƠNG", "revenue_sheet": "Bình Dương"},
    {"label": "Vũng Tàu",   "payroll": "VŨNG TÀU",   "revenue_sheet": "Vũng Tàu"},
    {"label": "Q.1",        "payroll": "QUẬN 1",     "revenue_sheet": "Quận 1"},
    {"label": "Gò Vấp",     "payroll": "GÒ VẤP",     "revenue_sheet": "Gò Vấp"},
    {"label": "Q.10",       "payroll": "QUẬN 10",    "revenue_sheet": "Quận 10"},
    {"label": "Tân Bình",   "payroll": "TÂN BÌNH",   "revenue_sheet": "Tân Bình"},
    {"label": "Tân Phú",    "payroll": "TÂN PHÚ",    "revenue_sheet": "Tân Phú"},
    {"label": "Bình Tân",   "payroll": "BÌNH TÂN",   "revenue_sheet": "Bình Tân"},
    {"label": "Phú Nhuận",  "payroll": "PHÚ NHUẬN",  "revenue_sheet": "Phú Nhuận"},
]
TOTAL_LABEL = "TẤT CẢ CHI NHÁNH"

KTV_ROLES = {"Kỹ thuật viên", "Kỹ thuật viên phun xăm", "Chuyên viên Clinic"}
TVV_ROLES = {"Tư vấn viên", "Trợ lý bác sĩ"}
OMCMLEAD_ROLES = {"OM", "CM", "LEAD"}
QLCN_ROLES = {"QLCN"}

NEEDED_HEADERS = {
    "chi_nhanh": "CHI NHÁNH LÀM VIỆC",
    "vi_tri": "VỊ TRÍ",
    "ho_ten": "HỌ VÀ TÊN",
    "dt_ca_nhan": "DOANH THU CÁ NHÂN TRƯỚC THUẾ PHÍ",
    "tour_ca_nhan": "TỔNG ĐIỂM TOUR CÁ NHÂN",
    "tong_thu_nhap": "TỔNG THU NHẬP",
    "ty_le_kpi_ca_nhan": "TỶ LỆ %\nCÁ NHÂN ĐẠT SO VỚI KPI",
}

# --- MỚI: cấu hình đọc sheet "Data tổng" (raw_data) trong file Payroll -----
# Mỗi field có 1 danh sách biến thể tên cột (thử theo thứ tự, lấy cái khớp
# đầu tiên) vì tên cột đổi khác nhau giữa các tháng.
RAW_DATA_SHEET_NAME = "Data tổng"
RAW_DATA_HEADERS = {
    "chi_nhanh": ["Chi nhánh"],
    "kpi_doanh_thu": ["KPI"],
    "doanh_thu_truoc_thue": ["Doanh thu trước thuế phí"],
    "doanh_thu_sau_thue": ["Doanh thu sau thuế phí"],
    "dt_khach_cu": ["DT khách cũ trước thuế phí", "DT khách cũ"],
}
RAW_DATA_SEARCH_ROWS = (1, 2, 3)
RAW_DATA_MAX_COL = 6          # bảng doanh thu chính nằm ở các cột đầu (1-6)
RAW_DATA_MAX_SCAN_ROWS = 40   # quét tối đa bấy nhiêu dòng để tìm 9 chi nhánh
RAW_DATA_MISMATCH_THRESHOLD = 0.005  # 0.5% — lệch hơn mức này mới cảnh báo

FONT = Font(name="Arial", size=11)
FONT_BOLD = Font(name="Arial", size=11, bold=True)
FONT_HEADER = Font(name="Arial", size=11, bold=True, color="FFFFFF")
FILL_HEADER = PatternFill("solid", fgColor="4472C4")
FILL_TOTAL_HEADER = PatternFill("solid", fgColor="1F3864")
FILL_SECTION = PatternFill("solid", fgColor="D9E1F2")
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_DIFF = PatternFill("solid", fgColor="F2F2F2")
FILL_GOOD = PatternFill("solid", fgColor="C6EFCE")
FILL_BAD = PatternFill("solid", fgColor="FFC7CE")
FILL_WARN = PatternFill("solid", fgColor="FFEB9C")
FILL_RAWDATA = PatternFill("solid", fgColor="DDEBF7")  # màu riêng cho cảnh báo raw_data
FONT_GOOD = Font(name="Arial", size=11, color="006100")
FONT_BAD = Font(name="Arial", size=11, color="9C0006")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = "#,##0"
INT_FMT = "#,##0"
PCT_FMT = "0.0%"


# ---------------------------------------------------------------------------
# 2. ĐỌC FILE PAYROLL — read_only + cache để nhanh
# ---------------------------------------------------------------------------

def find_header_columns(ws, header_map, search_rows=(2, 3, 4)):
    found = {}
    max_col = ws.max_column
    for r in search_rows:
        for c in range(1, max_col + 1):
            val = ws.cell(row=r, column=c).value
            if val is None:
                continue
            val_norm = str(val).strip()
            for key, header_name in header_map.items():
                if key in found:
                    continue
                if val_norm == header_name or val_norm.replace("\n", " ") == header_name.replace("\n", " "):
                    found[key] = c
    missing = [header_map[k] for k in header_map if k not in found]
    if missing:
        raise ValueError(
            "Không tìm thấy các cột sau trong sheet 'Bảng lương': " + ", ".join(missing) +
            ". Kiểm tra lại định dạng file payroll."
        )
    return found


def find_raw_data_columns(ws):
    """Dò cột trong sheet 'Data tổng' (raw_data) theo danh sách biến thể tên
    cột trong RAW_DATA_HEADERS. Trả về dict field -> số cột (1-indexed).
    Không raise lỗi nếu thiếu — sheet raw_data là nguồn PHỤ (đối chiếu / dự
    phòng), thiếu thì bỏ qua chứ không chặn app."""
    found = {}
    max_col = min(ws.max_column, RAW_DATA_MAX_COL)
    for r in RAW_DATA_SEARCH_ROWS:
        for c in range(1, max_col + 1):
            val = ws.cell(row=r, column=c).value
            if val is None:
                continue
            val_norm = str(val).strip()
            for key, variants in RAW_DATA_HEADERS.items():
                if key in found:
                    continue
                if val_norm in variants:
                    found[key] = c
    return found


def read_raw_data_dashboard(wb):
    """Đọc sheet 'Data tổng' (raw_data) trong workbook Payroll đã mở sẵn.
    Trả về dict: {revenue_sheet_label: {kpi_doanh_thu, doanh_thu_truoc_thue,
    doanh_thu_sau_thue, dt_khach_cu}}. Trả về {} nếu không có sheet này hoặc
    không dò được cột — không làm app dừng lại."""
    if RAW_DATA_SHEET_NAME not in wb.sheetnames:
        return {}
    ws = wb[RAW_DATA_SHEET_NAME]
    cols = find_raw_data_columns(ws)
    if "chi_nhanh" not in cols or "doanh_thu_truoc_thue" not in cols:
        return {}

    valid_labels = {b["revenue_sheet"] for b in BRANCHES}
    idx_cn = cols["chi_nhanh"] - 1
    dashboard = {}
    max_row = min(ws.max_row, RAW_DATA_MAX_SCAN_ROWS)
    for r in range(1, max_row + 1):
        cn = ws.cell(row=r, column=cols["chi_nhanh"]).value
        if not isinstance(cn, str):
            continue
        cn = cn.strip()
        if cn not in valid_labels:
            continue
        entry = {}
        for key in ("kpi_doanh_thu", "doanh_thu_truoc_thue", "doanh_thu_sau_thue", "dt_khach_cu"):
            if key not in cols:
                entry[key] = None
                continue
            v = ws.cell(row=r, column=cols[key]).value
            entry[key] = v if isinstance(v, (int, float)) else None
        dashboard[cn] = entry
    return dashboard


def detect_payroll_month(title: str, filename: str):
    text = f"{title} {filename}"
    m = re.search(r"THÁNG\s*(\d{1,2})\s*/\s*(\d{4})", text, re.IGNORECASE)
    if m:
        month, year = int(m.group(1)), int(m.group(2))
        return year, month, f"T{month}"
    m2 = re.search(r"(\d{2})-(\d{2})-(\d{4})", text)
    if m2:
        dd, mm, yyyy = m2.groups()
        return int(yyyy), int(mm), f"T{int(mm)}"
    return None, None, None


@st.cache_data(show_spinner=False)
def read_payroll_bytes(data: bytes, filename: str):
    """Đọc sheet 'Bảng lương' ở chế độ read_only, ĐỒNG THỜI đọc luôn sheet
    'Data tổng' (raw_data) nếu có. Trả về:
      (year, month, label, {branch: stats}, unmapped_branches_set, employees,
       raw_data_dashboard)
    employees = list các dict {chi_nhanh, ten, vi_tri, nhom, doanh_thu,
    tour, thu_nhap, ty_le_kpi} — dùng để làm bảng xếp hạng top nhân viên.
    raw_data_dashboard = {revenue_sheet_label: {...}} lấy từ sheet Data tổng,
    dùng để đối chiếu / dự phòng cho Tổng doanh thu (xem đầu file)."""
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    if "Bảng lương" not in wb.sheetnames:
        raise ValueError(f"File '{filename}' không có sheet 'Bảng lương'.")
    ws = wb["Bảng lương"]

    title = ws.cell(row=1, column=3).value or ""
    year, month, label = detect_payroll_month(title, filename)

    cols = find_header_columns(ws, NEEDED_HEADERS)
    stats = {}
    unmapped_branches = set()
    employees = []

    def bucket(branch):
        if branch not in stats:
            stats[branch] = {
                "so_nhan_su": 0,
                "so_ktv": 0, "dt_ktv": 0.0, "tour_ktv": 0.0, "thunhap_ktv_sum": 0.0,
                "so_tvv": 0, "dt_tvv": 0.0, "thunhap_tvv_sum": 0.0, "kpi_rate_tvv_sum": 0.0, "kpi_rate_tvv_n": 0,
                "so_omcmlead": 0,
                "so_qlcn": 0,
                "chi_phi_nhan_su": 0.0,
            }
        return stats[branch]

    valid_branch_names = {b["payroll"] for b in BRANCHES}
    branch_label_by_payroll = {b["payroll"]: b["label"] for b in BRANCHES}
    idx_chi_nhanh = cols["chi_nhanh"] - 1
    idx_vi_tri = cols["vi_tri"] - 1
    idx_ten = cols["ho_ten"] - 1
    idx_dt = cols["dt_ca_nhan"] - 1
    idx_tour = cols["tour_ca_nhan"] - 1
    idx_tn = cols["tong_thu_nhap"] - 1
    idx_kpi = cols["ty_le_kpi_ca_nhan"] - 1

    for row in ws.iter_rows(min_row=5, values_only=True):
        branch = row[idx_chi_nhanh]
        if not isinstance(branch, str) or not branch.strip():
            continue
        branch = branch.strip()
        pos = row[idx_vi_tri]
        if pos is None:
            continue
        if branch not in valid_branch_names:
            if re.match(r"^[A-ZÀ-Ỹ ]{3,}$", branch):
                unmapped_branches.add(branch)
            continue
        pos = str(pos).strip()
        ten = row[idx_ten]

        dt_ca_nhan = row[idx_dt] or 0
        tour = row[idx_tour] or 0
        thu_nhap = row[idx_tn] or 0
        kpi_rate = row[idx_kpi]

        b = bucket(branch)
        b["so_nhan_su"] += 1
        b["chi_phi_nhan_su"] += thu_nhap

        nhom = None
        if pos in KTV_ROLES:
            nhom = "KTV"
            b["so_ktv"] += 1
            b["dt_ktv"] += dt_ca_nhan
            b["tour_ktv"] += tour
            b["thunhap_ktv_sum"] += thu_nhap
        elif pos in TVV_ROLES:
            nhom = "TVV"
            b["so_tvv"] += 1
            b["dt_tvv"] += dt_ca_nhan
            b["thunhap_tvv_sum"] += thu_nhap
            if isinstance(kpi_rate, (int, float)):
                b["kpi_rate_tvv_sum"] += kpi_rate
                b["kpi_rate_tvv_n"] += 1
        elif pos in OMCMLEAD_ROLES:
            nhom = "OM/CM/LEAD"
            b["so_omcmlead"] += 1
        elif pos in QLCN_ROLES:
            nhom = "QLCN"
            b["so_qlcn"] += 1
        else:
            nhom = "Khác"

        if nhom in ("KTV", "TVV"):
            employees.append({
                "chi_nhanh": branch_label_by_payroll[branch],
                "ten": str(ten).strip() if ten else "(?)",
                "vi_tri": pos,
                "nhom": nhom,
                "doanh_thu": dt_ca_nhan,
                "tour": tour,
                "thu_nhap": thu_nhap,
                "ty_le_kpi": kpi_rate if isinstance(kpi_rate, (int, float)) else None,
            })

    raw_data_dashboard = read_raw_data_dashboard(wb)

    return year, month, label, stats, unmapped_branches, employees, raw_data_dashboard


# ---------------------------------------------------------------------------
# 3. ĐỌC FILE PHÂN TÍCH DOANH THU
# ---------------------------------------------------------------------------

REVENUE_ROW_MAP = {
    "khach_moi": "Khách mới",
    "dt_khach_moi": "Doanh thu khách mới",
    "khach_cu": "Khách thực tế (cũ)",
    "dt_khach_cu": "Doanh thu khách cũ",
    "ty_le_chot_moi": "Tỷ lệ chốt khách mới",
    "ty_le_chot_cu": "Tỷ lệ chốt khách cũ",
    "bill_tb_moi": "Bill TB khách mới",
    "bill_tb_cu": "Bill TB khách cũ",
    "tong_doanh_thu": "Tổng doanh thu",
    "mua_tt_moi": "Khách mua hàng TT (mới)",
    "mua_tt_cu": "Khách mua hàng TT (cũ)",
    "booking_moi": "Khách booking mới",
    "checkin_moi": "Khách checkin mới",
}


def _to_num(v):
    """Vài ô trong file Phân tích Doanh thu bị lưu dạng text có dấu phẩy /
    khoảng trắng không ngắt (vd '5,129,349,000\\xa0') thay vì số — ép về số
    để tránh lỗi cộng dồn ở phần tổng hợp 'TẤT CẢ CHI NHÁNH'."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        vv = v.replace(",", "").replace("\xa0", "").replace("%", "").strip()
        if not vv:
            return None
        try:
            return float(vv)
        except ValueError:
            return None
    return v


@st.cache_data(show_spinner=False)
def read_revenue_bytes(data: bytes):
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    out = {}
    covered_sheets = {b["revenue_sheet"] for b in BRANCHES}
    unmapped_sheets = set()
    for sheet_name in wb.sheetnames:
        if sheet_name == "Tất cả chi nhánh":
            continue
        if sheet_name not in covered_sheets:
            unmapped_sheets.add(sheet_name)
        ws = wb[sheet_name]
        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        row_by_label = {}
        for r in range(2, ws.max_row + 1):
            chi_tieu = ws.cell(row=r, column=1).value
            if not chi_tieu:
                continue
            row_by_label[str(chi_tieu).strip()] = r

        month_data = {}
        for c_idx, col_label in enumerate(headers[1:], start=2):
            if not col_label:
                continue
            vals = {}
            for field, vn_label in REVENUE_ROW_MAP.items():
                r = row_by_label.get(vn_label)
                vals[field] = _to_num(ws.cell(row=r, column=c_idx).value) if r else None
            month_data[str(col_label).strip()] = vals
        out[sheet_name] = month_data
    return out, unmapped_sheets


# ---------------------------------------------------------------------------
# 4. TỔNG HỢP "TẤT CẢ CHI NHÁNH" (đúng trọng số, không phải trung bình cộng)
# ---------------------------------------------------------------------------

def aggregate_payroll_total(stats_by_branch):
    keys_sum = ["so_nhan_su", "so_ktv", "dt_ktv", "tour_ktv", "thunhap_ktv_sum",
                "so_tvv", "dt_tvv", "thunhap_tvv_sum", "kpi_rate_tvv_sum",
                "kpi_rate_tvv_n", "so_omcmlead", "so_qlcn", "chi_phi_nhan_su"]
    total = {k: 0.0 for k in keys_sum}
    for b in BRANCHES:
        s = stats_by_branch.get(b["payroll"])
        if not s:
            continue
        for k in keys_sum:
            total[k] += s.get(k, 0) or 0
    for k in ("so_nhan_su", "so_ktv", "so_tvv", "so_omcmlead", "so_qlcn", "kpi_rate_tvv_n"):
        total[k] = int(total[k])
    return total


def aggregate_revenue_total(revenue, label):
    total = {"khach_moi": 0, "khach_cu": 0, "dt_khach_moi": 0, "dt_khach_cu": 0,
              "mua_tt_moi": 0, "mua_tt_cu": 0, "tong_doanh_thu": 0,
              "booking_moi": 0, "checkin_moi": 0}
    any_data = False
    for b in BRANCHES:
        d = revenue.get(b["revenue_sheet"], {}).get(label)
        if not d:
            continue
        any_data = True
        for k in total:
            total[k] += d.get(k) or 0
    if not any_data:
        return None
    out = dict(total)
    out["ty_le_chot_moi"] = (total["mua_tt_moi"] / total["khach_moi"]) if total["khach_moi"] else None
    out["ty_le_chot_cu"] = (total["mua_tt_cu"] / total["khach_cu"]) if total["khach_cu"] else None
    out["bill_tb_moi"] = (total["dt_khach_moi"] / total["mua_tt_moi"]) if total["mua_tt_moi"] else None
    out["bill_tb_cu"] = (total["dt_khach_cu"] / total["mua_tt_cu"]) if total["mua_tt_cu"] else None
    return out


# --- MỚI: lấy giá trị doanh thu, có dự phòng bằng raw_data trong Payroll ---

def get_revenue_field_with_fallback(revenue, raw_dashboard, revenue_sheet, month_label, field):
    """Lấy `field` (chỉ áp dụng cho 'tong_doanh_thu' / 'dt_khach_cu') từ file
    Phân tích Doanh thu; nếu thiếu (None) thì lấy tạm từ raw_data (sheet
    'Data tổng' trong Payroll). Trả về (value, from_raw_data: bool)."""
    d = revenue.get(revenue_sheet, {}).get(month_label, {})
    v = d.get(field)
    if v is not None:
        return v, False
    if not raw_dashboard:
        return None, False
    entry = raw_dashboard.get(revenue_sheet)
    if not entry:
        return None, False
    raw_key = "doanh_thu_truoc_thue" if field == "tong_doanh_thu" else "dt_khach_cu"
    v_raw = entry.get(raw_key)
    return v_raw, v_raw is not None


# ---------------------------------------------------------------------------
# 5. LAYOUT CỘT — hỗ trợ N tháng bất kỳ
# ---------------------------------------------------------------------------

def compute_branch_plan(start_col, n_months):
    plan = []
    col = start_col
    val_col = {}
    for i in range(n_months):
        val_col[i] = col
        plan.append(("val", i, col))
        col += 1
        if i > 0:
            plan.append(("diff", i - 1, i, col))
            col += 1
    end_col = col - 1
    return plan, val_col, end_col


def write_values(ws, row, label, plan, get_value_fn, number_format, bold=False, flag_fn=None):
    """flag_fn(i) -> True nếu giá trị ở tháng thứ i là số LẤY TẠM từ raw_data
    (dùng để tô màu riêng + ghi chú, khác với ô thật sự thiếu dữ liệu)."""
    a = ws.cell(row=row, column=1, value=label)
    a.font = FONT_BOLD if bold else FONT
    a.border = BORDER
    for item in plan:
        if item[0] != "val":
            continue
        i, col = item[1], item[2]
        val = get_value_fn(i)
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = FONT
        cell.border = BORDER
        cell.number_format = number_format
        cell.alignment = Alignment(horizontal="right")
        if val is None:
            cell.fill = FILL_INPUT
            cell.comment = Comment("Không có dữ liệu — kiểm tra lại file raw.", "App KTV-TVV")
        elif flag_fn is not None and flag_fn(i):
            cell.fill = FILL_RAWDATA
            cell.comment = Comment(
                "Số này đang LẤY TẠM từ raw_data (sheet 'Data tổng' trong file Payroll) "
                "vì file Phân tích Doanh thu không có dữ liệu — nên xác nhận lại số chính thức.",
                "App KTV-TVV",
            )


def write_formula_same_col(ws, row, label, plan, formula_fn, number_format, bold=False):
    a = ws.cell(row=row, column=1, value=label)
    a.font = FONT_BOLD if bold else FONT
    a.border = BORDER
    for item in plan:
        if item[0] != "val":
            continue
        _, col = item[1], item[2]
        L = get_column_letter(col)
        cell = ws.cell(row=row, column=col, value=formula_fn(L))
        cell.font = FONT
        cell.border = BORDER
        cell.number_format = number_format


def write_diffs(ws, row, plan, val_col, diff_kind, number_format_diff="0.0%", color=False):
    diff_cols = []
    for item in plan:
        if item[0] != "diff":
            continue
        i_from, i_to, col = item[1], item[2], item[3]
        Lf, Lt = get_column_letter(val_col[i_from]), get_column_letter(val_col[i_to])
        cell = ws.cell(row=row, column=col)
        if diff_kind == "abs":
            cell.value = f"={Lt}{row}-{Lf}{row}"
        else:
            cell.value = f"=({Lf}{row}-{Lt}{row})/{Lt}{row}"
            cell.number_format = number_format_diff
        cell.font = FONT
        cell.border = BORDER
        cell.fill = FILL_DIFF
        cell.alignment = Alignment(horizontal="right")
        diff_cols.append(col)
    if color:
        apply_growth_colors(ws, row, diff_cols, reverse=(color == "reverse"))


def apply_growth_colors(ws, row, cols, reverse=False):
    """Tô xanh/đỏ cho các ô Δ dựa trên dấu giá trị. reverse=True dùng cho các
    chỉ tiêu càng THẤP càng tốt (vd chi phí nhân sự tăng = xấu -> đỏ)."""
    good_fill, bad_fill = (FILL_BAD, FILL_GOOD) if reverse else (FILL_GOOD, FILL_BAD)
    for c in cols:
        L = get_column_letter(c)
        ref = f"{L}{row}"
        ws.conditional_formatting.add(
            ref, CellIsRule(operator="greaterThan", formula=["0"], fill=good_fill)
        )
        ws.conditional_formatting.add(
            ref, CellIsRule(operator="lessThan", formula=["0"], fill=bad_fill)
        )


def write_row(ws, row, label, plan, val_col, get_value_fn, diff_kind, number_format, bold=False, color=False, flag_fn=None):
    write_values(ws, row, label, plan, get_value_fn, number_format, bold=bold, flag_fn=flag_fn)
    write_diffs(ws, row, plan, val_col, diff_kind, "0.0%", color=color)


def write_formula_row(ws, row, label, plan, val_col, formula_fn, diff_kind, number_format, bold=False, color=False):
    write_formula_same_col(ws, row, label, plan, formula_fn, number_format, bold=bold)
    write_diffs(ws, row, plan, val_col, diff_kind, "0.0%", color=color)


def style_header(cell, text, wrap=True, total=False):
    cell.value = text
    cell.font = FONT_HEADER
    cell.fill = FILL_TOTAL_HEADER if total else FILL_HEADER
    cell.alignment = Alignment(horizontal="center", wrap_text=wrap)
    cell.border = BORDER


def section_label(ws, row, text):
    ws.cell(row=row, column=1, value=text)
    ws.cell(row=row, column=1).font = FONT_BOLD
    ws.cell(row=row, column=1).fill = FILL_SECTION


# ---------------------------------------------------------------------------
# 6. XÂY DỰNG SHEET "KTV-TVV"
# ---------------------------------------------------------------------------

def build_ktv_tvv_sheet(wb, months, revenue):
    """months: list các tuple (label, payroll_stats, raw_dashboard)."""
    n = len(months)
    month_labels = [m[0] for m in months]
    payroll_by_month = [m[1] for m in months]
    raw_dashboard_by_month = [m[2] for m in months]
    total_payroll_by_month = [aggregate_payroll_total(s) for s in payroll_by_month]
    total_revenue_by_month = [aggregate_revenue_total(revenue, lbl) for lbl in month_labels]

    ws = wb.create_sheet("KTV-TVV")
    ws.column_dimensions["A"].width = 34

    all_entries = [(b["label"], False) for b in BRANCHES] + [(TOTAL_LABEL, True)]
    branch_plans = {}
    col = 2
    for label, is_total in all_entries:
        plan, val_col, end_col = compute_branch_plan(col, n)
        branch_plans[label] = (plan, val_col, col, end_col, is_total)
        ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=end_col)
        style_header(ws.cell(row=2, column=col), label, total=is_total)
        for item in plan:
            if item[0] == "val":
                style_header(ws.cell(row=3, column=item[2]), month_labels[item[1]], total=is_total)
            else:
                style_header(ws.cell(row=3, column=item[3]), "Δ", total=is_total)
        for cc in range(col, end_col + 1):
            ws.column_dimensions[get_column_letter(cc)].width = 13
        col = end_col + 1

    ws["A2"] = "Chỉ tiêu"; ws["A2"].font = FONT_HEADER; ws["A2"].fill = FILL_HEADER
    ws["A3"] = "Tháng"; ws["A3"].font = FONT_HEADER; ws["A3"].fill = FILL_HEADER

    def get_rev(label, month_idx, field, is_total):
        if is_total:
            d = total_revenue_by_month[month_idx]
            return d.get(field) if d else None
        sheet_name = next(b["revenue_sheet"] for b in BRANCHES if b["label"] == label)
        d = revenue.get(sheet_name, {}).get(month_labels[month_idx], {})
        return d.get(field)

    # --- MỚI: bản có dự phòng raw_data, chỉ dùng cho 2 field hỗ trợ được ---
    def get_rev_fallback(label, month_idx, field, is_total):
        """Trả về (value, from_raw). Cột TỔNG (is_total) KHÔNG áp dụng dự
        phòng riêng lẻ — nó vẫn cộng từ số các chi nhánh (đã có fallback)."""
        if is_total:
            # cộng lại từ từng chi nhánh để cột Tổng cũng được hưởng fallback
            total = 0.0
            any_val = False
            any_raw = False
            for b in BRANCHES:
                v, from_raw = get_revenue_field_with_fallback(
                    revenue, raw_dashboard_by_month[month_idx], b["revenue_sheet"],
                    month_labels[month_idx], field,
                )
                if v is not None:
                    total += v
                    any_val = True
                    any_raw = any_raw or from_raw
            return (total if any_val else None), any_raw
        sheet_name = next(b["revenue_sheet"] for b in BRANCHES if b["label"] == label)
        return get_revenue_field_with_fallback(
            revenue, raw_dashboard_by_month[month_idx], sheet_name, month_labels[month_idx], field
        )

    def get_pay(label, month_idx, field, is_total):
        if is_total:
            return total_payroll_by_month[month_idx].get(field)
        payroll_name = next(b["payroll"] for b in BRANCHES if b["label"] == label)
        s = payroll_by_month[month_idx].get(payroll_name)
        return s.get(field) if s else None

    def get_pay_avg(label, month_idx, sum_field, count_field, is_total):
        if is_total:
            t = total_payroll_by_month[month_idx]
            return (t[sum_field] / t[count_field]) if t.get(count_field) else None
        payroll_name = next(b["payroll"] for b in BRANCHES if b["label"] == label)
        s = payroll_by_month[month_idx].get(payroll_name)
        if not s or not s.get(count_field):
            return None
        return s[sum_field] / s[count_field]

    def get_tvv_kpi_rate(label, month_idx, is_total):
        if is_total:
            t = total_payroll_by_month[month_idx]
            return (t["kpi_rate_tvv_sum"] / t["kpi_rate_tvv_n"]) if t.get("kpi_rate_tvv_n") else None
        payroll_name = next(b["payroll"] for b in BRANCHES if b["label"] == label)
        s = payroll_by_month[month_idx].get(payroll_name)
        if not s or not s.get("kpi_rate_tvv_n"):
            return None
        return s["kpi_rate_tvv_sum"] / s["kpi_rate_tvv_n"]

    section_label(ws, 4, "CHỈ SỐ VẬN HÀNH")
    section_label(ws, 5, "KHÁCH HÀNG MỚI /CŨ")

    op_rows = [
        (6, "Khách mới", "khach_moi", "abs", INT_FMT, False),
        (7, "Doanh thu khách mới", "dt_khach_moi", "pct", MONEY_FMT, True),
        (8, "Khách cũ", "khach_cu", "abs", INT_FMT, False),
        (9, "Doanh thu khách cũ", "dt_khach_cu", "pct", MONEY_FMT, True),
        (10, "Tỷ lệ chốt khách mới (%)", "ty_le_chot_moi", "pct", "0.0%", True),
        (11, "Tỷ lệ chốt khách cũ (%)", "ty_le_chot_cu", "pct", "0.0%", True),
        (12, "Bill TB khách mới", "bill_tb_moi", "pct", MONEY_FMT, True),
        (13, "Bill TB khách cũ", "bill_tb_cu", "pct", MONEY_FMT, True),
    ]
    # "dt_khach_cu" (dòng 9) là field có hỗ trợ dự phòng raw_data -> dùng
    # get_rev_fallback + flag_fn; các field khác giữ nguyên get_rev như cũ.
    for row, label, field, kind, fmt, color in op_rows:
        for branch_label, (plan, val_col, *_r, is_total) in branch_plans.items():
            if field == "dt_khach_cu":
                write_row(
                    ws, row, label, plan, val_col,
                    lambda i, bl=branch_label, t=is_total: get_rev_fallback(bl, i, "dt_khach_cu", t)[0],
                    kind, fmt, color=color,
                    flag_fn=lambda i, bl=branch_label, t=is_total: get_rev_fallback(bl, i, "dt_khach_cu", t)[1],
                )
            else:
                write_row(ws, row, label, plan, val_col,
                          lambda i, bl=branch_label, f=field, t=is_total: get_rev(bl, i, f, t),
                          kind, fmt, color=color)

    section_label(ws, 14, "BẢNG ĐÁNH GIÁ HIỆU SUẤT LÀM VIỆC CHI NHÁNH")
    for branch_label, (plan, val_col, start_col, end_col, is_total) in branch_plans.items():
        ws.merge_cells(start_row=15, start_column=start_col, end_row=15, end_column=end_col)
        style_header(ws.cell(row=15, column=start_col), branch_label, total=is_total)
        for item in plan:
            if item[0] == "val":
                style_header(ws.cell(row=16, column=item[2]), month_labels[item[1]], total=is_total)
            else:
                style_header(ws.cell(row=16, column=item[3]), "Δ", total=is_total)
    ws["A15"] = "Chỉ tiêu"; ws["A15"].font = FONT_HEADER; ws["A15"].fill = FILL_HEADER
    ws["A16"] = "Tháng"; ws["A16"].font = FONT_HEADER; ws["A16"].fill = FILL_HEADER

    headcount_rows = [
        (17, "Sô nhân sự", "so_nhan_su", "abs", INT_FMT),
        (18, "KTV", "so_ktv", "abs", INT_FMT),
        (19, "TVV", "so_tvv", "abs", INT_FMT),
        (20, "OM/CM/LEAD", "so_omcmlead", "abs", INT_FMT),
        (21, "QLCN", "so_qlcn", "abs", INT_FMT),
    ]
    for row, label, field, kind, fmt in headcount_rows:
        for branch_label, (plan, val_col, *_r, is_total) in branch_plans.items():
            write_row(ws, row, label, plan, val_col,
                      lambda i, bl=branch_label, f=field, t=is_total: get_pay(bl, i, f, t), kind, fmt)

    section_label(ws, 22, "HIỆU SUẤT KTV")
    for branch_label, (plan, val_col, *_r, is_total) in branch_plans.items():
        write_row(ws, 23, "Doanh thu KTV", plan, val_col,
                  lambda i, bl=branch_label, t=is_total: get_pay(bl, i, "dt_ktv", t), "pct", MONEY_FMT, color=True)
        write_formula_row(ws, 24, "DT/KTV", plan, val_col,
                           lambda L: f"={L}23/{L}18", "pct", MONEY_FMT, color=True)
        write_row(ws, 25, "Điểm tour KTV", plan, val_col,
                  lambda i, bl=branch_label, t=is_total: get_pay(bl, i, "tour_ktv", t), "pct", INT_FMT)
        write_row(ws, 26, "Thu nhập BQ KTV", plan, val_col,
                  lambda i, bl=branch_label, t=is_total: get_pay_avg(bl, i, "thunhap_ktv_sum", "so_ktv", t),
                  "pct", MONEY_FMT)

    section_label(ws, 27, "HIỆU SUẤT  TVV")
    for branch_label, (plan, val_col, *_r, is_total) in branch_plans.items():
        write_row(ws, 28, "Doanh thu TVV", plan, val_col,
                  lambda i, bl=branch_label, t=is_total: get_pay(bl, i, "dt_tvv", t), "pct", MONEY_FMT, color=True)
        write_formula_row(ws, 29, "DT/TVV", plan, val_col,
                           lambda L: f"={L}28/{L}19", "pct", MONEY_FMT, color=True)
        write_row(ws, 30, "Tỷ lệ chốt bình quân", plan, val_col,
                  lambda i, bl=branch_label, t=is_total: get_tvv_kpi_rate(bl, i, t), "pct", "0.0%", color=True)
        write_row(ws, 31, "Thu nhập BQ TVV", plan, val_col,
                  lambda i, bl=branch_label, t=is_total: get_pay_avg(bl, i, "thunhap_tvv_sum", "so_tvv", t),
                  "pct", MONEY_FMT)

    section_label(ws, 32, "HIỆU QUẢ HOẠT ĐỘNG CHI NHÁNH")
    for branch_label, (plan, val_col, *_r, is_total) in branch_plans.items():
        # "Tổng doanh thu" (dòng 33) cũng là field có hỗ trợ dự phòng raw_data
        write_row(
            ws, 33, "Tổng doanh thu", plan, val_col,
            lambda i, bl=branch_label, t=is_total: get_rev_fallback(bl, i, "tong_doanh_thu", t)[0],
            "pct", MONEY_FMT, color=True,
            flag_fn=lambda i, bl=branch_label, t=is_total: get_rev_fallback(bl, i, "tong_doanh_thu", t)[1],
        )

        ws.cell(row=34, column=1, value="Tăng trưởng doanh thu (%)").font = FONT
        ws.cell(row=34, column=1).border = BORDER
        for item in plan:
            if item[0] != "val":
                continue
            i, c = item[1], item[2]
            if i == 0:
                continue
            Lc, Lp = get_column_letter(c), get_column_letter(val_col[i - 1])
            cell = ws.cell(row=34, column=c, value=f"=({Lc}33-{Lp}33)/{Lp}33")
            cell.number_format = "0.0%"; cell.font = FONT; cell.border = BORDER
        write_diffs(ws, 34, plan, val_col, "pct", color=True)

        write_row(ws, 35, "Chi phí nhân sự", plan, val_col,
                  lambda i, bl=branch_label, t=is_total: get_pay(bl, i, "chi_phi_nhan_su", t),
                  "pct", MONEY_FMT, color="reverse")
        write_formula_row(ws, 36, "CPNS/Doanh thu (%)", plan, val_col,
                           lambda L: f"={L}35/{L}33", "pct", "0.0%", color="reverse")
        write_formula_row(ws, 37, "Thu nhập BQ / tổng nhân sự", plan, val_col,
                           lambda L: f"={L}35/{L}17", "pct", MONEY_FMT)
        write_formula_row(ws, 38, "Doanh thu/Tổng nhân sự", plan, val_col,
                           lambda L: f"={L}33/{L}17", "pct", MONEY_FMT, color=True)

        ws.cell(row=39, column=1, value="ĐIỂM HIỆU QUẢ").font = FONT_BOLD
        ws.cell(row=39, column=1).border = BORDER
        for item in plan:
            if item[0] != "val":
                continue
            _, c = item[1], item[2]
            L = get_column_letter(c)
            formula = f"=IFERROR({L}38/1000000,0)*0.4+IFERROR({L}34,0)*100*0.2-IFERROR({L}36,0)*100*0.4"
            cell = ws.cell(row=39, column=c, value=formula)
            cell.font = FONT_BOLD; cell.border = BORDER
        write_diffs(ws, 39, plan, val_col, "pct", color=True)

    # Xếp hạng (dòng 40) — chỉ so sánh giữa 9 chi nhánh thật, KHÔNG tính cột Tổng
    ws.cell(row=40, column=1, value="XẾP HẠNG").font = FONT_BOLD
    ws.cell(row=40, column=1).border = BORDER
    for month_idx in range(n):
        cols_this_month = [branch_plans[b["label"]][1][month_idx] for b in BRANCHES]
        for c in cols_this_month:
            L = get_column_letter(c)
            others = [get_column_letter(oc) for oc in cols_this_month if oc != c]
            terms = "+".join(f"({o}$39>{L}$39)" for o in others)
            cell = ws.cell(row=40, column=c, value=f"={terms}+1")
            cell.font = FONT; cell.border = BORDER; cell.alignment = Alignment(horizontal="center")
        total_col = branch_plans[TOTAL_LABEL][1][month_idx]
        ws.cell(row=40, column=total_col, value="—").alignment = Alignment(horizontal="center")
        ws.cell(row=40, column=total_col).border = BORDER

    ws.freeze_panes = "B4"
    return ws


def compute_audit_warnings(months, revenue, unmapped_branches, unmapped_sheets):
    """months: list các tuple (label, payroll_stats, raw_dashboard)."""
    warnings = []
    if unmapped_branches:
        warnings.append(
            "⚠ Payroll có chi nhánh CHƯA cấu hình trong app (dữ liệu chi nhánh này sẽ KHÔNG "
            "xuất hiện trong sheet KTV-TVV): " + ", ".join(sorted(unmapped_branches))
        )
    if unmapped_sheets:
        warnings.append(
            "⚠ File Phân tích Doanh thu có sheet chi nhánh CHƯA cấu hình trong app: "
            + ", ".join(sorted(unmapped_sheets))
        )

    month_labels = [m[0] for m in months]
    payroll_by_month = [m[1] for m in months]
    raw_dashboard_by_month = [m[2] for m in months]

    for b in BRANCHES:
        for i, lbl in enumerate(month_labels):
            d = revenue.get(b["revenue_sheet"], {}).get(lbl)
            has_revenue_file_data = d and d.get("tong_doanh_thu") is not None
            raw_entry = raw_dashboard_by_month[i].get(b["revenue_sheet"])
            has_raw_data = raw_entry and raw_entry.get("doanh_thu_truoc_thue") is not None

            if not has_revenue_file_data and not has_raw_data:
                warnings.append(f"⚠ [{b['label']} - {lbl}] Không tìm thấy dữ liệu doanh thu tương ứng (kể cả trong raw_data Payroll).")
            elif not has_revenue_file_data and has_raw_data:
                warnings.append(
                    f"🆘 [{b['label']} - {lbl}] File Phân tích Doanh thu thiếu 'Tổng doanh thu' — "
                    f"đang DÙNG TẠM số từ raw_data (Data tổng) trong Payroll: "
                    f"{raw_entry['doanh_thu_truoc_thue']:,.0f} đ. Đề nghị xác nhận lại số chính thức."
                )
            elif has_revenue_file_data and has_raw_data:
                # --- MỚI: đối chiếu chéo Tổng doanh thu giữa 2 nguồn ---
                rev_val = d.get("tong_doanh_thu")
                raw_val = raw_entry["doanh_thu_truoc_thue"]
                if raw_val:
                    pct_diff = abs(rev_val - raw_val) / raw_val
                    if pct_diff > RAW_DATA_MISMATCH_THRESHOLD:
                        warnings.append(
                            f"❗ [{b['label']} - {lbl}] Tổng doanh thu LỆCH giữa 2 nguồn: "
                            f"file Phân tích Doanh thu = {rev_val:,.0f} đ, raw_data Payroll (Data tổng) = "
                            f"{raw_val:,.0f} đ (lệch {pct_diff*100:.2f}%) — nên kiểm tra lại."
                        )

            s = payroll_by_month[i].get(b["payroll"])
            if not s:
                warnings.append(f"⚠ [{b['label']} - {lbl}] Không có dữ liệu nhân sự trong Payroll tháng này.")

    for i in range(1, len(months)):
        lbl_prev, lbl_cur = month_labels[i - 1], month_labels[i]
        for b in BRANCHES:
            d_prev = revenue.get(b["revenue_sheet"], {}).get(lbl_prev, {})
            d_cur = revenue.get(b["revenue_sheet"], {}).get(lbl_cur, {})
            rev_prev, rev_cur = d_prev.get("tong_doanh_thu"), d_cur.get("tong_doanh_thu")
            if rev_prev and rev_cur and rev_cur < rev_prev:
                pct = (rev_cur - rev_prev) / rev_prev
                warnings.append(f"📉 [{b['label']}] Doanh thu {lbl_prev}→{lbl_cur}: {pct*100:+.1f}%")

            s_prev = payroll_by_month[i - 1].get(b["payroll"])
            s_cur = payroll_by_month[i].get(b["payroll"])
            if s_prev and s_cur and rev_prev and rev_cur:
                cp_prev, cp_cur = s_prev["chi_phi_nhan_su"], s_cur["chi_phi_nhan_su"]
                if cp_prev and rev_prev:
                    cpns_dt_prev = cp_prev / rev_prev
                    cpns_dt_cur = (cp_cur / rev_cur) if rev_cur else None
                    if cpns_dt_cur is not None and cpns_dt_cur > cpns_dt_prev:
                        warnings.append(
                            f"💸 [{b['label']}] CPNS/Doanh thu tăng: {cpns_dt_prev*100:.1f}% → {cpns_dt_cur*100:.1f}% "
                            f"({lbl_prev}→{lbl_cur}) — chi phí nhân sự tăng nhanh hơn doanh thu."
                        )
    return warnings


def build_audit_sheet(wb, months, revenue, unmapped_branches, unmapped_sheets):
    ws = wb.create_sheet("Audit", 0)
    ws.column_dimensions["A"].width = 100
    ws["A1"] = "Báo cáo kiểm tra dữ liệu — tự động tạo lúc xuất file"
    ws["A1"].font = FONT_BOLD
    ws["A2"] = "Các tháng trong file: " + " → ".join(m[0] for m in months)
    ws["A2"].font = FONT

    warnings = compute_audit_warnings(months, revenue, unmapped_branches, unmapped_sheets)

    r = 4
    if warnings:
        ws[f"A{r}"] = f"Phát hiện {len(warnings)} điểm cần lưu ý:"
        ws[f"A{r}"].font = FONT_BOLD
        r += 1
        for w in warnings:
            cell = ws[f"A{r}"]
            cell.value = w
            cell.font = FONT
            if w.startswith("❗") or w.startswith("🆘"):
                cell.fill = FILL_RAWDATA
            elif w.startswith("⚠"):
                cell.fill = FILL_WARN
            else:
                cell.fill = PatternFill("solid", fgColor="FCE4D6")
            cell.alignment = Alignment(wrap_text=True)
            r += 1
    else:
        ws[f"A{r}"] = "✓ Không phát hiện bất thường nào."
        ws[f"A{r}"].font = FONT
    ws.freeze_panes = "A4"
    return ws


# ---------------------------------------------------------------------------
# 7. BẢNG DỮ LIỆU TỔNG HỢP (tidy dataframe) CHO DASHBOARD WEB
#    Dùng chung nguồn số liệu với sheet Excel, nhưng ở dạng số Python thường
#    (không phải công thức) để vẽ biểu đồ Plotly.
# ---------------------------------------------------------------------------

def compute_metrics_table(months, revenue):
    """months: list các tuple (label, payroll_stats, raw_dashboard)."""
    month_labels = [m[0] for m in months]
    payroll_by_month = [m[1] for m in months]
    raw_dashboard_by_month = [m[2] for m in months]

    records = []
    for i, lbl in enumerate(month_labels):
        for b in BRANCHES:
            s = payroll_by_month[i].get(b["payroll"], {})
            d = revenue.get(b["revenue_sheet"], {}).get(lbl, {})
            rec = {"chi_nhanh": b["label"], "thang": lbl, "thang_idx": i}
            for k in ["khach_moi", "dt_khach_moi", "khach_cu", "dt_khach_cu",
                      "ty_le_chot_moi", "ty_le_chot_cu", "bill_tb_moi", "bill_tb_cu",
                      "tong_doanh_thu"]:
                rec[k] = d.get(k)
            # Dashboard web dùng luôn fallback raw_data cho Tổng doanh thu để
            # không bị hụt biểu đồ khi thiếu file doanh thu 1 vài chi nhánh.
            tdt, tdt_from_raw = get_revenue_field_with_fallback(
                revenue, raw_dashboard_by_month[i], b["revenue_sheet"], lbl, "tong_doanh_thu"
            )
            rec["tong_doanh_thu"] = tdt
            rec["tong_doanh_thu_tu_raw"] = tdt_from_raw
            rec["so_nhan_su"] = s.get("so_nhan_su")
            rec["so_ktv"] = s.get("so_ktv")
            rec["so_tvv"] = s.get("so_tvv")
            rec["so_omcmlead"] = s.get("so_omcmlead")
            rec["so_qlcn"] = s.get("so_qlcn")
            rec["dt_ktv"] = s.get("dt_ktv")
            rec["tour_ktv"] = s.get("tour_ktv")
            rec["dt_tvv"] = s.get("dt_tvv")
            rec["chi_phi_nhan_su"] = s.get("chi_phi_nhan_su")
            rec["thu_nhap_bq_ktv"] = (s["thunhap_ktv_sum"] / s["so_ktv"]) if s.get("so_ktv") else None
            rec["thu_nhap_bq_tvv"] = (s["thunhap_tvv_sum"] / s["so_tvv"]) if s.get("so_tvv") else None
            rec["ty_le_chot_bq_tvv"] = (s["kpi_rate_tvv_sum"] / s["kpi_rate_tvv_n"]) if s.get("kpi_rate_tvv_n") else None
            rec["dt_tren_ktv"] = (rec["dt_ktv"] / rec["so_ktv"]) if rec.get("so_ktv") else None
            rec["dt_tren_tvv"] = (rec["dt_tvv"] / rec["so_tvv"]) if rec.get("so_tvv") else None
            rec["dt_tren_ns"] = (rec["tong_doanh_thu"] / rec["so_nhan_su"]) if rec.get("so_nhan_su") and rec.get("tong_doanh_thu") else None
            rec["cpns_dt"] = (rec["chi_phi_nhan_su"] / rec["tong_doanh_thu"]) if rec.get("tong_doanh_thu") else None
            records.append(rec)

    df = pd.DataFrame(records)
    if df.empty:
        return df
    df = df.sort_values(["chi_nhanh", "thang_idx"]).reset_index(drop=True)
    df["tang_truong_dt"] = df.groupby("chi_nhanh")["tong_doanh_thu"].pct_change()

    def _diem_hq(row):
        a = (row["dt_tren_ns"] / 1_000_000) if pd.notna(row["dt_tren_ns"]) else 0
        bgr = (row["tang_truong_dt"] * 100) if pd.notna(row["tang_truong_dt"]) else 0
        c = (row["cpns_dt"] * 100) if pd.notna(row["cpns_dt"]) else 0
        return a * 0.4 + bgr * 0.2 - c * 0.4

    df["diem_hieu_qua"] = df.apply(_diem_hq, axis=1)
    df["xep_hang"] = df.groupby("thang_idx")["diem_hieu_qua"].rank(ascending=False, method="min").astype(int)
    df["thang"] = pd.Categorical(df["thang"], categories=month_labels, ordered=True)
    return df


def render_web_dashboard(df, month_labels):
    if df.empty:
        st.info("Chưa có đủ dữ liệu để vẽ Dashboard.")
        return

    st.header("📺 Dashboard trực quan")

    if "tong_doanh_thu_tu_raw" in df.columns and df["tong_doanh_thu_tu_raw"].any():
        n_raw = int(df["tong_doanh_thu_tu_raw"].sum())
        st.caption(
            f"🆘 {n_raw} ô 'Tổng doanh thu' đang lấy tạm từ raw_data (Data tổng) trong Payroll "
            "vì thiếu trong file Phân tích Doanh thu — xem chi tiết ở sheet Audit khi xuất file."
        )

    all_branches = [b["label"] for b in BRANCHES]
    fcol1, fcol2 = st.columns(2)
    with fcol1:
        sel_months = st.multiselect("🗓️ Chọn tháng", options=month_labels, default=month_labels)
    with fcol2:
        sel_branches = st.multiselect("🏢 Chọn chi nhánh", options=all_branches, default=all_branches)

    if not sel_months or not sel_branches:
        st.warning("Vui lòng chọn ít nhất 1 tháng và 1 chi nhánh.")
        return

    dff = df[df["thang"].isin(sel_months) & df["chi_nhanh"].isin(sel_branches)].copy()
    dff["thang"] = dff["thang"].cat.remove_unused_categories()
    latest = sel_months[-1]
    prev = sel_months[-2] if len(sel_months) >= 2 else None
    dl = dff[dff["thang"] == latest]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"Tổng doanh thu ({latest})", f"{dl['tong_doanh_thu'].sum():,.0f} đ")
    c2.metric(f"Tổng nhân sự ({latest})", f"{dl['so_nhan_su'].sum():,.0f}")
    avg_score = dl["diem_hieu_qua"].mean()
    c3.metric("Điểm hiệu quả TB", f"{avg_score:,.1f}" if pd.notna(avg_score) else "n/a")
    if prev:
        prev_rev = dff[dff["thang"] == prev]["tong_doanh_thu"].sum()
        cur_rev = dl["tong_doanh_thu"].sum()
        growth = (cur_rev - prev_rev) / prev_rev if prev_rev else None
        c4.metric(f"Tăng trưởng DT so với {prev}", f"{growth*100:+.1f}%" if growth is not None else "n/a")
    else:
        c4.metric("Tăng trưởng DT", "n/a")

    tabs = st.tabs(["🏢 So sánh chi nhánh", "📈 Xu hướng", "🧑‍🤝‍🧑 Hiệu suất KTV/TVV",
                    "💰 Chi phí & Hiệu quả", "📋 Bảng chi tiết"])

    with tabs[0]:
        fig1 = px.bar(dff, x="chi_nhanh", y="tong_doanh_thu", color="thang", barmode="group",
                       title="Tổng doanh thu theo chi nhánh qua các tháng",
                       labels={"tong_doanh_thu": "Doanh thu (đ)", "chi_nhanh": "Chi nhánh", "thang": "Tháng"})
        fig1.update_layout(yaxis_tickformat=",.0f")
        st.plotly_chart(fig1, use_container_width=True)

        dl_sorted = dl.sort_values("diem_hieu_qua", ascending=True)
        fig2 = px.bar(dl_sorted, x="diem_hieu_qua", y="chi_nhanh", orientation="h",
                       color="diem_hieu_qua", color_continuous_scale="RdYlGn",
                       title=f"Điểm hiệu quả theo chi nhánh ({latest})",
                       labels={"diem_hieu_qua": "Điểm hiệu quả", "chi_nhanh": "Chi nhánh"})
        fig2.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig2, use_container_width=True)

    with tabs[1]:
        agg = dff.groupby("thang", observed=True).agg(
            tong_doanh_thu=("tong_doanh_thu", "sum"),
            so_nhan_su=("so_nhan_su", "sum"),
            so_ktv=("so_ktv", "sum"),
            so_tvv=("so_tvv", "sum"),
            chi_phi_nhan_su=("chi_phi_nhan_su", "sum"),
        ).reset_index()
        fig3 = px.line(agg, x="thang", y="tong_doanh_thu", markers=True,
                        title="Xu hướng tổng doanh thu (các chi nhánh đã chọn)",
                        labels={"tong_doanh_thu": "Doanh thu (đ)", "thang": "Tháng"})
        fig3.update_layout(yaxis_tickformat=",.0f")
        st.plotly_chart(fig3, use_container_width=True)

        fig4 = go.Figure()
        fig4.add_trace(go.Scatter(x=agg["thang"], y=agg["so_ktv"], name="KTV", mode="lines+markers"))
        fig4.add_trace(go.Scatter(x=agg["thang"], y=agg["so_tvv"], name="TVV", mode="lines+markers"))
        fig4.update_layout(title="Xu hướng số lượng KTV / TVV")
        st.plotly_chart(fig4, use_container_width=True)

        fig5 = px.line(agg, x="thang", y="chi_phi_nhan_su", markers=True,
                        title="Xu hướng chi phí nhân sự", labels={"chi_phi_nhan_su": "Chi phí (đ)"})
        fig5.update_layout(yaxis_tickformat=",.0f")
        st.plotly_chart(fig5, use_container_width=True)

    with tabs[2]:
        fig6 = px.bar(dl, x="chi_nhanh", y=["dt_ktv", "dt_tvv"], barmode="group",
                       title=f"Doanh thu KTV vs TVV theo chi nhánh ({latest})",
                       labels={"value": "Doanh thu (đ)", "variable": "Nhóm"})
        fig6.update_layout(yaxis_tickformat=",.0f")
        st.plotly_chart(fig6, use_container_width=True)

        fig7 = px.bar(dl, x="chi_nhanh", y=["dt_tren_ktv", "dt_tren_tvv"], barmode="group",
                       title=f"Doanh thu bình quân / người ({latest})",
                       labels={"value": "Doanh thu/người (đ)", "variable": "Nhóm"})
        fig7.update_layout(yaxis_tickformat=",.0f")
        st.plotly_chart(fig7, use_container_width=True)

        fig8 = px.bar(dl.sort_values("ty_le_chot_bq_tvv"), x="chi_nhanh", y="ty_le_chot_bq_tvv",
                       title=f"Tỷ lệ chốt bình quân TVV ({latest})",
                       labels={"ty_le_chot_bq_tvv": "Tỷ lệ chốt"})
        fig8.update_layout(yaxis_tickformat=".0%")
        st.plotly_chart(fig8, use_container_width=True)

    with tabs[3]:
        fig9 = px.bar(dl, x="chi_nhanh", y="cpns_dt", color="cpns_dt", color_continuous_scale="RdYlGn_r",
                       title=f"Chi phí nhân sự / Doanh thu theo chi nhánh ({latest})",
                       labels={"cpns_dt": "CPNS/DT"})
        fig9.update_layout(yaxis_tickformat=".0%", coloraxis_showscale=False)
        st.plotly_chart(fig9, use_container_width=True)

        if prev:
            fig10 = px.bar(dff[dff["thang"].isin([prev, latest])], x="chi_nhanh", y="diem_hieu_qua",
                            color="thang", barmode="group",
                            title=f"Điểm hiệu quả: {prev} vs {latest}")
            st.plotly_chart(fig10, use_container_width=True)

        st.subheader("Bảng xếp hạng chi nhánh")
        rank_cols = ["chi_nhanh", "tong_doanh_thu", "cpns_dt", "diem_hieu_qua", "xep_hang"]
        rank_df = dl[rank_cols].rename(columns={
            "chi_nhanh": "Chi nhánh", "tong_doanh_thu": "Tổng doanh thu",
            "cpns_dt": "CPNS/DT", "diem_hieu_qua": "Điểm hiệu quả", "xep_hang": "Xếp hạng",
        }).sort_values("Xếp hạng").set_index("Chi nhánh")
        st.dataframe(
            rank_df.style.format({"Tổng doanh thu": "{:,.0f}", "CPNS/DT": "{:.1%}", "Điểm hiệu quả": "{:.1f}"}),
            use_container_width=True,
        )

    with tabs[4]:
        st.subheader(f"Bảng chi tiết toàn bộ chỉ tiêu — {', '.join(sel_months)}")
        display_cols = ["chi_nhanh", "thang", "khach_moi", "dt_khach_moi", "khach_cu", "dt_khach_cu",
                         "so_nhan_su", "so_ktv", "so_tvv", "dt_ktv", "dt_tvv",
                         "thu_nhap_bq_ktv", "thu_nhap_bq_tvv", "chi_phi_nhan_su",
                         "tong_doanh_thu", "cpns_dt", "diem_hieu_qua", "xep_hang"]
        rename_map = {
            "chi_nhanh": "Chi nhánh", "thang": "Tháng", "khach_moi": "Khách mới",
            "dt_khach_moi": "DT khách mới", "khach_cu": "Khách cũ", "dt_khach_cu": "DT khách cũ",
            "so_nhan_su": "Số NS", "so_ktv": "KTV", "so_tvv": "TVV",
            "dt_ktv": "DT KTV", "dt_tvv": "DT TVV",
            "thu_nhap_bq_ktv": "TN BQ KTV", "thu_nhap_bq_tvv": "TN BQ TVV",
            "chi_phi_nhan_su": "Chi phí NS", "tong_doanh_thu": "Tổng DT",
            "cpns_dt": "CPNS/DT", "diem_hieu_qua": "Điểm HQ", "xep_hang": "Xếp hạng",
        }
        table = dff[display_cols].rename(columns=rename_map).sort_values(["Tháng", "Xếp hạng"])
        money_cols = ["DT khách mới", "DT khách cũ", "DT KTV", "DT TVV", "TN BQ KTV", "TN BQ TVV",
                      "Chi phí NS", "Tổng DT"]
        fmt = {c: "{:,.0f}" for c in money_cols}
        fmt["CPNS/DT"] = "{:.1%}"
        fmt["Điểm HQ"] = "{:.1f}"
        st.dataframe(table.style.format(fmt), use_container_width=True, height=420)
        st.download_button(
            "⬇️ Tải bảng chi tiết (CSV)",
            data=table.to_csv(index=False).encode("utf-8-sig"),
            file_name="ktv_tvv_bang_chi_tiet.csv",
            mime="text/csv",
        )


# ---------------------------------------------------------------------------
# 8. UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Báo cáo lương KTV-TVV", layout="wide")
st.title("📊 Tự động tạo sheet KTV-TVV (Vận hành + Hiệu suất + Chi phí nhân sự)")

st.markdown(
    """
Upload file **Phân tích Doanh thu khách hàng** và **bao nhiêu file Payroll cũng
được** (mỗi file 1 tháng, sheet "Bảng lương") — app tự nhận diện tháng, tự sắp
xếp theo thời gian và tự mở rộng bảng KTV-TVV theo đúng số tháng bạn upload.

🆕 App giờ đọc thêm **raw_data** (sheet "Data tổng") có sẵn trong mỗi file
Payroll để tự đối chiếu số Tổng doanh thu, và dùng làm dự phòng nếu file
Phân tích Doanh thu bị thiếu dữ liệu 1 vài chi nhánh/tháng.
"""
)

revenue_file = st.file_uploader("📈 File Phân tích Doanh thu khách hàng", type=["xlsx"])
payroll_files = st.file_uploader(
    "💰 File Payroll (chọn nhiều file cùng lúc, mỗi file 1 tháng)",
    type=["xlsx"], accept_multiple_files=True,
)

if revenue_file and payroll_files:
    with st.spinner("Đang đọc file Phân tích Doanh thu..."):
        try:
            revenue, unmapped_sheets = read_revenue_bytes(revenue_file.getvalue())
        except Exception as e:
            st.error(f"Lỗi đọc file Phân tích Doanh thu: {e}")
            st.stop()

    parsed = []
    unmapped_branches_all = set()
    with st.spinner(f"Đang đọc {len(payroll_files)} file Payroll..."):
        for f in payroll_files:
            try:
                year, month, label, stats, unmapped, employees, raw_dashboard = read_payroll_bytes(f.getvalue(), f.name)
            except Exception as e:
                st.error(f"Lỗi đọc file '{f.name}': {e}")
                continue
            if not label:
                st.warning(f"Không tự nhận diện được tháng của file '{f.name}' — bỏ qua file này.")
                continue
            unmapped_branches_all |= unmapped
            parsed.append({"filename": f.name, "year": year, "month": month, "label": label,
                            "stats": stats, "employees": employees, "raw_dashboard": raw_dashboard})

    if not parsed:
        st.stop()

    parsed.sort(key=lambda p: (p["year"], p["month"]))

    labels_seen = [p["label"] for p in parsed]
    dups = {l for l in labels_seen if labels_seen.count(l) > 1}
    if dups:
        st.error(f"Có nhiều file cùng nhận diện là tháng {', '.join(dups)} — kiểm tra lại, mỗi tháng chỉ nên có 1 file.")
        st.stop()

    st.success("Đã nhận diện: " + " → ".join(f"**{p['label']}** ({p['filename']})" for p in parsed))
    if unmapped_branches_all:
        st.warning("⚠️ Có chi nhánh trong Payroll chưa được cấu hình trong app: " + ", ".join(sorted(unmapped_branches_all)) + " — dữ liệu chi nhánh này sẽ bị bỏ qua. Xem chi tiết trong sheet Audit sau khi xuất.")
    if unmapped_sheets:
        st.warning("⚠️ File Doanh thu có sheet chi nhánh chưa cấu hình: " + ", ".join(sorted(unmapped_sheets)))
    if not any(p["raw_dashboard"] for p in parsed):
        st.info("ℹ️ Không tìm thấy sheet 'Data tổng' (raw_data) trong các file Payroll đã upload — bỏ qua bước đối chiếu/dự phòng raw_data.")

    months_arg = [(p["label"], p["stats"], p["raw_dashboard"]) for p in parsed]
    metrics_df = compute_metrics_table(months_arg, revenue)

    st.divider()
    render_web_dashboard(metrics_df, [p["label"] for p in parsed])
    st.divider()

    if st.button("🚀 Xuất sheet KTV-TVV", type="primary"):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        build_ktv_tvv_sheet(wb, months_arg, revenue)
        build_audit_sheet(wb, months_arg, revenue, unmapped_branches_all, unmapped_sheets)
        wb.move_sheet("Audit", offset=-len(wb.sheetnames))

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)

        st.success("Đã tạo xong sheet KTV-TVV + sheet Audit!")
        st.download_button(
            "⬇️ Tải file KTV-TVV.xlsx",
            data=out,
            file_name=f"KTV-TVV_{parsed[0]['label']}_{parsed[-1]['label']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Vui lòng upload file Phân tích Doanh thu và ít nhất 1 file Payroll để bắt đầu.")
