"""
Streamlit app: Tự động tạo sheet "KTV-TVV" TRỰC TIẾP TỪ RAW DASHBOARD doanh thu
(nhiều file, mỗi file 1 tháng) + file Payroll — KHÔNG cần file trung gian
"Phân tích Doanh thu khách hàng" đã xử lý sẵn (vì file đó chưa làm).

SO VỚI BẢN TRƯỚC:
  ✅ GỘP thẳng phần đọc "raw dashboard doanh thu" (read_raw_dashboard,
     detect_month, build_branch_values, sanity_check_single_month...) từ app
     "Phân tích Doanh thu khách hàng" vào ngay trong app này. App tự tính các
     chỉ tiêu doanh thu (Khách mới/cũ, Doanh thu, Tỷ lệ chốt, Bill TB, Tổng
     doanh thu...) từ raw dashboard, KHÔNG cần upload file Excel trung gian
     nữa.
  ❌ BỎ HẲN phần đọc/đối chiếu/dự phòng "raw_data" (sheet "Data tổng") có
     trong file Payroll — vì đó chỉ là giải pháp tạm khi chưa có số doanh thu
     chính thức; giờ doanh thu đã được tính trực tiếp từ raw dashboard nên
     không cần đối chiếu/dự phòng chéo với Payroll nữa.

Cách chạy:
    pip install streamlit openpyxl pandas plotly
    streamlit run app.py
"""

import io
import re
from datetime import date

import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# ============================================================================
# 0. CẤU HÌNH CHI NHÁNH & NHÃN CHỈ TIÊU DÙNG CHUNG
# ============================================================================

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

# Nhãn chỉ tiêu doanh thu dùng chung giữa raw dashboard và bảng KTV-TVV.
# Đây chính là "cầu nối" giúp gộp thẳng raw dashboard vào, không cần file
# trung gian: TEMPLATE_ROWS (raw dashboard) và REVENUE_ROW_MAP (KTV-TVV) dùng
# đúng cùng 1 bộ nhãn tiếng Việt nên map trực tiếp theo nhãn.
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
FONT_GOOD = Font(name="Arial", size=11, color="006100")
FONT_BAD = Font(name="Arial", size=11, color="9C0006")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = "#,##0"
INT_FMT = "#,##0"
PCT_FMT = "0.0%"


# ============================================================================
# 1. ĐỌC RAW DASHBOARD DOANH THU (gộp từ app "Phân tích Doanh thu khách hàng")
#    Không còn cần file trung gian — app tự tính chỉ tiêu doanh thu từ đây.
# ============================================================================

RAW_COLS = {
    "kpi": "KPI",
    "doanh_thu": "Doanh thu",
    "khach_moi": "Khách mới",
    "mua_tt_km": "Mua TT KM",
    "dt_khach_moi_all": "DT khách mới (all)",
    "bill_tb_km": "Bill TB KM",
    "dt_khach_moi_30d": "DT khách mới 30D",
    "booking_moi": "Booking Mới",
    "checkin_moi": "Checkin Mới",
    "khach_thuc_te": "Khách thực tế",   # dạng "30838 (3434 / 1361 / 26043)"
    "mua_tt_kc": "Mua TT KC",
    "dt_khach_cu": "DT khách cũ",
    "bill_tb_mua_kc": "Bill TB mua KC",
}

# (dòng gốc trong file Excel cũ, nhãn, loại, key trong RAW_COLS) — nhãn ở đây
# khớp 1-1 với REVENUE_ROW_MAP ở trên, dùng để map trực tiếp.
TEMPLATE_ROWS = [
    (2, "KPI", "raw", "kpi"),
    (3, "Tổng doanh thu", "raw", "doanh_thu"),
    (4, "Khách mới", "raw", "khach_moi"),
    (5, "Khách mua hàng TT (mới)", "raw", "mua_tt_km"),
    (6, "Tỷ lệ chốt khách mới", "formula_moi", None),
    (7, "Doanh thu khách mới", "raw", "dt_khach_moi_all"),
    (8, "Bill TB khách mới", "raw", "bill_tb_km"),
    (9, "Doanh thu khách mới 30 ngày", "raw", "dt_khach_moi_30d"),
    (10, "Bill TB khách mới 30 ngày", "missing", None),
    (11, "Khách booking mới", "raw", "booking_moi"),
    (12, "Khách checkin mới", "raw", "checkin_moi"),
    (13, "Khách thực tế (cũ)", "khach_cu", None),
    (14, "Khách mua hàng TT (cũ)", "raw", "mua_tt_kc"),
    (15, "Tỷ lệ chốt khách cũ", "formula_cu", None),
    (16, "Doanh thu khách cũ", "raw", "dt_khach_cu"),
    (17, "Bill TB khách cũ", "raw", "bill_tb_mua_kc"),
]

# Chi nhánh dùng KPI/doanh thu thật — loại các dòng hành chính không tính KPI
EXCLUDE_BRANCHES = {"Học Viện LGS", "Văn Phòng"}


def to_num(v):
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        vv = v.replace(",", "").replace("\xa0", "").strip()
        try:
            return float(vv)
        except ValueError:
            return None
    return None


def parse_khach_cu(v):
    if v is None:
        return None
    nums = re.findall(r"-?\d+", str(v))
    if len(nums) >= 4:
        return int(nums[3])
    return None


def read_raw_dashboard(file):
    """Trả về (title, {tên chi nhánh: {tên cột: giá trị}})."""
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb.active

    title = ws.cell(row=1, column=1).value or ""
    header_row_idx = None
    for r in range(1, 6):
        cell = ws.cell(row=r, column=1).value
        if cell and str(cell).strip() == "Chi nhánh":
            header_row_idx = r
            break
    if header_row_idx is None:
        raise ValueError("Không tìm thấy dòng tiêu đề 'Chi nhánh' trong file.")

    headers = [c.value for c in ws[header_row_idx]]
    data = {}
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        name = row[0]
        if not name:
            continue
        data[str(name).strip()] = dict(zip(headers, row))
    return title, data


def detect_month(title: str, filename: str):
    """Cố tìm ngày bắt đầu 'dd-mm-yyyy' trong tiêu đề hoặc tên file.
    Trả về (year, month, ngay_bat_dau) hoặc (None, None, None) nếu không tìm được."""
    text = f"{title} {filename}"
    m = re.search(r"(\d{2})-(\d{2})-(\d{4})", text)
    if m:
        dd, mm, yyyy = m.groups()
        try:
            return int(yyyy), int(mm), date(int(yyyy), int(mm), int(dd))
        except ValueError:
            pass
    return None, None, None


def build_branch_values(branch_row: dict) -> dict:
    """Trả về {row_idx (theo TEMPLATE_ROWS): giá trị} cho 1 chi nhánh."""
    raw_vals = {k: branch_row.get(colname) for k, colname in RAW_COLS.items()}
    out = {}
    for row_idx, _label, kind, key in TEMPLATE_ROWS:
        if kind == "raw":
            out[row_idx] = to_num(raw_vals[key])
        elif kind == "khach_cu":
            out[row_idx] = parse_khach_cu(raw_vals["khach_thuc_te"])
        else:
            out[row_idx] = None
    return out


def sanity_check_single_month(label: str, raw_data: dict) -> list[str]:
    """Cảnh báo KPI trùng nhau giữa các chi nhánh trong CÙNG 1 tháng — dấu
    hiệu copy nhầm dòng trong raw dashboard."""
    warnings = []
    kpi_by_branch = {
        b: to_num(v.get("KPI"))
        for b, v in raw_data.items()
        if b not in EXCLUDE_BRANCHES and b != "Tất cả chi nhánh" and to_num(v.get("KPI"))
    }
    seen = {}
    for b, kpi in kpi_by_branch.items():
        seen.setdefault(kpi, []).append(b)
    for kpi, branches in seen.items():
        if len(branches) > 1:
            warnings.append(
                f"⚠️ [{label}] KPI giống hệt nhau ({kpi:,.0f}) giữa: "
                f"{', '.join(branches)} — khả năng copy nhầm dòng trong raw dashboard."
            )
    return warnings


_LABEL_TO_FIELD = {v: k for k, v in REVENUE_ROW_MAP.items()}


def build_revenue_entry_for_month(raw_data_for_month: dict) -> dict:
    """raw_data_for_month: {chi nhánh: {tên cột raw: giá trị}} (từ
    read_raw_dashboard, 1 tháng). Trả về {chi nhánh: {field: value}} theo
    REVENUE_ROW_MAP — cấu trúc dùng thẳng cho bảng KTV-TVV bên dưới, không
    qua file Excel trung gian nữa."""
    out = {}
    for branch, row in raw_data_for_month.items():
        if branch in EXCLUDE_BRANCHES or branch == "Tất cả chi nhánh":
            continue
        computed = build_branch_values(row)  # {row_idx: value}
        vals = {}
        for row_idx, label, _kind, _key in TEMPLATE_ROWS:
            field = _LABEL_TO_FIELD.get(label)
            if field is None:
                continue
            vals[field] = computed.get(row_idx)
        # 2 dòng này vốn là công thức Excel (=mua_tt/khach) trong file cũ —
        # tính thẳng bằng Python vì giờ không xuất qua file trung gian nữa.
        vals["ty_le_chot_moi"] = (
            vals["mua_tt_moi"] / vals["khach_moi"]
            if vals.get("khach_moi") else None
        )
        vals["ty_le_chot_cu"] = (
            vals["mua_tt_cu"] / vals["khach_cu"]
            if vals.get("khach_cu") else None
        )
        out[branch] = vals
    return out


def build_revenue_dict(months_raw: list[dict]):
    """months_raw: [{'label': str, 'raw_data': {...}}, ...] đã sắp xếp theo
    thời gian và đã xác nhận tên cột (tháng). Trả về:
      revenue = {chi nhánh: {label: {field: value}}}
      unmapped_branches = set các chi nhánh xuất hiện trong raw dashboard mà
        chưa được cấu hình trong BRANCHES ở trên.
    """
    covered = {b["revenue_sheet"] for b in BRANCHES}
    revenue = {}
    unmapped_branches = set()
    for m in months_raw:
        entry = build_revenue_entry_for_month(m["raw_data"])
        for branch, vals in entry.items():
            if branch not in covered:
                unmapped_branches.add(branch)
            revenue.setdefault(branch, {})[m["label"]] = vals
    return revenue, unmapped_branches


# ============================================================================
# 2. ĐỌC FILE PAYROLL — read_only + cache để nhanh
#    (đã bỏ hoàn toàn phần đọc/đối chiếu "raw_data" / sheet "Data tổng")
# ============================================================================

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
    """Đọc sheet 'Bảng lương' ở chế độ read_only. Trả về:
      (year, month, label, {branch: stats}, unmapped_branches_set, employees)
    employees = list các dict {chi_nhanh, ten, vi_tri, nhom, doanh_thu,
    tour, thu_nhap, ty_le_kpi} — dùng để làm bảng xếp hạng top nhân viên."""
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

    return year, month, label, stats, unmapped_branches, employees


# ============================================================================
# 3. TỔNG HỢP "TẤT CẢ CHI NHÁNH" (đúng trọng số, không phải trung bình cộng)
# ============================================================================

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


# ============================================================================
# 4. LAYOUT CỘT — hỗ trợ N tháng bất kỳ
# ============================================================================

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


def write_values(ws, row, label, plan, get_value_fn, number_format, bold=False):
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


def write_row(ws, row, label, plan, val_col, get_value_fn, diff_kind, number_format, bold=False, color=False):
    write_values(ws, row, label, plan, get_value_fn, number_format, bold=bold)
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


# ============================================================================
# 5. XÂY DỰNG SHEET "KTV-TVV"
# ============================================================================

def build_ktv_tvv_sheet(wb, months, revenue):
    """months: list các tuple (label, payroll_stats)."""
    n = len(months)
    month_labels = [m[0] for m in months]
    payroll_by_month = [m[1] for m in months]
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
    for row, label, field, kind, fmt, color in op_rows:
        for branch_label, (plan, val_col, *_r, is_total) in branch_plans.items():
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
        write_row(ws, 33, "Tổng doanh thu", plan, val_col,
                  lambda i, bl=branch_label, t=is_total: get_rev(bl, i, "tong_doanh_thu", t),
                  "pct", MONEY_FMT, color=True)

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


def compute_audit_warnings(months, revenue, unmapped_branches_payroll, unmapped_branches_revenue):
    """months: list các tuple (label, payroll_stats)."""
    warnings = []
    if unmapped_branches_payroll:
        warnings.append(
            "⚠ Payroll có chi nhánh CHƯA cấu hình trong app (dữ liệu chi nhánh này sẽ KHÔNG "
            "xuất hiện trong sheet KTV-TVV): " + ", ".join(sorted(unmapped_branches_payroll))
        )
    if unmapped_branches_revenue:
        warnings.append(
            "⚠ Raw dashboard doanh thu có chi nhánh CHƯA cấu hình trong app: "
            + ", ".join(sorted(unmapped_branches_revenue))
        )

    month_labels = [m[0] for m in months]
    payroll_by_month = [m[1] for m in months]

    for b in BRANCHES:
        for lbl in month_labels:
            d = revenue.get(b["revenue_sheet"], {}).get(lbl)
            has_revenue_data = d and d.get("tong_doanh_thu") is not None
            if not has_revenue_data:
                warnings.append(f"⚠ [{b['label']} - {lbl}] Không tìm thấy dữ liệu doanh thu tương ứng trong raw dashboard.")

    for i, lbl in enumerate(month_labels):
        for b in BRANCHES:
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


def build_audit_sheet(wb, months, revenue, unmapped_branches_payroll, unmapped_branches_revenue):
    ws = wb.create_sheet("Audit", 0)
    ws.column_dimensions["A"].width = 100
    ws["A1"] = "Báo cáo kiểm tra dữ liệu — tự động tạo lúc xuất file"
    ws["A1"].font = FONT_BOLD
    ws["A2"] = "Các tháng trong file: " + " → ".join(m[0] for m in months)
    ws["A2"].font = FONT

    warnings = compute_audit_warnings(months, revenue, unmapped_branches_payroll, unmapped_branches_revenue)

    r = 4
    if warnings:
        ws[f"A{r}"] = f"Phát hiện {len(warnings)} điểm cần lưu ý:"
        ws[f"A{r}"].font = FONT_BOLD
        r += 1
        for w in warnings:
            cell = ws[f"A{r}"]
            cell.value = w
            cell.font = FONT
            if w.startswith("⚠"):
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


# ============================================================================
# 6. BẢNG DỮ LIỆU TỔNG HỢP (tidy dataframe) CHO DASHBOARD WEB
# ============================================================================

def compute_metrics_table(months, revenue):
    """months: list các tuple (label, payroll_stats)."""
    month_labels = [m[0] for m in months]
    payroll_by_month = [m[1] for m in months]

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


# ============================================================================
# 7. UI
# ============================================================================

st.set_page_config(page_title="Báo cáo lương KTV-TVV", layout="wide")
st.title("📊 Tự động tạo sheet KTV-TVV — từ RAW DASHBOARD + Payroll")

st.markdown(
    """
Upload **raw dashboard doanh thu** (bao nhiêu file/tháng cũng được) và
**file Payroll** (bao nhiêu file cũng được, mỗi file 1 tháng, sheet "Bảng
lương"). App tự tính toàn bộ chỉ tiêu doanh thu trực tiếp từ raw dashboard
(không cần file "Phân tích Doanh thu khách hàng" làm sẵn), tự nhận diện
tháng, tự sắp xếp theo thời gian, và tự mở rộng bảng KTV-TVV theo đúng số
tháng bạn upload.
"""
)

st.subheader("Bước 1 — Upload raw dashboard doanh thu (nhiều file, mỗi file 1 tháng)")
raw_revenue_files = st.file_uploader(
    "Raw dashboard doanh thu",
    type=["xlsx"],
    accept_multiple_files=True,
    key="raw_revenue_files",
)

revenue = {}
unmapped_branches_revenue = set()
months_raw_confirmed = False

if raw_revenue_files:
    parsed_rev = []
    for f in raw_revenue_files:
        try:
            title, raw_data = read_raw_dashboard(f)
        except Exception as e:
            st.error(f"Lỗi đọc file '{f.name}': {e}")
            continue
        year, month, start_date = detect_month(title, f.name)
        default_label = f"T{month}" if month else f.name
        parsed_rev.append({
            "filename": f.name, "title": title, "raw_data": raw_data,
            "year": year, "month": month, "start_date": start_date,
            "default_label": default_label,
        })

    if parsed_rev:
        st.caption(
            "App tự đoán tên tháng từ tiêu đề file — sửa lại nếu cần (vd 2 file "
            "cùng là 'T1' nhưng khác năm thì nên sửa thành 'T1/26', 'T1/27'...)."
        )
        parsed_rev.sort(key=lambda p: (p["start_date"] is None, p["start_date"]))

        rev_labels = []
        for i, p in enumerate(parsed_rev):
            cols = st.columns([3, 2, 3])
            cols[0].write(f"📄 {p['filename']}")
            cols[1].write(p["title"][:40] + ("..." if len(p["title"]) > 40 else ""))
            label = cols[2].text_input("Tên cột (tháng)", value=p["default_label"], key=f"rev_label_{i}")
            rev_labels.append(label)

        if len(set(rev_labels)) != len(rev_labels):
            st.error("❌ Có 2 file raw dashboard đang trùng tên tháng — vui lòng sửa lại cho khác nhau.")
        else:
            months_raw = [{"label": rev_labels[i], "raw_data": p["raw_data"]} for i, p in enumerate(parsed_rev)]
            months_raw_confirmed = True

            sanity_warnings = []
            for i, p in enumerate(parsed_rev):
                sanity_warnings.extend(sanity_check_single_month(rev_labels[i], p["raw_data"]))
            if sanity_warnings:
                for w in sanity_warnings:
                    st.warning(w)

            revenue, unmapped_branches_revenue = build_revenue_dict(months_raw)
else:
    st.info("Vui lòng upload ít nhất 1 file raw dashboard doanh thu.")

st.divider()
st.subheader("Bước 2 — Upload file Payroll (nhiều file, mỗi file 1 tháng)")
payroll_files = st.file_uploader(
    "File Payroll",
    type=["xlsx"], accept_multiple_files=True, key="payroll_files",
)

if months_raw_confirmed and payroll_files:
    parsed = []
    unmapped_branches_payroll = set()
    with st.spinner(f"Đang đọc {len(payroll_files)} file Payroll..."):
        for f in payroll_files:
            try:
                year, month, label, stats, unmapped, employees = read_payroll_bytes(f.getvalue(), f.name)
            except Exception as e:
                st.error(f"Lỗi đọc file '{f.name}': {e}")
                continue
            if not label:
                st.warning(f"Không tự nhận diện được tháng của file '{f.name}' — bỏ qua file này.")
                continue
            unmapped_branches_payroll |= unmapped
            parsed.append({"filename": f.name, "year": year, "month": month, "label": label,
                            "stats": stats, "employees": employees})

    if not parsed:
        st.stop()

    parsed.sort(key=lambda p: (p["year"], p["month"]))

    labels_seen = [p["label"] for p in parsed]
    dups = {l for l in labels_seen if labels_seen.count(l) > 1}
    if dups:
        st.error(f"Có nhiều file Payroll cùng nhận diện là tháng {', '.join(dups)} — kiểm tra lại, mỗi tháng chỉ nên có 1 file.")
        st.stop()

    st.success("Đã nhận diện Payroll: " + " → ".join(f"**{p['label']}** ({p['filename']})" for p in parsed))
    if unmapped_branches_payroll:
        st.warning("⚠️ Có chi nhánh trong Payroll chưa được cấu hình trong app: " + ", ".join(sorted(unmapped_branches_payroll)) + " — dữ liệu chi nhánh này sẽ bị bỏ qua. Xem chi tiết trong sheet Audit sau khi xuất.")
    if unmapped_branches_revenue:
        st.warning("⚠️ Có chi nhánh trong raw dashboard doanh thu chưa được cấu hình trong app: " + ", ".join(sorted(unmapped_branches_revenue)))

    months_arg = [(p["label"], p["stats"]) for p in parsed]
    metrics_df = compute_metrics_table(months_arg, revenue)

    st.divider()
    render_web_dashboard(metrics_df, [p["label"] for p in parsed])
    st.divider()

    if st.button("🚀 Xuất sheet KTV-TVV", type="primary"):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        build_ktv_tvv_sheet(wb, months_arg, revenue)
        build_audit_sheet(wb, months_arg, revenue, unmapped_branches_payroll, unmapped_branches_revenue)
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
elif months_raw_confirmed:
    st.info("Vui lòng upload ít nhất 1 file Payroll để tiếp tục.")
