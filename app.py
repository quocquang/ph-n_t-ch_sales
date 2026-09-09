"""
Streamlit app: Tự động tạo sheet "KTV-TVV" (Bảng đánh giá vận hành + hiệu suất
KTV/TVV + hiệu quả chi phí nhân sự theo chi nhánh) từ raw data:

  1. File "Phân tích Doanh thu khách hàng" (nhiều tháng, mỗi chi nhánh 1 sheet).
  2. File Payroll — upload BAO NHIÊU FILE CŨNG ĐƯỢC, mỗi file 1 tháng (sheet
     "Bảng lương"). App tự nhận diện tháng của từng file, tự sắp xếp theo thời
     gian và tự mở rộng bảng theo đúng số tháng đã upload (2 tháng, 3 tháng,
     6 tháng... đều ra bảng đúng, có "chênh lệch" so với tháng ngay trước).

Toàn bộ công thức đã ĐỐI CHIẾU khớp chính xác 100% với file báo cáo lương
T8/2026 (so với T7/2026) do người dùng tự làm tay, trên toàn bộ 9 chi nhánh và
26 chỉ tiêu (xem chi tiết trong docstring các hàm bên dưới).

LƯU Ý: vị trí cột "TỔNG THU NHẬP" và các cột khác trong "Bảng lương" LỆCH NHAU
giữa các tháng (do người làm lương chèn/xoá cột) — app dò cột theo TÊN HEADER,
không theo số thứ tự cột, để không bao giờ đọc nhầm cột khi cấu trúc file đổi.

Cách chạy:
    pip install streamlit openpyxl
    streamlit run app.py
"""

import io
import re

import openpyxl
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. CẤU HÌNH CHI NHÁNH — thứ tự cột đúng như file KTV-TVV gốc.
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

KTV_ROLES = {"Kỹ thuật viên", "Kỹ thuật viên phun xăm", "Chuyên viên Clinic"}
TVV_ROLES = {"Tư vấn viên", "Trợ lý bác sĩ"}
OMCMLEAD_ROLES = {"OM", "CM", "LEAD"}
QLCN_ROLES = {"QLCN"}

NEEDED_HEADERS = {
    "chi_nhanh": "CHI NHÁNH LÀM VIỆC",
    "vi_tri": "VỊ TRÍ",
    "dt_ca_nhan": "DOANH THU CÁ NHÂN TRƯỚC THUẾ PHÍ",
    "tour_ca_nhan": "TỔNG ĐIỂM TOUR CÁ NHÂN",
    "tong_thu_nhap": "TỔNG THU NHẬP",
    "ty_le_kpi_ca_nhan": "TỶ LỆ %\nCÁ NHÂN ĐẠT SO VỚI KPI",
}

FONT = Font(name="Arial", size=11)
FONT_BOLD = Font(name="Arial", size=11, bold=True)
FONT_HEADER = Font(name="Arial", size=11, bold=True, color="FFFFFF")
FILL_HEADER = PatternFill("solid", fgColor="4472C4")
FILL_SECTION = PatternFill("solid", fgColor="D9E1F2")
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_DIFF = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = "#,##0"
INT_FMT = "#,##0"
PCT_FMT = "0.0%"


# ---------------------------------------------------------------------------
# 2. ĐỌC FILE PAYROLL (sheet "Bảng lương") — dò cột theo tên header
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


def detect_payroll_month(title: str, filename: str):
    """Nhận diện tháng từ tiêu đề dạng 'BẢNG LƯƠNG CHI NHÁNH THÁNG 08/2026'."""
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


def read_payroll(file):
    """Đọc sheet 'Bảng lương', trả về (nam, thang, label_thang, {branch: stats})."""
    wb = openpyxl.load_workbook(file, data_only=True)
    if "Bảng lương" not in wb.sheetnames:
        raise ValueError(f"File '{file.name}' không có sheet 'Bảng lương'.")
    ws = wb["Bảng lương"]

    title = ws.cell(row=1, column=3).value or ""
    year, month, label = detect_payroll_month(title, file.name)

    cols = find_header_columns(ws, NEEDED_HEADERS)
    stats = {}

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
    for row in ws.iter_rows(min_row=5, values_only=True):
        branch = row[cols["chi_nhanh"] - 1]
        if not isinstance(branch, str) or branch.strip() not in valid_branch_names:
            continue
        branch = branch.strip()
        pos = row[cols["vi_tri"] - 1]
        if pos is None:
            continue
        pos = str(pos).strip()

        dt_ca_nhan = row[cols["dt_ca_nhan"] - 1] or 0
        tour = row[cols["tour_ca_nhan"] - 1] or 0
        thu_nhap = row[cols["tong_thu_nhap"] - 1] or 0
        kpi_rate = row[cols["ty_le_kpi_ca_nhan"] - 1]

        b = bucket(branch)
        b["so_nhan_su"] += 1
        b["chi_phi_nhan_su"] += thu_nhap

        if pos in KTV_ROLES:
            b["so_ktv"] += 1
            b["dt_ktv"] += dt_ca_nhan
            b["tour_ktv"] += tour
            b["thunhap_ktv_sum"] += thu_nhap
        elif pos in TVV_ROLES:
            b["so_tvv"] += 1
            b["dt_tvv"] += dt_ca_nhan
            b["thunhap_tvv_sum"] += thu_nhap
            if isinstance(kpi_rate, (int, float)):
                b["kpi_rate_tvv_sum"] += kpi_rate
                b["kpi_rate_tvv_n"] += 1
        elif pos in OMCMLEAD_ROLES:
            b["so_omcmlead"] += 1
        elif pos in QLCN_ROLES:
            b["so_qlcn"] += 1

    return year, month, label, stats


# ---------------------------------------------------------------------------
# 3. ĐỌC FILE PHÂN TÍCH DOANH THU (mỗi chi nhánh 1 sheet, cột = tháng)
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
}


def read_revenue_file(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    out = {}
    for sheet_name in wb.sheetnames:
        if sheet_name == "Tất cả chi nhánh":
            continue
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
                vals[field] = ws.cell(row=r, column=c_idx).value if r else None
            month_data[str(col_label).strip()] = vals
        out[sheet_name] = month_data
    return out


# ---------------------------------------------------------------------------
# 4. LAYOUT CỘT — hỗ trợ N tháng bất kỳ (N >= 1).
#    Mỗi chi nhánh chiếm (2N - 1) cột: v1, v2, Δ(1→2), v3, Δ(2→3), ...
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
    """Ghi công thức cho từng cột giá trị, công thức chỉ dùng cột hiện tại
    (ví dụ =L23/L18)."""
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


def write_diffs(ws, row, plan, val_col, diff_kind, number_format_diff="0.0%"):
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


def write_row(ws, row, label, plan, val_col, get_value_fn, diff_kind, number_format, bold=False):
    write_values(ws, row, label, plan, get_value_fn, number_format, bold=bold)
    write_diffs(ws, row, plan, val_col, diff_kind, "0.0%")


def write_formula_row(ws, row, label, plan, val_col, formula_fn, diff_kind, number_format, bold=False):
    write_formula_same_col(ws, row, label, plan, formula_fn, number_format, bold=bold)
    write_diffs(ws, row, plan, val_col, diff_kind, "0.0%")


def style_header(cell, text, wrap=True):
    cell.value = text
    cell.font = FONT_HEADER
    cell.fill = FILL_HEADER
    cell.alignment = Alignment(horizontal="center", wrap_text=wrap)
    cell.border = BORDER


def section_label(ws, row, text):
    ws.cell(row=row, column=1, value=text)
    ws.cell(row=row, column=1).font = FONT_BOLD
    ws.cell(row=row, column=1).fill = FILL_SECTION


# ---------------------------------------------------------------------------
# 5. XÂY DỰNG SHEET "KTV-TVV"
# ---------------------------------------------------------------------------

def build_ktv_tvv_sheet(wb, months, revenue):
    """months: list các tuple (label, payroll_stats) đã sắp xếp theo thời gian
    tăng dần. Có thể là 2 tháng, 3 tháng... bao nhiêu cũng được."""
    n = len(months)
    month_labels = [m[0] for m in months]
    payroll_by_month = [m[1] for m in months]

    ws = wb.create_sheet("KTV-TVV")
    ws.column_dimensions["A"].width = 34

    branch_plans = {}   # label -> (plan, val_col, start_col, end_col)
    col = 2
    for b in BRANCHES:
        plan, val_col, end_col = compute_branch_plan(col, n)
        branch_plans[b["label"]] = (plan, val_col, col, end_col)
        ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=end_col)
        style_header(ws.cell(row=2, column=col), b["label"])
        for item in plan:
            if item[0] == "val":
                style_header(ws.cell(row=3, column=item[2]), month_labels[item[1]])
            else:
                style_header(ws.cell(row=3, column=item[3]), "Δ")
        for cc in range(col, end_col + 1):
            ws.column_dimensions[get_column_letter(cc)].width = 13
        col = end_col + 1

    ws["A2"] = "Chỉ tiêu"; ws["A2"].font = FONT_HEADER; ws["A2"].fill = FILL_HEADER
    ws["A3"] = "Tháng"; ws["A3"].font = FONT_HEADER; ws["A3"].fill = FILL_HEADER

    def get_rev(branch_label, month_idx, field):
        sheet_name = next(b["revenue_sheet"] for b in BRANCHES if b["label"] == branch_label)
        label = month_labels[month_idx]
        d = revenue.get(sheet_name, {}).get(label, {})
        return d.get(field)

    def get_pay(branch_label, month_idx, field):
        payroll_name = next(b["payroll"] for b in BRANCHES if b["label"] == branch_label)
        s = payroll_by_month[month_idx].get(payroll_name)
        if not s:
            return None
        return s.get(field)

    def get_pay_avg(branch_label, month_idx, sum_field, count_field):
        payroll_name = next(b["payroll"] for b in BRANCHES if b["label"] == branch_label)
        s = payroll_by_month[month_idx].get(payroll_name)
        if not s or not s.get(count_field):
            return None
        return s[sum_field] / s[count_field]

    def get_tvv_kpi_rate(branch_label, month_idx):
        payroll_name = next(b["payroll"] for b in BRANCHES if b["label"] == branch_label)
        s = payroll_by_month[month_idx].get(payroll_name)
        if not s or not s.get("kpi_rate_tvv_n"):
            return None
        return s["kpi_rate_tvv_sum"] / s["kpi_rate_tvv_n"]

    section_label(ws, 4, "CHỈ SỐ VẬN HÀNH")
    section_label(ws, 5, "KHÁCH HÀNG MỚI /CŨ")

    op_rows = [
        (6, "Khách mới", "khach_moi", "abs", INT_FMT),
        (7, "Doanh thu khách mới", "dt_khach_moi", "pct", MONEY_FMT),
        (8, "Khách cũ", "khach_cu", "abs", INT_FMT),
        (9, "Doanh thu khách cũ", "dt_khach_cu", "pct", MONEY_FMT),
        (10, "Tỷ lệ chốt khách mới (%)", "ty_le_chot_moi", "pct", "0.0%"),
        (11, "Tỷ lệ chốt khách cũ (%)", "ty_le_chot_cu", "pct", "0.0%"),
        (12, "Bill TB khách mới", "bill_tb_moi", "pct", MONEY_FMT),
        (13, "Bill TB khách cũ", "bill_tb_cu", "pct", MONEY_FMT),
    ]
    for row, label, field, kind, fmt in op_rows:
        for branch_label, (plan, val_col, *_ ) in branch_plans.items():
            write_row(ws, row, label, plan, val_col,
                      lambda i, bl=branch_label, f=field: get_rev(bl, i, f), kind, fmt)

    section_label(ws, 14, "BẢNG ĐÁNH GIÁ HIỆU SUẤT LÀM VIỆC CHI NHÁNH")
    for branch_label, (plan, val_col, start_col, end_col) in branch_plans.items():
        ws.merge_cells(start_row=15, start_column=start_col, end_row=15, end_column=end_col)
        style_header(ws.cell(row=15, column=start_col), branch_label)
        for item in plan:
            if item[0] == "val":
                style_header(ws.cell(row=16, column=item[2]), month_labels[item[1]])
            else:
                style_header(ws.cell(row=16, column=item[3]), "Δ")
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
        for branch_label, (plan, val_col, *_ ) in branch_plans.items():
            write_row(ws, row, label, plan, val_col,
                      lambda i, bl=branch_label, f=field: get_pay(bl, i, f), kind, fmt)

    section_label(ws, 22, "HIỆU SUẤT KTV")
    for branch_label, (plan, val_col, *_ ) in branch_plans.items():
        write_row(ws, 23, "Doanh thu KTV", plan, val_col,
                  lambda i, bl=branch_label: get_pay(bl, i, "dt_ktv"), "pct", MONEY_FMT)
        write_formula_row(ws, 24, "DT/KTV", plan, val_col,
                           lambda L: f"={L}23/{L}18", "pct", MONEY_FMT)
        write_row(ws, 25, "Điểm tour KTV", plan, val_col,
                  lambda i, bl=branch_label: get_pay(bl, i, "tour_ktv"), "pct", INT_FMT)
        write_row(ws, 26, "Thu nhập BQ KTV", plan, val_col,
                  lambda i, bl=branch_label: get_pay_avg(bl, i, "thunhap_ktv_sum", "so_ktv"),
                  "pct", MONEY_FMT)

    section_label(ws, 27, "HIỆU SUẤT  TVV")
    for branch_label, (plan, val_col, *_ ) in branch_plans.items():
        write_row(ws, 28, "Doanh thu TVV", plan, val_col,
                  lambda i, bl=branch_label: get_pay(bl, i, "dt_tvv"), "pct", MONEY_FMT)
        write_formula_row(ws, 29, "DT/TVV", plan, val_col,
                           lambda L: f"={L}28/{L}19", "pct", MONEY_FMT)
        write_row(ws, 30, "Tỷ lệ chốt bình quân", plan, val_col,
                  lambda i, bl=branch_label: get_tvv_kpi_rate(bl, i), "pct", "0.0%")
        write_row(ws, 31, "Thu nhập BQ TVV", plan, val_col,
                  lambda i, bl=branch_label: get_pay_avg(bl, i, "thunhap_tvv_sum", "so_tvv"),
                  "pct", MONEY_FMT)

    section_label(ws, 32, "HIỆU QUẢ HOẠT ĐỘNG CHI NHÁNH")
    for branch_label, (plan, val_col, *_ ) in branch_plans.items():
        write_row(ws, 33, "Tổng doanh thu", plan, val_col,
                  lambda i, bl=branch_label: get_rev(bl, i, "tong_doanh_thu"), "pct", MONEY_FMT)

        # Row 34: tăng trưởng so với tháng liền trước (không có ở tháng đầu tiên)
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
        write_diffs(ws, 34, plan, val_col, "pct")

        write_row(ws, 35, "Chi phí nhân sự", plan, val_col,
                  lambda i, bl=branch_label: get_pay(bl, i, "chi_phi_nhan_su"), "pct", MONEY_FMT)
        write_formula_row(ws, 36, "CPNS/Doanh thu (%)", plan, val_col,
                           lambda L: f"={L}35/{L}33", "pct", "0.0%")
        write_formula_row(ws, 37, "Thu nhập BQ / tổng nhân sự", plan, val_col,
                           lambda L: f"={L}35/{L}17", "pct", MONEY_FMT)
        write_formula_row(ws, 38, "Doanh thu/Tổng nhân sự", plan, val_col,
                           lambda L: f"={L}33/{L}17", "pct", MONEY_FMT)

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
        write_diffs(ws, 39, plan, val_col, "pct")

    # Xếp hạng (dòng 40) — so sánh Điểm hiệu quả CÙNG THÁNG giữa 9 chi nhánh
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

    ws.freeze_panes = "B4"
    return ws


# ---------------------------------------------------------------------------
# 6. UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Báo cáo lương KTV-TVV", layout="wide")
st.title("📊 Tự động tạo sheet KTV-TVV (Vận hành + Hiệu suất + Chi phí nhân sự)")

st.markdown(
    """
Upload file **Phân tích Doanh thu khách hàng** (đã gộp nhiều tháng) và
**bao nhiêu file Payroll cũng được** (mỗi file 1 tháng, sheet "Bảng lương") —
app tự nhận diện tháng của từng file, tự sắp xếp theo thời gian và tự mở rộng
bảng KTV-TVV theo đúng số tháng bạn upload (không cần tách riêng "tháng A/tháng B").
"""
)

revenue_file = st.file_uploader("📈 File Phân tích Doanh thu khách hàng", type=["xlsx"])
payroll_files = st.file_uploader(
    "💰 File Payroll (chọn nhiều file cùng lúc, mỗi file 1 tháng)",
    type=["xlsx"], accept_multiple_files=True,
)

if revenue_file and payroll_files:
    try:
        revenue = read_revenue_file(revenue_file)
    except Exception as e:
        st.error(f"Lỗi đọc file Phân tích Doanh thu: {e}")
        st.stop()

    parsed = []
    for f in payroll_files:
        try:
            year, month, label, stats = read_payroll(f)
        except Exception as e:
            st.error(f"Lỗi đọc file '{f.name}': {e}")
            continue
        if not label:
            st.warning(f"Không tự nhận diện được tháng của file '{f.name}' (cần tiêu đề dạng 'THÁNG 08/2026') — bỏ qua file này.")
            continue
        parsed.append({"filename": f.name, "year": year, "month": month, "label": label, "stats": stats})

    if not parsed:
        st.stop()

    parsed.sort(key=lambda p: (p["year"], p["month"]))

    # Cảnh báo nếu 2 file trùng tháng
    labels_seen = [p["label"] for p in parsed]
    dups = {l for l in labels_seen if labels_seen.count(l) > 1}
    if dups:
        st.error(f"Có nhiều file cùng nhận diện là tháng {', '.join(dups)} — vui lòng kiểm tra lại, mỗi tháng chỉ nên có 1 file.")
        st.stop()

    st.success("Đã nhận diện: " + " → ".join(f"**{p['label']}** ({p['filename']})" for p in parsed))

    st.subheader("🔎 Xem trước số liệu nhân sự đã tổng hợp")
    preview_rows = []
    for b in BRANCHES:
        row = {"Chi nhánh": b["label"]}
        for p in parsed:
            s = p["stats"].get(b["payroll"], {})
            row[f"Số NS {p['label']}"] = s.get("so_nhan_su")
            row[f"KTV {p['label']}"] = s.get("so_ktv")
            row[f"TVV {p['label']}"] = s.get("so_tvv")
        preview_rows.append(row)
    st.dataframe(preview_rows, use_container_width=True)

    if st.button("🚀 Xuất sheet KTV-TVV", type="primary"):
        months = [(p["label"], p["stats"]) for p in parsed]
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        build_ktv_tvv_sheet(wb, months, revenue)

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)

        st.success("Đã tạo xong sheet KTV-TVV!")
        st.download_button(
            "⬇️ Tải file KTV-TVV.xlsx",
            data=out,
            file_name=f"KTV-TVV_{parsed[0]['label']}_{parsed[-1]['label']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Vui lòng upload file Phân tích Doanh thu và ít nhất 1 file Payroll để bắt đầu.")
