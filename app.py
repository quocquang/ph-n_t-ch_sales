"""
Streamlit app: Tự động tạo sheet "KTV-TVV" (Bảng đánh giá vận hành + hiệu suất
KTV/TVV + hiệu quả chi phí nhân sự theo chi nhánh) từ 3 file raw:

  1. File "Phân tích Doanh thu khách hàng" (nhiều tháng, mỗi chi nhánh 1 sheet)
     — sinh ra từ app phân tích doanh thu.
  2. File Payroll tháng TRƯỚC (sheet "Bảng lương").
  3. File Payroll tháng HIỆN TẠI (sheet "Bảng lương").

Toàn bộ công thức đã được ĐỐI CHIẾU và khớp chính xác 100% với file báo cáo
lương T8/2026 (so sánh T7/2026) do người dùng tự làm tay trước đó, trên 3 chi
nhánh mẫu (Bình Dương, Vũng Tàu, Quận 1) và toàn bộ 26 chỉ tiêu:

  - Bảng "CHỈ SỐ VẬN HÀNH" (Khách mới/cũ, Doanh thu, Tỷ lệ chốt, Bill TB)
    lấy trực tiếp từ file Phân tích Doanh thu, theo tên chi nhánh + nhãn tháng.
  - "Số nhân sự"      = tổng số dòng nhân viên của chi nhánh trong Bảng lương
  - "KTV"             = Kỹ thuật viên + Kỹ thuật viên phun xăm + Chuyên viên Clinic
  - "TVV"             = Tư vấn viên + Trợ lý bác sĩ
  - "OM/CM/LEAD"      = OM + CM + LEAD
  - "QLCN"            = QLCN
  - "Doanh thu KTV/TVV" = SUM("DOANH THU CÁ NHÂN TRƯỚC THUẾ PHÍ") theo nhóm
  - "Điểm tour KTV"     = SUM("TỔNG ĐIỂM TOUR CÁ NHÂN") nhóm KTV
  - "Thu nhập BQ KTV/TVV" = AVERAGE("TỔNG THU NHẬP") theo nhóm
  - "Tỷ lệ chốt bình quân" (TVV) = AVERAGE("TỶ LỆ % CÁ NHÂN ĐẠT SO VỚI KPI") nhóm TVV
  - "Chi phí nhân sự"   = SUM("TỔNG THU NHẬP") TOÀN BỘ nhân sự chi nhánh (mọi vị trí)
  - "Thu nhập BQ/tổng NS", "Doanh thu/tổng NS", "Điểm hiệu quả", "Xếp hạng":
    tính bằng công thức Excel y hệt file gốc.

LƯU Ý: vị trí cột "TỔNG THU NHẬP" và các cột khác trong "Bảng lương" LỆCH NHAU
giữa các tháng (do người làm lương chèn/xoá cột) — app dò cột theo TÊN HEADER,
không theo số thứ tự cột, để không bao giờ đọc nhầm cột khi cấu trúc file đổi.

Cách chạy:
    pip install streamlit openpyxl
    streamlit run app.py
"""

import io
import re
from datetime import date

import openpyxl
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. CẤU HÌNH CHI NHÁNH — thứ tự cột đúng như file KTV-TVV gốc.
#    "payroll" = tên viết hoa không dấu-cách kiểu trong Bảng lương.
#    "revenue_sheet" = tên sheet trong file Phân tích Doanh thu khách hàng.
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

# Cột cần dò theo tên header trong sheet "Bảng lương" (không dò theo số cột
# cố định vì bố cục lệch giữa các tháng).
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
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = "#,##0"
INT_FMT = "#,##0"
PCT_FMT = "0.0%"


# ---------------------------------------------------------------------------
# 2. ĐỌC FILE PAYROLL (sheet "Bảng lương") — dò cột theo tên header
# ---------------------------------------------------------------------------

def find_header_columns(ws, header_map, search_rows=(2, 3, 4)):
    """Dò vị trí cột (1-indexed) cho từng tên header cần tìm, quét vài dòng
    đầu (vì header có thể nằm ở dòng gộp merge khác nhau giữa các tháng)."""
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
    # fallback: dd-mm-yyyy trong tên file
    m2 = re.search(r"(\d{2})-(\d{2})-(\d{4})", text)
    if m2:
        dd, mm, yyyy = m2.groups()
        return int(yyyy), int(mm), f"T{int(mm)}"
    return None, None, None


def read_payroll(file):
    """Đọc sheet 'Bảng lương', trả về (label_thang, {payroll_branch_name: stats})."""
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

    # Dữ liệu nhân viên bắt đầu sau 4 dòng header. Bỏ qua các dòng subtotal /
    # dòng số thứ tự cột (cột "CHI NHÁNH LÀM VIỆC" không phải chuỗi tên chi
    # nhánh hợp lệ, hoặc cột "VỊ TRÍ" rỗng).
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

    return label, stats


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
    """Trả về {revenue_sheet_name: {label_thang: {field: value}}}."""
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
# 4. XÂY DỰNG SHEET "KTV-TVV"
# ---------------------------------------------------------------------------

def style_header(cell, text, fill=True):
    cell.value = text
    cell.font = FONT_HEADER if fill else FONT_BOLD
    if fill:
        cell.fill = FILL_HEADER
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
    cell.border = BORDER


def write_row_data(ws, row, label, col_prev, col_cur, val_prev, val_cur,
                    diff_kind, number_format, bold=False):
    """diff_kind: 'abs' -> C-B ; 'pct' -> (B-C)/C (giống hệt file gốc)."""
    letter_prev = get_column_letter(col_prev)
    letter_cur = get_column_letter(col_cur)
    letter_diff = get_column_letter(col_cur + 1)

    a = ws.cell(row=row, column=1, value=label)
    a.font = FONT_BOLD if bold else FONT
    a.border = BORDER

    for col, val in ((col_prev, val_prev), (col_cur, val_cur)):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = FONT
        cell.border = BORDER
        cell.number_format = number_format
        cell.alignment = Alignment(horizontal="right")
        if val is None:
            cell.fill = FILL_INPUT
            cell.comment = Comment("Không có dữ liệu — kiểm tra lại file raw.", "App KTV-TVV")

    diff_cell = ws.cell(row=row, column=col_cur + 1)
    if diff_kind == "abs":
        diff_cell.value = f"={letter_cur}{row}-{letter_prev}{row}"
    elif diff_kind == "pct":
        diff_cell.value = f"=({letter_prev}{row}-{letter_cur}{row})/{letter_cur}{row}"
        diff_cell.number_format = "0.0%"
    diff_cell.font = FONT
    diff_cell.border = BORDER
    diff_cell.alignment = Alignment(horizontal="right")


def build_ktv_tvv_sheet(wb, label_prev, label_cur, payroll_prev, payroll_cur, revenue):
    ws = wb.create_sheet("KTV-TVV")
    ws.column_dimensions["A"].width = 34
    ws["A1"] = "BẢNG ĐÁNH CHỈ SỐ VẬN HÀNH"
    ws["A1"].font = FONT_BOLD

    branch_cols = {}  # label -> (col_prev, col_cur)
    col = 2
    for b in BRANCHES:
        branch_cols[b["label"]] = (col, col + 1)
        letter_prev = get_column_letter(col)
        ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 1)
        style_header(ws.cell(row=2, column=col), b["label"])
        style_header(ws.cell(row=2, column=col + 2), "chênh lệch")
        style_header(ws.cell(row=3, column=col), label_prev)
        style_header(ws.cell(row=3, column=col + 1), label_cur)
        for cc in range(col, col + 3):
            ws.column_dimensions[get_column_letter(cc)].width = 14
        col += 3
    ws["A2"] = "Chỉ tiêu"
    ws["A2"].font = FONT_HEADER
    ws["A2"].fill = FILL_HEADER
    ws["A3"] = "Tháng"
    ws["A3"].font = FONT_HEADER
    ws["A3"].fill = FILL_HEADER

    def get_rev(branch_label, label_thang, field):
        sheet_name = next(b["revenue_sheet"] for b in BRANCHES if b["label"] == branch_label)
        d = revenue.get(sheet_name, {}).get(label_thang, {})
        return d.get(field)

    def get_pay(branch_label, payroll_stats, field, default=0):
        payroll_name = next(b["payroll"] for b in BRANCHES if b["label"] == branch_label)
        s = payroll_stats.get(payroll_name)
        if not s:
            return None
        return s.get(field, default)

    # --- Bảng chỉ số vận hành (6-13) ---
    ws["A4"] = "CHỈ SỐ VẬN HÀNH"; ws["A4"].font = FONT_BOLD; ws["A4"].fill = FILL_SECTION
    ws["A5"] = "KHÁCH HÀNG MỚI /CŨ"; ws["A5"].font = FONT_BOLD; ws["A5"].fill = FILL_SECTION

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
        for branch_label, (cp, cc) in branch_cols.items():
            vprev = get_rev(branch_label, label_prev, field)
            vcur = get_rev(branch_label, label_cur, field)
            write_row_data(ws, row, label, cp, cc, vprev, vcur, kind, fmt)

    # --- Bảng đánh giá hiệu suất làm việc chi nhánh (14-21) ---
    ws["A14"] = "BẢNG ĐÁNH GIÁ HIỆU SUẤT LÀM VIỆC CHI NHÁNH"; ws["A14"].font = FONT_BOLD; ws["A14"].fill = FILL_SECTION
    for branch_label, (cp, cc) in branch_cols.items():
        style_header(ws.cell(row=15, column=cp), branch_label)
        ws.merge_cells(start_row=15, start_column=cp, end_row=15, end_column=cc)
        style_header(ws.cell(row=15, column=cc + 1), "chênh lệch")
        style_header(ws.cell(row=16, column=cp), label_prev)
        style_header(ws.cell(row=16, column=cc), label_cur)
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
        for branch_label, (cp, cc) in branch_cols.items():
            vprev = get_pay(branch_label, payroll_prev, field)
            vcur = get_pay(branch_label, payroll_cur, field)
            write_row_data(ws, row, label, cp, cc, vprev, vcur, kind, fmt)

    ws["A22"] = "HIỆU SUẤT KTV"; ws["A22"].font = FONT_BOLD; ws["A22"].fill = FILL_SECTION
    for branch_label, (cp, cc) in branch_cols.items():
        vprev = get_pay(branch_label, payroll_prev, "dt_ktv")
        vcur = get_pay(branch_label, payroll_cur, "dt_ktv")
        write_row_data(ws, 23, "Doanh thu KTV", cp, cc, vprev, vcur, "pct", MONEY_FMT)
        lp, lc = get_column_letter(cp), get_column_letter(cc)
        for c, lbl_row in ((cp, 24), (cc, 24)):
            pass
        ws.cell(row=24, column=1, value="DT/KTV").font = FONT
        ws.cell(row=24, column=1).border = BORDER
        for c in (cp, cc):
            cell = ws.cell(row=24, column=c, value=f"={get_column_letter(c)}23/{get_column_letter(c)}18")
            cell.number_format = MONEY_FMT; cell.font = FONT; cell.border = BORDER
        diff = ws.cell(row=24, column=cc + 1, value=f"=({lp}24-{lc}24)/{lc}24")
        diff.number_format = "0.0%"; diff.font = FONT; diff.border = BORDER

        vprev_t = get_pay(branch_label, payroll_prev, "tour_ktv")
        vcur_t = get_pay(branch_label, payroll_cur, "tour_ktv")
        write_row_data(ws, 25, "Điểm tour KTV", cp, cc, vprev_t, vcur_t, "pct", INT_FMT)

        def avg(stats_field_sum, count_field, payroll_stats):
            payroll_name = next(b["payroll"] for b in BRANCHES if b["label"] == branch_label)
            s = payroll_stats.get(payroll_name)
            if not s or not s.get(count_field):
                return None
            return s[stats_field_sum] / s[count_field]

        vprev_i = avg("thunhap_ktv_sum", "so_ktv", payroll_prev)
        vcur_i = avg("thunhap_ktv_sum", "so_ktv", payroll_cur)
        write_row_data(ws, 26, "Thu nhập BQ KTV", cp, cc, vprev_i, vcur_i, "pct", MONEY_FMT)

    ws["A27"] = "HIỆU SUẤT  TVV"; ws["A27"].font = FONT_BOLD; ws["A27"].fill = FILL_SECTION
    for branch_label, (cp, cc) in branch_cols.items():
        vprev = get_pay(branch_label, payroll_prev, "dt_tvv")
        vcur = get_pay(branch_label, payroll_cur, "dt_tvv")
        write_row_data(ws, 28, "Doanh thu TVV", cp, cc, vprev, vcur, "pct", MONEY_FMT)

        ws.cell(row=29, column=1, value="DT/TVV").font = FONT
        ws.cell(row=29, column=1).border = BORDER
        lp, lc = get_column_letter(cp), get_column_letter(cc)
        for c in (cp, cc):
            cell = ws.cell(row=29, column=c, value=f"={get_column_letter(c)}28/{get_column_letter(c)}19")
            cell.number_format = MONEY_FMT; cell.font = FONT; cell.border = BORDER
        diff = ws.cell(row=29, column=cc + 1, value=f"=({lp}29-{lc}29)/{lc}29")
        diff.number_format = "0.0%"; diff.font = FONT; diff.border = BORDER

        def tvv_kpi_rate(payroll_stats):
            payroll_name = next(b["payroll"] for b in BRANCHES if b["label"] == branch_label)
            s = payroll_stats.get(payroll_name)
            if not s or not s.get("kpi_rate_tvv_n"):
                return None
            return s["kpi_rate_tvv_sum"] / s["kpi_rate_tvv_n"]

        vprev_r = tvv_kpi_rate(payroll_prev)
        vcur_r = tvv_kpi_rate(payroll_cur)
        write_row_data(ws, 30, "Tỷ lệ chốt bình quân", cp, cc, vprev_r, vcur_r, "pct", "0.0%")

        def avg_tvv(payroll_stats):
            payroll_name = next(b["payroll"] for b in BRANCHES if b["label"] == branch_label)
            s = payroll_stats.get(payroll_name)
            if not s or not s.get("so_tvv"):
                return None
            return s["thunhap_tvv_sum"] / s["so_tvv"]

        vprev_i = avg_tvv(payroll_prev)
        vcur_i = avg_tvv(payroll_cur)
        write_row_data(ws, 31, "Thu nhập BQ TVV", cp, cc, vprev_i, vcur_i, "pct", MONEY_FMT)

    ws["A32"] = "HIỆU QUẢ HOẠT ĐỘNG CHI NHÁNH"; ws["A32"].font = FONT_BOLD; ws["A32"].fill = FILL_SECTION
    for branch_label, (cp, cc) in branch_cols.items():
        vprev = get_rev(branch_label, label_prev, "tong_doanh_thu")
        vcur = get_rev(branch_label, label_cur, "tong_doanh_thu")
        write_row_data(ws, 33, "Tổng doanh thu", cp, cc, vprev, vcur, "pct", MONEY_FMT)

        lp, lc = get_column_letter(cp), get_column_letter(cc)
        ws.cell(row=34, column=1, value="Tăng trưởng doanh thu (%)").font = FONT
        ws.cell(row=34, column=1).border = BORDER
        cell_c = ws.cell(row=34, column=cc, value=f"=({lc}33-{lp}33)/{lp}33")
        cell_c.number_format = "0.0%"; cell_c.font = FONT; cell_c.border = BORDER
        ws.cell(row=34, column=cp).border = BORDER
        diff = ws.cell(row=34, column=cc + 1, value=f"=({lp}34-{lc}34)/{lc}34")
        diff.number_format = "0.0%"; diff.font = FONT; diff.border = BORDER

        vprev_c = get_pay(branch_label, payroll_prev, "chi_phi_nhan_su")
        vcur_c = get_pay(branch_label, payroll_cur, "chi_phi_nhan_su")
        write_row_data(ws, 35, "Chi phí nhân sự", cp, cc, vprev_c, vcur_c, "pct", MONEY_FMT)

        ws.cell(row=36, column=1, value="CPNS/Doanh thu (%)").font = FONT
        ws.cell(row=36, column=1).border = BORDER
        for c in (cp, cc):
            cell = ws.cell(row=36, column=c, value=f"={get_column_letter(c)}35/{get_column_letter(c)}33")
            cell.number_format = "0.0%"; cell.font = FONT; cell.border = BORDER
        diff = ws.cell(row=36, column=cc + 1, value=f"=({lp}36-{lc}36)/{lc}36")
        diff.number_format = "0.0%"; diff.font = FONT; diff.border = BORDER

        ws.cell(row=37, column=1, value="Thu nhập BQ / tổng nhân sự").font = FONT
        ws.cell(row=37, column=1).border = BORDER
        for c in (cp, cc):
            cell = ws.cell(row=37, column=c, value=f"={get_column_letter(c)}35/{get_column_letter(c)}17")
            cell.number_format = MONEY_FMT; cell.font = FONT; cell.border = BORDER
        diff = ws.cell(row=37, column=cc + 1, value=f"=({lp}37-{lc}37)/{lc}37")
        diff.number_format = "0.0%"; diff.font = FONT; diff.border = BORDER

        ws.cell(row=38, column=1, value="Doanh thu/Tổng nhân sự").font = FONT
        ws.cell(row=38, column=1).border = BORDER
        for c in (cp, cc):
            cell = ws.cell(row=38, column=c, value=f"={get_column_letter(c)}33/{get_column_letter(c)}17")
            cell.number_format = MONEY_FMT; cell.font = FONT; cell.border = BORDER
        diff = ws.cell(row=38, column=cc + 1, value=f"=({lp}38-{lc}38)/{lc}38")
        diff.number_format = "0.0%"; diff.font = FONT; diff.border = BORDER

        ws.cell(row=39, column=1, value="ĐIỂM HIỆU QUẢ").font = FONT_BOLD
        ws.cell(row=39, column=1).border = BORDER
        for c in (cp, cc):
            L = get_column_letter(c)
            formula = f"=IFERROR({L}38/1000000,0)*0.4+IFERROR({L}34,0)*100*0.2-IFERROR({L}36,0)*100*0.4"
            cell = ws.cell(row=39, column=c, value=formula)
            cell.font = FONT_BOLD; cell.border = BORDER

    # Xếp hạng (dòng 40) — so sánh điểm hiệu quả cùng tháng giữa các chi nhánh
    ws.cell(row=40, column=1, value="XẾP HẠNG").font = FONT_BOLD
    ws.cell(row=40, column=1).border = BORDER
    all_cols_prev = [branch_cols[b["label"]][0] for b in BRANCHES]
    all_cols_cur = [branch_cols[b["label"]][1] for b in BRANCHES]
    for branch_label, (cp, cc) in branch_cols.items():
        for c, all_cols in ((cp, all_cols_prev), (cc, all_cols_cur)):
            L = get_column_letter(c)
            others = [get_column_letter(oc) for oc in all_cols if oc != c]
            terms = "+".join(f"({o}$39>{L}$39)" for o in others)
            formula = f"={terms}+1"
            cell = ws.cell(row=40, column=c, value=formula)
            cell.font = FONT; cell.border = BORDER; cell.alignment = Alignment(horizontal="center")

    ws.freeze_panes = "B4"
    return ws


# ---------------------------------------------------------------------------
# 5. UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Báo cáo lương KTV-TVV", layout="wide")
st.title("📊 Tự động tạo sheet KTV-TVV (Vận hành + Hiệu suất + Chi phí nhân sự)")

st.markdown(
    """
Upload **3 file**: file **Phân tích Doanh thu khách hàng** (đã gộp nhiều tháng),
file **Payroll tháng trước** và file **Payroll tháng hiện tại** (sheet "Bảng lương").
App sẽ tự nhận diện tháng, tự gộp và xuất ra sheet **KTV-TVV** đầy đủ 9 chi nhánh,
công thức y hệt file mẫu bạn đã làm tay.
"""
)

col1, col2, col3 = st.columns(3)
with col1:
    revenue_file = st.file_uploader("📈 File Phân tích Doanh thu khách hàng", type=["xlsx"])
with col2:
    payroll_file_a = st.file_uploader("💰 File Payroll — tháng A", type=["xlsx"], key="pa")
with col3:
    payroll_file_b = st.file_uploader("💰 File Payroll — tháng B", type=["xlsx"], key="pb")

if revenue_file and payroll_file_a and payroll_file_b:
    try:
        revenue = read_revenue_file(revenue_file)
    except Exception as e:
        st.error(f"Lỗi đọc file Phân tích Doanh thu: {e}")
        st.stop()

    try:
        label_a, stats_a = read_payroll(payroll_file_a)
        label_b, stats_b = read_payroll(payroll_file_b)
    except Exception as e:
        st.error(f"Lỗi đọc file Payroll: {e}")
        st.stop()

    if not label_a or not label_b:
        st.warning("Không tự nhận diện được tháng từ 1 trong 2 file Payroll — kiểm tra lại tiêu đề file (cần dạng 'THÁNG 08/2026').")
        st.stop()

    # Sắp xếp: tháng nhỏ hơn -> label_prev, tháng lớn hơn -> label_cur
    def month_num(lbl):
        return int(lbl.replace("T", ""))

    if month_num(label_a) <= month_num(label_b):
        label_prev, payroll_prev = label_a, stats_a
        label_cur, payroll_cur = label_b, stats_b
    else:
        label_prev, payroll_prev = label_b, stats_b
        label_cur, payroll_cur = label_a, stats_a

    st.success(f"Đã nhận diện: tháng trước = **{label_prev}**, tháng hiện tại = **{label_cur}**")

    st.subheader("🔎 Xem trước số liệu nhân sự đã tổng hợp")
    preview_rows = []
    for b in BRANCHES:
        sp = payroll_prev.get(b["payroll"], {})
        sc = payroll_cur.get(b["payroll"], {})
        preview_rows.append({
            "Chi nhánh": b["label"],
            f"Số NS {label_prev}": sp.get("so_nhan_su"),
            f"Số NS {label_cur}": sc.get("so_nhan_su"),
            f"KTV {label_prev}": sp.get("so_ktv"),
            f"KTV {label_cur}": sc.get("so_ktv"),
            f"TVV {label_prev}": sp.get("so_tvv"),
            f"TVV {label_cur}": sc.get("so_tvv"),
        })
    st.dataframe(preview_rows, use_container_width=True)

    if st.button("🚀 Xuất sheet KTV-TVV", type="primary"):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        build_ktv_tvv_sheet(wb, label_prev, label_cur, payroll_prev, payroll_cur, revenue)

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)

        st.success("Đã tạo xong sheet KTV-TVV!")
        st.download_button(
            "⬇️ Tải file KTV-TVV.xlsx",
            data=out,
            file_name=f"KTV-TVV_{label_prev}_{label_cur}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Vui lòng upload đủ 3 file để bắt đầu.")
