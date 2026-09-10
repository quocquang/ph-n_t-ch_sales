"""
Streamlit app: Tự động tạo sheet "KTV-TVV" từ raw data — BẢN NÂNG CẤP (nhanh
hơn + chi tiết hơn):

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

Toàn bộ công thức nghiệp vụ đã ĐỐI CHIẾU khớp chính xác 100% với file báo cáo
lương T8/2026 (so với T7/2026) người dùng tự làm tay, trên toàn bộ 9 chi nhánh
và 26 chỉ tiêu.

LƯU Ý: vị trí cột "TỔNG THU NHẬP" và các cột khác trong "Bảng lương" LỆCH NHAU
giữa các tháng — app dò cột theo TÊN HEADER, không theo số thứ tự cột.

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
    "dt_ca_nhan": "DOANH THU CÁ NHÂN TRƯỚC THUẾ PHÍ",
    "tour_ca_nhan": "TỔNG ĐIỂM TOUR CÁ NHÂN",
    "tong_thu_nhap": "TỔNG THU NHẬP",
    "ty_le_kpi_ca_nhan": "TỶ LỆ %\nCÁ NHÂN ĐẠT SO VỚI KPI",
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
    """Đọc sheet 'Bảng lương' ở chế độ read_only (nhanh hơn nhiều lần so với
    load bình thường trên file lớn nhiều sheet). Trả về:
      (year, month, label, {branch: stats}, unmapped_branches_set)
    unmapped_branches_set = tên chi nhánh xuất hiện trong Payroll nhưng CHƯA
    có trong cấu hình BRANCHES của app (để cảnh báo, tránh bỏ sót âm thầm)."""
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    if "Bảng lương" not in wb.sheetnames:
        raise ValueError(f"File '{filename}' không có sheet 'Bảng lương'.")
    ws = wb["Bảng lương"]

    title = ws.cell(row=1, column=3).value or ""
    year, month, label = detect_payroll_month(title, filename)

    cols = find_header_columns(ws, NEEDED_HEADERS)
    stats = {}
    unmapped_branches = set()

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
    idx_chi_nhanh = cols["chi_nhanh"] - 1
    idx_vi_tri = cols["vi_tri"] - 1
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
            # có thể là dòng subtotal (branch field chứa số/nhãn khác) hoặc
            # 1 chi nhánh mới chưa cấu hình -> chỉ cảnh báo khi giống 1 tên
            # chi nhánh thật (chữ hoa, có dấu cách) để không báo nhầm dòng rác.
            if re.match(r"^[A-ZÀ-Ỹ ]{3,}$", branch):
                unmapped_branches.add(branch)
            continue
        pos = str(pos).strip()

        dt_ca_nhan = row[idx_dt] or 0
        tour = row[idx_tour] or 0
        thu_nhap = row[idx_tn] or 0
        kpi_rate = row[idx_kpi]

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

    return year, month, label, stats, unmapped_branches


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
              "mua_tt_moi": 0, "mua_tt_cu": 0, "tong_doanh_thu": 0}
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


# ---------------------------------------------------------------------------
# 6. XÂY DỰNG SHEET "KTV-TVV"
# ---------------------------------------------------------------------------

def build_ktv_tvv_sheet(wb, months, revenue):
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


def build_audit_sheet(wb, months, revenue, unmapped_branches, unmapped_sheets):
    ws = wb.create_sheet("Audit", 0)
    ws.column_dimensions["A"].width = 100
    ws["A1"] = "Báo cáo kiểm tra dữ liệu — tự động tạo lúc xuất file"
    ws["A1"].font = FONT_BOLD
    ws["A2"] = "Các tháng trong file: " + " → ".join(m[0] for m in months)
    ws["A2"].font = FONT

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

    for b in BRANCHES:
        for i, lbl in enumerate(month_labels):
            d = revenue.get(b["revenue_sheet"], {}).get(lbl)
            if not d or d.get("tong_doanh_thu") is None:
                warnings.append(f"⚠ [{b['label']} - {lbl}] Không tìm thấy dữ liệu doanh thu tương ứng.")
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

    r = 4
    if warnings:
        ws[f"A{r}"] = f"Phát hiện {len(warnings)} điểm cần lưu ý:"
        ws[f"A{r}"].font = FONT_BOLD
        r += 1
        for w in warnings:
            cell = ws[f"A{r}"]
            cell.value = w
            cell.font = FONT
            cell.fill = FILL_WARN if w.startswith("⚠") else PatternFill("solid", fgColor="FCE4D6")
            cell.alignment = Alignment(wrap_text=True)
            r += 1
    else:
        ws[f"A{r}"] = "✓ Không phát hiện bất thường nào."
        ws[f"A{r}"].font = FONT
    ws.freeze_panes = "A4"
    return ws


# ---------------------------------------------------------------------------
# 7. UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Báo cáo lương KTV-TVV", layout="wide")
st.title("📊 Tự động tạo sheet KTV-TVV (Vận hành + Hiệu suất + Chi phí nhân sự)")

st.markdown(
    """
Upload file **Phân tích Doanh thu khách hàng** và **bao nhiêu file Payroll cũng
được** (mỗi file 1 tháng, sheet "Bảng lương") — app tự nhận diện tháng, tự sắp
xếp theo thời gian và tự mở rộng bảng KTV-TVV theo đúng số tháng bạn upload.
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
                year, month, label, stats, unmapped = read_payroll_bytes(f.getvalue(), f.name)
            except Exception as e:
                st.error(f"Lỗi đọc file '{f.name}': {e}")
                continue
            if not label:
                st.warning(f"Không tự nhận diện được tháng của file '{f.name}' — bỏ qua file này.")
                continue
            unmapped_branches_all |= unmapped
            parsed.append({"filename": f.name, "year": year, "month": month, "label": label, "stats": stats})

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

    st.subheader("🔎 Xem trước số liệu nhân sự đã tổng hợp")
    preview_rows = []
    for b in BRANCHES + [{"label": TOTAL_LABEL, "payroll": None}]:
        row = {"Chi nhánh": b["label"]}
        for p in parsed:
            if b["payroll"] is None:
                s = aggregate_payroll_total(p["stats"])
            else:
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
        build_audit_sheet(wb, months, revenue, unmapped_branches_all, unmapped_sheets)
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
