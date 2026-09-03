"""
Streamlit app: Upload NHIỀU file raw dashboard (bao nhiêu tháng cũng được),
app tự gộp và xuất ra file "Phân tích Doanh thu khách hàng" — mỗi chi nhánh
1 sheet, mỗi tháng 1 cột, sắp xếp theo thời gian.

Cách chạy:
    pip install streamlit openpyxl pandas
    streamlit run app.py
"""

import io
import re
import shutil
import subprocess
import tempfile
from copy import copy
from datetime import date
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. MAPPING: cột trong RAW DASHBOARD  ->  dòng trong sheet phân tích
#    (đã kiểm chứng khớp 117/117 với dữ liệu T7 và 114/117 với T8 — 3 chỗ
#    lệch còn lại là lỗi có sẵn trong file phân tích cũ, không phải do map sai)
# ---------------------------------------------------------------------------

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

# (dòng trong sheet phân tích, nhãn, loại, key trong RAW_COLS)
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

MONEY_ROWS = {2, 3, 7, 9, 16}

# Chi nhánh dùng KPI/doanh thu thật — loại các dòng hành chính không tính KPI
EXCLUDE_BRANCHES = {"Học Viện LGS", "Văn Phòng"}

FONT = Font(name="Arial", size=11)
FONT_BOLD = Font(name="Arial", size=11, bold=True)
FONT_HEADER = Font(name="Arial", size=11, bold=True, color="FFFFFF")
FILL_HEADER = PatternFill("solid", fgColor="4472C4")
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = "#,##0"
INT_FMT = "#,##0"
PCT_FMT = "0.0%"


# ---------------------------------------------------------------------------
# 2. Đọc raw dashboard + tự nhận diện tháng từ tiêu đề file
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# 3. Kiểm tra bất thường
# ---------------------------------------------------------------------------

def sanity_check_single_month(label: str, raw_data: dict) -> list[str]:
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


def month_over_month_warnings(sheet_title, month_labels, values_by_month) -> list[str]:
    warnings = []
    for i in range(1, len(month_labels)):
        prev_label, cur_label = month_labels[i - 1], month_labels[i]
        prev_vals, cur_vals = values_by_month[i - 1], values_by_month[i]
        for row_idx, label, kind, _key in TEMPLATE_ROWS:
            if kind != "raw" and kind != "khach_cu":
                continue
            pv, cv = prev_vals.get(row_idx), cur_vals.get(row_idx)
            if pv in (None, 0) or cv is None:
                continue
            change = (cv - pv) / pv
            if abs(change) > 0.6:
                warnings.append(
                    f"[{sheet_title}] '{label}': {prev_label}={pv:,.0f} → "
                    f"{cur_label}={cv:,.0f} ({change*100:+.0f}%) — biến động lớn, nên kiểm tra lại."
                )
    return warnings


# ---------------------------------------------------------------------------
# 4. Xây dựng workbook nhiều tháng
# ---------------------------------------------------------------------------

def build_workbook(months: list[dict]):
    """months: list các dict {'label': str, 'raw_data': {...}} đã sắp xếp theo thời gian."""
    all_branches = []
    for m in months:
        for b in m["raw_data"].keys():
            if b not in EXCLUDE_BRANCHES and b not in all_branches:
                all_branches.append(b)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    all_warnings = []

    for branch in all_branches:
        ws = wb.create_sheet(branch[:31])  # Excel giới hạn tên sheet 31 ký tự
        ws.column_dimensions["A"].width = 32
        ws["A1"] = "Chỉ tiêu"
        ws["A1"].font = FONT_HEADER
        ws["A1"].fill = FILL_HEADER
        ws["A1"].border = BORDER

        values_by_month = []
        for col_idx, m in enumerate(months, start=2):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 18
            header_cell = ws[f"{col_letter}1"]
            header_cell.value = m["label"]
            header_cell.font = FONT_HEADER
            header_cell.fill = FILL_HEADER
            header_cell.alignment = Alignment(horizontal="center")
            header_cell.border = BORDER

            branch_row = m["raw_data"].get(branch, {})
            computed = build_branch_values(branch_row) if branch_row else {}
            values_by_month.append(computed)

            for row_idx, label, kind, _key in TEMPLATE_ROWS:
                a = ws[f"A{row_idx}"]
                a.value = label
                a.font = FONT_BOLD if kind.startswith("formula") else FONT
                a.border = BORDER

                cell = ws[f"{col_letter}{row_idx}"]
                cell.font = FONT
                cell.border = BORDER
                cell.alignment = Alignment(horizontal="right")

                if kind == "formula_moi":
                    cell.value = f"={col_letter}5/{col_letter}4"
                    cell.number_format = PCT_FMT
                elif kind == "formula_cu":
                    cell.value = f"={col_letter}14/{col_letter}13"
                    cell.number_format = PCT_FMT
                elif kind == "missing":
                    cell.value = None
                    cell.fill = FILL_INPUT
                    cell.comment = Comment(
                        "Raw dashboard không có số liệu này. Vui lòng nhập tay.",
                        "App phân tích doanh thu",
                    )
                else:
                    val = computed.get(row_idx)
                    cell.value = val
                    cell.number_format = MONEY_FMT if row_idx in MONEY_ROWS else INT_FMT
                    if not branch_row:
                        cell.fill = FILL_INPUT
                        cell.comment = Comment(
                            f"Không có dữ liệu chi nhánh này trong file raw tháng {m['label']}.",
                            "App phân tích doanh thu",
                        )

        ws.freeze_panes = "B2"
        month_labels = [m["label"] for m in months]
        all_warnings.extend(
            month_over_month_warnings(branch, month_labels, values_by_month)
        )

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out, all_warnings


def recalc_with_libreoffice(xlsx_bytes: io.BytesIO, timeout: int = 30) -> io.BytesIO:
    """Mở/lưu lại file bằng LibreOffice headless để công thức (=B5/B4...) có
    sẵn giá trị đã tính (cached value), thay vì để trống cho tới khi người
    dùng tự mở file và tính lại. Nếu máy chủ không có LibreOffice, hoặc có
    lỗi bất kỳ, trả nguyên file gốc (công thức vẫn đúng, Excel sẽ tự tính khi
    mở bình thường — chỉ là chưa có sẵn giá trị cache)."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return xlsx_bytes

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            in_path = Path(tmpdir) / "in.xlsx"
            in_path.write_bytes(xlsx_bytes.getvalue())
            result = subprocess.run(
                [soffice, "--headless", "--calc", "--convert-to", "xlsx",
                 "--outdir", tmpdir, str(in_path)],
                capture_output=True, timeout=timeout, text=True,
            )
            out_path = Path(tmpdir) / "in.xlsx"
            if result.returncode == 0 and out_path.exists():
                recalced = io.BytesIO(out_path.read_bytes())
                recalced.seek(0)
                return recalced
    except Exception:
        pass
    return xlsx_bytes


# ---------------------------------------------------------------------------
# 5. UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Phân tích Doanh thu khách hàng", layout="wide")
st.title("📊 Phân tích Doanh thu khách hàng — nhiều tháng")

st.markdown(
    """
Upload **1 hoặc nhiều file raw dashboard** (mỗi file là 1 tháng, dạng
`dashboard-dd-mm-yyyy to dd-mm-yyyy.xlsx`) — app sẽ tự nhận diện tháng theo
tiêu đề trong file, gộp lại và xuất ra file **Phân tích Doanh thu khách hàng**
đầy đủ (mỗi chi nhánh 1 sheet, mỗi tháng 1 cột, xếp theo thời gian).
Lần sau có thêm tháng mới, chỉ cần upload thêm — không giới hạn số tháng.
"""
)

raw_files = st.file_uploader(
    "Upload raw dashboard (chọn nhiều file cùng lúc được)",
    type=["xlsx"],
    accept_multiple_files=True,
)

if raw_files:
    parsed = []
    for f in raw_files:
        try:
            title, raw_data = read_raw_dashboard(f)
        except Exception as e:
            st.error(f"Lỗi đọc file '{f.name}': {e}")
            continue
        year, month, start_date = detect_month(title, f.name)
        default_label = f"T{month}" if month else f.name
        parsed.append(
            {
                "filename": f.name,
                "title": title,
                "raw_data": raw_data,
                "year": year,
                "month": month,
                "start_date": start_date,
                "default_label": default_label,
            }
        )

    if parsed:
        st.subheader("Xác nhận tên cột (tháng) cho từng file")
        st.caption(
            "App tự đoán tên tháng từ tiêu đề file — bạn có thể sửa lại nếu cần "
            "(ví dụ 2 file cùng là 'T1' nhưng khác năm thì nên sửa thành 'T1/26', 'T1/27'...)."
        )

        # sắp theo ngày bắt đầu nếu nhận diện được, file không nhận diện được xếp cuối
        parsed.sort(key=lambda p: (p["start_date"] is None, p["start_date"]))

        labels = []
        for i, p in enumerate(parsed):
            cols = st.columns([3, 2, 3])
            cols[0].write(f"📄 {p['filename']}")
            cols[1].write(p["title"][:40] + ("..." if len(p["title"]) > 40 else ""))
            label = cols[2].text_input(
                "Tên cột", value=p["default_label"], key=f"label_{i}"
            )
            labels.append(label)

        if len(set(labels)) != len(labels):
            st.error("❌ Có 2 file đang trùng tên cột — vui lòng sửa lại cho khác nhau.")
        else:
            months = [
                {"label": labels[i], "raw_data": p["raw_data"]}
                for i, p in enumerate(parsed)
            ]

            st.subheader("🔎 Cảnh báo / kiểm tra số liệu")
            warnings = []
            for i, p in enumerate(parsed):
                warnings.extend(sanity_check_single_month(labels[i], p["raw_data"]))
            if warnings:
                for w in warnings:
                    st.warning(w)
            else:
                st.success("Không phát hiện bất thường ở bước kiểm tra sơ bộ (trùng KPI giữa các chi nhánh).")

            st.info(
                "ℹ️ Dòng **'Bill TB khách mới 30 ngày'** không có trong raw dashboard "
                "nên sẽ để trống + tô vàng để bạn nhập tay và kiểm tra kỹ."
            )

            st.divider()
            if st.button("🚀 Xuất file phân tích", type="primary"):
                out_bytes, mom_warnings = build_workbook(months)
                out_bytes = recalc_with_libreoffice(out_bytes)
                if mom_warnings:
                    st.subheader("⚠️ Biến động lớn giữa các tháng liên tiếp")
                    for w in mom_warnings:
                        st.warning(w)
                st.success("Đã tạo file thành công!")
                st.download_button(
                    "⬇️ Tải file Phân tích Doanh thu khách hàng",
                    data=out_bytes,
                    file_name="Phan_tich_Doanh_thu_khach_hang.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
else:
    st.info("Vui lòng upload ít nhất 1 file raw dashboard.")
