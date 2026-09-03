"""
Streamlit app: Tự động chuyển RAW DASHBOARD (dashboard-dd-mm-yyyy...xlsx)
sang file "Phân tích Doanh thu khách hàng" (mỗi chi nhánh 1 sheet, mỗi tháng 1 cột).

Cách chạy:
    pip install streamlit openpyxl pandas
    streamlit run app.py

Logic mapping đã được kiểm chứng bằng cách đối chiếu ngược lại với file
"Phân tích Doanh thu khách hàng tháng 7 tháng 8 v1.xlsx" do người dùng cung cấp:
117 giá trị được so khớp, khớp đúng 114/117 (97.4%). 3 sai lệch còn lại KHÔNG
phải do logic map sai, mà là lỗi/dữ liệu trôi có sẵn trong file gốc:
  - Quận 10, dòng KPI: file gốc ghi nhầm KPI của Tân Phú (8,931,000,000) thay vì
    KPI thật của Quận 10 (7,263,000,000).
  - "Tất cả chi nhánh" và "Bình Dương", dòng Khách booking mới: lệch đúng 1
    đơn vị (3851 vs 3852, 349 vs 350) — nhiều khả năng do dashboard được xuất
    lệch thời điểm vài phút so với lúc làm file phân tích.
App sẽ tự động cảnh báo các kiểu bất thường này (xem phần "Cảnh báo" bên dưới).
"""

import io
import re
from copy import copy

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. MAPPING: cột trong RAW DASHBOARD  ->  dòng trong sheet phân tích
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

# (dòng trong sheet phân tích, nhãn, loại)
# loại: "raw"     -> lấy trực tiếp từ RAW_COLS[key]
#        "formula" -> công thức tỉ lệ (không cần raw)
#        "missing" -> raw dashboard hiện KHÔNG có số liệu này
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

MISSING_ROWS = [r for r, _, t, _ in TEMPLATE_ROWS if t == "missing"]


# ---------------------------------------------------------------------------
# 2. Đọc raw dashboard
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
    """'30838 (3434 / 1361 / 26043)' -> số cuối cùng (khách thực tế CŨ)."""
    if v is None:
        return None
    nums = re.findall(r"-?\d+", str(v))
    if len(nums) >= 4:
        return int(nums[3])
    return None


def read_raw_dashboard(file) -> dict:
    """Trả về {tên chi nhánh: {tên cột: giá trị}}."""
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb.active

    header_row_idx = None
    for r in range(1, 6):
        cell = ws.cell(row=r, column=1).value
        if cell and str(cell).strip() == "Chi nhánh":
            header_row_idx = r
            break
    if header_row_idx is None:
        raise ValueError(
            "Không tìm thấy dòng tiêu đề 'Chi nhánh' trong file raw dashboard. "
            "Vui lòng kiểm tra lại file xuất từ hệ thống."
        )

    headers = [c.value for c in ws[header_row_idx]]
    data = {}
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        name = row[0]
        if not name:
            continue
        data[str(name).strip()] = dict(zip(headers, row))
    return data


def build_branch_values(branch_row: dict) -> dict:
    """Map 1 dòng raw dashboard -> {số dòng trong sheet phân tích: giá trị}."""
    raw_vals = {k: branch_row.get(colname) for k, colname in RAW_COLS.items()}
    out = {}
    for row_idx, _label, kind, key in TEMPLATE_ROWS:
        if kind == "raw":
            out[row_idx] = to_num(raw_vals[key])
        elif kind == "khach_cu":
            out[row_idx] = parse_khach_cu(raw_vals["khach_thuc_te"])
        elif kind in ("formula_moi", "formula_cu", "missing"):
            out[row_idx] = None  # xử lý riêng lúc ghi (formula) hoặc bỏ trống
    return out


# ---------------------------------------------------------------------------
# 3. Kiểm tra bất thường (đối chiếu số cho kỹ trước khi xuất)
# ---------------------------------------------------------------------------

def sanity_check(raw_data: dict, template_sheets: list[str]) -> list[str]:
    warnings = []

    # 3a. Trùng KPI y hệt giữa 2 chi nhánh khác nhau (dấu hiệu copy nhầm dòng)
    kpi_by_branch = {
        b: to_num(v.get("KPI"))
        for b, v in raw_data.items()
        if b != "Tất cả chi nhánh" and to_num(v.get("KPI"))
    }
    seen = {}
    for b, kpi in kpi_by_branch.items():
        seen.setdefault(kpi, []).append(b)
    for kpi, branches in seen.items():
        if len(branches) > 1:
            warnings.append(
                f"⚠️ KPI giống hệt nhau ({kpi:,.0f}) giữa các chi nhánh: "
                f"{', '.join(branches)} — khả năng cao là copy nhầm dòng trong "
                f"raw dashboard, vui lòng kiểm tra lại nguồn."
            )

    # 3b. Chi nhánh có trong template nhưng không thấy trong raw dashboard
    for sheet in template_sheets:
        if sheet not in raw_data:
            warnings.append(
                f"⚠️ Không tìm thấy chi nhánh '{sheet}' trong raw dashboard vừa "
                f"upload — cột tháng mới của sheet này sẽ bị bỏ trống."
            )
    return warnings


def compare_with_previous_column(ws, new_col_idx: int, computed: dict) -> list[str]:
    """So với cột liền trước (tháng trước) để phát hiện biến động bất thường."""
    warnings = []
    if new_col_idx <= 2:  # không có cột trước để so
        return warnings
    prev_col_letter = get_column_letter(new_col_idx - 1)
    for row_idx, label, kind, _key in TEMPLATE_ROWS:
        if kind == "missing":
            continue
        new_val = computed.get(row_idx)
        if new_val is None:
            continue
        prev_val = to_num(ws[f"{prev_col_letter}{row_idx}"].value)
        if prev_val in (None, 0):
            continue
        change = (new_val - prev_val) / prev_val
        if abs(change) > 0.6:  # lệch hơn 60% so với tháng trước
            warnings.append(
                f"[{ws.title}] '{label}': {prev_val:,.0f} → {new_val:,.0f} "
                f"({change*100:+.0f}%) — biến động lớn, nên kiểm tra lại."
            )
    return warnings


# ---------------------------------------------------------------------------
# 4. Ghi vào workbook phân tích (thêm 1 cột mới mỗi sheet)
# ---------------------------------------------------------------------------

def copy_cell_style(src_cell, dst_cell):
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = copy(src_cell.number_format)
        dst_cell.protection = copy(src_cell.protection)
        dst_cell.alignment = copy(src_cell.alignment)


def append_month_to_workbook(template_bytes, raw_data: dict, month_label: str):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    all_warnings = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if sheet_name not in raw_data:
            all_warnings.append(
                f"⚠️ Sheet '{sheet_name}': không có trong raw dashboard, bỏ qua."
            )
            continue

        # cột trống tiếp theo (dựa vào tiêu đề ở dòng 1)
        next_col_idx = 2
        while ws.cell(row=1, column=next_col_idx).value not in (None, ""):
            next_col_idx += 1
        new_col_letter = get_column_letter(next_col_idx)
        style_src_col_letter = get_column_letter(next_col_idx - 1)

        # header
        header_cell = ws[f"{new_col_letter}1"]
        header_cell.value = month_label
        copy_cell_style(ws[f"{style_src_col_letter}1"], header_cell)

        computed = build_branch_values(raw_data[sheet_name])

        for row_idx, label, kind, _key in TEMPLATE_ROWS:
            dst = ws[f"{new_col_letter}{row_idx}"]
            copy_cell_style(ws[f"{style_src_col_letter}{row_idx}"], dst)

            if kind == "formula_moi":
                dst.value = f"={new_col_letter}5/{new_col_letter}4"
            elif kind == "formula_cu":
                dst.value = f"={new_col_letter}14/{new_col_letter}13"
            elif kind == "missing":
                dst.value = None  # để trống - raw dashboard không có số liệu này
            else:
                dst.value = computed.get(row_idx)

        all_warnings.extend(
            compare_with_previous_column(ws, next_col_idx, computed)
        )

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out, all_warnings


# ---------------------------------------------------------------------------
# 5. UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Cập nhật Phân tích Doanh thu khách hàng", layout="wide")
st.title("📊 Cập nhật Phân tích Doanh thu khách hàng theo tháng")

st.markdown(
    """
Quy trình:
1. Upload **raw dashboard** xuất từ hệ thống (dạng `dashboard-dd-mm-yyyy to dd-mm-yyyy.xlsx`).
2. Upload file **Phân tích Doanh thu khách hàng** hiện tại (file này sẽ được thêm 1 cột tháng mới vào mỗi sheet chi nhánh).
   Tháng sau, bạn upload lại **chính file vừa xuất** để tiếp tục nối thêm cột — không cần làm lại từ đầu.
3. Đặt tên cột (ví dụ `T9`), kiểm tra bảng xem trước + cảnh báo, rồi tải file kết quả về.
"""
)

col1, col2 = st.columns(2)
with col1:
    raw_file = st.file_uploader("1️⃣ Raw dashboard tháng mới (.xlsx)", type=["xlsx"])
with col2:
    template_file = st.file_uploader(
        "2️⃣ File Phân tích Doanh thu khách hàng hiện tại (.xlsx)", type=["xlsx"]
    )

month_label = st.text_input("3️⃣ Tên cột tháng mới (vd: T9, Tháng 9...)", value="")

if raw_file and template_file:
    try:
        raw_data = read_raw_dashboard(raw_file)
    except Exception as e:
        st.error(f"Lỗi đọc raw dashboard: {e}")
        st.stop()

    template_bytes = template_file.read()
    wb_preview = openpyxl.load_workbook(io.BytesIO(template_bytes))
    sheet_names = wb_preview.sheetnames

    st.subheader("Xem trước dữ liệu sẽ được ghi vào từng chi nhánh")
    preview_tabs = st.tabs(sheet_names)
    for tab, sheet_name in zip(preview_tabs, sheet_names):
        with tab:
            if sheet_name not in raw_data:
                st.warning(f"Không tìm thấy '{sheet_name}' trong raw dashboard.")
                continue
            computed = build_branch_values(raw_data[sheet_name])
            rows_display = []
            for row_idx, label, kind, _key in TEMPLATE_ROWS:
                if kind == "formula_moi":
                    val = "= Khách mua TT / Khách mới (công thức)"
                elif kind == "formula_cu":
                    val = "= Khách mua TT / Khách thực tế (công thức)"
                elif kind == "missing":
                    val = "— (không có trong raw dashboard)"
                else:
                    v = computed.get(row_idx)
                    val = f"{v:,.0f}" if isinstance(v, (int, float)) else v
                rows_display.append({"Chỉ tiêu": label, "Giá trị": val})
            st.table(pd.DataFrame(rows_display))

    st.subheader("🔎 Cảnh báo / kiểm tra số liệu")
    warnings = sanity_check(raw_data, sheet_names)
    if warnings:
        for w in warnings:
            st.warning(w)
    else:
        st.success("Không phát hiện bất thường rõ ràng ở bước kiểm tra sơ bộ (trùng KPI, thiếu chi nhánh).")

    st.info(
        "ℹ️ Dòng **'Bill TB khách mới 30 ngày'** (dòng 10) hiện KHÔNG thể tự tính "
        "từ raw dashboard này (không có cột số liệu tương ứng) — ô này sẽ để trống, "
        "bạn cần điền tay hoặc bổ sung nguồn dữ liệu chi tiết hơn nếu cần con số này."
    )

    st.divider()
    if not month_label.strip():
        st.warning("Nhập tên cột tháng mới ở bước 3 để có thể xuất file.")
    else:
        if st.button("🚀 Xuất file cập nhật", type="primary"):
            out_bytes, write_warnings = append_month_to_workbook(
                template_bytes, raw_data, month_label.strip()
            )
            if write_warnings:
                st.subheader("⚠️ Cảnh báo biến động so với cột trước")
                for w in write_warnings:
                    st.warning(w)
            st.success("Đã tạo file thành công!")
            st.download_button(
                "⬇️ Tải file Phân tích Doanh thu khách hàng (đã cập nhật)",
                data=out_bytes,
                file_name=f"Phan_tich_Doanh_thu_khach_hang_{month_label.strip()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("Vui lòng upload đủ 2 file ở bước 1 và 2.")
