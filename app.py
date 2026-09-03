"""
Streamlit app: Upload NHIỀU file raw dashboard (bao nhiêu tháng cũng được),
app TỰ ĐỘNG nhận diện tháng, tự gộp và xuất ra file "Phân tích Doanh thu
khách hàng" — mỗi chi nhánh 1 sheet, mỗi tháng 1 cột, sắp xếp theo thời gian.

Điểm khác so với bản trước:
  1. Tên cột (tháng) được TỰ ĐỘNG điền và sắp xếp — không cần gõ tay.
     Chỉ khi nào app không tự đoán được ngày tháng, hoặc 2 file trùng
     nhãn tháng, mới hiện ô để bạn sửa tay (đúng chỗ bị lỗi thôi).
  2. Kiểm tra số liệu kỹ hơn — ngoài check trùng KPI và biến động
     tháng-qua-tháng, còn có thêm bước ĐỐI SOÁT TỔNG: so khớp
     "Tất cả chi nhánh" với tổng cộng dồn từng chi nhánh, để phát hiện
     copy nhầm dòng / thiếu chi nhánh trong raw dashboard.
  3. File xuất ra có thêm 1 sheet "Audit" ghi lại toàn bộ kết quả kiểm
     tra (đối soát tổng, KPI trùng, biến động lớn) để lưu vết lâu dài,
     không chỉ hiện trên UI rồi mất.

Cách chạy:
    pip install streamlit openpyxl pandas
    streamlit run app.py
"""

import io
import re
import shutil
import subprocess
import tempfile
from datetime import date
from pathlib import Path

import openpyxl
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. MAPPING: cột trong RAW DASHBOARD  ->  dòng trong sheet phân tích
#    (đã đối chiếu khớp 100% với dữ liệu T6/T7/T8-2026 thực tế của bạn)
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

# Chi nhánh không tính vào KPI kinh doanh (đào tạo / hành chính) -> không ra
# sheet riêng. "Học Viện LGS" cũng KHÔNG được cộng vào "Tất cả chi nhánh"
# trong raw dashboard, nên loại luôn khi đối soát tổng bên dưới.
EXCLUDE_BRANCHES = {"Học Viện LGS", "Văn Phòng"}
EXCLUDE_FROM_TOTAL_RECONCILIATION = {"Học Viện LGS"}

FONT = Font(name="Arial", size=11)
FONT_BOLD = Font(name="Arial", size=11, bold=True)
FONT_HEADER = Font(name="Arial", size=11, bold=True, color="FFFFFF")
FILL_HEADER = PatternFill("solid", fgColor="4472C4")
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_WARN = PatternFill("solid", fgColor="FFC7CE")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = "#,##0"
INT_FMT = "#,##0"
PCT_FMT = "0.0%"

MOM_THRESHOLD = 0.6          # ngưỡng cảnh báo biến động tháng-qua-tháng
RECONCILE_TOLERANCE = 1000   # sai số cho phép khi đối soát tổng (VNĐ)


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
    missing_cols = [name for name in RAW_COLS.values() if name not in headers]
    if missing_cols:
        raise ValueError(
            "File raw thiếu các cột cần thiết: " + ", ".join(missing_cols) +
            ". Kiểm tra lại định dạng file export từ dashboard."
        )

    data = {}
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        name = row[0]
        if not name:
            continue
        data[str(name).strip()] = dict(zip(headers, row))
    return title, data


def detect_month(title: str, filename: str):
    """Tìm khoảng ngày 'dd-mm-yyyy to dd-mm-yyyy' hoặc 'dd-mm-yyyy' đầu tiên
    trong tiêu đề/tên file. Trả về (year, month, ngay_bat_dau) hoặc
    (None, None, None) nếu không tìm được."""
    text = f"{title} {filename}"
    m = re.search(r"(\d{2})-(\d{2})-(\d{4})", text)
    if m:
        dd, mm, yyyy = m.groups()
        try:
            return int(yyyy), int(mm), date(int(yyyy), int(mm), int(dd))
        except ValueError:
            pass
    return None, None, None


def auto_generate_labels(parsed: list[dict]) -> list[str]:
    """Tự động sinh nhãn cột tháng cho từng file, không cần người dùng gõ tay.
    - Nếu nhận diện được tháng/năm: nhãn mặc định là 'T{thang}'.
    - Nếu 2+ file trùng nhãn (vd cùng là T1 nhưng khác năm), tự thêm hậu tố
      năm 2 số: 'T1/26', 'T1/27' để phân biệt.
    - Nếu không nhận diện được ngày tháng, dùng tên file (không dấu .xlsx)
      làm nhãn tạm — trường hợp này sẽ được đánh dấu để người dùng xác nhận.
    """
    base_labels = []
    for p in parsed:
        if p["month"]:
            base_labels.append(f"T{p['month']}")
        else:
            base_labels.append(Path(p["filename"]).stem)

    # Đếm số lần xuất hiện của từng nhãn cơ bản -> nếu >1 thì cần thêm năm
    from collections import Counter
    counts = Counter(base_labels)

    final_labels = []
    for p, base in zip(parsed, base_labels):
        if counts[base] > 1 and p["year"]:
            final_labels.append(f"{base}/{str(p['year'])[-2:]}")
        else:
            final_labels.append(base)
    return final_labels


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
# 3. Kiểm tra bất thường (audit) — 3 lớp kiểm tra độc lập
# ---------------------------------------------------------------------------

def check_duplicate_kpi(label: str, raw_data: dict) -> list[str]:
    """Lớp 1: 2 chi nhánh có KPI giống hệt nhau -> khả năng copy nhầm dòng."""
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
                f"[{label}] KPI giống hệt nhau ({kpi:,.0f}) giữa: "
                f"{', '.join(branches)} — khả năng copy nhầm dòng trong raw dashboard."
            )
    return warnings


def check_total_reconciliation(label: str, raw_data: dict) -> list[str]:
    """Lớp 2 (MỚI): đối soát dòng 'Tất cả chi nhánh' với tổng cộng dồn từng
    chi nhánh riêng lẻ. Đây là cách hiệu quả nhất để bắt lỗi thiếu chi nhánh
    hoặc copy sai số trong raw dashboard, vì HQ tổng và tổng từng chi nhánh
    phải luôn khớp nhau (trừ Học Viện LGS không tính vào tổng kinh doanh)."""
    warnings = []
    if "Tất cả chi nhánh" not in raw_data:
        warnings.append(f"[{label}] Không tìm thấy dòng 'Tất cả chi nhánh' để đối soát.")
        return warnings

    for field_key, field_name in [("doanh_thu", "Doanh thu"), ("khach_moi", "Khách mới")]:
        col = RAW_COLS[field_key]
        reported = to_num(raw_data["Tất cả chi nhánh"].get(col))
        included_sum = sum(
            to_num(v.get(col)) or 0
            for b, v in raw_data.items()
            if b not in ({"Tất cả chi nhánh"} | EXCLUDE_FROM_TOTAL_RECONCILIATION)
        )
        if reported is None:
            continue
        diff = reported - included_sum
        if abs(diff) > RECONCILE_TOLERANCE:
            warnings.append(
                f"[{label}] '{field_name}' TẤT CẢ CHI NHÁNH = {reported:,.0f} nhưng "
                f"tổng cộng dồn từng chi nhánh = {included_sum:,.0f} "
                f"(lệch {diff:,.0f}) — kiểm tra lại raw dashboard có thiếu/thừa "
                f"chi nhánh nào không."
            )
    return warnings


def check_month_over_month(sheet_title, month_labels, values_by_month) -> list[str]:
    """Lớp 3: biến động bất thường giữa các tháng liên tiếp trong CÙNG 1 chi nhánh."""
    warnings = []
    for i in range(1, len(month_labels)):
        prev_label, cur_label = month_labels[i - 1], month_labels[i]
        prev_vals, cur_vals = values_by_month[i - 1], values_by_month[i]
        for row_idx, label, kind, _key in TEMPLATE_ROWS:
            if kind not in ("raw", "khach_cu"):
                continue
            pv, cv = prev_vals.get(row_idx), cur_vals.get(row_idx)
            if pv in (None, 0) or cv is None:
                continue
            change = (cv - pv) / pv
            if abs(change) > MOM_THRESHOLD:
                warnings.append(
                    f"[{sheet_title}] '{label}': {prev_label}={pv:,.0f} → "
                    f"{cur_label}={cv:,.0f} ({change*100:+.0f}%) — biến động lớn, nên kiểm tra lại."
                )
    return warnings


# ---------------------------------------------------------------------------
# 4. Xây dựng workbook nhiều tháng (+ sheet Audit lưu vết kiểm tra)
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

    # --- Lớp 1 + 2: kiểm tra trong nội bộ từng tháng (không cần build sheet) ---
    for m in months:
        all_warnings.extend(check_duplicate_kpi(m["label"], m["raw_data"]))
        all_warnings.extend(check_total_reconciliation(m["label"], m["raw_data"]))

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
            check_month_over_month(branch, month_labels, values_by_month)
        )

    # --- Sheet Audit: lưu toàn bộ kết quả kiểm tra vào chính file xuất ra ---
    audit_ws = wb.create_sheet("Audit", 0)
    audit_ws.column_dimensions["A"].width = 100
    audit_ws["A1"] = f"Báo cáo kiểm tra số liệu — tự động tạo lúc xuất file"
    audit_ws["A1"].font = FONT_BOLD
    audit_ws["A2"] = f"Các tháng trong file: {', '.join(m['label'] for m in months)}"
    audit_ws["A2"].font = FONT
    r = 4
    if all_warnings:
        audit_ws[f"A{r}"] = f"⚠ Phát hiện {len(all_warnings)} điểm cần kiểm tra lại:"
        audit_ws[f"A{r}"].font = FONT_BOLD
        r += 1
        for w in all_warnings:
            cell = audit_ws[f"A{r}"]
            cell.value = w
            cell.font = FONT
            cell.fill = FILL_WARN
            cell.alignment = Alignment(wrap_text=True)
            r += 1
    else:
        audit_ws[f"A{r}"] = "✓ Không phát hiện bất thường nào ở tất cả các bước kiểm tra."
        audit_ws[f"A{r}"].font = FONT
        r += 1
    audit_ws.freeze_panes = "A4"

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
            in_dir = Path(tmpdir) / "in"
            out_dir = Path(tmpdir) / "out"
            in_dir.mkdir()
            out_dir.mkdir()
            in_path = in_dir / "in.xlsx"
            in_path.write_bytes(xlsx_bytes.getvalue())
            # outdir PHẢI khác thư mục chứa file gốc, nếu không LibreOffice
            # sẽ báo lỗi ghi đè (SfxBaseModel) và không recalc được.
            result = subprocess.run(
                [soffice, "--headless", "--calc", "--convert-to", "xlsx",
                 "--outdir", str(out_dir), str(in_path)],
                capture_output=True, timeout=timeout, text=True,
            )
            out_path = out_dir / "in.xlsx"
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
Upload **1 hoặc nhiều file raw dashboard** (mỗi file là 1 tháng) — app tự
nhận diện tháng, tự đặt tên cột, gộp lại và xuất ra file **Phân tích Doanh
thu khách hàng** đầy đủ (mỗi chi nhánh 1 sheet, mỗi tháng 1 cột, xếp theo
thời gian), kèm 1 sheet **Audit** ghi lại toàn bộ kết quả kiểm tra số liệu.
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
        parsed.append(
            {
                "filename": f.name,
                "title": title,
                "raw_data": raw_data,
                "year": year,
                "month": month,
                "start_date": start_date,
            }
        )

    if parsed:
        # sắp theo ngày bắt đầu nếu nhận diện được, file không nhận diện được xếp cuối
        parsed.sort(key=lambda p: (p["start_date"] is None, p["start_date"]))

        # === TỰ ĐỘNG sinh nhãn cột — không cần người dùng gõ tay ===
        labels = auto_generate_labels(parsed)
        undetected = [p["filename"] for p in parsed if p["month"] is None]

        st.subheader("✅ Đã tự động nhận diện tháng")
        preview_cols = st.columns([3, 2, 2])
        preview_cols[0].markdown("**File**")
        preview_cols[1].markdown("**Tiêu đề trong file**")
        preview_cols[2].markdown("**Tên cột (tự động)**")
        for p, label in zip(parsed, labels):
            c = st.columns([3, 2, 2])
            c[0].write(f"📄 {p['filename']}")
            c[1].write(p["title"][:45] + ("..." if len(p["title"]) > 45 else ""))
            c[2].write(f"`{label}`")

        # Chỉ hiện ô sửa tay khi thực sự cần: nhãn trùng nhau hoặc không
        # nhận diện được ngày tháng — mọi trường hợp bình thường đều tự động.
        needs_manual_fix = undetected or (len(set(labels)) != len(labels))
        if needs_manual_fix:
            st.warning(
                "⚠️ Một số file app không tự đặt tên cột chắc chắn được — "
                "vui lòng xác nhận/sửa lại tên cột bên dưới."
            )
            fixed_labels = []
            for i, (p, label) in enumerate(zip(parsed, labels)):
                new_label = st.text_input(
                    f"Tên cột cho {p['filename']}", value=label, key=f"label_fix_{i}"
                )
                fixed_labels.append(new_label)
            labels = fixed_labels
            if len(set(labels)) != len(labels):
                st.error("❌ Vẫn còn 2 file trùng tên cột — vui lòng sửa lại cho khác nhau.")
                st.stop()

        months = [
            {"label": labels[i], "raw_data": p["raw_data"]}
            for i, p in enumerate(parsed)
        ]

        st.divider()
        st.subheader("🔎 Kiểm tra số liệu (chạy tự động)")

        prelim_warnings = []
        for m in months:
            prelim_warnings.extend(check_duplicate_kpi(m["label"], m["raw_data"]))
            prelim_warnings.extend(check_total_reconciliation(m["label"], m["raw_data"]))

        if prelim_warnings:
            for w in prelim_warnings:
                st.warning(w)
        else:
            st.success(
                "Không phát hiện bất thường: không trùng KPI giữa các chi nhánh, "
                "và 'Tất cả chi nhánh' khớp với tổng cộng dồn từng chi nhánh."
            )

        st.info(
            "ℹ️ Dòng **'Bill TB khách mới 30 ngày'** không có trong raw dashboard "
            "nên sẽ để trống + tô vàng để bạn nhập tay và kiểm tra kỹ."
        )

        st.divider()
        if st.button("🚀 Xuất file phân tích", type="primary"):
            out_bytes, mom_warnings = build_workbook(months)
            out_bytes = recalc_with_libreoffice(out_bytes)
            new_only = [w for w in mom_warnings if w not in prelim_warnings]
            if new_only:
                st.subheader("⚠️ Biến động lớn giữa các tháng liên tiếp")
                for w in new_only:
                    st.warning(w)
            st.success(
                "Đã tạo file thành công! Toàn bộ cảnh báo (nếu có) cũng đã được "
                "lưu vào sheet 'Audit' trong file để tiện tra cứu sau này."
            )
            st.download_button(
                "⬇️ Tải file Phân tích Doanh thu khách hàng",
                data=out_bytes,
                file_name="Phan_tich_Doanh_thu_khach_hang.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("Vui lòng upload ít nhất 1 file raw dashboard.")
