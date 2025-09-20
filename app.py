import os
import pandas as pd
import streamlit as st
from pathlib import Path
from datetime import datetime

# ========================
# Core processing helpers
# ========================
def read_excel_safely(path, sheet, header_row):
    """Đọc Excel theo đuôi file, fallback khi lỗi."""
    p = Path(path)
    suf = p.suffix.lower()

    if suf == ".xlsb":
        engine = "pyxlsb"
    elif suf in [".xlsx", ".xlsm"]:
        engine = "openpyxl"
    elif suf == ".xls":
        engine = "xlrd"
    else:
        raise ValueError(f"Định dạng không hỗ trợ: {suf}")

    # Lấy sheet đầu nếu để trống
    if sheet == "" or sheet is None:
        xls = pd.ExcelFile(p, engine=engine)
        sheet = xls.sheet_names[0]

    try:
        return pd.read_excel(p, sheet_name=sheet, header=header_row, engine=engine)
    except Exception as e:
        if suf == ".xlsb":
            raise RuntimeError(
                "❌ Không đọc được .xlsb. "
                "Hãy mở file này trong Excel rồi Save As thành .xlsx để xử lý."
            ) from e
        raise


def convert_headers_to_yyyyww(cols: pd.Index):
    """Đổi tên cột: nếu parse được ngày → đổi sang dạng YYYYWW"""
    s = pd.Index(cols).astype(str)
    is_yyyyww = s.str.fullmatch(r"\d{6}", na=False)
    to_parse = s.where(~is_yyyyww, None)
    dt = pd.to_datetime(to_parse, errors="coerce", dayfirst=True)
    is_date = dt.notna()

    new = s.copy().to_series()
    if is_date.any():
        iso = dt[is_date].isocalendar()
        new_vals = iso["year"].astype(str) + iso["week"].astype(int).map("{:02d}".format)
        new.loc[is_date] = new_vals.to_numpy()
    new = pd.Index(new)

    week_mask = is_yyyyww | is_date
    return new, week_mask


def consolidate_weeks_fast(df: pd.DataFrame, week_mask: pd.Index, sort_week_cols=True):
    non = df.loc[:, ~week_mask]
    wk = df.loc[:, week_mask]
    if wk.shape[1] == 0:
        return df

    wk_num = wk.apply(pd.to_numeric, errors="coerce")
    wk_sum = wk_num.groupby(wk_num.columns, axis=1).sum(min_count=1)

    # Nếu toàn NaN thì giữ cột gốc
    groups = {}
    for c in wk.columns:
        groups.setdefault(c, []).append(c)
    for name, cols in groups.items():
        sub_num = wk_num[cols]
        if not sub_num.notna().any().any():
            wk_sum[name] = wk[cols[0]]

    wk_sum = wk_sum.loc[:, ~wk_sum.columns.duplicated(keep="last")]

    if sort_week_cols:
        def wkey(x):
            xs = str(x)
            return (0, int(xs)) if xs.isdigit() and len(xs) == 6 else (1, xs)
        wk_sum = wk_sum[sorted(wk_sum.columns, key=wkey)]

    return pd.concat([non, wk_sum], axis=1)


def filter_firm_forecast_colB(df: pd.DataFrame) -> pd.DataFrame:
    """Giữ lại dòng có colB = Firm hoặc Forecast"""
    if df.shape[1] <= 1:
        return df
    col = df.iloc[:, 1].astype(str).str.strip().str.lower()
    mask = col.isin(["firm", "forecast"])
    return df.loc[mask].copy()


def process_excel(file, sheet_name, header_row):
    df = read_excel_safely(file, sheet_name, header_row)
    df = filter_firm_forecast_colB(df)

    new_cols, week_mask = convert_headers_to_yyyyww(pd.Index(df.columns))
    df.columns = new_cols
    df = consolidate_weeks_fast(df, week_mask=week_mask, sort_week_cols=True)
    return df


# ========================
# Streamlit App
# ========================
st.set_page_config(page_title="Convert Header to YYYYWW", layout="wide")
st.title("📊 Convert Header to YYYYWW")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xlsm", "xls", "xlsb"])
sheet_name = st.text_input("Sheet name (để trống = sheet đầu)", value="")
header_row = st.number_input("Header row (0-based)", min_value=0, max_value=100, value=0, step=1)

if uploaded_file:
    if st.button("Process"):
        try:
            # Lưu file upload vào tạm
            suffix = Path(uploaded_file.name).suffix
            temp_path = Path("temp_input" + suffix)
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.read())

            # Xử lý
            df = process_excel(temp_path, sheet_name.strip(), int(header_row))

            st.success("✅ Xử lý thành công!")
            st.dataframe(df.head(50))  # hiển thị preview

            # Xuất ra file Excel tải về
            today_str = datetime.today().strftime("%Y%m%d")
            out_name = f"{today_str}.xlsx"
            df.to_excel(out_name, index=False, engine="xlsxwriter")

            with open(out_name, "rb") as f:
                st.download_button(
                    label="📥 Download output",
                    data=f,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"❌ Lỗi: {e}")
