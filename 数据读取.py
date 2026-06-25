"""Excel data loading helpers for the risk attribution workflow."""

from pathlib import Path
from typing import Any, Optional

import numpy as np
import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_MARKET_FILE = BASE_DIR / "市场数据-指数收盘价&无风险利率.xlsx"
DEFAULT_DURATION_FILE = BASE_DIR / "市场数据-指数久期.xlsx"
DEFAULT_PRODUCT_FILE = BASE_DIR / "产品价值序列.xlsx"


def read_market_price(
    code: str,
    market_file: Optional[Any] = None,
    value_name: str = "close",
) -> pd.DataFrame:
    """Read one index close-price sheet from the market Excel file."""
    path = Path(market_file) if market_file else DEFAULT_MARKET_FILE
    df = read_ths_export_sheet(path, sheet_name=code)
    return normalize_single_value_frame(df, value_name=value_name)


def read_risk_free_yield(
    market_file: Optional[Any] = None,
    sheet_name: str = "L001619604",
) -> pd.DataFrame:
    """Read 10Y government bond yield from the market Excel file."""
    path = Path(market_file) if market_file else DEFAULT_MARKET_FILE
    df = read_ths_export_sheet(path, sheet_name=sheet_name)
    return normalize_single_value_frame(df, value_name="yield_10y")


def read_duration_data(duration_file: Optional[Any] = None) -> pd.DataFrame:
    """Read duration data and normalize it to the two required columns."""
    path = Path(duration_file) if duration_file else DEFAULT_DURATION_FILE

    # Try the common THS export shape first: row 1 has codes/names, row 2 has fields.
    try:
        raw = read_excel_any(path, sheet_name=0, header=1)
    except Exception:
        raw = read_excel_any(path, sheet_name=0, header=0)

    date_col = guess_date_column(raw, None)
    numeric_cols = numeric_candidate_columns(raw, exclude={date_col})
    if len(numeric_cols) < 2:
        # Some exports have multiple sheets with one duration series each.
        xl = excel_file_any(path)
        frames = []
        for sheet in xl.sheet_names[:2]:
            sheet_df = read_ths_export_sheet(path, sheet_name=sheet)
            frames.append(normalize_single_value_frame(sheet_df, value_name=str(sheet)))
        if len(frames) >= 2:
            result = pd.concat(frames[:2], axis=1).dropna()
            result.columns = ["CBA00602_duration", "CBA00722_duration"]
            return result
        raise ValueError("Duration file must contain two numeric duration columns.")

    result = raw[[date_col, numeric_cols[0], numeric_cols[1]]].copy()
    result[date_col] = pd.to_datetime(result[date_col])
    for col in numeric_cols[:2]:
        result[col] = pd.to_numeric(result[col], errors="coerce")
    result = result.dropna(subset=[date_col, numeric_cols[0], numeric_cols[1]])
    result = result.set_index(date_col).sort_index()
    result.columns = ["CBA00602_duration", "CBA00722_duration"]
    return result


def read_product_returns(product_file: Optional[Any] = None) -> pd.Series:
    """Read product value Excel and convert it to W-FRI log returns."""
    path = Path(product_file) if product_file else DEFAULT_PRODUCT_FILE
    raw = read_ths_export_sheet(path, sheet_name=0)
    value_df = normalize_single_value_frame(raw, value_name="value")
    weekly_value = value_df["value"].resample("W-FRI").last().dropna()
    return np.log(weekly_value / weekly_value.shift(1)).dropna().rename("期间收益率")


def read_ths_export_sheet(path: Any, sheet_name: Any = 0) -> pd.DataFrame:
    """Read THS export sheets whose first row is metadata and second row is fields."""
    return read_excel_any(path, sheet_name=sheet_name, header=1)


def read_excel_any(
    path: Any,
    sheet_name: Any = 0,
    header: Optional[int] = 0,
) -> pd.DataFrame:
    """Read Excel with pandas first; for .xls fallback to Windows Excel COM."""
    path = Path(path)
    try:
        return pd.read_excel(path, sheet_name=sheet_name, header=header)
    except ImportError as exc:
        if path.suffix.lower() != ".xls":
            raise
        try:
            return read_excel_with_com(path, sheet_name=sheet_name, header=header)
        except Exception as com_exc:
            raise ImportError(
                "读取 .xls 文件需要 xlrd，或可用的 Windows Excel COM。"
                "请安装 xlrd>=2.0.1，或将文件另存为 .xlsx。"
            ) from com_exc
    except ValueError:
        raise


def excel_file_any(path: Any):
    """Return a pandas ExcelFile-like object where possible."""
    path = Path(path)
    try:
        return pd.ExcelFile(path)
    except ImportError:
        return ComExcelFile(path)


def read_excel_with_com(
    path: Any,
    sheet_name: Any = 0,
    header: Optional[int] = 0,
) -> pd.DataFrame:
    """Read an Excel sheet via Windows Excel COM for legacy .xls files."""
    import win32com.client

    path = str(Path(path).resolve())
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(path, ReadOnly=True)
        if isinstance(sheet_name, int):
            sheet = workbook.Worksheets(sheet_name + 1)
        else:
            sheet = workbook.Worksheets(str(sheet_name))
        values = sheet.UsedRange.Value
        if values is None:
            return pd.DataFrame()
        rows = [list(row) for row in values]
        rows = _trim_empty_rows(rows)
        if header is None:
            return pd.DataFrame(rows)
        columns = rows[header]
        data = rows[header + 1 :]
        return pd.DataFrame(data, columns=columns)
    finally:
        if workbook is not None:
            workbook.Close(False)
        excel.Quit()


class ComExcelFile:
    """Minimal ExcelFile-like object exposing sheet_names for legacy .xls."""

    def __init__(self, path: Any):
        import win32com.client

        self.path = str(Path(path).resolve())
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = None
        try:
            workbook = excel.Workbooks.Open(self.path, ReadOnly=True)
            self.sheet_names = [
                workbook.Worksheets(i).Name for i in range(1, workbook.Worksheets.Count + 1)
            ]
        finally:
            if workbook is not None:
                workbook.Close(False)
            excel.Quit()


def normalize_single_value_frame(
    data: pd.DataFrame,
    value_name: str,
    date_col: Optional[str] = None,
    value_col: Optional[str] = None,
) -> pd.DataFrame:
    """Normalize a dated one-value table to DatetimeIndex + value_name."""
    df = data.copy()
    df = df.dropna(how="all")
    date_col = guess_date_column(df, date_col)
    value_col = guess_value_column(df, value_col, exclude={date_col})
    result = df[[date_col, value_col]].copy()
    result[date_col] = pd.to_datetime(result[date_col])
    result[value_col] = pd.to_numeric(result[value_col], errors="coerce")
    result = result.dropna(subset=[date_col, value_col]).set_index(date_col).sort_index()
    result.columns = [value_name]
    return result


def guess_date_column(df: pd.DataFrame, explicit_col: Optional[str]) -> str:
    """Detect a date column unless the caller provides one."""
    if explicit_col:
        if explicit_col not in df.columns:
            raise ValueError("Date column '{}' not found.".format(explicit_col))
        return explicit_col

    normalized = {str(col).strip().lower(): col for col in df.columns}
    for name in ["time", "date", "日期", "时间", "交易日期", "净值日期"]:
        key = name.lower()
        if key in normalized:
            return normalized[key]

    for col in df.columns:
        converted = pd.to_datetime(df[col], errors="coerce")
        if converted.notna().sum() >= max(1, int(len(df) * 0.6)):
            return col
    raise ValueError("Could not detect date column.")


def guess_value_column(
    df: pd.DataFrame,
    explicit_col: Optional[str],
    exclude: Optional[set] = None,
) -> str:
    """Detect a numeric value column unless the caller provides one."""
    exclude = exclude or set()
    if explicit_col:
        if explicit_col not in df.columns:
            raise ValueError("Value column '{}' not found.".format(explicit_col))
        return explicit_col

    preferred_names = [
        "adjustment_nv",
        "close",
        "收盘价",
        "复权单位净值",
        "累计净值",
        "单位净值",
        "净值",
        "数值",
        "value",
        "yield_10y",
    ]
    normalized = {str(col).strip().lower(): col for col in df.columns}
    for name in preferred_names:
        key = name.lower()
        if key in normalized and normalized[key] not in exclude:
            return normalized[key]

    candidates = numeric_candidate_columns(df, exclude=exclude)
    if not candidates:
        raise ValueError("Could not detect value column.")
    return candidates[0]


def numeric_candidate_columns(df: pd.DataFrame, exclude: Optional[set] = None):
    """Return columns that can be interpreted as numeric."""
    exclude = exclude or set()
    candidates = []
    for col in df.columns:
        if col in exclude:
            continue
        numeric = pd.to_numeric(df[col], errors="coerce")
        if numeric.notna().sum() > 0:
            candidates.append(col)
    return candidates


def _trim_empty_rows(rows):
    return [row for row in rows if any(cell is not None for cell in row)]
