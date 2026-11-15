"""報表自動整理工具：把多家門市格式不一的銷售報表合併成一份整理好的 Excel。

用法：
    python merge_reports.py                       # 讀 sample_data/，輸出 整理結果.xlsx
    python merge_reports.py 資料夾 -o 輸出檔.xlsx

功能：
- 讀取資料夾內所有 .xlsx / .xls / .csv（略過 Excel 開啟時產生的 ~$ 暫存檔）
- 欄位名稱自動對應（日期 / 銷售日期 / date …）
- 清理空白列、重複列、千分位文字、前後空白，統一日期格式
- 從檔名取得門市名稱（台北店_2026-10.xlsx → 台北店）
- 無法解析的資料列放到「異常資料」工作表，並註明原因
- 輸出：明細、門市彙總、品項彙總、每日營收、異常資料
"""
import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ALIASES = {
    "日期": ["日期", "銷售日期", "交易日期", "date"],
    "品項": ["品項", "商品名稱", "商品", "品名", "item", "product"],
    "數量": ["數量", "銷售數量", "qty", "quantity"],
    "單價": ["單價", "單價(元)", "售價", "price", "unit_price"],
}
COLUMNS = list(ALIASES)


def read_file(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        for enc in ("utf-8-sig", "big5"):
            try:
                return pd.read_csv(path, dtype=str, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise ValueError("無法判斷 CSV 編碼")
    return pd.read_excel(path, dtype=str)


def normalize_columns(df: pd.DataFrame, path: Path) -> pd.DataFrame:
    lookup = {a.lower(): std for std, names in ALIASES.items() for a in names}
    df = df.rename(columns=lambda c: lookup.get(str(c).strip().lower(), str(c).strip()))
    missing = [c for c in COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name} 缺少欄位：{'、'.join(missing)}")
    return df[COLUMNS]


def to_number(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.str.replace(",", "", regex=False).str.strip(), errors="coerce")


def to_date(s: pd.Series) -> pd.Series:
    out = pd.Series(pd.NaT, index=s.index)
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S"):
        todo = out.isna()
        out[todo] = pd.to_datetime(s[todo], format=fmt, errors="coerce")
    return out


def load_all(folder: Path):
    files = sorted(p for p in folder.iterdir()
                   if p.suffix.lower() in (".xlsx", ".xls", ".csv") and not p.name.startswith("~$"))
    if not files:
        raise SystemExit(f"在 {folder} 找不到任何報表檔（.xlsx / .xls / .csv）")

    good, bad = [], []
    for path in files:
        df = normalize_columns(read_file(path), path)
        df = df.dropna(how="all")
        df["品項"] = df["品項"].str.strip()
        df["門市"] = path.stem.split("_")[0]
        df["來源檔案"] = path.name

        date, qty, price = to_date(df["日期"]), to_number(df["數量"]), to_number(df["單價"])
        reason = pd.Series("", index=df.index)
        reason[date.isna()] += "日期無法辨識；"
        reason[qty.isna()] += "數量不是數字；"
        reason[price.isna()] += "單價不是數字；"
        reason[df["品項"].fillna("") == ""] += "缺少品項；"

        ok = reason == ""
        bad.append(df[~ok].assign(異常原因=reason[~ok].str.rstrip("；")))
        clean = df[ok].assign(日期=date[ok], 數量=qty[ok].astype(int), 單價=price[ok])
        good.append(clean)
        print(f"  {path.name}：{ok.sum()} 筆可用，{(~ok).sum()} 筆異常")

    detail = pd.concat(good, ignore_index=True)
    dup = detail.duplicated(subset=["門市", *COLUMNS])
    if dup.any():
        bad.append(detail[dup].assign(日期=detail.loc[dup, "日期"].dt.strftime("%Y-%m-%d"),
                                      異常原因="重複資料（已排除）"))
        detail = detail[~dup]
    detail["金額"] = detail["數量"] * detail["單價"]
    detail = detail.sort_values(["日期", "門市", "品項"]).reset_index(drop=True)
    bad = pd.concat(bad, ignore_index=True).reindex(columns=["來源檔案", "門市", *COLUMNS, "異常原因"])
    return detail[["日期", "門市", "品項", "數量", "單價", "金額", "來源檔案"]], bad


def summarize(detail: pd.DataFrame) -> dict:
    by_store = (detail.groupby("門市").agg(筆數=("金額", "size"), 總數量=("數量", "sum"), 營收=("金額", "sum"))
                .sort_values("營收", ascending=False).reset_index())
    by_store["營收占比"] = by_store["營收"] / by_store["營收"].sum()

    by_item = (detail.pivot_table(index="品項", columns="門市", values="金額", aggfunc="sum", fill_value=0)
               .assign(合計=lambda d: d.sum(axis=1)).sort_values("合計", ascending=False).reset_index())
    by_item.columns.name = None

    daily = (detail.pivot_table(index="日期", columns="門市", values="金額", aggfunc="sum", fill_value=0)
             .assign(合計=lambda d: d.sum(axis=1)).reset_index())
    daily.columns.name = None
    return {"門市彙總": by_store, "品項彙總": by_item, "每日營收": daily}


def style(ws, money_cols=(), pct_cols=(), date_cols=()):
    header = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), header
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    for i, col in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row), start=1):
        name = col[0].value
        fmt = ("#,##0" if name in money_cols else "0.0%" if name in pct_cols
               else "yyyy-mm-dd" if name in date_cols else None)
        if fmt:
            for c in col[1:]:
                c.number_format = fmt
        width = max(len(str(c.value or "")) * (2 if any(ord(ch) > 127 for ch in str(c.value or "")) else 1)
                    for c in col)
        ws.column_dimensions[get_column_letter(i)].width = min(max(width + 2, 10), 40)


def main():
    ap = argparse.ArgumentParser(description="合併多家門市銷售報表")
    ap.add_argument("folder", nargs="?", default="sample_data", help="報表所在資料夾（預設 sample_data）")
    ap.add_argument("-o", "--output", default="整理結果.xlsx", help="輸出檔名（預設 整理結果.xlsx）")
    args = ap.parse_args()

    folder = Path(args.folder)
    if not folder.is_dir():
        sys.exit(f"找不到資料夾：{folder}")

    print(f"讀取 {folder}/ …")
    detail, bad = load_all(folder)
    sheets = {"明細": detail, **summarize(detail), "異常資料": bad}

    with pd.ExcelWriter(args.output, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name, index=False)
            money = [c for c in df.columns if c not in ("日期", "門市", "品項", "數量", "筆數", "總數量",
                                                         "營收占比", "來源檔案", "異常原因")]
            style(xw.sheets[name], money_cols=money, pct_cols=["營收占比"], date_cols=["日期"])

    print(f"完成：{len(detail)} 筆明細、{len(bad)} 筆異常，總營收 {detail['金額'].sum():,.0f} 元")
    print(f"已輸出 {args.output}")


if __name__ == "__main__":
    main()
