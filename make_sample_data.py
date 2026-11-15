"""產生虛構的範例報表（三家門市，格式刻意不一致，用來示範整理功能）。

執行：python make_sample_data.py
"""
import random
from pathlib import Path

import pandas as pd

OUT = Path(__file__).parent / "sample_data"
ITEMS = {"美式咖啡": 80, "拿鐵": 110, "桂花冷萃": 130, "可頌": 65, "肉桂捲": 75}
random.seed(42)


def rows(n):
    seen = set()  # 避免隨機產生出一模一樣的列，只保留刻意放入的重複資料
    while len(seen) < n:
        item = random.choice(list(ITEMS))
        key = (random.randint(1, 31), item, random.randint(1, 6))
        if key not in seen:
            seen.add(key)
            yield *key, ITEMS[item]


def main():
    OUT.mkdir(exist_ok=True)

    # 台北店：最標準的格式
    df = pd.DataFrame([{"日期": f"2026-10-{d:02d}", "品項": i, "數量": q, "單價": p}
                       for d, i, q, p in rows(40)])
    df.to_excel(OUT / "台北店_2026-10.xlsx", index=False)

    # 台中店：欄位名稱不同、數字是含千分位的文字、夾雜空白列與重複列
    recs = [{"銷售日期": f"2026/10/{d}", "商品名稱": f" {i} ", "銷售數量": str(q),
             "單價(元)": f"{p:,}"} for d, i, q, p in rows(35)]
    recs.insert(10, {k: None for k in recs[0]})
    recs.append(dict(recs[3]))
    recs.append({"銷售日期": "2026/10/15", "商品名稱": "拿鐵", "銷售數量": "兩杯", "單價(元)": "110"})
    pd.DataFrame(recs).to_excel(OUT / "台中店_2026-10.xlsx", index=False)

    # 高雄店：CSV、英文欄位、日期格式又不一樣
    df = pd.DataFrame([{"date": f"10/{d:02d}/2026", "item": i, "qty": q, "price": p}
                       for d, i, q, p in rows(30)])
    df.to_csv(OUT / "高雄店_2026-10.csv", index=False, encoding="utf-8-sig")
    print(f"已產生範例資料於 {OUT}")


if __name__ == "__main__":
    main()
