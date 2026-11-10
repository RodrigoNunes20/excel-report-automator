"""Generate deliberately messy sample files in ./sample_data (deterministic).

Run:  python scripts/make_sample_data.py
Each file uses a different layout on purpose: delimiters, number formats,
date formats and column names, plus some bad rows, so the cleaning is visible.
"""
from __future__ import annotations

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

OUT = Path(__file__).resolve().parent.parent / "sample_data"
rng = random.Random(42)
_seen: set[tuple] = set()  # keeps the sample free of accidental duplicates

CUSTOMERS = ["Acme Corp", "Blue Ocean Ltd", "Casa Verde", "Delta Foods", "Eco Market", "Fast Wheels", "Green Leaf", "Helios Tech"]
PRODUCTS = {
    "Electronics": [("USB-C Cable", 12.9), ("Wireless Mouse", 39.9), ("Webcam HD", 89.0), ("Headphones", 129.5)],
    "Office": [("Notebook A5", 6.5), ("Desk Lamp", 45.0), ("Ergonomic Chair", 349.0)],
    "Home": [("Coffee Maker", 199.0), ("Air Fryer", 289.9), ("Blender", 149.0)],
}


def make_rows(start: date, days: int, n: int) -> list[dict]:
    rows = []
    while len(rows) < n:
        category = rng.choice(list(PRODUCTS))
        product, price = rng.choice(PRODUCTS[category])
        qty = rng.randint(1, 6)
        row = {
            "day": start + timedelta(days=rng.randint(0, days - 1)),
            "customer": rng.choice(CUSTOMERS),
            "category": category,
            "product": product,
            "qty": qty,
            "amount": round(price * qty, 2),
        }
        key = (row["day"], row["customer"], row["product"], qty)
        if key in _seen:
            continue
        _seen.add(key)
        rows.append(row)
    return rows


def br_number(value: float) -> str:
    return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def main() -> None:
    OUT.mkdir(exist_ok=True)

    # 1) January: comma CSV, ISO dates, US numbers, plus an exact duplicate and an empty line.
    jan = make_rows(date(2026, 1, 1), 31, 40)
    df = pd.DataFrame(
        {
            "Date": [r["day"].isoformat() for r in jan],
            "Customer": [r["customer"] for r in jan],
            "Category": [r["category"] for r in jan],
            "Product": [r["product"] for r in jan],
            "Qty": [r["qty"] for r in jan],
            "Amount": [f"{r['amount']:.2f}" for r in jan],
        }
    )
    df = pd.concat([df, df.iloc[[3]]], ignore_index=True)  # duplicate
    df.loc[len(df)] = [""] * 6  # blank row
    df.to_csv(OUT / "sales_2026_01.csv", index=False)

    # 2) February: semicolon CSV, dd/mm/yyyy dates, Brazilian numbers with currency symbol.
    feb = make_rows(date(2026, 2, 1), 28, 40)
    df = pd.DataFrame(
        {
            "Data": [r["day"].strftime("%d/%m/%Y") for r in feb],
            "Cliente": [f"  {r['customer'].upper()} " if i % 7 == 0 else r["customer"] for i, r in enumerate(feb)],
            "Categoria": [r["category"].lower() for r in feb],
            "Produto": [r["product"] for r in feb],
            "Quantidade": [r["qty"] for r in feb],
            "Valor": ["R$ " + br_number(r["amount"]) for r in feb],
        }
    )
    df.loc[5, "Data"] = "31/02/2026"  # impossible date
    df.loc[11, "Valor"] = "a combinar"  # not a number
    df.loc[17, "Categoria"] = ""  # missing category
    df.to_csv(OUT / "sales_2026_02.csv", index=False, sep=";", encoding="utf-8")

    # 3) March: Excel workbook with one sheet per region and real date cells.
    with pd.ExcelWriter(OUT / "sales_2026_03.xlsx") as writer:
        for region in ("North", "South"):
            mar = make_rows(date(2026, 3, 1), 31, 30)
            pd.DataFrame(
                {
                    "Order Date": [pd.Timestamp(r["day"]) for r in mar],
                    "Client": [r["customer"] for r in mar],
                    "Product Category": [r["category"] for r in mar],
                    "Item": [r["product"] for r in mar],
                    "Units": [r["qty"] for r in mar],
                    "Total": [r["amount"] for r in mar],
                }
            ).to_excel(writer, sheet_name=region, index=False)

    print(f"Sample files written to {OUT}")


if __name__ == "__main__":
    main()
