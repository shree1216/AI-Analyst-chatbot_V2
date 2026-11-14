"""Generate a simple, synthetic 100-row retail sales dataset for one month (USD)."""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

ROWS = 100
random.seed(2026)

# Product -> (category, unit price)
PRODUCTS = {
    "Wireless Mouse": ("Electronics", 25.00),
    "USB-C Charger": ("Electronics", 30.00),
    "Bluetooth Headphones": ("Electronics", 120.00),
    "Smart Watch": ("Electronics", 250.00),
    "Laptop Backpack": ("Accessories", 60.00),
    "Leather Wallet": ("Accessories", 40.00),
    "Coffee Maker": ("Home & Kitchen", 90.00),
    "Air Fryer": ("Home & Kitchen", 140.00),
    "Desk Lamp": ("Home & Kitchen", 45.00),
    "Running Shoes": ("Clothing", 85.00),
    "Denim Jacket": ("Clothing", 70.00),
    "Cotton T-Shirt": ("Clothing", 20.00),
}
REGIONS = {"North": 30, "West": 28, "South": 22, "East": 20}


def build_rows():
    start = date(2026, 10, 1)
    dates = sorted(start + timedelta(days=random.randint(0, 30)) for _ in range(ROWS))
    rows = []
    for number, order_date in enumerate(dates, start=1):
        product = random.choice(list(PRODUCTS))
        category, price = PRODUCTS[product]
        max_quantity = 3 if price >= 200 else 6 if price >= 80 else 10
        quantity = random.randint(1, max_quantity)
        rows.append({
            "Order ID": f"ORD-{number:04d}",
            "Order Date": order_date.isoformat(),
            "Product": product,
            "Category": category,
            "Region": random.choices(list(REGIONS), weights=REGIONS.values())[0],
            "Quantity": quantity,
            "Unit Price": f"{price:.2f}",
            "Sales Amount": f"{quantity * price:.2f}",
        })
    return rows


def main():
    rows = build_rows()
    output = Path(__file__).with_name("retail_sales_oct_2026.csv")
    with output.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    total = sum(float(r["Sales Amount"]) for r in rows)
    print(f"{len(rows)} rows -> {output.name} | total sales {total:,.2f}")


if __name__ == "__main__":
    main()
