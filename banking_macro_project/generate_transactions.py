"""Generate a simple, synthetic 100-row banking transaction dataset (USD)."""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

ROWS = 100
random.seed(2026)

CUSTOMERS = [
    ("1002345671", "James Carter", "New York"),
    ("1002345688", "Emily Johnson", "New York"),
    ("1002345695", "Michael Brown", "Chicago"),
    ("1002345702", "Sarah Davis", "Chicago"),
    ("1002345719", "David Wilson", "Houston"),
    ("1002345726", "Olivia Martinez", "Houston"),
    ("1002345733", "Daniel Anderson", "Dallas"),
    ("1002345740", "Sophia Thomas", "Dallas"),
    ("1002345757", "Robert Taylor", "Boston"),
    ("1002345764", "Emma Moore", "Boston"),
    ("1002345771", "William Jackson", "Seattle"),
    ("1002345788", "Ava White", "Seattle"),
]

# Transaction type -> (side, minimum, maximum)
TYPES = {
    "Cash Deposit": ("Credit", 200, 5000),
    "Salary Credit": ("Credit", 3000, 8000),
    "Transfer In": ("Credit", 100, 4000),
    "Interest Credit": ("Credit", 5, 80),
    "Cash Withdrawal": ("Debit", 100, 2000),
    "Bill Payment": ("Debit", 50, 600),
    "Loan EMI": ("Debit", 500, 2500),
    "Transfer Out": ("Debit", 100, 3500),
    "Card Purchase": ("Debit", 20, 900),
}
WEIGHTS = [14, 6, 10, 4, 14, 16, 8, 12, 16]


def build_rows():
    start = date(2026, 11, 2)
    business_days = [start + timedelta(days=i) for i in range(12)]
    business_days = [d for d in business_days if d.weekday() < 5][:10]
    dates = sorted(random.choice(business_days) for _ in range(ROWS))
    rows = []
    for number, txn_date in enumerate(dates, start=1):
        account, name, branch = random.choice(CUSTOMERS)
        txn_type = random.choices(list(TYPES), weights=WEIGHTS)[0]
        side, low, high = TYPES[txn_type]
        amount = round(random.uniform(low, high), 2)
        if txn_type == "Salary Credit":
            amount = float(round(amount, -2))
        rows.append({
            "Transaction ID": f"TXN{number:05d}",
            "Date": txn_date.isoformat(),
            "Account Number": account,
            "Customer Name": name,
            "Branch": branch,
            "Transaction Type": txn_type,
            "Debit": f"{amount:.2f}" if side == "Debit" else "",
            "Credit": f"{amount:.2f}" if side == "Credit" else "",
        })
    return rows


def main():
    rows = build_rows()
    output = Path(__file__).with_name("banking_transactions_100.csv")
    with output.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    debit = sum(float(r["Debit"] or 0) for r in rows)
    credit = sum(float(r["Credit"] or 0) for r in rows)
    print(f"{len(rows)} rows -> {output.name}")
    print(f"Total Debit {debit:,.2f} | Total Credit {credit:,.2f} | Net {credit - debit:,.2f}")


if __name__ == "__main__":
    main()
