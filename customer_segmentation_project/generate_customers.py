"""Generate a 100-row synthetic banking customer dataset with realistic data-quality issues.

The raw file intentionally contains duplicates, incomplete records and inconsistent formatting
so the VBA macro has something to clean.
"""
import csv
import random
from pathlib import Path

ROWS = 100
DUPLICATES = 8
random.seed(2026)

FIRST = ["James", "Emily", "Michael", "Sarah", "David", "Olivia", "Daniel", "Sophia", "Robert", "Emma",
         "William", "Ava", "Joseph", "Mia", "Thomas", "Grace", "Charles", "Chloe", "Andrew", "Lily",
         "Ryan", "Hannah", "Kevin", "Zoe", "Brian", "Nancy", "George", "Linda", "Edward", "Susan"]
LAST = ["Carter", "Johnson", "Brown", "Davis", "Wilson", "Martinez", "Anderson", "Thomas", "Taylor",
        "Moore", "Jackson", "White", "Harris", "Clark", "Lewis", "Walker", "Young", "King"]
CITIES = ["New York", "Chicago", "Houston", "Dallas", "Boston", "Seattle"]
ACCOUNT_TYPES = ["Savings", "Current", "Fixed Deposit"]
# Loan type -> (weight, minimum amount, maximum amount)
LOANS = {"Home Loan": (30, 80000, 450000), "Car Loan": (25, 10000, 60000),
         "Personal Loan": (20, 2000, 25000), "None": (25, 0, 0)}
# Messy spellings found in the raw extract, standardized by the macro.
LOAN_VARIANTS = {"Home Loan": ["home loan", "HOME LOAN", "Housing Loan", "Home"],
                 "Car Loan": ["car loan", "CAR LOAN", "Auto Loan", "Car"],
                 "Personal Loan": ["personal loan", "PERSONAL LOAN"],
                 "None": ["", "none", "No Loan"]}


def messy(value, variants, chance=0.35):
    return random.choice(variants) if random.random() < chance else value


def phone_text(digits, chance=0.4):
    formats = [f"({digits[:3]}) {digits[3:6]}-{digits[6:]}", f"{digits[:3]}.{digits[3:6]}.{digits[6:]}",
               digits, f"+1 {digits[:3]} {digits[3:6]} {digits[6:]}"]
    return random.choice(formats) if random.random() < chance else f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"


def build_customer(number):
    first, last = random.choice(FIRST), random.choice(LAST)
    loan = random.choices(list(LOANS), weights=[v[0] for v in LOANS.values()])[0]
    low, high = LOANS[loan][1:]
    age = random.choice([random.randint(21, 59), random.randint(21, 59), random.randint(60, 82)])
    digits = f"555{random.randint(1000000, 9999999)}"
    return {
        "Customer ID": f"CUST{number:04d}",
        "Customer Name": messy(f"{first} {last}", [f"{first.lower()} {last.lower()}", f"{first.upper()} {last.upper()}", f"  {first}  {last} "]),
        "Age": age,
        "City": messy(random.choice(CITIES), [random.choice(CITIES).lower(), random.choice(CITIES).upper() + " "]),
        "Phone": phone_text(digits),
        "Email": messy(f"{first.lower()}.{last.lower()}{number % 97}@example.com", [f"{first.upper()}.{last.upper()}{number % 97}@EXAMPLE.COM"]),
        "Account Type": messy(random.choice(ACCOUNT_TYPES), [random.choice(ACCOUNT_TYPES).lower(), random.choice(ACCOUNT_TYPES).upper()]),
        "Loan Type": messy(loan, LOAN_VARIANTS[loan], 0.45),
        "Loan Amount": round(random.uniform(low, high) / 500) * 500 if high else "",
    }


def main():
    customers = [build_customer(number) for number in range(1001, 1001 + ROWS - DUPLICATES)]
    # Incomplete records: one required field missing or invalid.
    problems = [("Customer Name", ""), ("Age", ""), ("Phone", ""), ("City", ""),
                ("Age", "N/A"), ("Phone", "555-12"), ("Customer Name", "  ")]
    incomplete = random.sample(range(len(customers)), len(problems))
    for index, (field, value) in zip(incomplete, problems):
        customers[index][field] = value
    # Duplicates: the same customer entered again later with different formatting.
    complete = [c for i, c in enumerate(customers) if i not in incomplete]
    for original in random.sample(complete, DUPLICATES):
        copy = dict(original, **{"Customer Name": original["Customer Name"].strip().upper()})
        position = random.randint(customers.index(original) + 1, len(customers))
        customers.insert(position, copy)

    output = Path(__file__).with_name("bank_customers_raw_100.csv")
    with output.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(customers[0]))
        writer.writeheader()
        writer.writerows(customers)
    print(f"{len(customers)} rows -> {output.name} ({DUPLICATES} duplicates, {len(problems)} incomplete)")


if __name__ == "__main__":
    main()
