"""Generate 100 synthetic employee records for the HR Employee Management workbook.

Run:  python generate_employees.py      (writes employees.csv next to this file)
"""

from __future__ import annotations

import csv
import random
from datetime import date, timedelta
from pathlib import Path

SEED = 2026
ROWS = 100

DEPARTMENTS = {
    "Human Resources": [("HR Executive", 420_000, 650_000), ("HR Manager", 900_000, 1_400_000),
                        ("Recruiter", 380_000, 600_000)],
    "Finance": [("Accountant", 450_000, 700_000), ("Financial Analyst", 650_000, 1_000_000),
                ("Finance Manager", 1_200_000, 1_900_000)],
    "IT": [("Software Engineer", 700_000, 1_500_000), ("System Administrator", 550_000, 900_000),
           ("IT Manager", 1_500_000, 2_400_000)],
    "Sales": [("Sales Executive", 380_000, 650_000), ("Account Manager", 700_000, 1_100_000),
              ("Sales Manager", 1_100_000, 1_800_000)],
    "Marketing": [("Marketing Executive", 400_000, 650_000), ("Content Strategist", 550_000, 850_000),
                  ("Marketing Manager", 1_100_000, 1_700_000)],
    "Operations": [("Operations Executive", 380_000, 600_000), ("Logistics Coordinator", 450_000, 700_000),
                   ("Operations Manager", 1_100_000, 1_700_000)],
}
FIRST_NAMES = {
    "Male": ["Aarav", "Vihaan", "Arjun", "Rohan", "Karan", "Rahul", "Siddharth", "Aditya", "Nikhil", "Varun",
             "Manish", "Suresh", "Kiran", "Pranav", "Deepak", "Amit", "Rajesh", "Harsh", "Yash", "Gaurav"],
    "Female": ["Ananya", "Diya", "Priya", "Sneha", "Kavya", "Meera", "Isha", "Pooja", "Neha", "Riya",
               "Shreya", "Divya", "Lakshmi", "Nandini", "Aishwarya", "Tanvi", "Sanjana", "Pallavi", "Swati", "Anjali"],
}
LAST_NAMES = ["Sharma", "Verma", "Iyer", "Reddy", "Nair", "Patel", "Gupta", "Rao", "Menon", "Kulkarni",
              "Joshi", "Mehta", "Pillai", "Desai", "Chopra", "Bose", "Singh", "Kapoor", "Shetty", "Agarwal"]
STATUSES = ["Active", "On Leave", "Resigned"]
STATUS_WEIGHTS = [85, 7, 8]
COLUMNS = ["Employee ID", "First Name", "Last Name", "Gender", "Department", "Job Title", "Email",
           "Phone", "Hire Date", "Annual Salary", "Status"]


def build_rows(seed: int = SEED, count: int = ROWS) -> list[dict[str, object]]:
    rng = random.Random(seed)
    used_emails: set[str] = set()
    rows = []
    start, end = date(2015, 1, 1), date(2025, 12, 31)
    for index in range(count):
        gender = rng.choice(["Male", "Female"])
        first, last = rng.choice(FIRST_NAMES[gender]), rng.choice(LAST_NAMES)
        department = rng.choice(list(DEPARTMENTS))
        title, low, high = rng.choice(DEPARTMENTS[department])
        email = f"{first}.{last}".lower() + "@company.com"
        suffix = 2
        while email in used_emails:
            email = f"{first}.{last}{suffix}".lower() + "@company.com"
            suffix += 1
        used_emails.add(email)
        rows.append({
            "Employee ID": f"EMP{1001 + index}",
            "First Name": first,
            "Last Name": last,
            "Gender": gender,
            "Department": department,
            "Job Title": title,
            "Email": email,
            "Phone": f"9{rng.randint(100_000_000, 999_999_999)}",
            "Hire Date": (start + timedelta(days=rng.randint(0, (end - start).days))).isoformat(),
            "Annual Salary": round(rng.randint(low, high), -3),
            "Status": rng.choices(STATUSES, STATUS_WEIGHTS)[0],
        })
    return rows


def main() -> None:
    target = Path(__file__).with_name("employees.csv")
    with target.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(build_rows())
    print(f"Wrote {ROWS} employees to {target}")


if __name__ == "__main__":
    main()
