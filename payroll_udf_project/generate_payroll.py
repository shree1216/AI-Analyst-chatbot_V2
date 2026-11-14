"""Generate a simple, synthetic 50-employee payroll dataset (monthly salaries in USD)."""
import csv
import random
from pathlib import Path

EMPLOYEES = 50
random.seed(2026)

FIRST_NAMES = ["James", "Emily", "Michael", "Sarah", "David", "Olivia", "Daniel", "Sophia", "Robert",
               "Emma", "William", "Ava", "Joseph", "Mia", "Thomas", "Grace", "Charles", "Chloe",
               "Andrew", "Lily", "Ryan", "Hannah", "Kevin", "Zoe", "Brian"]
LAST_NAMES = ["Carter", "Johnson", "Brown", "Davis", "Wilson", "Martinez", "Anderson", "Thomas",
              "Taylor", "Moore", "Jackson", "White", "Harris", "Clark", "Lewis", "Walker"]

# Department -> [(job title, minimum monthly salary, maximum monthly salary)]
DEPARTMENTS = {
    "Finance": [("Financial Analyst", 4500, 7000), ("Finance Manager", 8000, 11000)],
    "Human Resources": [("HR Executive", 3500, 5500), ("HR Manager", 7000, 9500)],
    "IT": [("Software Engineer", 5500, 9000), ("IT Manager", 9000, 12000)],
    "Sales": [("Sales Executive", 3000, 5000), ("Sales Manager", 7000, 10000)],
    "Operations": [("Operations Associate", 3000, 4800), ("Operations Manager", 7500, 10000)],
}
# Rating 1 (poor) .. 5 (outstanding); most employees meet or exceed expectations.
RATING_WEIGHTS = {1: 5, 2: 10, 3: 40, 4: 30, 5: 15}


def build_rows():
    names = random.sample([f"{f} {l}" for f in FIRST_NAMES for l in LAST_NAMES], EMPLOYEES)
    rows = []
    for number, name in enumerate(names, start=1):
        department = random.choice(list(DEPARTMENTS))
        title, low, high = random.choices(DEPARTMENTS[department], weights=[3, 1])[0]
        rows.append({
            "Employee ID": f"EMP{number:03d}",
            "Employee Name": name,
            "Department": department,
            "Job Title": title,
            "Monthly Salary": round(random.uniform(low, high) / 50) * 50,
            "Performance Rating": random.choices(list(RATING_WEIGHTS), weights=RATING_WEIGHTS.values())[0],
        })
    return rows


def main():
    rows = build_rows()
    output = Path(__file__).with_name("employee_payroll_50.csv")
    with output.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    print(f"{len(rows)} employees -> {output.name}")


if __name__ == "__main__":
    main()
