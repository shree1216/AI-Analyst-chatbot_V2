"""Generate a simple, synthetic 100-employee HR database."""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

EMPLOYEES = 100
random.seed(2026)

MALE = ["James", "Michael", "David", "Daniel", "Robert", "William", "Joseph", "Thomas", "Charles",
        "Andrew", "Ryan", "Kevin", "Brian", "George", "Edward", "Steven"]
FEMALE = ["Emily", "Sarah", "Olivia", "Sophia", "Emma", "Ava", "Mia", "Grace", "Chloe", "Lily",
          "Hannah", "Zoe", "Nancy", "Linda", "Susan", "Laura"]
LAST = ["Carter", "Johnson", "Brown", "Davis", "Wilson", "Martinez", "Anderson", "Thomas", "Taylor",
        "Moore", "Jackson", "White", "Harris", "Clark", "Lewis", "Walker", "Young", "King", "Scott", "Green"]
# Department -> [(job title, minimum annual salary, maximum annual salary)]
DEPARTMENTS = {
    "Finance": [("Accountant", 55000, 75000), ("Financial Analyst", 65000, 90000), ("Finance Manager", 95000, 130000)],
    "Human Resources": [("HR Executive", 45000, 60000), ("Recruiter", 50000, 68000), ("HR Manager", 85000, 110000)],
    "IT": [("Software Engineer", 80000, 120000), ("Systems Administrator", 65000, 90000), ("IT Manager", 110000, 145000)],
    "Marketing": [("Marketing Executive", 48000, 65000), ("Content Specialist", 50000, 70000), ("Marketing Manager", 90000, 120000)],
    "Operations": [("Operations Associate", 40000, 55000), ("Logistics Coordinator", 48000, 62000), ("Operations Manager", 85000, 115000)],
    "Sales": [("Sales Executive", 42000, 60000), ("Account Manager", 60000, 85000), ("Sales Manager", 90000, 125000)],
}
STATUS = {"Active": 88, "On Leave": 7, "Resigned": 5}


def main():
    rows, emails = [], set()
    for number in range(1, EMPLOYEES + 1):
        gender = random.choice(["Male", "Female"])
        first = random.choice(MALE if gender == "Male" else FEMALE)
        last = random.choice(LAST)
        email = f"{first.lower()}.{last.lower()}@company.com"
        suffix = 2
        while email in emails:
            email = f"{first.lower()}.{last.lower()}{suffix}@company.com"
            suffix += 1
        emails.add(email)
        department = random.choice(list(DEPARTMENTS))
        title, low, high = random.choices(DEPARTMENTS[department], weights=[5, 3, 1])[0]
        hire_date = date(2015, 1, 1) + timedelta(days=random.randint(0, (date(2026, 10, 31) - date(2015, 1, 1)).days))
        rows.append({
            "Employee ID": f"EMP{number:03d}",
            "Full Name": f"{first} {last}",
            "Gender": gender,
            "Department": department,
            "Job Title": title,
            "Email": email,
            "Phone": f"555-{random.randint(100, 999)}-{random.randint(1000, 9999)}",
            "Hire Date": hire_date.isoformat(),
            "Annual Salary": round(random.uniform(low, high) / 500) * 500,
            "Status": random.choices(list(STATUS), weights=STATUS.values())[0],
        })
    output = Path(__file__).with_name("employees_100.csv")
    with output.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    print(f"{len(rows)} employees -> {output.name}")


if __name__ == "__main__":
    main()
