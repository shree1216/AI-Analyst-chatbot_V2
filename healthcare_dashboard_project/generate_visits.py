"""Generate a 100-row synthetic healthcare visits dataset (Aug-Oct 2026, USD).

The raw extract intentionally contains a few duplicates, incomplete records and inconsistent
text so the VBA macro has something to clean.
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

ROWS = 100
DUPLICATES = 4
random.seed(2026)

# Department -> (doctors, minimum bill, maximum bill, weight)
DEPARTMENTS = {
    "Cardiology": (["Dr. Alan Reed", "Dr. Maria Lopez"], 250, 600, 18),
    "Orthopedics": (["Dr. Kevin Shaw", "Dr. Priya Nair"], 200, 500, 20),
    "Pediatrics": (["Dr. Laura Chen", "Dr. Omar Haddad"], 80, 200, 22),
    "Dermatology": (["Dr. Rachel Green", "Dr. Samuel Ortiz"], 100, 300, 16),
    "General Medicine": (["Dr. John Miller", "Dr. Emily Stone"], 50, 150, 24),
}
MONTHS = {(2026, 8): 28, (2026, 9): 33, (2026, 10): 39}   # rising patient volume
STATUS = {"Completed": 80, "Cancelled": 12, "No-Show": 8}
FIRST = ["James", "Emily", "Michael", "Sarah", "David", "Olivia", "Daniel", "Sophia", "Robert", "Emma",
         "William", "Ava", "Joseph", "Mia", "Thomas", "Grace", "Charles", "Chloe", "Andrew", "Lily"]
LAST = ["Carter", "Johnson", "Brown", "Davis", "Wilson", "Martinez", "Anderson", "Taylor", "Moore",
        "Jackson", "White", "Harris", "Clark", "Lewis", "Walker", "Young"]


def random_date():
    year, month = random.choices(list(MONTHS), weights=MONTHS.values())[0]
    first = date(year, month, 1)
    next_month = date(year + (month == 12), month % 12 + 1, 1)
    return first + timedelta(days=random.randrange((next_month - first).days))


def messy(value, chance=0.25):
    if random.random() >= chance:
        return value
    return random.choice([value.lower(), value.upper(), f" {value}  "])


def build_visits(count):
    patients = [(f"PT{1000 + i}", f"{random.choice(FIRST)} {random.choice(LAST)}") for i in range(60)]
    visits = []
    for visit_date in sorted(random_date() for _ in range(count)):
        patient_id, name = random.choice(patients)
        department = random.choices(list(DEPARTMENTS), weights=[d[3] for d in DEPARTMENTS.values()])[0]
        doctors, low, high, _ = DEPARTMENTS[department]
        status = random.choices(list(STATUS), weights=STATUS.values())[0]
        bill = f"{round(random.uniform(low, high) / 5) * 5:.2f}" if status == "Completed" else random.choice(["", "0.00"])
        visits.append({
            "Visit ID": "",
            "Patient ID": patient_id,
            "Patient Name": messy(name),
            "Department": messy(department),
            "Doctor": random.choice(doctors),
            "Appointment Date": visit_date.isoformat(),
            "Visit Type": messy(random.choice(["New", "Follow-up"]), 0.15),
            "Status": messy(status),
            "Billing Amount": bill,
        })
    for number, visit in enumerate(visits, start=1):
        visit["Visit ID"] = f"V{number:04d}"
    return visits


def main():
    visits = build_visits(ROWS - DUPLICATES)
    completed = [v for v in visits if v["Status"].strip().lower() == "completed"]
    # Incomplete records: missing department, missing date, completed visit without a bill.
    targets = random.sample(completed, 3)
    targets[0]["Department"] = ""
    targets[1]["Appointment Date"] = ""
    targets[2]["Billing Amount"] = ""
    # Duplicates: the same visit exported twice.
    for original in random.sample([v for v in visits if v not in targets], DUPLICATES):
        visits.insert(visits.index(original) + random.randint(1, 5), dict(original))

    output = Path(__file__).with_name("healthcare_visits_raw_100.csv")
    with output.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(visits[0]))
        writer.writeheader()
        writer.writerows(visits)
    print(f"{len(visits)} rows -> {output.name} ({DUPLICATES} duplicates, {len(targets)} incomplete)")


if __name__ == "__main__":
    main()
