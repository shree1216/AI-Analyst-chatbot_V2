import random
from pathlib import Path

import pandas as pd
from faker import Faker

ROW_COUNT = 150
OUTPUT_FILE = Path(__file__).with_name("employees_data.xlsx")
DEPARTMENTS = ["Sales", "IT", "HR", "Finance", "Marketing", "Operations"]
GENDERS = ["Male", "Female"]
CITIES = ["Delhi", "Mumbai", "Bengaluru", "Chennai", "Hyderabad", "Kolkata", "Pune"]
STATUSES = ["Active", "On Leave", "Inactive"]

fake = Faker()
Faker.seed(42)
random.seed(42)

rows = []
for employee_number in range(1, ROW_COUNT + 1):
    rows.append(
        {
            "emp_id": 1000 + employee_number,
            "name": fake.name(),
            "age": random.randint(20, 60),
            "gender": random.choice(GENDERS),
            "department": random.choice(DEPARTMENTS),
            "city": random.choice(CITIES),
            "status": STATUSES[(employee_number - 1) % len(STATUSES)],
            "job_title": fake.job(),
            "salary": random.randint(30000, 150000),
            "join_date": fake.date_between(start_date="-10y", end_date="today"),
            "email": fake.email(),
            "phone_number": fake.phone_number(),
            "rating": round(random.uniform(1, 5), 1),
        }
    )

employees_data = pd.DataFrame(rows)
employees_data.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")

print(
    f"Generated {len(employees_data)} rows x {len(employees_data.columns)} columns "
    f"-> {OUTPUT_FILE.name}"
)