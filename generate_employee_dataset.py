import random

import pandas as pd
from faker import Faker

fake = Faker()
Faker.seed(42)
random.seed(42)

DEPARTMENTS = ["Sales", "IT", "HR", "Finance", "Marketing", "Operations"]
GENDERS = ["Male", "Female"]
CITIES = ["Delhi", "Mumbai", "Bengaluru", "Chennai", "Hyderabad", "Kolkata", "Pune"]
STATUSES = ["Active", "On Leave", "Inactive"]

rows = []
for i in range(1, 151):
    rows.append({
        "Employee ID": f"EMP{i:03d}",
        "Name": fake.name(),
        "Age": random.randint(20, 60),
        "Gender": random.choice(GENDERS),
        "Department": random.choice(DEPARTMENTS),
        "City": random.choice(CITIES),
        "Status": STATUSES[(i - 1) % len(STATUSES)],
        "Job Title": fake.job(),
        "Salary": random.randint(30000, 150000),
        "Joining Date": fake.date_between(start_date="-10y", end_date="today"),
        "Email": fake.email(),
        "Phone Number": fake.phone_number(),
    })

df = pd.DataFrame(rows)
df.loc[df["Employee ID"] == "EMP101", "Name"] = "Jaah"
df.loc[df["Employee ID"] == "EMP102", "Name"] = "Shree"
df.loc[df["Employee ID"] == "EMP103", "Name"] = "Jahnavi"
df.loc[df["Employee ID"] == "EMP104", "Name"] = (
    "  " + df.loc[df["Employee ID"] == "EMP104", "Name"]
)
df.loc[df["Employee ID"] == "EMP105", "Name"] = (
    df.loc[df["Employee ID"] == "EMP105", "Name"] + "  "
)
df.loc[df["Employee ID"] == "EMP106", "Name"] = (
    "  " + df.loc[df["Employee ID"] == "EMP106", "Name"] + "  "
)
df.to_excel("employee_dataset.xlsx", index=False, engine="openpyxl")

print(f"Generated {df.shape[0]} rows x {df.shape[1]} columns -> employee_dataset.xlsx")
