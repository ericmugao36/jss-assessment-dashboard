
from dotenv import load_dotenv
import os
from pathlib import Path
from sqlalchemy import select

load_dotenv()

from app import DB, User, engine

print("Database URL:", engine.url)
print("Database file:", Path("school.db").resolve())
print("Teacher email in .env:", os.getenv("TEACHER_EMAIL"))

with DB() as db:
    users = db.scalars(select(User)).all()

    if not users:
        print("No user accounts found in this database.")
    else:
        print("\nAccounts found:")
        for user in users:
            print(f"Email: {user.email} | Role: {user.role}")
