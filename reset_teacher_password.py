from getpass import getpass

from dotenv import load_dotenv
from sqlalchemy import select
from pwdlib import PasswordHash

from app import DB, User

load_dotenv()
passwords = PasswordHash.recommended()

email = input("Enter teacher email: ").strip().lower()
new_password = getpass(
    "Enter new password (at least 12 characters): "
)

if len(new_password) < 12:
    print("Password must contain at least 12 characters.")
else:
    with DB() as db:
        teacher = db.scalar(
            select(User).where(
                User.email == email,
                User.role == "teacher"
            )
        )

        if teacher is None:
            print("Teacher account not found. Check the email.")
        else:
            teacher.password_hash = passwords.hash(new_password)
            db.commit()
            print("Teacher password updated successfully.")