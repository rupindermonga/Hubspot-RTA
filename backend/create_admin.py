"""One-time admin bootstrap. Run with: python create_admin.py

Reads ADMIN_USERNAME / ADMIN_PASSWORD / ADMIN_EMAIL from env (Doppler in prd,
.env in dev). Idempotent: if the user already exists, prints a message and exits.
"""
import os
import sys

from dotenv import load_dotenv

load_dotenv()
sys.path.insert(0, os.path.dirname(__file__))

from app.database import Base, SessionLocal, engine
from app.models import User
from passlib.context import CryptContext

Base.metadata.create_all(bind=engine)

USERNAME = os.getenv("ADMIN_USERNAME", "admin")
PASSWORD = os.getenv("ADMIN_PASSWORD", "")
EMAIL = os.getenv("ADMIN_EMAIL", "admin@finel.ai")

if not PASSWORD or len(PASSWORD) < 8:
    print("ERROR: Set ADMIN_PASSWORD in .env / Doppler (minimum 8 characters).")
    sys.exit(1)

pwd = CryptContext(schemes=["bcrypt"], deprecated="auto")
db = SessionLocal()

existing = db.query(User).filter(User.username == USERNAME).first()
if existing:
    if not existing.is_admin:
        existing.is_admin = True
        db.commit()
        print(f"User '{USERNAME}' already exists — is_admin flag set to True.")
    else:
        print(f"User '{USERNAME}' already exists.")
else:
    user = User(
        username=USERNAME,
        email=EMAIL,
        hashed_password=pwd.hash(PASSWORD),
        is_admin=True,
    )
    db.add(user)
    db.commit()
    db.refresh(user)
    print("Admin user created.")

db.close()
print(f"\n  Username : {USERNAME}")
print(f"  Password : {'*' * len(PASSWORD)}")
print(f"\n  Open: http://localhost:8000\n")
