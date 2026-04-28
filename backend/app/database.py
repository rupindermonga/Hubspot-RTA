from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

_raw_url = os.getenv("DATABASE_URL", "sqlite:///./hubspot_rta.db")

# Resolve relative SQLite paths to absolute (relative to *this* file's directory)
# so the DB location is stable regardless of the process's working directory.
if _raw_url.startswith("sqlite:///./") or _raw_url.startswith("sqlite:///hubspot_rta"):
    _rel = _raw_url.replace("sqlite:///", "")
    _abs = str(Path(__file__).resolve().parent.parent / _rel)
    DATABASE_URL = f"sqlite:///{_abs}"
else:
    DATABASE_URL = _raw_url

engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False} if "sqlite" in DATABASE_URL else {}
)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
