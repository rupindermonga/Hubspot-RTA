"""SQLAlchemy engine wired through SQLCipher (encrypted SQLite at rest).

The DB file is opaque without DB_ENCRYPTION_KEY. Set it via .env in dev or
Doppler in prd. The key is read once at process start and held in this
module's `engine`; decrypt happens transparently on every query.

If the key is missing or wrong, opening the DB raises and the service won't start.
"""
import os
from pathlib import Path

import sqlcipher3
from dotenv import load_dotenv
from sqlalchemy import create_engine, event
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

load_dotenv()

_raw_url = os.getenv("DATABASE_URL", "sqlite:///./hubspot_rta.db")

# Resolve relative SQLite paths to absolute (relative to *this* file's directory)
# so the DB location is stable regardless of the process's working directory.
if _raw_url.startswith("sqlite:///./") or _raw_url.startswith("sqlite:///hubspot_rta"):
    _rel = _raw_url.replace("sqlite:///", "")
    DB_PATH = str(Path(__file__).resolve().parent.parent / _rel)
elif _raw_url.startswith("sqlite:////"):
    # Absolute path form: sqlite:////var/lib/.../db.db
    DB_PATH = _raw_url.replace("sqlite:///", "/")
else:
    raise RuntimeError(
        f"DATABASE_URL must be a sqlite:// URL for SQLCipher (got {_raw_url!r}). "
        "If you need Postgres later, the SQLCipher path here will need to be conditionalized."
    )

DB_ENCRYPTION_KEY = os.getenv("DB_ENCRYPTION_KEY", "").strip()
if not DB_ENCRYPTION_KEY or DB_ENCRYPTION_KEY == "CHANGE_ME_to_64_char_hex":
    raise RuntimeError(
        "DB_ENCRYPTION_KEY is not set. Generate with: "
        'python -c "import secrets; print(secrets.token_hex(32))"  '
        "and put it in .env (dev) or Doppler (prd)."
    )


def _connect():
    """Open the encrypted DB. Called by SQLAlchemy on every new connection."""
    conn = sqlcipher3.connect(DB_PATH, check_same_thread=False)
    # Quote the key as a SQL literal — single quote it; SQLCipher wants
    # `PRAGMA key = '<value>'`. Hex-only keys can't contain single quotes,
    # but escape defensively in case someone uses a passphrase later.
    safe_key = DB_ENCRYPTION_KEY.replace("'", "''")
    conn.execute(f"PRAGMA key = '{safe_key}'")
    # Touch the DB to force the key check; if wrong, this raises.
    conn.execute("SELECT count(*) FROM sqlite_master")
    return conn


# Use creator= so SQLAlchemy doesn't try to open the DB itself —
# we hand it pre-keyed connections.
engine = create_engine(
    "sqlite://",
    creator=_connect,
    # SQLite (and SQLCipher) wants this so the engine can be used across threads.
    connect_args={"check_same_thread": False},
)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
