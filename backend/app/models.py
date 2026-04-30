from sqlalchemy import (
    Boolean,
    Column,
    DateTime,
    ForeignKey,
    Integer,
    JSON,
    LargeBinary,
    String,
)
from sqlalchemy.orm import relationship
from datetime import datetime
from .database import Base


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True, nullable=False)
    email = Column(String, unique=True, index=True, nullable=False)
    hashed_password = Column(String, nullable=False)
    is_active = Column(Boolean, default=True)
    is_admin = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.utcnow)

    datasets = relationship("Dataset", back_populates="user", cascade="all, delete-orphan")


class Dataset(Base):
    """A single Hubspot ↔ RTA match run, persisted for later search + re-download.

    Storage is encrypted at rest via SQLCipher (DB_ENCRYPTION_KEY in env). The
    original uploaded xlsx/csv blobs are NOT stored — only the parsed DataFrames
    (with normalization columns) and the generated Excel result.
    """

    __tablename__ = "datasets"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), index=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, index=True)
    label = Column(String, nullable=True)  # auto-generated from timestamp if null

    # Mappings + cached result data — JSON (TEXT under the hood)
    col_map = Column(JSON, nullable=False)
    stats = Column(JSON, nullable=False)
    conflicts = Column(JSON, nullable=False, default=list)
    flagged = Column(JSON, nullable=False, default=list)
    rta_not_in_hubspot_preview = Column(JSON, nullable=False, default=list)
    hub_output_preview = Column(JSON, nullable=False, default=list)

    # Excel for re-download
    excel_bytes = Column(LargeBinary, nullable=False)

    # Pickled DataFrames — used by the search endpoint to rerun normalization-key lookups
    df_hub_pickle = Column(LargeBinary, nullable=False)
    df_rta_pickle = Column(LargeBinary, nullable=False)

    user = relationship("User", back_populates="datasets")
