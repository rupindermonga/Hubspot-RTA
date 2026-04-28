from pydantic import BaseModel, EmailStr, Field, field_validator
from typing import Optional, List
from datetime import datetime
import re

_USERNAME_RE = re.compile(r"^[a-zA-Z0-9_\-]+$")


# ─── Auth ────────────────────────────────────────────────────────────────────

class UserLogin(BaseModel):
    username: str = Field(..., min_length=1, max_length=50)
    password: str = Field(..., min_length=1, max_length=128)


class UserOut(BaseModel):
    id: int
    username: str
    email: str
    is_active: bool
    is_admin: bool = False
    created_at: datetime

    class Config:
        from_attributes = True


class Token(BaseModel):
    access_token: str
    token_type: str
    user: UserOut


class ChangePassword(BaseModel):
    current_password: str = Field(..., min_length=1)
    new_password: str = Field(..., min_length=8, max_length=128)


# ─── Matcher (Hubspot vs RTA) ────────────────────────────────────────────────

class MatcherStats(BaseModel):
    hubspot_total: int
    hubspot_matched: int
    hubspot_unmatched: int
    rta_total: int
    rta_in_hubspot: int
    rta_not_in_hubspot: int
    exact: int
    fuzzy: int
    conflict: int
    risky_no_pc: int


class MatcherConflictRow(BaseModel):
    address: str
    postal_code: str
    status: str
    key: str


class MatcherFlaggedRow(BaseModel):
    street_address: str
    postal_code: str
    rta_address: str
    rta_status: str
    match_type: str  # "exact" | "fuzzy" | "direction_strip" | "conflict" | "no_pc"


class MatcherResponse(BaseModel):
    download_token: str  # opaque key the client uses to GET the Excel
    stats: MatcherStats
    conflicts: List[MatcherConflictRow]
    flagged: List[MatcherFlaggedRow]
    rta_not_in_hubspot_preview: List[dict]


# ─── Compare (Old RTA vs New RTA) ────────────────────────────────────────────

class CompareStats(BaseModel):
    old_total: int
    new_total: int
    removed: int
    added: int
    status_changed: int
    in_construction_to_rta: int
    rta_to_in_construction: int
    other_status_changes: int
    conflict_addresses: int
    conflict_rows: int


class CompareStatusChangeRow(BaseModel):
    address_number: str
    street_name: str
    locality: str
    postal_code: str
    old_status: str
    new_status: str


class CompareResponse(BaseModel):
    download_token: str
    stats: CompareStats
    removed_preview: List[dict]
    added_preview: List[dict]
    status_changed_preview: List[CompareStatusChangeRow]
    conflicts_preview: List[dict]
