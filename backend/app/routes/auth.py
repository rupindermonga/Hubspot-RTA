from fastapi import APIRouter, Depends, HTTPException, Request
from sqlalchemy.orm import Session
from datetime import datetime, timedelta
from collections import defaultdict
from jose import jwt
from passlib.context import CryptContext
import os
import time

from ..database import get_db
from ..models import User
from ..schemas import UserLogin, UserOut, Token, ChangePassword
from ..dependencies import get_current_user, SECRET_KEY, ALGORITHM

router = APIRouter(prefix="/api/auth", tags=["auth"])

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
EXPIRE_MINUTES = int(os.getenv("JWT_EXPIRE_MINUTES", "10080"))

# ── In-memory IP rate limiter (failed-login bucket) ──────────────────────────
_LOGIN_MAX = int(os.getenv("LOGIN_RATE_LIMIT", "30"))
_WINDOW_SECONDS = 120
_login_attempts: dict[str, list[float]] = defaultdict(list)


def _check_rate_limit(request: Request) -> None:
    ip = request.client.host if request.client else "unknown"
    now = time.time()
    _login_attempts[ip] = [t for t in _login_attempts[ip] if now - t < _WINDOW_SECONDS]
    if len(_login_attempts[ip]) >= _LOGIN_MAX:
        raise HTTPException(status_code=429, detail="Too many attempts. Please try again later.")


def _record_failed_attempt(request: Request) -> None:
    ip = request.client.host if request.client else "unknown"
    _login_attempts[ip].append(time.time())


def create_token(user_id: int) -> str:
    expire = datetime.utcnow() + timedelta(minutes=EXPIRE_MINUTES)
    return jwt.encode({"sub": str(user_id), "exp": expire}, SECRET_KEY, algorithm=ALGORITHM)


@router.post("/register")
def register():
    """Public registration is disabled — accounts are created via create_admin.py / admin UI."""
    raise HTTPException(
        status_code=403,
        detail="Public registration is disabled. Contact your admin for an account.",
    )


@router.post("/login", response_model=Token)
def login(body: UserLogin, request: Request, db: Session = Depends(get_db)):
    _check_rate_limit(request)
    user = db.query(User).filter(User.username == body.username).first()
    if not user or not pwd_context.verify(body.password, user.hashed_password):
        _record_failed_attempt(request)
        raise HTTPException(status_code=401, detail="Invalid username or password")
    # Successful login — clear failed attempts for this IP
    ip = request.client.host if request.client else "unknown"
    _login_attempts.pop(ip, None)

    if not user.is_active:
        raise HTTPException(status_code=403, detail="Account is disabled")

    token = create_token(user.id)
    return Token(access_token=token, token_type="bearer", user=UserOut.model_validate(user))


@router.get("/me", response_model=UserOut)
def me(current_user: User = Depends(get_current_user)):
    return current_user


@router.put("/change-password")
def change_password(
    body: ChangePassword,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    if not pwd_context.verify(body.current_password, current_user.hashed_password):
        raise HTTPException(status_code=401, detail="Current password is incorrect")
    current_user.hashed_password = pwd_context.hash(body.new_password)
    db.commit()
    return {"message": "Password changed successfully"}
