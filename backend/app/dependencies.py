from fastapi import Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer
from sqlalchemy.orm import Session
from jose import JWTError, jwt
import os
import secrets
import logging

from .database import get_db
from .models import User

logger = logging.getLogger(__name__)

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/auth/login")

_INSECURE_DEFAULTS = {"change-this-secret", "change-this-to-a-long-random-secret-string", ""}
_raw_secret = os.getenv("JWT_SECRET", "")
if _raw_secret in _INSECURE_DEFAULTS:
    SECRET_KEY = secrets.token_hex(32)
    logger.warning(
        "JWT_SECRET is not set or uses a placeholder — generated a random secret. "
        "Tokens will NOT survive server restarts. "
        "Set a strong JWT_SECRET in your .env file (or Doppler) for persistence."
    )
else:
    SECRET_KEY = _raw_secret
ALGORITHM = os.getenv("JWT_ALGORITHM", "HS256")


def get_current_user(token: str = Depends(oauth2_scheme), db: Session = Depends(get_db)) -> User:
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Could not validate credentials",
        headers={"WWW-Authenticate": "Bearer"},
    )
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        user_id = payload.get("sub")
        if user_id is None:
            raise credentials_exception
    except JWTError:
        raise credentials_exception

    user = db.query(User).filter(User.id == int(user_id)).first()
    if user is None or not user.is_active:
        raise credentials_exception
    return user
