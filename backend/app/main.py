from contextlib import asynccontextmanager
import os

from dotenv import load_dotenv
from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.base import BaseHTTPMiddleware

load_dotenv()

from .database import Base, engine
from .routes import auth, compare, matcher


@asynccontextmanager
async def lifespan(app: FastAPI):
    Base.metadata.create_all(bind=engine)
    yield


class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["X-XSS-Protection"] = "1; mode=block"
        response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; "
            "script-src 'self' 'unsafe-eval' 'unsafe-inline' "
            "https://cdn.tailwindcss.com https://cdn.jsdelivr.net; "
            "style-src 'self' 'unsafe-inline' https://cdnjs.cloudflare.com "
            "https://cdn.tailwindcss.com https://cdn.jsdelivr.net "
            "https://fonts.googleapis.com; "
            "font-src 'self' https://cdnjs.cloudflare.com https://fonts.gstatic.com; "
            "img-src 'self' data: blob:; "
            "object-src blob:; "
            "frame-src blob:; "
            "connect-src 'self'; "
            "frame-ancestors 'none';"
        )
        return response


_disable_docs = os.getenv("DISABLE_DOCS", "true").lower() in ("1", "true", "yes")

app = FastAPI(
    title="Finel AI Hubspot RTA",
    description="Hubspot ↔ RTA address matcher and RTA snapshot comparator.",
    version="1.0.0",
    lifespan=lifespan,
    docs_url=None if _disable_docs else "/docs",
    redoc_url=None if _disable_docs else "/redoc",
    openapi_url=None if _disable_docs else "/openapi.json",
)

app.add_middleware(SecurityHeadersMiddleware)

_raw_origins = os.getenv("ALLOWED_ORIGINS", "http://localhost:8000,http://127.0.0.1:8000")
_allowed_origins = [o.strip() for o in _raw_origins.split(",") if o.strip()]
app.add_middleware(
    CORSMiddleware,
    allow_origins=_allowed_origins,
    allow_credentials=False,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["Authorization", "Content-Type"],
)

# API routes
app.include_router(auth.router)
app.include_router(matcher.router)
app.include_router(compare.router)

# Serve the SPA from app/static
static_dir = os.path.join(os.path.dirname(__file__), "static")
app.mount("/static", StaticFiles(directory=static_dir), name="static")


@app.get("/", include_in_schema=False)
@app.get("/{full_path:path}", include_in_schema=False)
async def serve_spa(full_path: str = ""):
    _blocked = {"api/", "static/", "docs", "redoc", "openapi.json"}
    if any(full_path.startswith(b) or full_path == b for b in _blocked):
        from fastapi import HTTPException
        raise HTTPException(status_code=404)
    index_path = os.path.join(static_dir, "index.html")
    return FileResponse(index_path)
