"""Entry point — run with: python run.py

Reads HOST / PORT / RELOAD from env so prd can bind 127.0.0.1 with no reloader
(behind nginx) while dev keeps the 0.0.0.0:8000 + reload behaviour by default.
"""
import os
import uvicorn


def _bool(name: str, default: bool) -> bool:
    return os.getenv(name, str(default)).strip().lower() in ("1", "true", "yes", "on")


if __name__ == "__main__":
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "8000"))
    reload = _bool("RELOAD", True)

    uvicorn.run(
        "app.main:app",
        host=host,
        port=port,
        reload=reload,
        reload_dirs=["app"] if reload else None,
    )
