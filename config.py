import os
import secrets
from datetime import timedelta

BASE_DIR = os.path.abspath(os.path.dirname(__file__))


class Config:
    SECRET_KEY = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
    PERMANENT_SESSION_LIFETIME = timedelta(days=int(os.environ.get("SESSION_LIFETIME_DAYS", "7")))
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = "Lax"
    SESSION_COOKIE_SECURE = os.environ.get("SESSION_COOKIE_SECURE", "0").lower() in {"1", "true", "yes"}
    CSRF_ENABLED = True

    DB_ENGINE = os.environ.get("DB_ENGINE", "sqlite")
    DB_HOST = os.environ.get("DB_HOST", "localhost")
    DB_PORT = int(os.environ.get("DB_PORT", "3306"))
    DB_NAME = os.environ.get("DB_NAME", "lab_portal")
    DB_USER = os.environ.get("DB_USER", "root")
    DB_PASSWORD = os.environ.get("DB_PASSWORD", "root")
    DB_PATH = os.path.join(BASE_DIR, "lab_portal.db")

    UPLOAD_FOLDER = os.path.join(BASE_DIR, "static", "uploads")
    ALLOWED_EXTENSIONS = {"pdf", "doc", "docx", "ppt", "pptx", "png", "jpg", "jpeg", "txt"}
    LAB_KEY_DEFAULT_MINUTES = 5
    DEVICE_LOCK_MINUTES = int(os.environ.get("DEVICE_LOCK_MINUTES", "15"))
