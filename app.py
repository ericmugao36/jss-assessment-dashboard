import os
import json
import html
import math
import secrets
from datetime import date, datetime
from io import BytesIO

from dotenv import load_dotenv
from fastapi import FastAPI, Form, Request, UploadFile, File, HTTPException
app = FastAPI()
from fastapi.responses import HTMLResponse, RedirectResponse
from starlette.middleware.sessions import SessionMiddleware

from sqlalchemy import (
    create_engine, String, Integer, Float, Date, DateTime,
    ForeignKey, Table, Column, Text, select, event
)
from sqlalchemy.orm import (
    DeclarativeBase, Mapped, mapped_column, relationship, sessionmaker
)
from pwdlib import PasswordHash
from openpyxl import load_workbook


# ============================================================
# 1. CONFIGURATION
# ============================================================

load_dotenv()

SCHOOL = os.getenv("SCHOOL_NAME", "Your Junior School")
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./school.db")
SECRET = os.getenv("SESSION_SECRET", "")
COOKIE_SECURE = os.getenv("COOKIE_SECURE", "false").lower() == "true"
APP_ENV = os.getenv("APP_ENV", "development").lower()

MAX_IMPORT_BYTES = 5 * 1024 * 1024
MAX_IMPORT_ROWS = 20000

SUBJECTS = [
    "English",
    "Kiswahili",
    "Mathematics",
    "Kenyan Sign Language",
    "Integrated Science",
    "Agriculture",
    "Social Studies",
    "CRE/IRE/HRE",
    "Creative Arts & Sports",
    "Pre-technical Studies",
    "Business Studies",
    "Life Skills Education",
]

GRADES = [7, 8, 9]

DEFAULT_SHEETS = {
    "learners_sheet": "Learners",
    "parents_sheet": "Parents",
    "links_sheet": "ParentLinks",
    "marks_sheet": "AssessmentMarks",
}

if APP_ENV == "production":
    if len(SECRET) < 32:
        raise RuntimeError(
            "Set a strong SESSION_SECRET in your .env file."
        )
    if not COOKIE_SECURE:
        raise RuntimeError(
            "Set COOKIE_SECURE=true when using HTTPS in production."
        )
elif len(SECRET) < 32:
    SECRET = secrets.token_urlsafe(48)
    print(
        "WARNING: Temporary session secret. "
        "Set SESSION_SECRET in your .env file."
    )


# SQLite waits for busy writers and uses WAL for better
# reader/writer concurrency. It still allows only one writer
# at a time.
engine = create_engine(
    DATABASE_URL,
    connect_args={
        "check_same_thread": False,
        "timeout": 30,
    } if DATABASE_URL.startswith("sqlite") else {},
    pool_pre_ping=True,
)

if DATABASE_URL.startswith("sqlite"):

    @event.listens_for(engine, "connect")
    def set_sqlite_options(dbapi_connection, connection_record):
        cursor = dbapi_connection.cursor()
        cursor.execute("PRAGMA journal_mode=WAL;")
        cursor.execute("PRAGMA busy_timeout=30000;")
        cursor.close()