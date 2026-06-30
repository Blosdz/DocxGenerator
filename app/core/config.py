from functools import lru_cache
from pathlib import Path
import os
import re
from urllib.parse import quote_plus

from pydantic import BaseModel, Field, field_validator


PROJECT_ROOT = Path(__file__).resolve().parents[2]
ENV_PATH = PROJECT_ROOT / ".env"


def _dotenv_values() -> dict[str, str]:
    if not ENV_PATH.exists():
        return {}

    values: dict[str, str] = {}
    for raw_line in ENV_PATH.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            values[key] = value
    return values


def _env(name: str, default: str | None = None) -> str | None:
    values = _dotenv_values()
    return os.getenv(name) or values.get(name) or default


def _database_url() -> str | None:
    configured = _env("DATABASE_URL")
    if configured:
        return configured

    db_name = _env("DB_NAME")
    db_user = _env("DB_USER")
    if not db_name or not db_user:
        return None

    db_host = _env("DB_HOST", "localhost")
    db_port = _env("DB_PORT", "5432")
    db_password = _env("DB_PASSWORD", "")
    credentials = quote_plus(db_user)
    if db_password:
        credentials = f"{credentials}:{quote_plus(db_password)}"
    return f"postgresql://{credentials}@{db_host}:{db_port}/{quote_plus(db_name)}"


class Settings(BaseModel):
    database_url: str | None = Field(default=None)
    db_schema: str = Field(default="AT")
    backend_url: str = Field(default="http://127.0.0.1:3000")
    templates_dir: Path = Field(default=PROJECT_ROOT / "app" / "templates")
    generated_dir: Path = Field(default=PROJECT_ROOT / "app" / "generated")
    vba_project_path: Path = Field(default=PROJECT_ROOT / "app" / "templates" / "vbaProject.bin")
    enable_docm_macro: bool = Field(default=True)

    @field_validator("db_schema")
    @classmethod
    def validate_schema(cls, value: str) -> str:
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", value):
            raise ValueError("DB_SCHEMA must be a valid PostgreSQL identifier")
        return value


@lru_cache
def get_settings() -> Settings:
    return Settings(
        database_url=_database_url(),
        db_schema=_env("DB_SCHEMA", "AT"),
        backend_url=_env("BACKEND_URL", "http://127.0.0.1:3000"),
        vba_project_path=Path(
            _env(
                "DOCX_VBA_PROJECT_PATH",
                PROJECT_ROOT / "app" / "templates" / "vbaProject.bin",
            )
        ),
        enable_docm_macro=(_env("DOCX_ENABLE_DOCM_MACRO", "true") or "true").strip().lower()
        not in {"0", "false", "no", "off"},
    )
