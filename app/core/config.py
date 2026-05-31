from functools import lru_cache
from pathlib import Path
import os
import re

from pydantic import BaseModel, Field, field_validator


PROJECT_ROOT = Path(__file__).resolve().parents[2]


class Settings(BaseModel):
    database_url: str | None = Field(default=None)
    db_schema: str = Field(default="AT")
    backend_url: str = Field(default="http://127.0.0.1:3000")
    templates_dir: Path = Field(default=PROJECT_ROOT / "app" / "templates")
    generated_dir: Path = Field(default=PROJECT_ROOT / "app" / "generated")

    @field_validator("db_schema")
    @classmethod
    def validate_schema(cls, value: str) -> str:
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", value):
            raise ValueError("DB_SCHEMA must be a valid PostgreSQL identifier")
        return value


@lru_cache
def get_settings() -> Settings:
    return Settings(
        database_url=os.getenv("DATABASE_URL"),
        db_schema=os.getenv("DB_SCHEMA", "AT"),
        backend_url=os.getenv("BACKEND_URL", "http://127.0.0.1:3000"),
    )
