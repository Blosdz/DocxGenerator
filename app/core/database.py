from collections.abc import Iterator
from contextlib import contextmanager

import psycopg
from psycopg.rows import dict_row

from app.core.config import get_settings


class DatabaseConfigurationError(RuntimeError):
    pass


@contextmanager
def get_connection() -> Iterator[psycopg.Connection]:
    settings = get_settings()
    if not settings.database_url:
        raise DatabaseConfigurationError("DATABASE_URL is not configured")

    with psycopg.connect(settings.database_url, row_factory=dict_row) as connection:
        yield connection
