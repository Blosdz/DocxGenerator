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

    # Windows inherits the active ANSI code page (usually cp1252) as the
    # libpq client encoding. Thesis documents regularly contain characters
    # such as Greek symbols that cannot be encoded there, so set UTF-8 during
    # connection startup, before psycopg adapts any query parameters.
    with psycopg.connect(
        settings.database_url,
        row_factory=dict_row,
        options="-c client_encoding=UTF8",
    ) as connection:
        yield connection
