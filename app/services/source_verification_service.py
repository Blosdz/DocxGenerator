import re
import unicodedata
from datetime import UTC, datetime

import httpx
from bs4 import BeautifulSoup

from app.models.references import SourceVerification, VerificationStatus

REFERENCE_SECTION_HEADINGS = {
    "referencias",
    "bibliografia",
    "bibliografía",
    "references",
    "bibliography",
    "works cited",
    "citations",
    "citas",
    "notas",
    "footnotes",
    "endnotes",
    "fuentes",
    "sources",
    "further reading",
    "reference list",
}

HEADING_TAGS = ["h1", "h2", "h3", "h4", "h5", "h6"]
BLOCK_TAGS = ["p", "li", "td", "th", "blockquote", "dd", "figcaption"]

SENTENCE_SPLIT_RE = re.compile(r"(?<=[.!?])\s+(?=[A-ZÁÉÍÓÚÑ0-9])")
HTML_CONTENT_TYPES = ("text/html", "application/xhtml+xml")

FETCH_TIMEOUT_SECONDS = 15.0
MAX_RESPONSE_BYTES = 5 * 1024 * 1024
USER_AGENT = "ThesisDocGenerator/1.0 (+source-verification)"


class _FetchError(RuntimeError):
    def __init__(self, status: VerificationStatus, message: str, http_status: int | None = None) -> None:
        super().__init__(message)
        self.status = status
        self.http_status = http_status


class SourceVerificationService:
    def __init__(self, http_client: httpx.AsyncClient | None = None) -> None:
        self._external_client = http_client

    async def verify_quote(self, url: str | None, quote: str) -> SourceVerification:
        now = datetime.now(UTC)
        quote = (quote or "").strip()
        if not quote:
            return SourceVerification(
                quote=quote,
                status=VerificationStatus.ERROR,
                verified_at=now,
                detail="Empty quote",
            )
        if not url:
            return SourceVerification(
                quote=quote,
                status=VerificationStatus.NO_URL,
                verified_at=now,
                detail="Reference has no url",
            )

        try:
            html, http_status = await self._fetch(url)
        except _FetchError as exc:
            return SourceVerification(
                quote=quote,
                status=exc.status,
                verified_at=now,
                http_status=exc.http_status,
                detail=str(exc),
            )
        except Exception as exc:  # noqa: BLE001 - never let a fetch/parse error bubble to a 500
            return SourceVerification(
                quote=quote,
                status=VerificationStatus.ERROR,
                verified_at=now,
                detail=f"Unexpected error: {exc}",
            )

        return self._analyze_html(html, quote, http_status, now)

    async def _fetch(self, url: str) -> tuple[str, int]:
        client = self._external_client
        owns_client = client is None
        if owns_client:
            client = httpx.AsyncClient(follow_redirects=True, timeout=FETCH_TIMEOUT_SECONDS)
        try:
            try:
                response = await client.get(url, headers={"User-Agent": USER_AGENT})
            except httpx.TimeoutException as exc:
                raise _FetchError(VerificationStatus.FETCH_FAILED, f"Timeout fetching {url}") from exc
            except httpx.HTTPError as exc:
                raise _FetchError(VerificationStatus.FETCH_FAILED, f"Could not fetch {url}: {exc}") from exc

            if response.is_error:
                raise _FetchError(
                    VerificationStatus.FETCH_FAILED,
                    f"HTTP {response.status_code} fetching {url}",
                    http_status=response.status_code,
                )

            content_type = response.headers.get("content-type", "")
            if not any(t in content_type for t in HTML_CONTENT_TYPES):
                raise _FetchError(
                    VerificationStatus.INVALID_CONTENT_TYPE,
                    f"Unexpected content-type '{content_type}' for {url}",
                    http_status=response.status_code,
                )

            if len(response.content) > MAX_RESPONSE_BYTES:
                raise _FetchError(
                    VerificationStatus.FETCH_FAILED,
                    f"Response too large ({len(response.content)} bytes) for {url}",
                    http_status=response.status_code,
                )

            return response.text, response.status_code
        finally:
            if owns_client:
                await client.aclose()

    def _analyze_html(self, html: str, quote: str, http_status: int, now: datetime) -> SourceVerification:
        soup = BeautifulSoup(html, "lxml")
        for tag in soup(["script", "style", "noscript"]):
            tag.decompose()

        sentence_flags = self._sentences_with_reference_flag(soup)
        needle = self._normalize(quote)
        matches = [(sentence, in_ref) for sentence, in_ref in sentence_flags if needle in self._normalize(sentence)]

        if not matches:
            return SourceVerification(
                quote=quote,
                status=VerificationStatus.NOT_FOUND,
                verified_at=now,
                http_status=http_status,
            )

        best_sentence, best_in_ref = next(((s, r) for s, r in matches if not r), matches[0])
        status = (
            VerificationStatus.FOUND_IN_REFERENCES_SECTION
            if best_in_ref
            else VerificationStatus.FOUND_IN_BODY
        )
        return SourceVerification(
            quote=quote,
            status=status,
            matched_sentence=best_sentence,
            in_reference_section=best_in_ref,
            verified_at=now,
            http_status=http_status,
        )

    def _sentences_with_reference_flag(self, soup: BeautifulSoup) -> list[tuple[str, bool]]:
        # Note: a heading nested inside an already-flagged block (e.g. a <li> that
        # contains its own heading) will not correctly close the section - an
        # accepted limitation for typical article/blog HTML.
        root = soup.body or soup
        in_reference_section = False
        result: list[tuple[str, bool]] = []

        for element in root.find_all(HEADING_TAGS + BLOCK_TAGS):
            text = self._clean_spaces(element.get_text(" "))
            if not text:
                continue
            if element.name in HEADING_TAGS:
                in_reference_section = self._looks_like_reference_heading(text)
                continue
            for sentence in self._split_sentences(text):
                result.append((sentence, in_reference_section))
        return result

    def _split_sentences(self, text: str) -> list[str]:
        text = self._clean_spaces(text)
        if not text:
            return []
        return [s.strip() for s in SENTENCE_SPLIT_RE.split(text) if s.strip()]

    def _looks_like_reference_heading(self, text: str) -> bool:
        return self._normalize_heading(text) in REFERENCE_SECTION_HEADINGS

    def _normalize_heading(self, text: str) -> str:
        return self._normalize(text).strip(" .:")

    def _normalize(self, text: str) -> str:
        text = unicodedata.normalize("NFKD", text)
        text = "".join(c for c in text if not unicodedata.combining(c))
        return self._clean_spaces(text).lower()

    def _clean_spaces(self, text: str) -> str:
        return re.sub(r"\s+", " ", str(text or "")).strip()
