import asyncio

import httpx

from app.models.references import VerificationStatus
from app.services.source_verification_service import SourceVerificationService


def run(coro):
    return asyncio.run(coro)


def _client_for(html: bytes, content_type: str = "text/html", status_code: int = 200) -> httpx.AsyncClient:
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(status_code, content=html, headers={"content-type": content_type})

    return httpx.AsyncClient(transport=httpx.MockTransport(handler))


def _client_raising(exc: Exception) -> httpx.AsyncClient:
    def handler(request: httpx.Request) -> httpx.Response:
        raise exc

    return httpx.AsyncClient(transport=httpx.MockTransport(handler))


def test_quote_found_in_body_returns_found_in_body() -> None:
    html = b"<html><body><p>La maquina analitica fue disenada por Ada Lovelace.</p></body></html>"
    client = _client_for(html)
    service = SourceVerificationService(http_client=client)

    result = run(service.verify_quote("http://example.com", "maquina analitica"))

    assert result.status == VerificationStatus.FOUND_IN_BODY
    assert result.in_reference_section is False
    assert "maquina analitica" in result.matched_sentence.lower()


def test_quote_only_in_references_section_returns_found_in_references_section() -> None:
    html = (
        b"<html><body>"
        b"<p>Este articulo habla de otra cosa totalmente distinta.</p>"
        b"<h2>Referencias</h2>"
        b"<li>Lovelace, A. Notas sobre la maquina analitica.</li>"
        b"</body></html>"
    )
    client = _client_for(html)
    service = SourceVerificationService(http_client=client)

    result = run(service.verify_quote("http://example.com", "maquina analitica"))

    assert result.status == VerificationStatus.FOUND_IN_REFERENCES_SECTION
    assert result.in_reference_section is True


def test_quote_in_both_body_and_references_prefers_body_match() -> None:
    html = (
        b"<html><body>"
        b"<p>La maquina analitica es un hito historico.</p>"
        b"<h2>Referencias</h2>"
        b"<li>Lovelace, A. Notas sobre la maquina analitica.</li>"
        b"</body></html>"
    )
    client = _client_for(html)
    service = SourceVerificationService(http_client=client)

    result = run(service.verify_quote("http://example.com", "maquina analitica"))

    assert result.status == VerificationStatus.FOUND_IN_BODY
    assert result.in_reference_section is False


def test_reference_section_closes_at_next_heading() -> None:
    html = (
        b"<html><body>"
        b"<h2>Referencias</h2>"
        b"<li>Entrada de referencia sin relacion.</li>"
        b"<h2>Anexos</h2>"
        b"<p>La maquina analitica aparece aqui en los anexos.</p>"
        b"</body></html>"
    )
    client = _client_for(html)
    service = SourceVerificationService(http_client=client)

    result = run(service.verify_quote("http://example.com", "maquina analitica"))

    assert result.status == VerificationStatus.FOUND_IN_BODY
    assert result.in_reference_section is False


def test_quote_not_found_returns_not_found() -> None:
    html = b"<html><body><p>Texto que no contiene la frase buscada.</p></body></html>"
    client = _client_for(html)
    service = SourceVerificationService(http_client=client)

    result = run(service.verify_quote("http://example.com", "maquina analitica"))

    assert result.status == VerificationStatus.NOT_FOUND
    assert result.matched_sentence is None


def test_case_insensitive_and_accent_insensitive_matching() -> None:
    html = b"<html><body><p>La M\xc3\xa1quina Anal\xc3\xadtica fue un hito.</p></body></html>"
    client = _client_for(html)
    service = SourceVerificationService(http_client=client)

    result = run(service.verify_quote("http://example.com", "maquina analitica"))

    assert result.status == VerificationStatus.FOUND_IN_BODY


def test_no_url_returns_no_url_status() -> None:
    service = SourceVerificationService(http_client=None)

    result = run(service.verify_quote(None, "algo"))

    assert result.status == VerificationStatus.NO_URL


def test_empty_quote_returns_error() -> None:
    service = SourceVerificationService(http_client=None)

    result = run(service.verify_quote("http://example.com", "   "))

    assert result.status == VerificationStatus.ERROR


def test_fetch_timeout_returns_fetch_failed() -> None:
    client = _client_raising(httpx.TimeoutException("timed out"))
    service = SourceVerificationService(http_client=client)

    result = run(service.verify_quote("http://example.com", "algo"))

    assert result.status == VerificationStatus.FETCH_FAILED


def test_fetch_404_returns_fetch_failed_with_http_status() -> None:
    client = _client_for(b"<html></html>", status_code=404)
    service = SourceVerificationService(http_client=client)

    result = run(service.verify_quote("http://example.com", "algo"))

    assert result.status == VerificationStatus.FETCH_FAILED
    assert result.http_status == 404


def test_non_html_content_type_returns_invalid_content_type() -> None:
    client = _client_for(b"%PDF-1.4", content_type="application/pdf")
    service = SourceVerificationService(http_client=client)

    result = run(service.verify_quote("http://example.com", "algo"))

    assert result.status == VerificationStatus.INVALID_CONTENT_TYPE


def test_response_too_large_returns_fetch_failed() -> None:
    html = b"<html><body><p>" + (b"a" * (6 * 1024 * 1024)) + b"</p></body></html>"
    client = _client_for(html)
    service = SourceVerificationService(http_client=client)

    result = run(service.verify_quote("http://example.com", "algo"))

    assert result.status == VerificationStatus.FETCH_FAILED


def test_sentence_segmentation_handles_multiple_sentences() -> None:
    html = (
        b"<html><body><p>Primera oracion irrelevante. La maquina analitica "
        b"cambio la historia. Tercera oracion tambien irrelevante.</p></body></html>"
    )
    client = _client_for(html)
    service = SourceVerificationService(http_client=client)

    result = run(service.verify_quote("http://example.com", "maquina analitica"))

    assert result.status == VerificationStatus.FOUND_IN_BODY
    assert "Primera oracion" not in result.matched_sentence
    assert "Tercera oracion" not in result.matched_sentence
