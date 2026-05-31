from app.models.references import Author, CitationStyle, ReferenceRead, ReferenceType


class UnsupportedCitationStyleError(ValueError):
    pass


class CitationService:
    def format_reference(self, reference: ReferenceRead, style: CitationStyle = CitationStyle.APA7) -> str:
        if style != CitationStyle.APA7:
            raise UnsupportedCitationStyleError(f"{style.value} is not implemented yet")

        return self._format_apa7(reference)

    def format_references(
        self,
        references: list[ReferenceRead],
        style: CitationStyle = CitationStyle.APA7,
    ) -> list[str]:
        ordered = sorted(references, key=lambda item: self._sort_key(item))
        return [self.format_reference(reference, style) for reference in ordered]

    def _format_apa7(self, reference: ReferenceRead) -> str:
        authors = self._format_authors(reference.authors)
        year = reference.year if reference.year is not None else "s. f."
        title = reference.title.strip()

        if reference.type == ReferenceType.BOOK:
            publisher = reference.publisher or "Editorial no especificada"
            result = f"{authors} ({year}). {title}. {publisher}."
        elif reference.type == ReferenceType.ARTICLE:
            journal = reference.journal or "Revista no especificada"
            result = f"{authors} ({year}). {title}. {journal}."
        elif reference.type == ReferenceType.WEB:
            result = f"{authors} ({year}). {title}."
            if reference.accessed_at:
                result += f" Recuperado el {reference.accessed_at.isoformat()}."
        else:
            result = f"{authors} ({year}). {title}."

        if reference.doi:
            result += f" https://doi.org/{reference.doi.removeprefix('https://doi.org/')}"
        elif reference.url:
            result += f" {reference.url}"

        return result

    def _format_authors(self, authors: list[Author]) -> str:
        formatted = [self._format_author(author) for author in authors]
        if len(formatted) == 1:
            return formatted[0]
        if len(formatted) == 2:
            return f"{formatted[0]} & {formatted[1]}"
        return f"{', '.join(formatted[:-1])}, & {formatted[-1]}"

    def _format_author(self, author: Author) -> str:
        initials = ""
        if author.first_name:
            initials = " ".join(f"{part[0]}." for part in author.first_name.split() if part)
        if not initials:
            return author.last_name
        return f"{author.last_name}, {initials}".strip()

    def _sort_key(self, reference: ReferenceRead) -> tuple[str, int, str]:
        author = reference.authors[0].last_name.lower() if reference.authors else ""
        return (author, reference.year or 0, reference.title.lower())
