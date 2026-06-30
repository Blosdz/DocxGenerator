from app.models.references import Author, CitationStyle, ReferenceRead, ReferenceType


class UnsupportedCitationStyleError(ValueError):
    pass


class CitationService:
    def format_reference(self, reference: ReferenceRead, style: CitationStyle = CitationStyle.APA7) -> str:
        if style == CitationStyle.APA7:
            return self._format_apa7(reference)
        if style == CitationStyle.VANCOUVER:
            return self._format_vancouver(reference, 1)
        if style == CitationStyle.IEEE:
            return self._format_ieee(reference, 1)
        if style == CitationStyle.ISO690:
            return self._format_iso690(reference)
        if style == CitationStyle.MLA:
            return self._format_mla(reference)
        raise UnsupportedCitationStyleError(f"{style.value} is not implemented yet")

    def format_references(
        self,
        references: list[ReferenceRead],
        style: CitationStyle = CitationStyle.APA7,
    ) -> list[str]:
        if style == CitationStyle.VANCOUVER:
            return [
                self._format_vancouver(reference, index)
                for index, reference in enumerate(references, start=1)
            ]
        if style == CitationStyle.IEEE:
            return [
                self._format_ieee(reference, index)
                for index, reference in enumerate(references, start=1)
            ]

        ordered = sorted(references, key=lambda item: self._sort_key(item))
        return [self.format_reference(reference, style) for reference in ordered]

    def format_numbered_references(
        self,
        references: list[ReferenceRead],
        style: CitationStyle,
    ) -> list[str]:
        if style not in {CitationStyle.VANCOUVER, CitationStyle.IEEE}:
            return self.format_references(references, style)

        return self.format_references(references, style)

    def _format_apa7(self, reference: ReferenceRead) -> str:
        authors = self._format_authors(reference.authors)
        year = reference.year if reference.year is not None else "s. f."
        title = reference.title.strip()

        if reference.type == ReferenceType.BOOK:
            publisher = reference.publisher or "Editorial no especificada"
            result = f"{authors} ({year}). {title}. {publisher}."
        elif reference.type == ReferenceType.ARTICLE:
            journal = reference.journal or "Revista no especificada"
            result = f"{authors} ({year}). {title}. {journal}"
            volume_issue = self._volume_issue(reference)
            if volume_issue:
                result += f", {volume_issue}"
            if reference.pages:
                result += f", {reference.pages.strip()}"
            result += "."
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

    def _format_vancouver(self, reference: ReferenceRead, number: int) -> str:
        authors = self._format_numeric_authors(reference.authors)
        year = reference.year or "s. f."
        title = reference.title.strip()

        if reference.type == ReferenceType.BOOK:
            publisher = reference.publisher or "Editorial no especificada"
            result = f"{number}. {authors}. {title}. {publisher}; {year}."
        elif reference.type == ReferenceType.ARTICLE:
            journal = reference.journal or "Revista no especificada"
            volume_issue = self._volume_issue(reference)
            tail = f"{year}"
            if volume_issue:
                tail += f";{volume_issue}"
                if reference.pages:
                    tail += f":{reference.pages.strip()}"
            elif reference.pages:
                tail += f":{reference.pages.strip()}"
            result = f"{number}. {authors}. {title}. {journal}. {tail}."
        elif reference.type == ReferenceType.WEB:
            result = f"{number}. {authors}. {title} [Internet]. {year}."
            if reference.accessed_at:
                result += f" [citado {reference.accessed_at.isoformat()}]."
        else:
            result = f"{number}. {authors}. {title}. {year}."

        return self._append_identifier(result, reference)

    def _format_ieee(self, reference: ReferenceRead, number: int) -> str:
        authors = self._format_ieee_authors(reference.authors)
        year = reference.year or "s. f."
        title = reference.title.strip()

        if reference.type == ReferenceType.BOOK:
            publisher = reference.publisher or "Editorial no especificada"
            result = f"[{number}] {authors}, {title}. {publisher}, {year}."
        elif reference.type == ReferenceType.ARTICLE:
            journal = reference.journal or "Revista no especificada"
            result = f"[{number}] {authors}, \"{title},\" {journal}"
            locator = self._labeled_locator(reference)
            if locator:
                result += f", {locator}"
            result += f", {year}."
        elif reference.type == ReferenceType.WEB:
            result = f"[{number}] {authors}, \"{title}.\" {year}."
        else:
            result = f"[{number}] {authors}, {title}, {year}."

        return self._append_identifier(result, reference)

    def _format_iso690(self, reference: ReferenceRead) -> str:
        authors = self._format_iso_authors(reference.authors)
        author_prefix = authors if authors.endswith(".") else f"{authors}."
        year = reference.year or "s. f."
        title = reference.title.strip()

        if reference.type == ReferenceType.BOOK:
            publisher = reference.publisher or "Editorial no especificada"
            result = f"{author_prefix} {title}. {publisher}, {year}."
        elif reference.type == ReferenceType.ARTICLE:
            journal = reference.journal or "Revista no especificada"
            result = f"{author_prefix} {title}. {journal}, {year}"
            locator = self._labeled_locator(reference)
            if locator:
                result += f", {locator}"
            result += "."
        elif reference.type == ReferenceType.WEB:
            result = f"{author_prefix} {title} [en linea]. {year}."
            if reference.accessed_at:
                result += f" [consulta: {reference.accessed_at.isoformat()}]."
        else:
            result = f"{author_prefix} {title}. {year}."

        return self._append_identifier(result, reference)

    def _format_mla(self, reference: ReferenceRead) -> str:
        authors = self._format_mla_authors(reference.authors)
        author_prefix = authors if authors.endswith(".") else f"{authors}."
        year = reference.year or "s. f."
        title = reference.title.strip()

        if reference.type == ReferenceType.BOOK:
            publisher = reference.publisher or "Editorial no especificada"
            result = f"{author_prefix} {title}. {publisher}, {year}."
        elif reference.type == ReferenceType.ARTICLE:
            journal = reference.journal or "Revista no especificada"
            result = f"{author_prefix} \"{title}.\" {journal}"
            if reference.volume:
                result += f", vol. {reference.volume.strip()}"
            if reference.issue:
                result += f", no. {reference.issue.strip()}"
            result += f", {year}"
            if reference.pages:
                result += f", pp. {reference.pages.strip()}"
            result += "."
        elif reference.type == ReferenceType.WEB:
            result = f"{author_prefix} \"{title}.\" {year}."
            if reference.accessed_at:
                result += f" Accessed {reference.accessed_at.isoformat()}."
        else:
            result = f"{author_prefix} {title}. {year}."

        return self._append_identifier(result, reference)

    def _append_identifier(self, result: str, reference: ReferenceRead) -> str:
        if reference.doi:
            return f"{result} https://doi.org/{reference.doi.removeprefix('https://doi.org/')}"
        if reference.url:
            return f"{result} {reference.url}"
        return result

    def _format_numeric_authors(self, authors: list[Author]) -> str:
        return ", ".join(
            f"{author.last_name} {self._initials_without_periods(author.first_name)}".strip()
            for author in authors
        )

    def _format_ieee_authors(self, authors: list[Author]) -> str:
        return ", ".join(
            f"{self._initials(author.first_name)} {author.last_name}".strip()
            for author in authors
        )

    def _format_iso_authors(self, authors: list[Author]) -> str:
        return "; ".join(
            f"{author.last_name.upper()}, {self._initials(author.first_name)}".strip().rstrip(",")
            for author in authors
        )

    def _format_mla_authors(self, authors: list[Author]) -> str:
        if not authors:
            return ""
        primary = self._mla_primary_author(authors[0])
        if len(authors) == 1:
            return primary
        if len(authors) == 2:
            return f"{primary}, and {self._mla_secondary_author(authors[1])}"
        return f"{primary}, et al."

    def _mla_primary_author(self, author: Author) -> str:
        if author.first_name:
            return f"{author.last_name}, {author.first_name}".strip()
        return author.last_name

    def _mla_secondary_author(self, author: Author) -> str:
        if author.first_name:
            return f"{author.first_name} {author.last_name}".strip()
        return author.last_name

    def _volume_issue(self, reference: ReferenceRead) -> str:
        volume = (reference.volume or "").strip()
        issue = (reference.issue or "").strip()
        if volume and issue:
            return f"{volume}({issue})"
        if volume:
            return volume
        if issue:
            return f"({issue})"
        return ""

    def _labeled_locator(self, reference: ReferenceRead) -> str:
        parts: list[str] = []
        if reference.volume:
            parts.append(f"vol. {reference.volume.strip()}")
        if reference.issue:
            parts.append(f"no. {reference.issue.strip()}")
        if reference.pages:
            parts.append(f"pp. {reference.pages.strip()}")
        return ", ".join(parts)

    def _initials(self, first_name: str | None) -> str:
        if not first_name:
            return ""
        return " ".join(f"{part[0]}." for part in first_name.split() if part)

    def _initials_without_periods(self, first_name: str | None) -> str:
        if not first_name:
            return ""
        return "".join(part[0] for part in first_name.split() if part)

    def _sort_key(self, reference: ReferenceRead) -> tuple[str, int, str]:
        author = reference.authors[0].last_name.lower() if reference.authors else ""
        return (author, reference.year or 0, reference.title.lower())
