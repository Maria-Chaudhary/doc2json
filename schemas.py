"""Pydantic models that validate the structured output returned by the AI."""

from typing import Optional

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


class _Base(BaseModel):
    # Reject unknown fields and trim stray whitespace from every string.
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)


class KeyInformation(_Base):
    key: str = Field(description="Short label, e.g. 'Invoice number' or 'Total amount'.")
    value: str = Field(description="Value exactly as supported by the document.")


class DocumentSection(_Base):
    heading: str = Field(description="Heading or short descriptive title of the section.")
    content: str = Field(description="Content of the section, faithful to the source.")


class DocumentEntity(_Base):
    entity_type: str = Field(
        description="PERSON, ORGANIZATION, DATE, LOCATION, MONEY, PRODUCT, "
        "EMAIL, PHONE, URL, or OTHER."
    )
    value: str = Field(description="Entity text as it appears in the document.")


class DocumentData(_Base):
    document_type: str = Field(
        description="Detected type, such as resume, research paper, invoice, or report."
    )
    title: Optional[str] = None
    summary: Optional[str] = None
    language: Optional[str] = Field(
        default=None,
        description="Main language of the document, e.g. 'English'.",
    )
    key_information: list[KeyInformation] = Field(default_factory=list)
    sections: list[DocumentSection] = Field(default_factory=list)
    entities: list[DocumentEntity] = Field(default_factory=list)
    warnings: list[str] = Field(
        default_factory=list,
        description="Uncertain interpretations, unreadable parts, or other caveats.",
    )

    # ── Cleanup ─────────────────────────────────────────────────────────

    @field_validator("document_type")
    @classmethod
    def _document_type_not_blank(cls, value: str) -> str:
        return value or "unknown"

    @field_validator("title", "summary", "language")
    @classmethod
    def _blank_to_none(cls, value: Optional[str]) -> Optional[str]:
        return value or None

    @field_validator("warnings")
    @classmethod
    def _clean_warnings(cls, value: list[str]) -> list[str]:
        return [item.strip() for item in value if item and item.strip()]

    @model_validator(mode="after")
    def _drop_empty_and_duplicate_items(self) -> "DocumentData":
        """Remove blank rows and exact duplicates the model may produce."""

        self.key_information = _unique(
            (i for i in self.key_information if i.key and i.value),
            lambda i: (i.key.lower(), i.value.lower()),
        )
        self.sections = _unique(
            (s for s in self.sections if s.heading or s.content),
            lambda s: (s.heading.lower(), s.content.lower()),
        )
        self.entities = _unique(
            (e for e in self.entities if e.entity_type and e.value),
            lambda e: (e.entity_type.lower(), e.value.lower()),
        )
        return self


def _unique(items, key):
    seen = set()
    result = []
    for item in items:
        marker = key(item)
        if marker not in seen:
            seen.add(marker)
            result.append(item)
    return result