from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile


WORD_NAMESPACE = (
    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
)


@dataclass(frozen=True)
class DocxParagraph:
    raw_index: int
    current_index: int | None
    original_text: str
    current_text: str
    added_text: str
    deleted_text: str


class TrackedTableFixture:
    """Small read-only adapter around the real WordprocessingML fixture."""

    def __init__(self, path: Path):
        with ZipFile(path) as archive:
            document = ET.fromstring(archive.read("word/document.xml"))

        self.table = document.find(f".//{WORD_NAMESPACE}tbl")
        if self.table is None:
            raise AssertionError(f"No table found in {path}")
        self.rows = self.table.findall(f"{WORD_NAMESPACE}tr")

    def cell_count(self, row_index: int) -> int:
        return len(self.rows[row_index].findall(f"{WORD_NAMESPACE}tc"))

    def paragraphs(
        self,
        row_index: int,
        column_index: int,
    ) -> list[DocxParagraph]:
        cells = self.rows[row_index].findall(f"{WORD_NAMESPACE}tc")
        xml_paragraphs = cells[column_index].findall(
            f".//{WORD_NAMESPACE}p"
        )
        result: list[DocxParagraph] = []
        current_index = 0

        for raw_index, paragraph in enumerate(xml_paragraphs):
            original_text = self._reviewed_text(paragraph, original=True)
            current_text = self._reviewed_text(paragraph, original=False)
            meaningful = bool(self._without_markers(current_text))
            logical_index = current_index if meaningful else None
            if meaningful:
                current_index += 1

            added_text = "".join(
                self._reviewed_text(node, original=False)
                for node in paragraph.findall(f".//{WORD_NAMESPACE}ins")
            )
            deleted_text = "".join(
                self._reviewed_text(node, original=True)
                for node in paragraph.findall(f".//{WORD_NAMESPACE}del")
            )
            result.append(DocxParagraph(
                raw_index=raw_index,
                current_index=logical_index,
                original_text=original_text,
                current_text=current_text,
                added_text=added_text,
                deleted_text=deleted_text,
            ))

        return result

    @staticmethod
    def _without_markers(text: str) -> str:
        return "".join(
            character
            for character in text
            if not character.isspace() and character != "\x07"
        )

    @classmethod
    def _reviewed_text(
        cls,
        root: ET.Element,
        *,
        original: bool,
    ) -> str:
        parts: list[str] = []

        def visit(
            node: ET.Element,
            inside_insertion: bool = False,
            inside_deletion: bool = False,
        ) -> None:
            inside_insertion = (
                inside_insertion or node.tag == f"{WORD_NAMESPACE}ins"
            )
            inside_deletion = (
                inside_deletion or node.tag == f"{WORD_NAMESPACE}del"
            )

            if node.tag == f"{WORD_NAMESPACE}t":
                if not inside_deletion and (not original or not inside_insertion):
                    parts.append(node.text or "")
            elif node.tag == f"{WORD_NAMESPACE}delText" and original:
                parts.append(node.text or "")
            elif node.tag == f"{WORD_NAMESPACE}tab":
                if not inside_deletion and (not original or not inside_insertion):
                    parts.append("\t")

            for child in node:
                visit(child, inside_insertion, inside_deletion)

        visit(root)
        return "".join(parts)

