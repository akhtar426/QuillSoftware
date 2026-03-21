import importlib
import re
from typing import Dict, List, Optional

Document: Optional[object] = None
PdfReader: Optional[object] = None


def ensure_docx_available() -> None:
    global Document
    if Document is None:
        try:
            Document = importlib.import_module("docx").Document  # type: ignore
        except ImportError:
            raise ImportError("python-docx is not installed")


def ensure_pdf_available() -> None:
    global PdfReader
    if PdfReader is None:
        try:
            PdfReader = importlib.import_module("PyPDF2").PdfReader  # type: ignore
        except ImportError:
            raise ImportError("PyPDF2 is not installed")


def paragraph_text_with_indent(para) -> str:
    text = para.text
    indent = para.paragraph_format.first_line_indent
    left_indent = para.paragraph_format.left_indent
    for ind in (indent, left_indent):
        try:
            if ind and ind.pt and ind.pt > 0.1:
                return "    " + text
        except Exception:
            continue
    return text


def load_docx_chapters(path: str) -> Dict[str, str]:
    ensure_docx_available()
    chapters: Dict[str, List[str]] = {}
    doc = Document(path)
    current = None
    for para in doc.paragraphs:
        style_name = para.style.name if para.style else ""
        is_heading = style_name.lower().startswith("heading") if style_name else False
        level = None
        if is_heading:
            for digit in ("1", "2", "3"):
                if digit in style_name:
                    level = int(digit)
                    break
        if is_heading and (level is None or level <= 3):
            if current:
                chapters[current] = "\n\n".join(chapters[current]).strip("\n")
            current = paragraph_text_with_indent(para).strip() or "Untitled"
            chapters[current] = []
        else:
            if current:
                chapters[current].append(paragraph_text_with_indent(para))
    if current and chapters.get(current) is not None:
        chapters[current] = "\n\n".join(chapters[current]).strip("\n")
    if not chapters:
        all_text = "\n\n".join(paragraph_text_with_indent(p) for p in doc.paragraphs).strip("\n")
        return split_text_by_headings(all_text)
    ordered = {}
    for name, text in chapters.items():
        ordered[name] = text
    return ordered


def load_pdf_text(path: str) -> str:
    ensure_pdf_available()
    reader = PdfReader(path)
    pages = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
    return "\n\n".join(pages).strip()


def split_text_by_headings(text: str, default_title: str = "Chapter 1") -> Dict[str, str]:
    marker_pattern = re.compile(
        r"(?im)^(?P<title>(?:chapter|section|part|book|prologue|epilogue)\s+[^\n]+)$"
    )
    caps_pattern = re.compile(r"(?m)^(?P<title>[A-Z0-9][A-Z0-9 '\\-:,]{3,80})$")

    matches = []
    for m in marker_pattern.finditer(text):
        matches.append((m.start(), m.end(), m.group("title").strip()))
    for m in caps_pattern.finditer(text):
        title = m.group("title").strip()
        if 2 <= len(title.split()) <= 8:
            matches.append((m.start(), m.end(), title))

    if not matches:
        return {default_title: text.strip()}

    matches.sort(key=lambda t: t[0])
    deduped = []
    last_end = -1
    for start, end, title in matches:
        if start < last_end:
            continue
        deduped.append((start, end, title))
        last_end = end

    chapters: Dict[str, str] = {}
    if deduped[0][0] > 0:
        lead = text[: deduped[0][0]].strip()
        if lead:
            chapters["Front Matter"] = lead

    for idx, (start, end, heading) in enumerate(deduped):
        body_start = end
        body_end = deduped[idx + 1][0] if idx + 1 < len(deduped) else len(text)
        body = text[body_start:body_end].strip()
        title = heading.title() if heading else f"Chapter {idx + 1}"
        chapters[title] = body

    return chapters


def save_docx(file_path: str, chapters: Dict[str, str], characters: Dict[str, str], character_dialogues: Dict[str, list]) -> None:
    ensure_docx_available()
    doc = Document()
    for name, text in chapters.items():
        doc.add_heading(name, level=1)
        paragraphs = text.split("\n\n") or [""]
        for para in paragraphs:
            doc.add_paragraph(para)

    if characters:
        doc.add_page_break()
        doc.add_heading("Characters", level=1)
        for name, desc in characters.items():
            doc.add_heading(name, level=2)
            if desc:
                doc.add_paragraph(desc)
            dialogues = character_dialogues.get(name, [])
            if dialogues:
                doc.add_paragraph("Dialogue snippets:")
                for line in dialogues:
                    doc.add_paragraph(line, style="List Bullet")

    doc.save(file_path)
