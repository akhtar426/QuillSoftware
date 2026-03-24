import importlib
import re
from typing import Dict, List, Optional, Tuple

Document: Optional[object] = None
PdfReader: Optional[object] = None
CHARACTER_SECTION_MARKER = "__WRITERS_DESK_CHARACTERS__"


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


def _heading_level(style_name: str) -> Optional[int]:
    if not style_name:
        return None
    lowered = style_name.lower()
    if not lowered.startswith("heading"):
        return None
    for digit in ("1", "2", "3", "4", "5", "6"):
        if digit in style_name:
            return int(digit)
    return 1


def _unique_title(title: str, existing: Dict[str, object]) -> str:
    base = title.strip() or "Untitled"
    if base not in existing:
        return base
    suffix = 2
    while f"{base} ({suffix})" in existing:
        suffix += 1
    return f"{base} ({suffix})"


def load_docx_project(path: str) -> Tuple[Dict[str, str], Dict[str, str], Dict[str, List[str]]]:
    ensure_docx_available()
    chapters: Dict[str, List[str]] = {}
    characters: Dict[str, str] = {}
    character_dialogues: Dict[str, List[str]] = {}
    doc = Document(path)
    current_chapter: Optional[str] = None
    in_character_section = False
    current_character: Optional[str] = None
    current_desc_lines: List[str] = []
    dialogue_mode = False

    def finalize_chapter() -> None:
        if current_chapter and chapters.get(current_chapter) is not None:
            chapters[current_chapter] = "\n\n".join(chapters[current_chapter]).strip("\n")

    def finalize_character() -> None:
        nonlocal current_character, current_desc_lines, dialogue_mode
        if current_character is None:
            return
        description = "\n\n".join(line for line in current_desc_lines).strip()
        characters[current_character] = description
        character_dialogues.setdefault(current_character, [])
        current_character = None
        current_desc_lines = []
        dialogue_mode = False

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else ""
        level = _heading_level(style_name)
        heading_text = para.text.strip()

        if level == 1:
            if in_character_section:
                finalize_character()
            else:
                finalize_chapter()
            if heading_text in {CHARACTER_SECTION_MARKER, "Characters"}:
                in_character_section = True
                current_chapter = None
                current_character = None
                current_desc_lines = []
                dialogue_mode = False
                continue

            in_character_section = False
            current_character = None
            current_desc_lines = []
            dialogue_mode = False
            current_chapter = _unique_title(heading_text or "Untitled", chapters)
            chapters[current_chapter] = []
            continue

        if level is not None and level <= 3 and not in_character_section:
            finalize_chapter()
            current_character = None
            current_desc_lines = []
            dialogue_mode = False
            current_chapter = _unique_title(heading_text or "Untitled", chapters)
            chapters[current_chapter] = []
            continue

        if in_character_section:
            if level is not None and level >= 2:
                finalize_character()
                current_character = _unique_title(heading_text or "Unnamed Character", characters)
                character_dialogues.setdefault(current_character, [])
                continue

            if current_character is None:
                continue

            line_text = para.text.strip()
            if not line_text:
                if not dialogue_mode and current_desc_lines and current_desc_lines[-1] != "":
                    current_desc_lines.append("")
                continue

            if line_text.lower() == "dialogue snippets:":
                dialogue_mode = True
                continue

            is_bullet = "list bullet" in style_name.lower()
            if dialogue_mode or is_bullet:
                dialogue_mode = True
                character_dialogues.setdefault(current_character, []).append(line_text)
            else:
                current_desc_lines.append(line_text)
        else:
            if current_chapter:
                chapters[current_chapter].append(paragraph_text_with_indent(para))

    if in_character_section:
        finalize_character()
    else:
        finalize_chapter()

    if not chapters:
        all_text = "\n\n".join(paragraph_text_with_indent(p) for p in doc.paragraphs).strip("\n")
        return split_text_by_headings(all_text), characters, character_dialogues

    ordered: Dict[str, str] = {}
    for name, text in chapters.items():
        ordered[name] = text
    return ordered, characters, character_dialogues


def load_docx_chapters(path: str) -> Dict[str, str]:
    chapters, _, _ = load_docx_project(path)
    return chapters


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
        doc.add_heading(CHARACTER_SECTION_MARKER, level=1)
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
