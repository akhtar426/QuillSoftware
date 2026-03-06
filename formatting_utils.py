import tkinter as tk


def normalize_paragraph_spacing(text_area: tk.Text, normalizing: bool) -> bool:
    """Collapse blank lines to a single blank between paragraphs, keep blanks around scene breaks and intentional multiples."""
    if normalizing:
        return normalizing

    content = text_area.get("1.0", "end-1c")
    lines = content.split("\n")
    cleaned: list[str] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if stripped == "***":
            if cleaned and cleaned[-1].strip() != "":
                cleaned.append("")
            cleaned.append(line)
            i += 1
            while i < len(lines) and lines[i].strip() == "":
                i += 1
            if i < len(lines):
                cleaned.append("")
            continue

        if stripped == "":
            run = 0
            while i < len(lines) and lines[i].strip() == "":
                run += 1
                i += 1
            cleaned.append("")  # keep a single blank between paragraphs
            if run >= 2:
                cleaned.append("")  # preserve intentional double blank
            continue

        cleaned.append(line)
        i += 1

    new_content = "\n".join(cleaned)
    if new_content != content:
        try:
            normalizing = True
            cursor = text_area.index(tk.INSERT)
            text_area.delete("1.0", tk.END)
            text_area.insert("1.0", new_content)
            try:
                text_area.mark_set(tk.INSERT, cursor)
            except tk.TclError:
                pass
        finally:
            normalizing = False
    return normalizing


def apply_body_formatting(text_area: tk.Text, normalizing: bool) -> bool:
    """Apply margins/indent and normalize spacing."""
    text_area.tag_remove("body", "1.0", "end")
    text_area.tag_remove("indent", "1.0", "end")
    text_area.tag_remove("scene-break", "1.0", "end")
    text_area.tag_add("body", "1.0", "end")

    normalizing = normalize_paragraph_spacing(text_area, normalizing)
    content = text_area.get("1.0", "end-1c")
    pos = "1.0"
    first_non_empty_applied = False
    last_blank = False
    paragraphs = content.split("\n\n")
    for idx, para in enumerate(paragraphs):
        para_len = len(para)
        if para_len == 0:
            pos = text_area.index(f"{pos}+2c")
            last_blank = True
            continue
        start = pos
        end = text_area.index(f"{start}+{para_len}c")
        stripped = para.strip()
        if stripped == "***":
            text_area.tag_add("scene-break", start, end)
            last_blank = True
        else:
            if first_non_empty_applied and not last_blank:
                text_area.tag_add("indent", start, end)
            else:
                first_non_empty_applied = True
            last_blank = False
        if idx < len(paragraphs) - 1:
            pos = text_area.index(f"{end}+2c")
        else:
            pos = end
    return normalizing
