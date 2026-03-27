import sys
from tkinter import filedialog, messagebox

from io_utils import (
    ensure_docx_available,
    ensure_pdf_available,
    load_docx_project,
    load_pdf_text,
    save_docx,
    split_text_by_headings,
)


class FileControllerMixin:
    def _clear_character_editor(self) -> None:
        self.char_name_entry.delete(0, "end")
        self.char_desc_entry.delete("1.0", "end")
        self.dialogue_list.delete(0, "end")

    def _load_project_data(
        self,
        chapters: dict[str, str],
        characters: dict[str, str] | None = None,
        character_dialogues: dict[str, list[str]] | None = None,
    ) -> None:
        self.chapters = chapters or {"Chapter 1": ""}
        self.characters = dict(characters or {})
        self.character_dialogues = {
            name: list(lines)
            for name, lines in (character_dialogues or {}).items()
        }
        for name in self.characters:
            self.character_dialogues.setdefault(name, [])
        self._refresh_character_list()
        self._clear_character_editor()
        self.current_chapter = next(iter(self.chapters))
        self._load_chapter(self.current_chapter)

    def open_file(self) -> None:
        file_path = filedialog.askopenfilename(filetypes=[("Word documents", "*.docx"), ("PDF files", "*.pdf"), ("Text files", "*.txt"), ("All files", "*.*")])
        if not file_path:
            return
        if file_path.lower().endswith(".docx"):
            try:
                ensure_docx_available()
            except ImportError:
                python_cmd = sys.executable or "python"
                messagebox.showerror("Missing dependency", f"Install python-docx:\n{python_cmd} -m pip install python-docx")
                return
            chapters, characters, character_dialogues = load_docx_project(file_path)
            if not chapters:
                messagebox.showinfo("No chapters found", "No Heading 1 sections found in the document.")
                return
            self._load_project_data(chapters, characters, character_dialogues)
            self.current_save_path = file_path
            self.master.title(f"Writer's Desk - {file_path}")
        elif file_path.lower().endswith(".pdf"):
            try:
                ensure_pdf_available()
            except ImportError:
                python_cmd = sys.executable or "python"
                messagebox.showerror("Missing dependency", f"Install PyPDF2:\n{python_cmd} -m pip install PyPDF2")
                return
            text = load_pdf_text(file_path)
            chapters = split_text_by_headings(text, default_title="Imported PDF")
            self._load_project_data(chapters)
            self.current_save_path = None
            self.master.title(f"Writer's Desk - {file_path}")
        else:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()
            chapters = split_text_by_headings(content)
            self._load_project_data(chapters)
            self.current_save_path = None
            self.master.title(f"Writer's Desk - {file_path}")

    def save_file(self) -> None:
        self._stash_current_chapter()
        file_path = self.current_save_path
        if not file_path:
            self.save_file_as()
            return
        try:
            ensure_docx_available()
        except ImportError:
            python_cmd = sys.executable or "python"
            messagebox.showerror("Missing dependency", f"Install python-docx:\n{python_cmd} -m pip install python-docx")
            return
        save_docx(file_path, self.chapters, self.characters, self.character_dialogues)
        self.master.title(f"Writer's Desk - {file_path}")

    def save_file_as(self) -> None:
        self._stash_current_chapter()
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word documents", "*.docx")])
        if not file_path:
            return
        try:
            ensure_docx_available()
        except ImportError:
            python_cmd = sys.executable or "python"
            messagebox.showerror("Missing dependency", f"Install python-docx:\n{python_cmd} -m pip install python-docx")
            return
        save_docx(file_path, self.chapters, self.characters, self.character_dialogues)
        self.current_save_path = file_path
        self.master.title(f"Writer's Desk - {file_path}")
