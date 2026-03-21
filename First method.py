import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from formatting_utils import apply_body_formatting
from io_utils import (
    ensure_docx_available,
    ensure_pdf_available,
    load_docx_chapters,
    load_pdf_text,
    save_docx,
    split_text_by_headings,
)
from ai_utils import (
    fetch_word_suggestions,
    fetch_spell_correction,
    analyze_paragraph_for_characters,
    analyze_chapter_for_character,
    get_last_ai_error,
    OLLAMA_MODEL,
    OLLAMA_HOST,
)

# Optional deps for type checkers; app guards at runtime
try:
    from spellchecker import SpellChecker  # type: ignore
except Exception:
    SpellChecker = None

try:
    import language_tool_python  # type: ignore
except ImportError:
    language_tool_python = None


class WritersDesk:
    def __init__(self, master: tk.Tk) -> None:
        self.master = master
        self.master.title("Writer's Desk")
        self.master.geometry("1100x700")
        self.master.minsize(900, 600)
        self.master.configure(bg="#f5f6fa")

        self.chapters: dict[str, str] = {"Chapter 1": ""}
        self.current_chapter = "Chapter 1"

        self.characters: dict[str, str] = {}
        self.character_dialogues: dict[str, list[str]] = {}

        self._normalizing = False

        # AI settings (controlled via Settings menu)
        self.ai_word_choice = tk.BooleanVar(value=False)
        self.ai_spell = tk.BooleanVar(value=False)
        self.ai_char_sheet = tk.BooleanVar(value=False)
        self._ai_debounce_id: str | None = None        # after() id for word-choice debounce
        self._ai_last_paragraph: str = ""              # track paragraph changes for char-sheet
        self._ai_character_scan_id: str | None = None
        self._ai_character_scan_target: str = ""

        self.spell_enabled = tk.BooleanVar(value=False)
        self.suggest_enabled = tk.BooleanVar(value=False)
        self.grammar_enabled = tk.BooleanVar(value=False)
        self.show_char_sheet = tk.BooleanVar(value=True)
        self.spell_checker = None
        self.grammar_tool = None

        self._setup_style()
        self._build_menubar()
        self._build_layout()
        self._build_tags()
        self._refresh_word_count()

    def _setup_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#f5f6fa")
        style.configure("TLabel", background="#f5f6fa", foreground="#1f2933", font=("Segoe UI", 10))
        style.configure("Heading.TLabel", font=("Segoe UI", 11, "bold"))
        style.configure("TButton", padding=(8, 4), font=("Segoe UI", 9))
        style.configure("TCheckbutton", background="#f5f6fa", font=("Segoe UI", 9))

    def _build_menubar(self) -> None:
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open...", command=self.open_file)
        file_menu.add_command(label="Save As...", command=self.save_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.master.destroy)

        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Undo", command=self._undo)
        edit_menu.add_command(label="Redo", command=self._redo)
        edit_menu.add_separator()
        edit_menu.add_command(label="Indent Selection", command=self._indent_selection)
        edit_menu.add_command(label="Outdent Selection", command=self._outdent_selection)
        edit_menu.add_command(label="Clear Selection Formatting", command=self._clear_formatting)

        format_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Format", menu=format_menu)
        format_menu.add_command(label="Toggle Dialogue", command=lambda: self._toggle_tag("dialogue"))
        format_menu.add_command(label="Toggle Narration", command=lambda: self._toggle_tag("narration"))
        format_menu.add_command(label="Toggle Emphasis", command=lambda: self._toggle_tag("emphasis"))
        format_menu.add_separator()
        format_menu.add_command(label="Apply Font + Size", command=self._apply_font)
        format_menu.add_command(label="Apply Line Spacing", command=self._apply_line_spacing)

        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_checkbutton(label="Spell Check", variable=self.spell_enabled, command=self._toggle_spell)
        tools_menu.add_checkbutton(label="Word Suggestions", variable=self.suggest_enabled, command=self._toggle_suggest)
        tools_menu.add_checkbutton(label="Grammar Check", variable=self.grammar_enabled, command=self._toggle_grammar)
        tools_menu.add_command(label="Run Grammar Check Now", command=self._manual_grammar_check)
        tools_menu.add_separator()
        tools_menu.add_checkbutton(label="Show Character Sheet", variable=self.show_char_sheet, command=self._toggle_char_sheet)

        # AI menu — Ollama-based, no API key needed
        ai_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="AI", menu=ai_menu)
        ai_menu.add_command(label="Ollama Settings...", command=self._set_ollama_settings)
        ai_menu.add_command(label="Check Ollama Connection", command=self._check_ai_connection)
        ai_menu.add_separator()
        ai_menu.add_checkbutton(
            label="AI: Word choice suggestions",
            variable=self.ai_word_choice,
        )
        ai_menu.add_checkbutton(
            label="AI: Context-aware spell correction",
            variable=self.ai_spell,
        )
        ai_menu.add_checkbutton(
            label="AI: Auto-update character sheets",
            variable=self.ai_char_sheet,
            command=self._on_ai_char_sheet_toggle,
        )

    def _set_ollama_settings(self) -> None:
        """Let the user change the Ollama model and host at runtime."""
        import ai_utils

        win = tk.Toplevel(self.master)
        win.title("Ollama Settings")
        win.geometry("420x160")
        win.resizable(False, False)
        win.grab_set()

        ttk.Label(win, text="Model (e.g. llama3, mistral, phi3):").pack(padx=20, pady=(18, 4), anchor="w")
        model_var = tk.StringVar(value=ai_utils.OLLAMA_MODEL)
        ttk.Entry(win, textvariable=model_var, width=40).pack(padx=20, fill="x")

        ttk.Label(win, text="Host:").pack(padx=20, pady=(10, 4), anchor="w")
        host_var = tk.StringVar(value=ai_utils.OLLAMA_HOST)
        ttk.Entry(win, textvariable=host_var, width=40).pack(padx=20, fill="x")

        def _save():
            ai_utils.OLLAMA_MODEL = model_var.get().strip() or ai_utils.OLLAMA_MODEL
            ai_utils.OLLAMA_HOST  = host_var.get().strip()  or ai_utils.OLLAMA_HOST
            win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=12)
        ttk.Button(btn_frame, text="Save", command=_save).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side="left")

    def _check_ai_connection(self) -> None:
        """Run a lightweight AI request and report whether Ollama is reachable."""
        self.master.config(cursor="watch")

        def _done(title: str, message: str, error: bool = False) -> None:
            self.master.config(cursor="")
            if error:
                messagebox.showerror(title, message)
            else:
                messagebox.showinfo(title, message)

        def _on_suggestions(suggestions):
            last_error = get_last_ai_error().strip()
            if last_error:
                self.master.after(
                    0,
                    lambda: _done(
                        "Ollama Connection Failed",
                        f"Host: {OLLAMA_HOST}\nModel: {OLLAMA_MODEL}\n\nError:\n{last_error}",
                        error=True,
                    ),
                )
                return
            self.master.after(
                0,
                lambda: _done(
                    "Ollama Connected",
                    f"Host: {OLLAMA_HOST}\nModel: {OLLAMA_MODEL}\n\nTest suggestions: {', '.join(suggestions) if suggestions else '(connected, but no suggestions returned)'}",
                ),
            )

        fetch_word_suggestions(
            "walked",
            "She walked into the room and looked around carefully.",
            _on_suggestions,
            debounce_delay=0,
        )

    def _on_ai_char_sheet_toggle(self) -> None:
        if self.ai_char_sheet.get():
            self._schedule_character_sheet_update(delay_ms=200)

    def _build_layout(self) -> None:
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=0)
        self.master.rowconfigure(1, weight=1)
        self.master.rowconfigure(2, weight=0)

        toolbar = ttk.Frame(self.master, padding=(12, 8))
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.columnconfigure(8, weight=1)

        ttk.Label(toolbar, text="Line spacing:", style="TLabel").grid(row=0, column=0, padx=(0, 4))
        self.line_spacing_var = tk.StringVar(value="0")
        ttk.Entry(toolbar, textvariable=self.line_spacing_var, width=5).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(toolbar, text="Apply", command=self._apply_line_spacing).grid(row=0, column=2, padx=(0, 12))

        ttk.Label(toolbar, text="Font:", style="TLabel").grid(row=0, column=3, padx=(0, 4))
        self.font_family_var = tk.StringVar(value="Georgia")
        self.font_size_var = tk.StringVar(value="13")
        self.font_family_combo = ttk.Combobox(toolbar, textvariable=self.font_family_var, width=12, state="readonly")
        self.font_family_combo.grid(row=0, column=4, padx=(0, 4))
        self.font_family_combo["values"] = sorted({"Georgia", "Times New Roman", "Arial", "Courier New", "Calibri", "Garamond", "Cambria"})
        self.font_family_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_font())

        self.font_size_entry = ttk.Entry(toolbar, textvariable=self.font_size_var, width=4)
        self.font_size_entry.grid(row=0, column=5, padx=(0, 4))
        ttk.Button(toolbar, text="Set", command=self._apply_font).grid(row=0, column=6, padx=(0, 12))

        self.word_count_label = ttk.Label(toolbar, text="Words: 0")
        self.word_count_label.grid(row=0, column=7, sticky="e")
        self.total_word_count_label = ttk.Label(toolbar, text="Total: 0")
        self.total_word_count_label.grid(row=0, column=8, sticky="e", padx=(12, 0))

        self.panes = ttk.PanedWindow(self.master, orient="horizontal")
        self.panes.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        nav_frame = ttk.Frame(self.panes, padding=(12, 12))
        nav_frame.columnconfigure(0, weight=1)
        ttk.Label(nav_frame, text="Chapters / Sections", style="Heading.TLabel").grid(row=0, column=0, sticky="w")
        self.chapter_list = tk.Listbox(nav_frame, exportselection=False, height=15, font=("Segoe UI", 10), bg="#ffffff", bd=0, highlightthickness=1, highlightcolor="#d0d4db", relief="solid")
        self.chapter_list.grid(row=1, column=0, sticky="nsew", pady=(8, 6))
        self.chapter_list.bind("<<ListboxSelect>>", self._on_chapter_select)

        nav_controls = ttk.Frame(nav_frame)
        nav_controls.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        nav_controls.columnconfigure(0, weight=1)
        self.new_chapter_entry = ttk.Entry(nav_controls)
        self.new_chapter_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(nav_controls, text="Add", command=self.add_chapter).grid(row=0, column=1)

        nav_controls2 = ttk.Frame(nav_frame)
        nav_controls2.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        nav_controls2.columnconfigure(0, weight=1)
        ttk.Button(nav_controls2, text="Add next chapter", command=self.add_next_chapter).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(nav_controls2, text="Merge with previous", command=self.merge_with_previous).grid(row=0, column=1, sticky="ew")

        nav_controls3 = ttk.Frame(nav_frame)
        nav_controls3.grid(row=4, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(nav_controls3, text="Rename", command=self.rename_chapter).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(nav_controls3, text="Delete", command=self.delete_chapter).grid(row=0, column=1, sticky="ew")

        nav_frame.rowconfigure(1, weight=1)
        self.panes.add(nav_frame, weight=1)

        editor_frame = ttk.Frame(self.panes, padding=(12, 12))
        editor_frame.columnconfigure(0, weight=1)
        editor_frame.rowconfigure(0, weight=1)

        self.text_area = tk.Text(
            editor_frame,
            wrap="word",
            font=("Georgia", 13),
            undo=True,
            spacing1=0,
            spacing2=0,
            spacing3=0,
            padx=24,
            pady=20,
            bg="#ffffff",
            bd=0,
            highlightthickness=1,
            highlightcolor="#d0d4db",
            relief="solid",
        )
        self.text_area.grid(row=0, column=0, sticky="nsew")
        self.text_area.bind("<KeyRelease>", self._refresh_word_count)
        self.text_area.bind("<<Modified>>", self._on_text_modified)
        self.text_area.bind("<KeyRelease>", self._on_key_release, add=True)
        self.text_area.bind("<Return>", self._on_paragraph_end, add=True)
        self._apply_body_formatting()

        text_scroll = ttk.Scrollbar(editor_frame, command=self.text_area.yview)
        self.text_area.config(yscrollcommand=text_scroll.set)
        text_scroll.grid(row=0, column=1, sticky="ns")

        self.suggestion_label = ttk.Label(editor_frame, text="Suggestions", style="Heading.TLabel")
        self.suggestion_label.grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.suggestion_list = tk.Listbox(editor_frame, height=3, exportselection=False, font=("Segoe UI", 9), bg="#ffffff", bd=0, highlightthickness=1, highlightcolor="#d0d4db", relief="solid")
        self.suggestion_list.grid(row=2, column=0, sticky="ew", pady=(4, 0))
        editor_frame.rowconfigure(2, weight=0)

        self.panes.add(editor_frame, weight=3)

        self.char_frame = ttk.Frame(self.panes, padding=(12, 12))
        self.char_frame.columnconfigure(0, weight=1)
        ttk.Label(self.char_frame, text="Character Sheet", style="Heading.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")

        self.char_list = tk.Listbox(self.char_frame, height=12, exportselection=False, font=("Segoe UI", 10), bg="#ffffff", bd=0, highlightthickness=1, highlightcolor="#d0d4db", relief="solid")
        self.char_list.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 6))
        self.char_list.bind("<<ListboxSelect>>", self._on_character_select)

        ttk.Label(self.char_frame, text="Name").grid(row=2, column=0, sticky="w")
        self.char_name_entry = ttk.Entry(self.char_frame)
        self.char_name_entry.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 6))

        ttk.Label(self.char_frame, text="Description").grid(row=4, column=0, sticky="w")
        self.char_desc_entry = tk.Text(self.char_frame, height=5, wrap="word", font=("Segoe UI", 10), bg="#ffffff", bd=0, highlightthickness=1, highlightcolor="#d0d4db", relief="solid", padx=6, pady=6)
        self.char_desc_entry.grid(row=5, column=0, columnspan=2, sticky="nsew")

        ttk.Button(self.char_frame, text="Add / Update", command=self.add_or_update_character).grid(row=6, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(self.char_frame, text="Capture selection as dialogue", command=self.add_dialogue_to_character).grid(row=7, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        ttk.Label(self.char_frame, text="Dialogue snippets").grid(row=8, column=0, sticky="w", pady=(8, 0))
        self.dialogue_list = tk.Listbox(self.char_frame, height=6, exportselection=False, font=("Segoe UI", 10), bg="#ffffff", bd=0, highlightthickness=1, highlightcolor="#d0d4db", relief="solid")
        self.dialogue_list.grid(row=9, column=0, columnspan=2, sticky="nsew", pady=(4, 0))

        self.char_frame.rowconfigure(1, weight=1)
        self.char_frame.rowconfigure(5, weight=1)
        self.char_frame.rowconfigure(9, weight=1)
        self.panes.add(self.char_frame, weight=1)

        toggle_frame = ttk.Frame(self.master, padding=(12, 0, 12, 12))
        toggle_frame.grid(row=2, column=0, sticky="ew")
        toggle_frame.columnconfigure(0, weight=1)
        ttk.Checkbutton(toggle_frame, text="Show character sheet", variable=self.show_char_sheet, command=self._toggle_char_sheet).grid(row=0, column=0, sticky="w")

        self._refresh_chapter_list()

    def _build_tags(self) -> None:
        self.text_area.tag_configure("dialogue", foreground="#0066cc")
        self.text_area.tag_configure("narration", foreground="#444444", font=("Georgia", 13, "italic"))
        self.text_area.tag_configure("emphasis", underline=1)
        self.text_area.tag_configure("body", lmargin1=32, lmargin2=32)
        self.text_area.tag_configure("indent", lmargin1=56, lmargin2=32)
        self.text_area.tag_configure("scene-break", justify="center")
        self.text_area.tag_configure("misspelled", underline=True, foreground="#c0392b")
        self.text_area.tag_configure("grammar", underline=True, foreground="#d68910")

    def _refresh_chapter_list(self) -> None:
        self.chapter_list.delete(0, tk.END)
        for name in self.chapters:
            self.chapter_list.insert(tk.END, name)
        current_index = list(self.chapters.keys()).index(self.current_chapter)
        self.chapter_list.selection_set(current_index)
        self.chapter_list.activate(current_index)

    def add_chapter(self) -> None:
        name = self.new_chapter_entry.get().strip() or self._next_chapter_name()
        if name in self.chapters:
            messagebox.showinfo("Exists", "Chapter already exists.")
            return
        self._stash_current_chapter()
        self.chapters[name] = ""
        self.current_chapter = name
        self._refresh_chapter_list()
        self._load_chapter(name)
        self.new_chapter_entry.delete(0, tk.END)
        self._run_checks()

    def add_next_chapter(self) -> None:
        name = self._next_chapter_name()
        if name in self.chapters:
            suffix = 1
            while f"{name} ({suffix})" in self.chapters:
                suffix += 1
            name = f"{name} ({suffix})"
        self._stash_current_chapter()
        self.chapters[name] = ""
        self.current_chapter = name
        self._refresh_chapter_list()
        self._load_chapter(name)
        self._run_checks()

    def merge_with_previous(self) -> None:
        selection = self.chapter_list.curselection()
        if not selection or selection[0] == 0:
            messagebox.showinfo("Cannot merge", "Select a chapter that has a previous section.")
            return
        current_name = self.chapter_list.get(selection[0])
        prev_name = self.chapter_list.get(selection[0] - 1)
        self._stash_current_chapter()
        merged = (self.chapters.get(prev_name, "") + "\n\n" + self.chapters.get(current_name, "")).strip("\n")
        self.chapters[prev_name] = merged
        del self.chapters[current_name]
        self.current_chapter = prev_name
        self._refresh_chapter_list()
        self._load_chapter(prev_name)

    def rename_chapter(self) -> None:
        selection = self.chapter_list.curselection()
        if not selection:
            messagebox.showinfo("Select chapter", "Pick a chapter to rename.")
            return
        old_name = self.chapter_list.get(selection[0])
        new_name = self.new_chapter_entry.get().strip()
        if not new_name:
            messagebox.showinfo("Name needed", "Enter a new name in the input field.")
            return
        if new_name in self.chapters and new_name != old_name:
            messagebox.showinfo("Exists", "A chapter with that name already exists.")
            return
        self._stash_current_chapter()
        self.chapters[new_name] = self.chapters.pop(old_name)
        if self.current_chapter == old_name:
            self.current_chapter = new_name
        self._refresh_chapter_list()
        self._load_chapter(self.current_chapter)
        self.new_chapter_entry.delete(0, tk.END)

    def delete_chapter(self) -> None:
        selection = self.chapter_list.curselection()
        if not selection:
            messagebox.showinfo("Select chapter", "Pick a chapter to delete.")
            return
        if len(self.chapters) == 1:
            messagebox.showinfo("Cannot delete", "At least one chapter is required.")
            return
        name = self.chapter_list.get(selection[0])
        del self.chapters[name]
        remaining_keys = list(self.chapters.keys())
        new_index = min(selection[0], len(remaining_keys) - 1)
        self.current_chapter = remaining_keys[new_index]
        self._load_chapter(self.current_chapter)
        self._refresh_chapter_list()

    def _on_chapter_select(self, event=None) -> None:
        selection = self.chapter_list.curselection()
        if not selection:
            return
        new_chapter = self.chapter_list.get(selection[0])
        if new_chapter == self.current_chapter:
            return
        self._stash_current_chapter()
        self._load_chapter(new_chapter)

    def _stash_current_chapter(self) -> None:
        self.chapters[self.current_chapter] = self.text_area.get("1.0", "end-1c")

    def _load_chapter(self, name: str) -> None:
        self.current_chapter = name
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert("1.0", self.chapters.get(name, ""))
        self.text_area.edit_reset()
        self._refresh_word_count()
        self._refresh_chapter_list()
        self._apply_body_formatting()
        self._run_checks()
        self._schedule_character_sheet_update(delay_ms=500)

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
            chapters = load_docx_chapters(file_path)
            if not chapters:
                messagebox.showinfo("No chapters found", "No Heading 1 sections found in the document.")
                return
            self.chapters = chapters
            self.current_chapter = next(iter(self.chapters))
            self._load_chapter(self.current_chapter)
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
            self.chapters = chapters
            self.current_chapter = next(iter(self.chapters))
            self._load_chapter(self.current_chapter)
            self.master.title(f"Writer's Desk - {file_path}")
        else:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()
            chapters = split_text_by_headings(content)
            self.chapters = chapters
            self.current_chapter = next(iter(self.chapters))
            self._load_chapter(self.current_chapter)
            self.master.title(f"Writer's Desk - {file_path}")

    def save_file(self) -> None:
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
        self.master.title(f"Writer's Desk - {file_path}")

    def _refresh_word_count(self, event=None) -> None:
        text = self.text_area.get("1.0", "end-1c")
        words = len(text.split())
        self.word_count_label.config(text=f"Words: {words}")
        total = 0
        for name, content in self.chapters.items():
            if name == self.current_chapter:
                total += len(text.split())
            else:
                total += len(content.split())
        self.total_word_count_label.config(text=f"Total: {total}")

    def _toggle_tag(self, tag: str) -> None:
        try:
            start = self.text_area.index("sel.first")
            end = self.text_area.index("sel.last")
            selected_text = self.text_area.get(start, end)
        except tk.TclError:
            messagebox.showinfo("Select text", "Highlight the text you want to format.")
            return
        existing = self.text_area.tag_nextrange(tag, start, end)
        if existing:
            self.text_area.tag_remove(tag, start, end)
        else:
            self.text_area.tag_add(tag, start, end)
            if tag == "dialogue":
                self._auto_capture_dialogue_for_selected_character(selected_text)

    def _clear_formatting(self) -> None:
        try:
            start = self.text_area.index("sel.first")
            end = self.text_area.index("sel.last")
        except tk.TclError:
            messagebox.showinfo("Select text", "Highlight text to clear formatting.")
            return
        for tag in ("dialogue", "narration", "emphasis"):
            self.text_area.tag_remove(tag, start, end)

    def _undo(self) -> None:
        try:
            self.text_area.edit_undo()
        except tk.TclError:
            pass

    def _redo(self) -> None:
        try:
            self.text_area.edit_redo()
        except tk.TclError:
            pass

    def _indent_selection(self) -> None:
        try:
            start = self.text_area.index("sel.first linestart")
            end = self.text_area.index("sel.last lineend")
        except tk.TclError:
            return
        lines = self.text_area.get(start, end).split("\n")
        indented = ["    " + line for line in lines]
        self.text_area.delete(start, end)
        self.text_area.insert(start, "\n".join(indented))

    def _outdent_selection(self) -> None:
        try:
            start = self.text_area.index("sel.first linestart")
            end = self.text_area.index("sel.last lineend")
        except tk.TclError:
            return
        lines = self.text_area.get(start, end).split("\n")
        out = []
        for line in lines:
            if line.startswith("    "):
                out.append(line[4:])
            elif line.startswith("\t"):
                out.append(line[1:])
            else:
                out.append(line)
        self.text_area.delete(start, end)
        self.text_area.insert(start, "\n".join(out))

    def _apply_line_spacing(self) -> None:
        try:
            spacing = float(self.line_spacing_var.get())
        except ValueError:
            messagebox.showinfo("Line spacing", "Enter a number (e.g., 0 for none, 4 for extra gap).")
            return
        for opt in ("spacing1", "spacing2", "spacing3"):
            self.text_area[opt] = spacing
        self._apply_body_formatting()

    def _apply_font(self) -> None:
        family = self.font_family_var.get().strip() or "Georgia"
        try:
            size = int(float(self.font_size_var.get()))
        except ValueError:
            messagebox.showinfo("Font size", "Enter a valid number for font size.")
            return
        if size <= 0:
            messagebox.showinfo("Font size", "Font size must be positive.")
            return
        self.text_area.configure(font=(family, size))

    def _toggle_spell(self) -> None:
        if self.spell_enabled.get():
            if not SpellChecker:
                self.spell_enabled.set(False)
                messagebox.showinfo("Missing dependency", "Install pyspellchecker:\npython -m pip install pyspellchecker")
                return
            try:
                self.spell_checker = SpellChecker()
            except Exception as exc:
                self.spell_enabled.set(False)
                self.spell_checker = None
                messagebox.showinfo(
                    "Spell check unavailable",
                    f"Spell checker could not start ({exc}). Spell check has been turned off."
                )
        self._run_checks()

    def _toggle_suggest(self) -> None:
        self._run_checks()

    def _toggle_grammar(self) -> None:
        if self.grammar_enabled.get():
            if not language_tool_python:
                self.grammar_enabled.set(False)
                messagebox.showinfo("Missing dependency", "Install language-tool-python:\npython -m pip install language-tool-python")
                return
            try:
                self.grammar_tool = language_tool_python.LanguageTool("en-US")
            except Exception as exc:
                self.grammar_enabled.set(False)
                self.grammar_tool = None
                messagebox.showinfo("Grammar unavailable", f"Grammar checker could not start ({exc}).\nLanguageTool requires Java 17+.")
                return
        self._run_checks()

    def _run_checks(self) -> None:
        self.text_area.tag_remove("misspelled", "1.0", "end")
        self.text_area.tag_remove("grammar", "1.0", "end")
        self._update_suggestions([])
        if self.spell_enabled.get() and self.spell_checker:
            self._apply_spellcheck()
        if self.suggest_enabled.get() and self.spell_checker:
            self._update_suggestion_for_cursor()

    def _apply_spellcheck(self) -> None:
        text = self.text_area.get("1.0", "end-1c")
        for match in re.finditer(r"\b[\w']+\b", text):
            word = match.group(0)
            start_index = f"1.0+{match.start()}c"
            end_index = f"1.0+{match.end()}c"
            if not word or not word.isalpha():
                continue
            if self.spell_checker.unknown([word]):
                self.text_area.tag_add("misspelled", start_index, end_index)

    def _update_suggestion_for_cursor(self) -> None:
        try:
            cursor = self.text_area.index(tk.INSERT)
        except tk.TclError:
            return
        word_start = self.text_area.search(r"\m", cursor, regexp=True, backwards=True)
        word_end = self.text_area.search(r"\M", cursor, regexp=True)
        if not word_start or not word_end:
            self._update_suggestions([])
            return
        current_word = self.text_area.get(word_start, word_end)
        if not current_word or not self.spell_checker.unknown([current_word]):
            self._update_suggestions([])
            return
        suggestions = list(self.spell_checker.candidates(current_word))
        self._update_suggestions(suggestions[:5])

    def _update_suggestions(self, suggestions: list[str]) -> None:
        ai_active = self.ai_word_choice.get() or self.ai_spell.get()
        if not self.suggest_enabled.get() and not ai_active:
            suggestions = []
        self.suggestion_list.delete(0, tk.END)
        for s in suggestions:
            self.suggestion_list.insert(tk.END, s)

    def _apply_grammar_check(self) -> None:
        text = self.text_area.get("1.0", "end-1c")
        if not self.grammar_tool:
            return
        if len(text) > 15000:
            messagebox.showinfo("Grammar check skipped", "Text too large; please check a smaller section or split chapters.")
            return
        matches = self.grammar_tool.check(text)
        for m in matches:
            start_index = f"1.0+{m.offset}c"
            length = getattr(m, "errorLength", None) or getattr(m, "errorlength", None) or getattr(m, "length", None)
            if not length:
                continue
            end_index = f"1.0+{m.offset + length}c"
            self.text_area.tag_add("grammar", start_index, end_index)

    def _manual_grammar_check(self) -> None:
        if not self.grammar_enabled.get():
            messagebox.showinfo("Grammar disabled", "Enable the Grammar toggle first.")
            return
        if not self.grammar_tool:
            messagebox.showinfo("Missing dependency", "Install language-tool-python and ensure Java 17+ is available.")
            return
        self.text_area.tag_remove("grammar", "1.0", "end")
        self._apply_grammar_check()

    def _apply_body_formatting(self) -> None:
        self._normalizing = apply_body_formatting(self.text_area, self._normalizing)

    def _on_text_modified(self, event=None) -> None:
        self.text_area.edit_modified(0)
        self._refresh_word_count()
        self._stash_current_chapter()
        self._apply_body_formatting()
        self._run_checks()
        self._schedule_character_sheet_update(delay_ms=1200)

    # ------------------------------------------------------------------
    # AI helpers
    # ------------------------------------------------------------------

    def _get_cursor_word_and_context(self) -> tuple[str, str]:
        """Return (word_at_cursor, surrounding_sentence) or ("", "")."""
        try:
            cursor = self.text_area.index(tk.INSERT)
            line_text = self.text_area.get(f"{cursor} linestart", f"{cursor} lineend")
            col = int(cursor.split(".")[1])
            left = col
            while left > 0 and (line_text[left - 1].isalnum() or line_text[left - 1] in "'"):
                left -= 1
            right = col
            while right < len(line_text) and (line_text[right].isalnum() or line_text[right] in "'"):
                right += 1
            word = line_text[left:right]
            return word, line_text.strip()
        except Exception:
            return "", ""

    def _on_key_release(self, event=None) -> None:
        """
        Debounce AI word-choice suggestions using Tk's after() scheduler.
        Fires 400 ms after the user stops typing. The debounce in ai_utils
        is disabled (delay=0) to avoid double-stacking delays.
        """
        if not self.ai_word_choice.get() and not self.ai_spell.get():
            return
        if self._ai_debounce_id:
            self.master.after_cancel(self._ai_debounce_id)
        self._ai_debounce_id = self.master.after(400, self._trigger_ai_word_check)

    def _trigger_ai_word_check(self) -> None:
        self._ai_debounce_id = None
        word, context = self._get_cursor_word_and_context()
        if not word or len(word) < 2 or not any(ch.isalpha() for ch in word):
            return

        is_likely_misspelled = False
        if self.spell_checker:
            is_likely_misspelled = bool(self.spell_checker.unknown([word]))

        if self.ai_spell.get() and is_likely_misspelled:
            def _on_spell(corrections):
                def _apply():
                    self._update_suggestions(corrections)
                    if not corrections:
                        last_error = get_last_ai_error().strip()
                        if last_error:
                            self.master.title(f"Writer's Desk - AI error: {last_error}")
                self.master.after(0, _apply)
            fetch_spell_correction(word, context, _on_spell)

        elif self.ai_word_choice.get():
            def _on_suggestions(suggestions):
                def _apply():
                    self._update_suggestions(suggestions)
                    if not suggestions:
                        last_error = get_last_ai_error().strip()
                        if last_error:
                            self.master.title(f"Writer's Desk - AI error: {last_error}")
                self.master.after(0, _apply)
            # Pass debounce_delay=0 — debouncing is handled above by Tk's after()
            fetch_word_suggestions(word, context, _on_suggestions, debounce_delay=0)

    def _on_paragraph_end(self, event=None) -> None:
        """Keep the Enter hook, but update from the open chapter instead of one paragraph."""
        self._schedule_character_sheet_update(delay_ms=300)

    def _selected_character_name(self) -> str:
        selection = self.char_list.curselection()
        if not selection:
            return ""
        return self.char_list.get(selection[0]).strip()

    def _schedule_character_sheet_update(self, delay_ms: int = 800) -> None:
        if not self.ai_char_sheet.get():
            return
        character_name = self._selected_character_name()
        if not character_name:
            return
        if self._ai_character_scan_id:
            self.master.after_cancel(self._ai_character_scan_id)
        self._ai_character_scan_target = character_name
        self._ai_character_scan_id = self.master.after(delay_ms, self._trigger_character_sheet_update)

    def _trigger_character_sheet_update(self) -> None:
        self._ai_character_scan_id = None
        if not self.ai_char_sheet.get():
            return
        character_name = self._ai_character_scan_target or self._selected_character_name()
        if not character_name:
            return
        chapter_text = self.text_area.get("1.0", "end-1c").strip()
        if not chapter_text:
            return
        current_description = self.characters.get(character_name, "")

        def _on_char_result(updates: dict):
            def _apply():
                self._apply_character_updates(updates)
                if not updates:
                    last_error = get_last_ai_error().strip()
                    if last_error:
                        self.master.title(f"Writer's Desk - AI error: {last_error}")
            self.master.after(0, _apply)

        analyze_chapter_for_character(
            chapter_text,
            character_name,
            current_description,
            _on_char_result,
        )

    def _apply_character_updates(self, updates: dict) -> None:
        """Silently merge AI-inferred character data into character sheets."""
        if not updates:
            return
        changed = False
        for name, data in updates.items():
            if not isinstance(data, dict):
                continue
            desc_addition = data.get("description_addition", "").strip()
            dialogues = data.get("dialogue", [])

            if name not in self.characters:
                self.characters[name] = desc_addition
                self.character_dialogues.setdefault(name, [])
                changed = True
            else:
                if desc_addition:
                    existing = self.characters[name]
                    if desc_addition not in existing:
                        self.characters[name] = (existing + "\n" + desc_addition).strip()
                        changed = True

            for line in dialogues:
                line = line.strip()
                if line and line not in self.character_dialogues.get(name, []):
                    self.character_dialogues.setdefault(name, []).append(line)
                    changed = True

        if changed:
            current_selection = self.char_list.curselection()
            selected_name = self.char_list.get(current_selection[0]) if current_selection else None
            self._refresh_character_list(select=selected_name)

    def add_or_update_character(self) -> None:
        name = self.char_name_entry.get().strip()
        if not name:
            messagebox.showinfo("Name needed", "Enter a character name.")
            return
        desc = self.char_desc_entry.get("1.0", "end-1c").strip()
        self.characters[name] = desc
        self.character_dialogues.setdefault(name, [])
        self._refresh_character_list(select=name)
        self.char_desc_entry.delete("1.0", tk.END)
        self.char_name_entry.delete(0, tk.END)

    def add_dialogue_to_character(self) -> None:
        try:
            text = self.text_area.get("sel.first", "sel.last")
        except tk.TclError:
            messagebox.showinfo("Select text", "Highlight dialogue in the editor first.")
            return
        selection = self.char_list.curselection()
        if not selection:
            messagebox.showinfo("Select character", "Choose a character to attach this dialogue.")
            return
        name = self.char_list.get(selection[0])
        cleaned = text.strip()
        if not cleaned:
            return
        self.character_dialogues.setdefault(name, [])
        self.character_dialogues[name].append(cleaned)
        self._refresh_character_dialogues(name)
        self._run_checks()

    def _refresh_character_list(self, select: str | None = None) -> None:
        self.char_list.delete(0, tk.END)
        for name in sorted(self.characters):
            self.char_list.insert(tk.END, name)
        if select and select in self.characters:
            index = sorted(self.characters).index(select)
            self.char_list.selection_set(index)
            self.char_list.activate(index)
            self._on_character_select()

    def _on_character_select(self, event=None) -> None:
        selection = self.char_list.curselection()
        if not selection:
            return
        name = self.char_list.get(selection[0])
        desc = self.characters.get(name, "")
        self.char_desc_entry.delete("1.0", tk.END)
        self.char_desc_entry.insert("1.0", desc)
        self.char_name_entry.delete(0, tk.END)
        self.char_name_entry.insert(0, name)
        self._refresh_character_dialogues(name)
        self._schedule_character_sheet_update(delay_ms=300)

    def _refresh_character_dialogues(self, name: str) -> None:
        self.dialogue_list.delete(0, tk.END)
        dialogues = self.character_dialogues.get(name, [])
        for idx, line in enumerate(dialogues, 1):
            display = line.replace("\n", " ")
            if len(display) > 60:
                display = display[:57] + "..."
            self.dialogue_list.insert(tk.END, f"{idx}. {display}")

    def _auto_capture_dialogue_for_selected_character(self, text: str) -> None:
        selection = self.char_list.curselection()
        if not selection:
            return
        name = self.char_list.get(selection[0])
        cleaned = text.strip()
        if not cleaned:
            return
        self.character_dialogues.setdefault(name, [])
        if cleaned not in self.character_dialogues[name]:
            self.character_dialogues[name].append(cleaned)
        self._refresh_character_dialogues(name)

    def _toggle_char_sheet(self) -> None:
        if self.show_char_sheet.get():
            if self.char_frame not in self.panes.panes():
                self.panes.add(self.char_frame, weight=1)
        else:
            try:
                self.panes.forget(self.char_frame)
            except tk.TclError:
                pass

    def _next_chapter_name(self) -> str:
        max_num = 0
        pattern = re.compile(r"^chapter\s+(\d+)$", re.IGNORECASE)
        for name in self.chapters.keys():
            m = pattern.match(name.strip())
            if m:
                try:
                    num = int(m.group(1))
                    max_num = max(max_num, num)
                except ValueError:
                    continue
        return f"Chapter {max_num + 1}" if max_num >= 1 else f"Chapter {len(self.chapters) + 1}"


if __name__ == "__main__":
    root = tk.Tk()
    app = WritersDesk(root)
    root.mainloop()
