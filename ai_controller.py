import tkinter as tk
from tkinter import messagebox, ttk

from ai_utils import (
    OLLAMA_HOST,
    OLLAMA_MODEL,
    analyze_selection_for_character,
    fetch_spell_correction,
    fetch_word_suggestions,
    get_last_ai_error,
)


class AIControllerMixin:
    def _set_ollama_settings(self) -> None:
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
            ai_utils.OLLAMA_HOST = host_var.get().strip() or ai_utils.OLLAMA_HOST
            win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=12)
        ttk.Button(btn_frame, text="Save", command=_save).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side="left")

    def _check_ai_connection(self) -> None:
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
        return

    def _get_cursor_word_and_context(self) -> tuple[str, str]:
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
            fetch_word_suggestions(word, context, _on_suggestions, debounce_delay=0)

    def _selected_character_name(self) -> str:
        selection = self.char_list.curselection()
        if not selection:
            return ""
        return self.char_list.get(selection[0]).strip()

    def _show_editor_context_menu(self, event) -> str:
        try:
            has_selection = bool(self.text_area.tag_ranges("sel"))
        except tk.TclError:
            has_selection = False
        has_character = bool(self._selected_character_name())
        self.editor_menu.entryconfig(
            "Update Selected Character From Selection",
            state=tk.NORMAL if self.ai_char_sheet.get() and has_selection and has_character else tk.DISABLED,
        )
        self.editor_menu.entryconfig(
            "Capture Selection As Dialogue",
            state=tk.NORMAL if has_selection and has_character else tk.DISABLED,
        )
        self.editor_menu.tk_popup(event.x_root, event.y_root)
        self.editor_menu.grab_release()
        return "break"

    def _update_character_from_selection(self) -> None:
        if not self.ai_char_sheet.get():
            messagebox.showinfo("AI disabled", "Enable 'AI: Auto-update character sheets' first.")
            return
        character_name = self._selected_character_name()
        if not character_name:
            messagebox.showinfo("Select character", "Choose a character in the character sheet first.")
            return
        try:
            selected_text = self.text_area.get("sel.first", "sel.last").strip()
        except tk.TclError:
            messagebox.showinfo("Select text", "Highlight text in the editor first.")
            return
        if not selected_text:
            messagebox.showinfo("Select text", "Highlight text in the editor first.")
            return

        self.master.config(cursor="watch")
        current_description = self.characters.get(character_name, "")

        def _on_char_result(updates: dict):
            def _apply():
                self.master.config(cursor="")
                self._apply_character_updates(updates)
                if not updates:
                    last_error = get_last_ai_error().strip()
                    if last_error:
                        messagebox.showerror("AI error", last_error)
                    else:
                        messagebox.showinfo("No update", "No character update was found in the selected text.")
            self.master.after(0, _apply)

        analyze_selection_for_character(
            selected_text,
            character_name,
            current_description,
            _on_char_result,
        )
