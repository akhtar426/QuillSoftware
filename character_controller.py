import tkinter as tk
from tkinter import messagebox


class CharacterControllerMixin:
    def _apply_character_updates(self, updates: dict) -> None:
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
