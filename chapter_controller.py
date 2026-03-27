import re
import tkinter as tk
from tkinter import messagebox


class ChapterControllerMixin:
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
