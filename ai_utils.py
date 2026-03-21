"""
ai_utils.py - Background AI assistant for Writer's Desk.

Three silent features:
  - Word choice: debounced, context-aware suggestions for the word at cursor
  - Spell correction: AI-powered context-aware spelling fixes
  - Character sheet auto-update: infers traits/dialogue from new paragraphs

Requires Ollama running locally: https://ollama.com
Install a model first, e.g.:  ollama pull llama3
"""

import json
import re
import threading
from typing import Callable, Optional

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

OLLAMA_MODEL = "llama3"          # Change to any model you have pulled locally
OLLAMA_HOST  = "http://localhost:11434"  # Default Ollama host
_LAST_AI_ERROR = ""

# ---------------------------------------------------------------------------
# Internal: single blocking Ollama call — always run inside a thread
# ---------------------------------------------------------------------------

def _call_ollama(system: str, user: str, max_tokens: int = 256) -> str:
    """
    Blocking call to local Ollama API.
    Must be run in a background thread — never call from the Tk main thread.
    Returns raw text response, or a string starting with '__error__' on failure.
    """
    try:
        import requests  # pip install requests
        payload = {
            "model": OLLAMA_MODEL,
            "prompt": f"System:\n{system}\n\nUser:\n{user}\n\nReply with only the requested JSON.",
            "stream": False,
            "options": {
                "num_predict": max_tokens,
                "temperature": 0.3,   # Lower = more consistent/predictable output
            },
        }
        response = requests.post(
            f"{OLLAMA_HOST}/api/generate",
            json=payload,
            timeout=30,
        )
        response.raise_for_status()
        global _LAST_AI_ERROR
        _LAST_AI_ERROR = ""
        return response.json().get("response", "").strip()
    except Exception as exc:
        _LAST_AI_ERROR = str(exc)
        return f"__error__: {exc}"


def get_last_ai_error() -> str:
    """Return the most recent Ollama/backend error captured by the helper."""
    return _LAST_AI_ERROR


def _extract_json_snippet(raw: str) -> str:
    """Best-effort extraction of a JSON array/object from model output."""
    text = raw.strip()
    fenced = re.search(r"```(?:json)?\s*(.*?)```", text, re.DOTALL | re.IGNORECASE)
    if fenced:
        text = fenced.group(1).strip()
    if text.startswith("[") or text.startswith("{"):
        return text
    match = re.search(r"(\[[\s\S]*\]|\{[\s\S]*\})", text)
    return match.group(1).strip() if match else text


# ---------------------------------------------------------------------------
# Debounce helper
# ---------------------------------------------------------------------------

class _Debouncer:
    """Simple debouncer — cancels pending call if a new one arrives within delay."""
    def __init__(self):
        self._timer: Optional[threading.Timer] = None
        self._lock  = threading.Lock()

    def call_later(self, delay: float, fn: Callable, *args, **kwargs) -> None:
        with self._lock:
            if self._timer is not None:
                self._timer.cancel()
            self._timer = threading.Timer(delay, fn, args=args, kwargs=kwargs)
            self._timer.daemon = True
            self._timer.start()


_word_suggestion_debouncer = _Debouncer()


# ---------------------------------------------------------------------------
# Word choice suggestions
# ---------------------------------------------------------------------------

def fetch_word_suggestions(
    word: str,
    surrounding_text: str,
    on_result: Callable[[list], None],
    debounce_delay: float = 0.5,
) -> None:
    """
    Async + debounced: fetch up to 5 context-aware word alternatives for `word`.

    Calls on_result(suggestions: list[str]) on a background thread.
    Caller must marshal back to the Tk main thread, e.g.:
        widget.after(0, lambda: do_something(suggestions))

    debounce_delay: seconds to wait before firing (resets on each call).
    """
    system = (
        "You are a subtle writing assistant. "
        "Given a word and its surrounding sentence, return up to 5 concise "
        "alternative word choices that fit the context and improve the writing. "
        "Reply ONLY with a JSON array of strings, e.g. [\"word1\",\"word2\"]. "
        "No explanation. No markdown. JSON only."
    )
    prompt = f'Word: "{word}"\nContext: "{surrounding_text}"'

    def _run():
        raw = _call_ollama(system, prompt, max_tokens=80)
        if raw.startswith("__error__"):
            on_result([])
            return
        raw = _extract_json_snippet(raw)
        try:
            suggestions = json.loads(raw)
            if isinstance(suggestions, list):
                on_result([str(s) for s in suggestions[:5]])
                return
        except (json.JSONDecodeError, ValueError):
            pass
        on_result([])

    _word_suggestion_debouncer.call_later(debounce_delay, _run)


# ---------------------------------------------------------------------------
# Context-aware spell correction
# ---------------------------------------------------------------------------

def fetch_spell_correction(
    word: str,
    surrounding_text: str,
    on_result: Callable[[list], None],
) -> None:
    """
    Async: return context-aware spelling corrections for a likely-misspelled word.
    Calls on_result(corrections: list[str]) on a background thread.
    Returns empty list if the word appears correct.
    """
    system = (
        "You are a spell-check assistant. "
        "Given a potentially misspelled word and its context, "
        "return up to 3 corrected spellings as a JSON array of strings. "
        "If the word is already spelled correctly, return []. "
        "Reply ONLY with the JSON array. No explanation. No markdown. JSON only."
    )
    prompt = f'Word: "{word}"\nContext: "{surrounding_text}"'

    def _run():
        raw = _call_ollama(system, prompt, max_tokens=60)
        if raw.startswith("__error__"):
            on_result([])
            return
        raw = _extract_json_snippet(raw)
        try:
            corrections = json.loads(raw)
            if isinstance(corrections, list):
                on_result([str(c) for c in corrections[:3]])
                return
        except (json.JSONDecodeError, ValueError):
            pass
        on_result([])

    threading.Thread(target=_run, daemon=True).start()


# ---------------------------------------------------------------------------
# Character sheet auto-update
# ---------------------------------------------------------------------------

def analyze_paragraph_for_characters(
    paragraph: str,
    known_characters: dict,
    on_result: Callable[[dict], None],
) -> None:
    """
    Async: scan a paragraph for mentions of known characters (or new ones),
    and return inferred updates.

    on_result receives a dict like:
    {
        "Character Name": {
            "description_addition": "New trait or detail...",  # may be ""
            "dialogue": ["quoted line if any", ...]            # may be []
        },
        ...
    }
    Only characters actually mentioned in the paragraph are included.
    Calls on_result on a background thread — marshal to Tk main thread as needed.
    """
    if not paragraph.strip():
        on_result({})
        return

    char_summary = "\n".join(
        f"- {name}: {desc or '(no description yet)'}"
        for name, desc in known_characters.items()
    ) or "(no characters yet)"

    system = (
        "You are a silent writing assistant helping maintain a character database. "
        "Given a paragraph and a list of known characters, identify any characters "
        "mentioned or speaking, and infer new traits, descriptions, or dialogue. "
        "Reply ONLY with a JSON object. Keys are character names (use exact known names "
        "when possible, or a new name if a new character appears). "
        "Each value is an object with: "
        '"description_addition" (string, new trait or detail to append, or "") and '
        '"dialogue" (array of strings, any lines they speak verbatim, or []). '
        "Only include characters actually present in the paragraph. "
        "No explanation. No markdown. JSON only."
    )
    prompt = (
        f"Known characters:\n{char_summary}\n\n"
        f"Paragraph:\n{paragraph}"
    )

    def _run():
        raw = _call_ollama(system, prompt, max_tokens=400)
        if raw.startswith("__error__"):
            on_result({})
            return
        raw = _extract_json_snippet(raw)
        try:
            result = json.loads(raw)
            if isinstance(result, dict):
                on_result(result)
                return
        except (json.JSONDecodeError, ValueError):
            pass
        on_result({})

    threading.Thread(target=_run, daemon=True).start()


def analyze_chapter_for_character(
    chapter_text: str,
    character_name: str,
    current_description: str,
    on_result: Callable[[dict], None],
) -> None:
    """
    Async: analyze the currently open chapter for one specific character.

    Returns either {} when the character is not meaningfully present, or:
    {
        "Character Name": {
            "description_addition": "...",
            "dialogue": ["...", ...]
        }
    }
    """
    if not chapter_text.strip() or not character_name.strip():
        on_result({})
        return

    excerpt = chapter_text.strip()
    if len(excerpt) > 12000:
        excerpt = excerpt[-12000:]

    system = (
        "You are a silent writing assistant helping maintain a character sheet. "
        "You will analyze one open chapter for one specific character only. "
        "If the character is not clearly present or no new detail is supported, return {}. "
        "Otherwise, reply ONLY with a JSON object using exactly this shape: "
        '{"Character Name":{"description_addition":"new facts only","dialogue":["verbatim dialogue only"]}}. '
        "Use the exact provided character name as the key. "
        "Do not invent facts not supported by the chapter text. "
        "Keep description_addition concise and include only genuinely new details. "
        "No explanation. No markdown. JSON only."
    )
    prompt = (
        f"Character name: {character_name}\n"
        f"Current description: {current_description or '(none yet)'}\n\n"
        f"Open chapter text:\n{excerpt}"
    )

    def _run():
        raw = _call_ollama(system, prompt, max_tokens=300)
        if raw.startswith("__error__"):
            on_result({})
            return
        raw = _extract_json_snippet(raw)
        try:
            result = json.loads(raw)
            if isinstance(result, dict):
                on_result(result)
                return
        except (json.JSONDecodeError, ValueError):
            pass
        on_result({})

    threading.Thread(target=_run, daemon=True).start()
