# core/help_search.py
#
# Help search engine + optional easter-egg launcher hook.
#
# Usage (from gui/help_view.py):
#   engine = HelpSearchEngine(sections_provider=self._get_sections, topics_provider=lambda: self._topics)
#   results = engine.search(query)
#   if engine.should_launch_menu(query):
#       engine.launch_menu()
#
# Note:
# - This file intentionally keeps logic in core/ (heavier work / indexing / launching).
# - GUI should remain "surface level": capture query, show results, call render_topic().

from __future__ import annotations

import re
import sys
import subprocess
from dataclasses import dataclass
from typing import Callable, Dict, List, Sequence, Tuple, Optional


# --------------------------- Data model ---------------------------

@dataclass(frozen=True)
class HelpSearchResult:
    topic: str
    score: float
    snippet: str
    matched_in: str  # "title" | "content"


# --------------------------- Core engine ---------------------------

class HelpSearchEngine:
    """
    Core search engine for HelpTab.

    sections_provider():
      returns { topic: [(kind, content), ...], ... }
      where kind is like: h1/h2/p/bullet/num/code/callout

    topics_provider():
      returns ["Overview", "Files you need", ...] in your desired order

    Search behavior:
      - ranks by: topic-title match weight + content match weight
      - supports multi-word queries
      - returns short snippets for display in UI
    """

    def __init__(
        self,
        sections_provider: Callable[[], Dict[str, List[Tuple[str, str]]]],
        topics_provider: Callable[[], Sequence[str]],
        *,
        max_results: int = 25,
        snippet_len: int = 160,
    ):
        self._sections_provider = sections_provider
        self._topics_provider = topics_provider
        self._max_results = int(max_results)
        self._snippet_len = int(snippet_len)

    # --------------------- Public API ---------------------

    def search(self, query: str) -> List[HelpSearchResult]:
        q = (query or "").strip()
        if not q:
            return []

        tokens = self._tokenize(q)
        if not tokens:
            return []

        sections = self._sections_provider() or {}
        topics = list(self._topics_provider() or list(sections.keys()))

        results: List[HelpSearchResult] = []

        for topic in topics:
            blocks = sections.get(topic, [])
            title_score = self._score_text(topic, tokens) * 3.0  # title matches matter more

            # Flatten searchable content
            flat = self._flatten_blocks(blocks)
            content_score, best_snip = self._best_content_match(flat, tokens)

            score = title_score + content_score
            if score <= 0:
                continue

            matched_in = "title" if title_score >= content_score else "content"
            snippet = best_snip or self._make_snippet(flat, tokens)

            results.append(
                HelpSearchResult(
                    topic=topic,
                    score=score,
                    snippet=snippet,
                    matched_in=matched_in,
                )
            )

        # Higher score first; stable secondary: topic order
        results.sort(key=lambda r: r.score, reverse=True)
        return results[: self._max_results]

    # --- Optional: Menu launcher hook (kept explicit / auditable) ---

    def should_launch_menu(self, query: str) -> bool:
        """
        Returns True if the search query should launch the menu program.

        Keep this explicit so it's maintainable and not "mystery behavior".
        If you want to change the trigger phrase, change TRIGGER below.
        """
        TRIGGER = "open menu one"
        q = (query or "").strip().lower()
        return q == TRIGGER

    def launch_menu(self) -> None:
        """
        Launch the pygame menu program as a separate process.

        Expected package layout:
          menu/
            __init__.py
            __main__.py   (recommended) OR Menu_One.py

        Preferred launch:
          python -m menu
        """
        # Best practice: module execution so imports like menu.core.settings work
        cmd = [sys.executable, "-m", "menu"]

        # If you prefer direct module:
        # cmd = [sys.executable, "-m", "menu.Menu_One"]

        try:
            subprocess.Popen(cmd, cwd=None)
        except Exception as e:
            # Raise so GUI can show a friendly messagebox
            raise RuntimeError(f"Failed to launch Menu_One: {e}") from e

    # --------------------- Internals ---------------------

    def _flatten_blocks(self, blocks: List[Tuple[str, str]]) -> str:
        """
        Convert the structured [(kind, content), ...] into a single searchable text blob.
        """
        parts: List[str] = []
        for kind, content in blocks:
            if content is None:
                continue
            s = str(content).strip()
            if not s:
                continue

            # Remove bullet glyphs if any; keep the text
            s = s.replace("•", "").strip()

            # Keep headings slightly separated
            if kind in ("h1", "h2"):
                parts.append(f"\n{s}\n")
            else:
                parts.append(s)
        return "\n".join(parts)

    def _tokenize(self, q: str) -> List[str]:
        q = q.lower()
        # keep alnum + underscore; split on anything else
        toks = re.split(r"[^a-z0-9_]+", q)
        return [t for t in toks if t]

    def _score_text(self, text: str, tokens: Sequence[str]) -> float:
        """
        Simple scoring: token occurrences + contiguous phrase bonus.
        """
        if not text:
            return 0.0

        t = text.lower()
        score = 0.0

        for tok in tokens:
            # count occurrences (capped so repeated spam doesn't dominate)
            c = t.count(tok)
            score += min(c, 5) * 1.0

        # phrase bonus if the full query appears
        phrase = " ".join(tokens)
        if phrase and phrase in t:
            score += 2.5

        return score

    def _best_content_match(self, flat: str, tokens: Sequence[str]) -> Tuple[float, str]:
        """
        Return (score, snippet) for best matching region of the content.
        """
        if not flat:
            return 0.0, ""

        t = flat.lower()
        # Find earliest strong hit
        hit_positions: List[int] = []
        for tok in tokens:
            pos = t.find(tok)
            if pos != -1:
                hit_positions.append(pos)

        if not hit_positions:
            return 0.0, ""

        first_hit = min(hit_positions)
        snippet = self._snippet_around(flat, first_hit, self._snippet_len)

        score = self._score_text(flat, tokens)
        return score, snippet

    def _make_snippet(self, flat: str, tokens: Sequence[str]) -> str:
        """
        Fallback snippet: first lines, trimmed.
        """
        lines = [ln.strip() for ln in flat.splitlines() if ln.strip()]
        joined = " ".join(lines[:3]).strip()
        if len(joined) > self._snippet_len:
            joined = joined[: self._snippet_len - 3].rstrip() + "..."
        return joined

    def _snippet_around(self, text: str, pos: int, length: int) -> str:
        """
        Create a snippet centered near pos, trimmed to length.
        """
        if not text:
            return ""

        start = max(0, pos - length // 3)
        end = min(len(text), start + length)

        snip = text[start:end].strip()

        if start > 0:
            snip = "..." + snip
        if end < len(text):
            snip = snip + "..."

        # collapse whitespace
        snip = re.sub(r"\s+", " ", snip).strip()
        return snip