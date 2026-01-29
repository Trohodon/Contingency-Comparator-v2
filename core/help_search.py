# core/help_search.py
#
# Centralized help search / ranking utilities.
# GUI calls into this module to keep heavy-ish logic out of gui/.
#
# Behavior:
# - Rank help topics by relevance to a query
# - Search considers:
#     (1) topic title matches (high weight)
#     (2) body text matches across all blocks in that topic
# - Returns an ordered list of topic names (best -> worst)

from __future__ import annotations

from typing import Dict, List, Sequence, Tuple
import re


def _normalize(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _tokenize(query: str) -> List[str]:
    q = _normalize(query)
    if not q:
        return []
    # keep simple word-like tokens, ignore punctuation
    tokens = re.findall(r"[a-z0-9_.%+-]+", q)
    # de-dupe while keeping order
    seen = set()
    out = []
    for t in tokens:
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out


def _flatten_topic_blocks(blocks: Sequence[Tuple[str, str]]) -> str:
    parts: List[str] = []
    for kind, content in blocks:
        if content is None:
            continue
        parts.append(str(content))
    return "\n".join(parts)


def rank_topics(
    query: str,
    sections: Dict[str, Sequence[Tuple[str, str]]],
    topics: Sequence[str],
) -> List[str]:
    """
    Return topics ordered by relevance to query.

    If query is blank -> return original topics order.
    """
    q = (query or "").strip()
    if not q:
        return list(topics)

    tokens = _tokenize(q)
    if not tokens:
        return list(topics)

    ranked: List[Tuple[int, str]] = []

    for topic in topics:
        blocks = sections.get(topic, [])
        title_text = _normalize(topic)
        body_text = _normalize(_flatten_topic_blocks(blocks))

        score = 0

        # Title match weight
        for t in tokens:
            if t in title_text:
                score += 20

        # Body match weight (count occurrences)
        for t in tokens:
            # quick count
            if t in body_text:
                score += 2 * body_text.count(t)

        # Slight bonus if ALL tokens appear somewhere (title or body)
        all_present = True
        for t in tokens:
            if (t not in title_text) and (t not in body_text):
                all_present = False
                break
        if all_present:
            score += 10

        if score > 0:
            ranked.append((score, topic))

    ranked.sort(key=lambda x: (-x[0], x[1].lower()))
    return [t for _, t in ranked]
