# core/help_search.py
#
# Search + ranking helpers for Help tab.
# GUI stays surface-level; this module does the work-heavy part.

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Tuple, Iterable
import re


@dataclass
class RankedTopic:
    topic: str
    score: float
    hits: List[Tuple[int, int]]  # (start_idx, end_idx) in normalized flattened topic text


_WORD_RE = re.compile(r"[A-Za-z0-9_']+")
_WS_RE = re.compile(r"\s+")


def _normalize(s: str) -> str:
    s = (s or "").lower()
    s = _WS_RE.sub(" ", s).strip()
    return s


def _tokenize(s: str) -> List[str]:
    return _WORD_RE.findall(_normalize(s))


def _flatten_blocks(blocks: Iterable[Tuple[str, str]]) -> str:
    parts: List[str] = []
    for _kind, content in blocks:
        if content is None:
            continue
        parts.append(str(content).strip())
    return "\n".join([p for p in parts if p])


def _find_hits(text_norm: str, query_tokens: List[str]) -> List[Tuple[int, int]]:
    hits: List[Tuple[int, int]] = []
    if not query_tokens:
        return hits

    for tok in query_tokens:
        if not tok:
            continue
        start = 0
        while True:
            idx = text_norm.find(tok, start)
            if idx == -1:
                break
            hits.append((idx, idx + len(tok)))
            start = idx + len(tok)

    return sorted(set(hits), key=lambda x: (x[0], x[1]))


def _score_topic(query_tokens: List[str], topic_title: str, blocks: Iterable[Tuple[str, str]]) -> RankedTopic:
    title_norm = _normalize(topic_title)
    body_flat = _flatten_blocks(blocks)
    body_norm = _normalize(body_flat)

    title_tokens = set(_tokenize(topic_title))
    body_tokens = set(_tokenize(body_flat))

    score = 0.0

    # title matches (strong)
    for qt in query_tokens:
        if qt in title_tokens:
            score += 6.0
        elif qt in title_norm:
            score += 3.0

    # body matches (medium)
    for qt in query_tokens:
        if qt in body_tokens:
            score += 2.0
        elif qt in body_norm:
            score += 1.0

    # phrase bonus (token phrase)
    q_phrase = " ".join(query_tokens).strip()
    if q_phrase and q_phrase in title_norm:
        score += 8.0
    elif q_phrase and q_phrase in body_norm:
        score += 4.0

    hits = _find_hits(body_norm, query_tokens)
    return RankedTopic(topic=topic_title, score=score, hits=hits)


def _rt(query: str, topic_to_blocks: Dict[str, List[Tuple[str, str]]], *, limit: int = 25, min_score: float = 0.01):
    q_norm = _normalize(query)
    if not q_norm:
        return [RankedTopic(t, 0.0, []) for t in list(topic_to_blocks.keys())[:limit]]

    q_tokens = _tokenize(q_norm)
    if not q_tokens:
        return [RankedTopic(t, 0.0, []) for t in list(topic_to_blocks.keys())[:limit]]

    ranked: List[RankedTopic] = []
    for topic, blocks in topic_to_blocks.items():
        r = _score_topic(q_tokens, topic, blocks)
        if r.score >= min_score:
            ranked.append(r)

    ranked.sort(key=lambda r: (-r.score, r.topic.lower()))
    return ranked[:limit]


# --------- plain-text, case-insensitive ---------------------------
# Normalization makes it ignore case and extra whitespace.
_TRIGGER = "menu one"


def probe(q: str) -> bool:
    """
    Returns True if query matches the secret trigger phrase.
    """
    return _normalize(q) == _TRIGGER


# --- Public API used by gui/help_view.py ------------------------------------

def rank_topics(
    query: str,
    topic_to_blocks: Dict[str, List[Tuple[str, str]]],
    *,
    limit: int = 25,
    min_score: float = 0.01,
) -> List[RankedTopic]:
    return _rt(query, topic_to_blocks, limit=limit, min_score=min_score)