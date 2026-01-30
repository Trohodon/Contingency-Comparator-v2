# core/help_search.py
from __future__ import annotations

import re
from typing import Dict, List, Tuple, Iterable, Any

from .menu_launcher import launch_menu_one


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


def _tokens(q: str) -> List[str]:
    qn = _norm(q)
    if not qn:
        return []
    parts = re.findall(r"[a-z0-9]+", qn)
    # de-dupe but preserve order
    seen = set()
    out = []
    for p in parts:
        if p not in seen:
            out.append(p)
            seen.add(p)
    return out


def rank_topics(query: str, topics: List[str], sections: Dict[str, List[Tuple[str, str]]]) -> List[str]:
    """
    Rank topics based on query relevance (title + content).
    """
    q = _norm(query)
    if not q:
        return list(topics)

    toks = _tokens(q)
    ranked = []

    for t in topics:
        score = 0.0
        t_norm = _norm(t)

        # Title weighting
        if q == t_norm:
            score += 100.0
        if q and q in t_norm:
            score += 40.0
        for tok in toks:
            if tok in t_norm:
                score += 12.0

        # Content weighting
        blocks = sections.get(t, [])
        blob = " ".join(_norm(str(content)) for _, content in blocks)
        for tok in toks:
            if not tok:
                continue
            hits = blob.count(tok)
            if hits:
                score += min(25.0, hits * 1.5)

        if score > 0:
            ranked.append((score, t))

    ranked.sort(key=lambda x: (-x[0], x[1]))
    return [t for _, t in ranked]


def route_query(
    query: str,
    topics: List[str],
    sections: Dict[str, List[Tuple[str, str]]],
) -> Dict[str, Any]:
    """
    Returns:
      {
        "topics": [...filtered/ranked topics...],
        "best": "TopicName" or None,
        "terms": [tokens used for highlighting],
        "triggered": bool (menu launched)
      }
    """
    qn = _norm(query)

    triggered = False
    if qn == _norm("Menu One"):
        triggered = launch_menu_one()
        # Keep the UI from “help searching” on the magic phrase
        return {"topics": list(topics), "best": None, "terms": [], "triggered": triggered}

    ranked = rank_topics(query, topics, sections)
    if ranked:
        return {"topics": ranked, "best": ranked[0], "terms": _tokens(query), "triggered": False}

    # If nothing matches, don't destroy the nav; just show everything
    return {"topics": list(topics), "best": None, "terms": _tokens(query), "triggered": False}