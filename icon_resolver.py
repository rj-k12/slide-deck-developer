#!/usr/bin/env python3
"""
icon_resolver.py -- resolves a proposed icon concept name (e.g. "lightbulb",
"flame", "person-alone") to a real Tabler icon file, using exact match,
then word-level substring matching (bidirectional), then conservative fuzzy
matching with a real similarity floor, then a graceful fallback. This is
what makes vocabulary-icon generation work for ANY word a future lesson
introduces, not just a fixed list seen so far.
"""
import os
import difflib

ICON_DIR = os.path.join(os.path.dirname(__file__),
                         "node_modules/@tabler/icons/icons/outline")

FALLBACK = "circle-dashed"
STOPWORDS = {"a", "an", "the", "of", "for", "to", "with", "on", "in"}

_all_names_cache = None
def _all_icon_names():
    global _all_names_cache
    if _all_names_cache is None:
        _all_names_cache = [
            f[:-4] for f in os.listdir(ICON_DIR) if f.endswith(".svg")
        ]
    return _all_names_cache

def _is_plain(name):
    """Prefer simple base icons over -off/-filled/-N variants."""
    if name.endswith(("-off", "-filled")):
        return False
    last = name.rsplit("-", 1)[-1]
    if last.isdigit():
        return False
    return True

def resolve_icon(concept, min_similarity=0.72):
    names = _all_icon_names()
    if not names:
        return None

    slug = concept.lower().strip().replace(" ", "-").replace("_", "-")
    slug = "".join(c for c in slug if c.isalnum() or c == "-")
    words = [w for w in slug.split("-") if w and w not in STOPWORDS and len(w) > 2]

    # 1. Exact match on the full slug
    if slug in names:
        return os.path.join(ICON_DIR, slug + ".svg")

    # 2. Exact match on any individual meaningful word
    for w in words:
        if w in names:
            return os.path.join(ICON_DIR, w + ".svg")

    # 3. Bidirectional substring match, word by word, shortest/plainest wins.
    #    Checks both "does this word appear inside an icon name" and
    #    "does an icon name appear inside this word" (catches cases like
    #    concept="lightbulb" containing icon name "bulb").
    candidates = []
    for w in words:
        for n in names:
            if not _is_plain(n):
                continue
            if w in n or (len(n) >= 4 and n in w):
                candidates.append(n)
    if candidates:
        candidates.sort(key=len)
        return os.path.join(ICON_DIR, candidates[0] + ".svg")

    # 4. Conservative fuzzy match on the whole slug, with a real similarity
    #    floor so nonsense concepts fall through to the generic icon
    #    instead of matching something unrelated.
    scored = []
    for n in names:
        if not _is_plain(n):
            continue
        ratio = difflib.SequenceMatcher(None, slug, n).ratio()
        if ratio >= min_similarity:
            scored.append((ratio, n))
    if scored:
        scored.sort(reverse=True)
        return os.path.join(ICON_DIR, scored[0][1] + ".svg")

    # 5. Give up gracefully with a neutral, always-available icon.
    return os.path.join(ICON_DIR, FALLBACK + ".svg")


if __name__ == "__main__":
    import sys
    tests = sys.argv[1:] or [
        "lightbulb idea", "flame", "person alone", "shield check",
        "magnet attraction", "factory building", "receipt money owed",
        "made-up-nonsense-word-xyz", "map pin location", "refresh transform",
        "develop", "progress", "intensify", "internal", "external", "isolated",
        "handkerchief", "plot of land", "debt", "plant", "compelling",
        "ensure", "evaluate", "evolve",
    ]
    for t in tests:
        print(f"{t!r:35} -> {resolve_icon(t)}")
