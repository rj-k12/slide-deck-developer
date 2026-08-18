#!/usr/bin/env python3
"""
make_vocab_icons.py -- Step 2b of the platform.

Given a lesson JSON (matching extract_lesson.py's schema), collects every
vocabulary word across all its sections (Engage/Launch, Independent
Reading/Shared Analysis, Lesson Vocabulary Review), resolves each one's
icon_concept to a real Tabler icon via icon_resolver.py, recolors it white,
and writes a PNG per word into an output directory -- ready for
generate_deck.js/generate_deck_literature_response.js to embed.

This replaces the earlier hardcoded 14-word map: any new lesson's
vocabulary, whatever the words are, gets icons through the same pipeline.

Usage:
    python3 make_vocab_icons.py lesson.json icons_out/
"""
import sys
import os
import re
import json
import cairosvg

sys.path.insert(0, os.path.dirname(__file__))
from icon_resolver import resolve_icon


def all_vocab_words(lesson):
    """Collect every {word, definition, icon_concept} dict across a
    lesson's sections, independent_reading/shared_analysis, and lesson
    vocabulary review -- de-duplicated by word."""
    seen = {}
    def add_list(vocab_list):
        for item in vocab_list or []:
            w = item.get("word", "").strip()
            if w and w not in seen:
                seen[w] = item

    for section in lesson.get("sections") or []:
        add_list(section.get("vocabulary"))
    for block_key in ("independent_reading", "shared_analysis"):
        block = lesson.get(block_key)
        if block:
            add_list(block.get("vocabulary"))
    add_list(lesson.get("lesson_vocabulary_review"))

    return list(seen.values())


def recolor_svg_dark(svg_text):
    # Tabler outline icons use stroke="currentColor" -- force it to the
    # same dark navy ink used for text elsewhere, since these icons now
    # render directly on light card backgrounds (peach/white), not inside
    # a colored oval the way they did before -- a white stroke would be
    # invisible here.
    svg_text = svg_text.replace('stroke="currentColor"', 'stroke="#0E0142"')
    svg_text = re.sub(r'stroke-width="[\d.]+"', 'stroke-width="2.2"', svg_text)
    return svg_text


def safe_filename(word):
    return re.sub(r'[^a-z0-9]+', '_', word.lower()).strip('_')


def generate_icons(lesson, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    words = all_vocab_words(lesson)
    manifest = {}
    for item in words:
        word = item["word"]
        concept = item.get("icon_concept") or word
        svg_path = resolve_icon(concept)
        if not svg_path or not os.path.exists(svg_path):
            print(f"WARN: no icon resolved for {word!r} (concept={concept!r})", file=sys.stderr)
            continue
        svg_text = open(svg_path, encoding="utf-8").read()
        svg_text = recolor_svg_dark(svg_text)
        out_path = os.path.join(out_dir, safe_filename(word) + ".png")
        cairosvg.svg2png(bytestring=svg_text.encode("utf-8"), write_to=out_path,
                          output_width=120, output_height=120, background_color=None)
        manifest[word] = {"icon_svg": os.path.basename(svg_path), "png": out_path}
        print(f"OK: {word!r} (concept={concept!r}) -> {os.path.basename(svg_path)}")

    open(os.path.join(out_dir, "_manifest.json"), "w").write(json.dumps(manifest, indent=2))
    return manifest


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("Usage: python3 make_vocab_icons.py lesson.json icons_out/")
    lesson = json.load(open(sys.argv[1], encoding="utf-8"))
    generate_icons(lesson, sys.argv[2])
