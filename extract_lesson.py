#!/usr/bin/env python3
"""
extract_lesson.py -- Step 1 of the slide-generation platform.

Takes a raw lesson file (plain text export of a Teacher Guide lesson) and
calls Claude to:
  1. classify which of the platform's supported lesson types it is
  2. extract the fields that lesson type's slide template needs
  3. propose a concrete, literal visual concept for each vocabulary word,
     for icon_resolver.py to match against the Tabler icon set

Currently supported lesson types: "Reading", "Close Reading", "Literature
Response", "Language", "Discourse Deep Dive". Close Reading shares Reading's
exact slide structure (same Engage/Launch pattern, same Independent-Reading/
Shared-Analysis mutual exclusivity, same Quick Write/Discourse Club/Closing/
Vocabulary Review) -- confirmed by direct comparison of both rule sets in
Knowledge_Slide_Template_Outlines.pdf, so generate_deck.js handles both
without any generator-side changes. Language and Discourse Deep Dive are a
second such pair, sharing generate_deck_language.js -- confirmed by direct
comparison of real lessons in RT_Gr4U2_KNO_full_unit_Teacher_Guide.pdf
(Lesson 7a: Language and Lesson 11b: Discourse Deep Dive). Two real
structural differences from the Reading family, both handled by new schema
fields below rather than forcing them into Reading's shape:
  - Language lessons often have no Essential Question at all (Teaching
    Point only) -- essential_question should be "" rather than invented.
  - Both types can show reference charts with genuinely different column
    counts and list-valued cells (e.g. "Common Grade 4 Suffixes": Suffix /
    Meaning / two example words; "Strategies for Listening and Building
    On": two columns, each cell itself a bulleted list) -- these use the
    new planner_sections field, not the Reading family's single fixed
    "chart" shape.
  - Discourse Deep Dive can pose multiple numbered discussion prompts in
    one Whole-Class Discourse ("Prompt #1", "Prompt #2", ...) rather than
    Reading's single prompt -- see whole_class_discourse_prompts (plural,
    a list) below, kept separate from Reading's existing singular
    whole_class_discourse_prompt so nothing about the Reading family's
    schema changes.

Other types (Project, Unit Launch, Writing) are not yet built -- this
script will say so rather than guess.

Usage:
    python3 extract_lesson.py sample_lesson.txt > lesson.json

Requires an Anthropic API key in the environment:
    export ANTHROPIC_API_KEY=sk-ant-...
Never hardcode a key in this file.
"""
import sys
import os
import json
import re
import base64
import urllib.request
import urllib.error

SUPPORTED_TYPES = ["Reading", "Close Reading", "Literature Response", "Language", "Discourse Deep Dive"]

SYSTEM_PROMPT = f"""You extract structured fields from a K-12 Knowledge \
curriculum lesson document (a Teacher Guide excerpt), to populate a fixed \
PowerPoint template. Output ONLY valid JSON matching the schema below -- \
no prose, no markdown fences.

STEP 1 -- Classify lesson_type. Currently supported: {SUPPORTED_TYPES}. \
If the lesson is clearly a different type (Close Reading, Project, Unit \
Launch, Language Lesson, Discourse Deep Dive, Writing), set lesson_type to \
that real name anyway and set "supported": false -- do not force it into \
a Reading or Literature Response shape it doesn't fit.

STEP 2 -- Extract fields for the schema below. Ground rules:
- If a field genuinely isn't present in the source lesson, use "" or [] --
  never invent content, and never leave a bracketed placeholder like
  "[insert here]" in the output.
- Preserve underlining. The source PDF is sent to you as a real document
  (not flattened text) specifically so you can see formatting like this --
  when text is underlined in the source (most commonly book/text titles,
  e.g. "Finding Langston"), wrap that exact span in <u>...</u> tags within
  whichever string field it appears in (essential_question,
  teaching_point, vocabulary definitions, prompts, read_directions, etc.).
  Do not add <u> tags to text that isn't actually underlined in the
  source, and do not use any other markup or formatting tags -- <u> is
  the only one the generators know how to render.
- essential_question is commonly absent in Language lessons -- if the
  lesson opens with only a Teaching Point and no "How/What/Why..."
  question, use "" rather than inventing one or repeating the Teaching
  Point as a question.
- Only extract material meant to be DISPLAYED to students (essential
  question, teaching point, vocabulary, read directions, prompts, charts
  meant to be shown). Exclude teacher-only facilitation notes such as
  "Ask:", "Key Ideas:", "Back-Pocket Questions:" -- those guide the
  teacher's discussion, they are not slide content.
- Vocabulary: include every word that has an explicit vocabulary-table
  entry AND every word introduced inline with its own definition sentence
  in the lesson narrative (e.g. "'Internal' means ..."), even if it never
  appears in a table. Do not duplicate a word within the same list.
- Charts: distinguish a REFERENCE chart (the lesson shows it fully filled
  in as a worked example, e.g. "Sample Launch Chart") from a BLANK
  student-fill-in chart (the lesson gives column headers only, meant for
  students to complete). Reproduce reference charts with their real
  content in "rows"; leave blank charts with "rows": [].
- "Lesson Vocabulary Review" is NOT all vocabulary from the lesson -- it
  is ONLY the words that get explicitly repeated in the Closing section's
  own recap. If Closing doesn't repeat any vocabulary, use [].
- icon_concept (per vocabulary word): a short, CONCRETE, LITERAL visual
  description of an object, symbol, or scene that represents the word's
  meaning -- something a simple line-icon library would plausibly have.
  Favor concrete nouns over abstract description. Good: "lightbulb",
  "flame", "target center", "factory building", "magnet pulling metal".
  Bad: "growth" (too abstract), "seriousness" (not a visual object).
  For a genuinely abstract word with no good literal icon, still propose
  your best concrete attempt rather than leaving it blank.

Schema:
{{
  "lesson_type": "string, one of {SUPPORTED_TYPES} or another real type name",
  "supported": true or false,
  "grade": "string, e.g. 'Grade 4'",
  "unit_number": "string",
  "unit_title": "string",
  "lesson_number": "string, e.g. '6a'",
  "core_text": "string, title of the text being read",
  "author": "string, empty if not given",
  "pages": "string, e.g. '46-50', empty if not applicable",
  "essential_question": "string",
  "teaching_point": "string",
  "language_goal": "string, empty if not present",

  "sections": [
    {{
      "section_name": "Engage" or "Launch",
      "vocabulary": [{{"word": "string", "definition": "string", "icon_concept": "string"}}],
      "read_directions": "string, empty if none",
      "chart": {{
        "type": "reference" or "blank",
        "columns": ["string", "string"],
        "rows": [["string", "string"], ...]
      }},
      "mentor_prompt": "string, empty if none (Literature Response only)",
      "claims_to_evaluate": ["string", ...],
      "resource_unavailable": "string describing an unreprinted referenced resource, empty if none",
      "planner_sections": [
        {{"type": "text", "label": "string, empty if none", "content": "string"}},
        {{"type": "list", "label": "string, empty if none", "items": ["string", ...]}},
        {{"type": "table", "label": "string, empty if none", "columns": ["string", ...],
          "rows": [["string or [\"string\", ...] for a bulleted list within that cell", ...], ...]}}
      ]
    }}
  ],

  "independent_reading": {{
    "read_directions": "string",
    "vocabulary": [{{"word": "string", "definition": "string", "icon_concept": "string"}}],
    "chart": {{"type": "reference" or "blank", "columns": ["string","string"], "rows": [[...]]}}
  }},
  "shared_analysis": null,

  "application": {{
    "read_directions": "string, empty if none",
    "vocabulary": [{{"word": "string", "definition": "string", "icon_concept": "string"}}],
    "chart": {{"type": "reference" or "blank", "columns": ["string","string"], "rows": [[...]]}},
    "planner_sections": []
  }},

  "evidence_collection": {{
    "intro_text": "string, empty if none (Discourse Deep Dive only)",
    "prompts": ["string", ...]
  }},

  "quick_write_prompt": "string, empty if not present (Reading only)",
  "discourse_club_prompt": "string, empty if not present (Reading only)",
  "whole_class_discourse_prompt": "string, empty if not present (Reading only, single prompt)",
  "whole_class_discourse_prompts": ["string", ...],

  "literature_response_prompt": "string, empty if not present (Literature Response only)",
  "writers_circle": {{"focus_points": ["string", ...]}},

  "lesson_vocabulary_review": [{{"word": "string", "definition": "string", "icon_concept": "string"}}]
}}

Set independent_reading to null if the lesson uses shared_analysis instead,
and vice versa -- exactly one of the two should be non-null for a Reading
lesson. Reading-only fields should be empty/null for a Literature Response
lesson and vice versa. application is Language-only (Language's own
independent-work section, labeled "Application" rather than "Independent
Reading" in the source -- same shape as independent_reading, kept as a
separate field rather than reused so a Language lesson's own field name
maps directly to its own slide label). evidence_collection is Discourse
Deep Dive-only. whole_class_discourse_prompt (singular) stays Reading-only
as before; whole_class_discourse_prompts (plural, a list) is for Discourse
Deep Dive lessons that pose more than one numbered prompt in the same
Whole-Class Discourse section -- use whichever of the two actually matches
what the lesson shows, never populate both.

planner_sections (nested under a section, or under application) is for
reference charts and structured content that doesn't fit the single fixed
"chart" shape above -- e.g. a 3-column suffix reference chart, or a chart
where each cell is itself a bulleted list of multiple items rather than
one plain string. Prefer the plain "chart" field when content is a simple
flat 2-column table; use planner_sections when it isn't. Most sections
will have an empty planner_sections list -- only populate it for content
that genuinely needs the extra structure."""


def _build_content_block(lesson_path: str):
    """
    PDFs are sent to Claude as a real document content block (base64),
    not flattened to text first -- this lets Claude see the actual page
    layout, tables, and highlighting (the highlighted extraction-target
    boxes seen in real Teacher Guide excerpts are exactly the kind of
    visual signal that gets lost if the PDF is pre-flattened to text).
    """
    if lesson_path.lower().endswith(".pdf"):
        with open(lesson_path, "rb") as f:
            pdf_b64 = base64.standard_b64encode(f.read()).decode("ascii")
        return [
            {
                "type": "document",
                "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_b64},
            },
            {"type": "text", "text": "Extract this lesson per the schema and rules above."},
        ]
    with open(lesson_path, "r", encoding="utf-8") as f:
        return f.read()


def extract_from_path(lesson_path: str) -> dict:
    content_block = _build_content_block(lesson_path)
    return extract(content_block)


def extract(content) -> dict:
    """
    content: either a plain string (raw lesson text) or a list of
    Anthropic content blocks (e.g. a PDF document block + a text block).
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise SystemExit(
            "Set ANTHROPIC_API_KEY in the environment before running this "
            "script. Never hardcode a key in the file itself."
        )

    body = json.dumps({
        "model": "claude-sonnet-5",
        "max_tokens": 8000,
        "system": SYSTEM_PROMPT,
        "messages": [{"role": "user", "content": content}],
    }).encode("utf-8")

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=body,
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
        },
    )
    try:
        with urllib.request.urlopen(req) as resp:
            raw_response = resp.read()
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        raise SystemExit(
            f"Anthropic API returned HTTP {e.code}:\n{error_body}"
        ) from None

    data = json.loads(raw_response)

    # Real diagnostics instead of a bare crash: show what actually came
    # back (stop_reason, block types, a text snippet) whenever the
    # response isn't the clean JSON-only reply the system prompt asks
    # for, rather than letting json.loads fail with an opaque
    # "Expecting value" error that gives no hint why.
    content_blocks = data.get("content", [])
    raw_text = "".join(block.get("text", "") for block in content_blocks if block.get("type") == "text")

    if not raw_text.strip():
        block_types = [b.get("type") for b in content_blocks]
        raise SystemExit(
            "Extraction failed: Claude's response had no text content to "
            "parse.\n"
            f"  stop_reason: {data.get('stop_reason')!r}\n"
            f"  content block types: {block_types!r}\n"
            f"  full response: {json.dumps(data, indent=2)[:2000]}\n"
            "This usually means either the input was too large/unclear for "
            "the model to act on, or the request hit an unexpected stop "
            "condition -- the details above should show which."
        )

    # Defensive: the system prompt asks for raw JSON with no markdown
    # fences, but models sometimes wrap it in ```json ... ``` anyway.
    # Strip that before parsing rather than trust compliance blindly.
    cleaned_text = raw_text.strip()
    if cleaned_text.startswith("```"):
        cleaned_text = re.sub(r"^```[a-zA-Z]*\n?", "", cleaned_text)
        cleaned_text = re.sub(r"\n?```\s*$", "", cleaned_text)

    if data.get("stop_reason") == "max_tokens":
        raise SystemExit(
            "Extraction failed: Claude's response was cut off before "
            "finishing (hit the max_tokens limit), so the JSON is "
            "incomplete and can't be parsed.\n"
            f"  raw text so far (last 500 chars): {cleaned_text[-500:]!r}\n"
            "Fix: increase max_tokens in extract_lesson.py's extract() "
            "function -- this lesson's full extraction needs more room "
            "than what's currently configured."
        )

    try:
        result = json.loads(cleaned_text)
    except json.JSONDecodeError:
        raise SystemExit(
            "Extraction failed: Claude's response wasn't valid JSON despite "
            "the system prompt requiring it (checked for and stripped "
            "markdown code fences first, still didn't parse).\n"
            f"  stop_reason: {data.get('stop_reason')!r}\n"
            f"  raw text (first 2000 chars): {cleaned_text[:2000]!r}"
        ) from None

    if not result.get("supported", False):
        sys.stderr.write(
            f"NOTE: detected lesson_type={result.get('lesson_type')!r}, which "
            f"this platform doesn't have a slide generator for yet. "
            f"Currently supported: {SUPPORTED_TYPES}. Extraction still ran "
            f"and the JSON below reflects best-effort fields, but there is "
            f"no generate_deck script for this type.\n"
        )
    return result


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python3 extract_lesson.py <lesson_file.txt|.pdf>")
    result = extract_from_path(sys.argv[1])
    print(json.dumps(result, indent=2))
