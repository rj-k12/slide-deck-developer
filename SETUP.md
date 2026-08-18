# RedThread Slide Deck Platform — Local UI

## One-time setup

You need Node.js and Python 3 installed. Then, from this folder:

```
pip install flask cairosvg
npm install pptxgenjs @tabler/icons
```

## Running it

```
python3 app.py
```

Then open **http://localhost:5001** in your browser. That's the only
command-line step — uploading lessons, generating decks, and downloading
results all happen in the browser from here on.

## Using it

**Two modes, as tabs at the top of the page:**

- **Raw lesson text** — paste or upload a Teacher Guide lesson excerpt as
  plain text, enter your Anthropic API key, and it runs the real
  extraction step (calls Claude to classify the lesson type and pull out
  the structured fields) before generating the deck.
- **Pre-extracted JSON** — skip extraction entirely by uploading a JSON
  file already in the schema `extract_lesson.py` produces (see
  `lesson_6a_REAL.json` / `lesson_6b_REAL.json` for real examples). No
  API key needed for this path — useful for testing.

Click **Generate slide deck**. You'll see the pipeline's log output
(icon resolution, cover generation, QA checks) right in the page, and a
download link appears once it's done.

## About the API key field

Typed into the browser, sent to your own local server, held in memory for
exactly the one request, then cleared — never written to disk, never
logged. Same safety pattern used earlier in this project for the
in-browser AI key prompt. This is still a local tool for your own use,
not something to expose to other people over a network.

## Currently supported lesson types

Reading, Close Reading, Literature Response. Anything else will fail with
a clear message rather than producing a wrong deck — see `build_deck.py`
and `extract_lesson.py` for what's classified as supported.
