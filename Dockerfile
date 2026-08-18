# RedThread Slide Deck Platform
#
# Bundles everything the pipeline needs to actually run on a real server:
# Python (Flask + the generation/extraction scripts), Node (pptxgenjs +
# the icon set), and LibreOffice (needed only for rendering the template's
# cover slide to an image -- confirmed by direct testing that nothing
# else in the pipeline calls it; QA runs on pure python-pptx instead).

FROM node:20-slim

# --- System packages: Python + LibreOffice (Impress only, not the full
# suite, to keep the image reasonably sized) + poppler-utils for pdftoppm
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 \
    python3-pip \
    libreoffice-impress \
    poppler-utils \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# --- Node dependencies ---
COPY package.json .
RUN npm install --production

# --- Python dependencies ---
COPY requirements.txt .
RUN pip3 install --break-system-packages --no-cache-dir -r requirements.txt

# --- Application code ---
COPY . .

# Generated decks and webhook job logs live here -- see the cleanup note
# in app.py. Not a volume by default: if the platform restarts, old
# generated files are expected to be gone. Mount a real volume here if
# generated decks need to survive a restart.
RUN mkdir -p generated

EXPOSE 5001

# gunicorn instead of `python3 app.py` -- the Flask dev server used during
# local testing explicitly warns against production use. Single worker
# with threads (not multiple worker processes) because build_deck.py's
# background webhook processing uses in-process threading; multiple
# worker processes would each have their own thread pool and the
# /webhook/status endpoint could hit a different worker than the one
# doing the work.
# Shell form (not exec form) is required here specifically so $PORT gets
# expanded -- Cloud Run injects its own PORT env var at container start
# (commonly 8080, but never assume a fixed value) and requires the
# container to listen on whatever it provides. The exec-form array syntax
# used everywhere else in this file does NOT go through a shell, so
# "$PORT" would be passed to gunicorn as a literal, unexpanded string.
CMD gunicorn --bind 0.0.0.0:${PORT:-5001} --workers 1 --threads 4 --timeout 120 app:app
