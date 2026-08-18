# Deploying the Slide Deck Platform Online

## Where to host it

WP Engine can't run this — it's PHP/MySQL-only, and this needs Python,
Node, and LibreOffice. That's fine: this tool has no shared-database
dependency on RedThread in the first place, so there's no reason it needs
to live where RedThread lives.

Anything that runs Docker works. Two reasonable starting points:

- **Render** — connect the repo, choose "Docker" as the environment, set
  the environment variables below, deploy. Probably the least fiddly.
- **Railway** — same idea, similarly quick from a Dockerfile.

Both give you a public HTTPS URL immediately, which is what
`SLIDE_DECK_SERVICE_URL` in `config.php` needs to point at.

## Environment variables to set on the host

| Variable | Required? | What it's for |
|---|---|---|
| `SLIDE_DECK_SERVICE_TOKEN` | Yes | Shared secret `approve.php` sends as a bearer token. Generate a long random string; put the same value in `config.php`'s `SLIDE_DECK_SERVICE_TOKEN`. |
| `ANTHROPIC_API_KEY` | Yes, for real extraction | Server-side now, not typed into a browser field — needed for both the webhook path and the "raw lesson text" UI path. |
| `UI_USERNAME` / `UI_PASSWORD` | Strongly recommended | Basic auth on the browser UI. Without these, anyone with the URL can generate decks and burn your API credits. The `/webhook/*` routes have their own separate bearer-token auth regardless. |
| `RETENTION_HOURS` | Optional (default 24) | How long generated decks and webhook job logs stick around before automatic cleanup. |

## Building and running locally first (recommended before deploying)

```
docker build -t redthread-slide-deck .
docker run -p 5001:5001 \
  -e SLIDE_DECK_SERVICE_TOKEN=some-long-random-string \
  -e UI_USERNAME=teacher -e UI_PASSWORD=something-real \
  redthread-slide-deck
```

Then hit `http://localhost:5001` same as before. If this works locally,
it'll work identically on Render/Railway — same Dockerfile either way.

## Wiring up `approve.php`

Two constants need adding to `config.php` (not `approve.php` itself):

```php
define('SLIDE_DECK_SERVICE_URL', 'https://your-deployed-url.onrender.com');
define('SLIDE_DECK_SERVICE_TOKEN', 'the-same-long-random-string-as-above');
```

That's the only change needed on the RedThread side — the webhook-calling
code is already in `approve.php` from the last round.

## What's still true regardless of where this runs

- **The `lesson_text: null` gap is unchanged.** Deploying this doesn't
  resolve it — the webhook will still receive an empty `lesson_text`
  until the ingestion side can hand over actual lesson prose. Decks
  generated via the webhook path will sit at `waiting_on_lesson_text`
  until that's resolved; check `/webhook/status/<lesson_id>` to see.
- **Only 3 of 8 lesson types are supported** (Reading, Close Reading,
  Literature Response). Everything else fails with a clear message
  rather than guessing.
- **Generated files don't persist across a redeploy/restart** unless you
  attach a real persistent volume at `/app/generated` — by default,
  restarting the container loses whatever's in there (on top of the
  1-hour cleanup sweep already removing anything past `RETENTION_HOURS`
  regardless).
