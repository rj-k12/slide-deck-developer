# Running This on Replit (Temporary/Prototype Use)

## What I actually built vs. what I could verify

`replit.nix`, `.replit`, and `setup_replit.sh` are ready to try. **I have not
run any of this on Replit itself** -- I don't have access to it from this
sandbox. Everything below is a solid starting point based on how Replit's
Nix-based package system works, not a confirmed-working deployment.

## Real risks specific to Replit, stated plainly

1. **LibreOffice is heavy** (600MB+). Replit's storage limits vary by plan,
   and have changed over time -- I can't tell you with confidence what
   the current free-tier limit is or whether this fits inside it. If the
   Nix install fails or times out, this is almost certainly why.
2. **Free-tier repls have historically slept when idle**, and Replit's
   sleep/always-on policies change fairly often. If the webhook from
   `approve.php` needs to reach this reliably at any hour, don't assume
   an idle repl will be awake -- confirm current behavior on Replit's own
   pricing page before relying on it.
3. **This is genuinely fine for "let me try the pipeline end-to-end
   without setting up a real server"** -- prototyping, demoing, testing
   the webhook flow manually. I would not treat it as where this actually
   lives once real teachers depend on it.

## Setup

1. Create a new Replit, choose "import from GitHub" (or upload these
   files directly) -- make sure `replit.nix`, `.replit`, and
   `setup_replit.sh` are all in the root alongside everything else.
2. In Replit's **Secrets** panel (not a `.env` file -- Replit's own secret
   store), add:
   - `SLIDE_DECK_SERVICE_TOKEN`
   - `ANTHROPIC_API_KEY`
   - `UI_USERNAME` / `UI_PASSWORD`
   - `HOST` = `0.0.0.0` (required for Replit's webview/public URL to
     reach the app at all -- it binds to localhost-only by default,
     same as running it on your own laptop)
3. Click Run. First boot will be slow (installing LibreOffice via Nix,
   then `npm install`/`pip install` via `setup_replit.sh`) -- subsequent
   runs should be fast.
4. Replit gives you a public URL once it's running. That's what goes in
   `config.php`'s `SLIDE_DECK_SERVICE_URL`.

## If LibreOffice fails to install

That's the most likely failure point. At that point, the honest options
are: try a paid Replit tier with more storage, or fall back to one of the
already-verified-more-standard paths in `DEPLOY.md`/`DEPLOY_NO_DOCKER.md`
(Render, Railway, or a real VM) -- those give you a normal Linux
filesystem where installing LibreOffice via `apt` is a well-trodden path.
