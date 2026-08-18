# Deploying to Google Cloud Run

## What changed to make this Cloud-Run-ready

The Dockerfile's startup command was hardcoded to port 5001. Cloud Run
injects its own `PORT` environment variable at container start (commonly
8080, but never assume a fixed value) and requires the container to
listen on whatever it provides -- fixed it to read `$PORT` with a
fallback, and switched to shell-form `CMD` specifically because the
exec-form array syntax used everywhere else in this file does NOT expand
environment variables, so `$PORT` would have been passed to gunicorn as
a literal string.

## One-time setup

```
export GCP_PROJECT=your-project-id
export SLIDE_DECK_SERVICE_TOKEN=some-long-random-string
export ANTHROPIC_API_KEY=sk-ant-...
export UI_USERNAME=teacher
export UI_PASSWORD=something-real

bash deploy_cloudrun.sh
```

That's the whole deployment. `--source .` tells Cloud Run to build your
Dockerfile via Cloud Build directly -- no separate `docker build`/`docker
push` step, unlike the plain-Docker path in `DEPLOY.md`.

The script prints your service URL at the end. That's what goes in
`config.php`'s `SLIDE_DECK_SERVICE_URL`.

## Redeploying after a code change

Same command, same script -- Cloud Run builds a new revision from source
each time and shifts traffic to it automatically.

## Two things worth knowing before you rely on this

**Environment variables here are plain, not Secret Manager.** They're
visible to anyone with read access to the Cloud Run service in your GCP
console, and show up in `gcloud run services describe`. Fine for a
solo/small-team temporary setup; if this becomes long-lived with more
people having project access, moving `ANTHROPIC_API_KEY` specifically
into Secret Manager (`--set-secrets` instead of `--set-env-vars`) is the
safer next step -- happy to build that out if/when you get there.

**Cold starts will be slow, specifically because of LibreOffice.**
`--min-instances 0` means the container fully shuts down when idle (this
is what keeps it free/cheap) and has to cold-start on the next request.
LibreOffice is not a fast thing to boot from zero -- the first
cover-generation call after any idle period will likely take
noticeably longer than subsequent ones. If that first-request latency
ever becomes a real problem (e.g. `approve.php`'s webhook times out
waiting), the fix is `--min-instances 1`, which keeps one instance
always warm -- at the cost of it no longer being free, since you're now
paying for that instance's uptime continuously rather than per-request.

## Checking on it

```
gcloud run services describe redthread-slide-deck --region us-central1
gcloud run services logs read redthread-slide-deck --region us-central1
```
