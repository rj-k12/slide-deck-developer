# Easiest Way to Get This Online for Others to Test

## My actual recommendation: Render, via the dashboard

Of everything built across this whole conversation (Docker/Render,
Railway/Nixpasks, a bare VM, Replit, Cloud Run), **this is the one I'd
point you to first** if the goal is genuinely "least setup work, get a
URL, let people test it." The reason: it's the only path here that
needs zero command-line tooling installed on your end. Cloud Run needs
the `gcloud` CLI installed and authenticated first; a bare VM needs SSH
and comfort with systemd/nginx; Replit carries real, unverified risk
around whether LibreOffice even fits. Render is: push code to GitHub,
click around a web page, done.

`render.yaml` is now in this folder specifically to make this closer to
one click rather than filling out a form by hand.

## The actual steps

1. **Get this code into a GitHub repo** if it isn't already (Render
   deploys from a connected repo, not a local file upload).
2. Go to **render.com**, sign up/log in, click **New > Blueprint**.
3. Connect the GitHub repo. Render will detect `render.yaml`
   automatically and show you the service it's about to create,
   pre-filled with everything except the actual secret values.
4. Type in the four values it prompts for:
   - `SLIDE_DECK_SERVICE_TOKEN` — any long random string you make up
   - `ANTHROPIC_API_KEY` — your real key
   - `UI_USERNAME` / `UI_PASSWORD` — whatever you want people to log in
     with
5. Click **Apply**. Render builds your Dockerfile and deploys it.
6. You get a `https://your-service-name.onrender.com` URL when it's
   done — that's the link to share with people for testing, and also
   what goes in `config.php`'s `SLIDE_DECK_SERVICE_URL` if you're also
   wiring up the `approve.php` webhook.

## The one real tradeoff, stated plainly

Render's free tier **spins down after a period of inactivity** and has
to cold-start on the next visit — meaning if nobody's used it in a
while, the first person to try it will sit waiting for maybe 30-60+
seconds (spin-up time) before LibreOffice's own cold-start on top of
that. This is genuinely fine for "a few people testing this over a few
days" — it's not fine as a permanent home once people expect it to
always be instantly available. When you outgrow that, Render also has
a paid tier that stays always-on, and everything else here (Cloud Run,
the bare-VM path) still works exactly as already documented.

## If you'd rather not use GitHub

Cloud Run (`DEPLOY_CLOUDRUN.md`) is the next-best option — it does need
the `gcloud` CLI installed once, but after that it's a single script
(`deploy_cloudrun.sh`) with no GitHub step at all, deploying straight
from the files on your own machine.
