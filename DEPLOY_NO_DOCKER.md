# Deploying Without Docker

Two real options, depending on how much control you want.

## Option A: Railway with Nixpacks (closest to "just push it")

Render's non-Docker service type is single-runtime only (Python *or*
Node, not both) — that's specifically why Docker got recommended for
Render. Railway's default builder, Nixpacks, is different: it detects
multiple runtimes from what's in the repo (`requirements.txt` → Python,
`package.json` → Node) *and* lets you specify system packages via a
config file. `nixpacks.toml` is already set up to do exactly that.

1. Push this repo to GitHub (or wherever Railway pulls from).
2. Create a new Railway project from that repo. `nixpacks.toml` is
   picked up automatically — no dropdown to change, no Dockerfile needed.
3. Set the same four environment variables as before: `SLIDE_DECK_SERVICE_TOKEN`,
   `ANTHROPIC_API_KEY`, `UI_USERNAME`, `UI_PASSWORD`.
4. Deploy. Railway assigns the public port via `$PORT`, which the start
   command already binds to.

**Honest caveat**: I can't run Nixpacks in this environment to confirm
the actual build succeeds end-to-end — I validated the TOML syntax is
correct and that the install commands are identical to ones already
proven working here, but the real build (especially whether Railway's
Nixpacks image has `apt` access for the LibreOffice install the way I've
assumed) is unverified. If it fails on `aptPkgs`, that's the first thing
to check.

## Option B: A real Linux server — EC2, a DigitalOcean droplet, Linode, any VPS

Full manual control, no PaaS abstraction. Three new files handle this:

- **`setup_server.sh`** — one-time setup script: installs Python, Node,
  LibreOffice, and nginx via `apt`, sets up a virtual environment,
  installs both sets of dependencies, and installs the systemd service
  and nginx config below.
- **`redthread-slide-deck.service`** — a systemd unit so the app runs as
  a real background service, restarts automatically if it crashes, and
  survives a server reboot.
- **`nginx-redthread-slide-deck.conf`** — nginx as the actual
  internet-facing layer, handling TLS and proxying to gunicorn, which
  only listens on `127.0.0.1`. This is a meaningfully safer split than
  binding gunicorn straight to `0.0.0.0` the way the Docker version does
  — nginx is the only thing directly reachable from the internet.

**Steps:**
```
# On a fresh Ubuntu/Debian server, as root:
sudo bash setup_server.sh /path/to/this/project

# Then, following the script's own printed instructions:
sudo nano /opt/redthread-slide-deck/.env          # fill in real secrets
sudo nano /etc/nginx/sites-available/redthread-slide-deck  # your real domain
sudo systemctl restart nginx
sudo systemctl start redthread-slide-deck
sudo systemctl status redthread-slide-deck        # confirm it's running
sudo certbot --nginx -d your-domain.example.com   # free TLS cert
```

**Honest caveat**: I validated the shell script's syntax and the systemd
unit file's structure, but I don't have a bare Linux VM in this
environment to actually run the full setup end-to-end. The individual
pieces (the `apt install` package names, the gunicorn command, the nginx
proxy config) all match what's already confirmed working in the Docker
version — but "confirmed working in Docker" and "confirmed working via
this exact script on a fresh VM" aren't quite the same claim, and I want
to be direct about that difference rather than imply I tested something
I couldn't.

## What's identical either way

- Same four environment variables
- Same `config.php` change on the RedThread/WP Engine side
- Same known gap: `lesson_text` is still `null` from the webhook until
  the ingestion side can hand over real lesson prose — deployment method
  doesn't touch that at all
