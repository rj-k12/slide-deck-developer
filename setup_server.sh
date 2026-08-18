#!/bin/bash
# setup_server.sh -- one-time setup for a bare Ubuntu/Debian server
# (a fresh EC2 instance, DigitalOcean droplet, Linode, etc.). Run once
# as root or with sudo. Idempotent-ish -- safe to re-run if something
# fails partway through.
#
# Usage: sudo bash setup_server.sh

set -e

APP_DIR="/opt/redthread-slide-deck"
REPO_SOURCE="${1:-.}"  # pass the path to this project's files, or run
                       # from inside that directory and leave blank

echo "==> Installing system packages (Python, Node, LibreOffice)..."
apt-get update
apt-get install -y \
    python3 python3-venv python3-pip \
    nodejs npm \
    libreoffice-impress poppler-utils fonts-liberation \
    nginx

echo "==> Copying application files to $APP_DIR..."
mkdir -p "$APP_DIR"
cp -r "$REPO_SOURCE"/* "$APP_DIR"/
cd "$APP_DIR"

echo "==> Creating a dedicated system user (www-data already exists on Debian/Ubuntu)..."
chown -R www-data:www-data "$APP_DIR"

echo "==> Setting up Python virtual environment..."
python3 -m venv venv
./venv/bin/pip install --no-cache-dir -r requirements.txt

echo "==> Installing Node dependencies..."
npm install --production

echo "==> Creating .env for secrets (EDIT THIS FILE before starting the service)..."
if [ ! -f .env ]; then
    cat > .env << 'EOF'
SLIDE_DECK_SERVICE_TOKEN=REPLACE_WITH_A_LONG_RANDOM_STRING
ANTHROPIC_API_KEY=REPLACE_WITH_YOUR_KEY
UI_USERNAME=REPLACE_WITH_A_USERNAME
UI_PASSWORD=REPLACE_WITH_A_REAL_PASSWORD
RETENTION_HOURS=24
EOF
    chmod 600 .env
    chown www-data:www-data .env
    echo "    Wrote $APP_DIR/.env with placeholder values -- edit it before continuing."
fi

echo "==> Installing systemd service..."
cp redthread-slide-deck.service /etc/systemd/system/
systemctl daemon-reload
systemctl enable redthread-slide-deck

echo "==> Installing nginx site config..."
cp nginx-redthread-slide-deck.conf /etc/nginx/sites-available/redthread-slide-deck
ln -sf /etc/nginx/sites-available/redthread-slide-deck /etc/nginx/sites-enabled/
nginx -t

echo ""
echo "==> Setup complete. Remaining manual steps:"
echo "    1. Edit $APP_DIR/.env with real values."
echo "    2. Edit /etc/nginx/sites-available/redthread-slide-deck -- replace"
echo "       'your-domain.example.com' with your real domain."
echo "    3. sudo systemctl restart nginx"
echo "    4. sudo systemctl start redthread-slide-deck"
echo "    5. sudo systemctl status redthread-slide-deck   # confirm it's running"
echo "    6. sudo certbot --nginx -d your-domain.example.com   # real TLS cert"
