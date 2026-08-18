#!/bin/bash
# setup_replit.sh -- runs npm/pip installs on first boot. Nix (replit.nix)
# handles the system-level Python/Node/LibreOffice binaries; this handles
# the actual project dependencies on top of them.
set -e

if [ ! -d "node_modules" ]; then
    echo "Installing Node dependencies..."
    npm install --production
fi

if ! python3 -c "import flask" 2>/dev/null; then
    echo "Installing Python dependencies..."
    pip install --break-system-packages -r requirements.txt
fi
