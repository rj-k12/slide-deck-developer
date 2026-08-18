#!/bin/bash
# deploy_cloudrun.sh -- deploys this app to Google Cloud Run.
#
# Prerequisites (can't be done for you from here):
#   1. A Google Cloud project with billing enabled (Cloud Run's free tier
#      still requires a billing account on file, even though the free
#      allowance itself doesn't charge you).
#   2. gcloud CLI installed and authenticated: `gcloud auth login`
#   3. This directory contains the Dockerfile, app.py, package.json,
#      requirements.txt, and everything else from the platform.
#
# Usage:
#   export GCP_PROJECT=your-project-id
#   export SLIDE_DECK_SERVICE_TOKEN=some-long-random-string
#   export ANTHROPIC_API_KEY=sk-ant-...
#   export UI_USERNAME=teacher
#   export UI_PASSWORD=something-real
#   bash deploy_cloudrun.sh

set -e

: "${GCP_PROJECT:?Set GCP_PROJECT to your Google Cloud project ID}"
: "${SLIDE_DECK_SERVICE_TOKEN:?Set SLIDE_DECK_SERVICE_TOKEN}"
: "${ANTHROPIC_API_KEY:?Set ANTHROPIC_API_KEY}"

REGION="${REGION:-us-central1}"
SERVICE_NAME="${SERVICE_NAME:-redthread-slide-deck}"

echo "==> Setting project..."
gcloud config set project "$GCP_PROJECT"

echo "==> Enabling required APIs (safe to re-run, no-op if already enabled)..."
gcloud services enable run.googleapis.com cloudbuild.googleapis.com

echo "==> Building and deploying from source..."
# --source . tells Cloud Run to build the Dockerfile in this directory via
# Cloud Build itself -- no separate manual "docker build && docker push"
# step needed, unlike the plain-Docker deployment path in DEPLOY.md.
gcloud run deploy "$SERVICE_NAME" \
    --source . \
    --region "$REGION" \
    --platform managed \
    --memory 1Gi \
    --cpu 1 \
    --timeout 120 \
    --min-instances 0 \
    --max-instances 3 \
    --allow-unauthenticated \
    --set-env-vars "SLIDE_DECK_SERVICE_TOKEN=${SLIDE_DECK_SERVICE_TOKEN},ANTHROPIC_API_KEY=${ANTHROPIC_API_KEY},UI_USERNAME=${UI_USERNAME:-},UI_PASSWORD=${UI_PASSWORD:-}"

echo ""
echo "==> Done. The deployed URL is printed above -- that's what goes in"
echo "    config.php's SLIDE_DECK_SERVICE_URL."
