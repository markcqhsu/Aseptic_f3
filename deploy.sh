#!/bin/bash
set -e
cd "$(dirname "$0")"
echo "▶ Building image..."
gcloud builds submit --project=cap-ocr
echo "▶ Deploying to Cloud Run..."
gcloud run deploy aseptic-f3-api \
  --image=asia.gcr.io/cap-ocr/aseptic-f3-api:latest \
  --region=asia-east1 \
  --project=cap-ocr \
  --service-account=aseptic-f3-api-sa@cap-ocr.iam.gserviceaccount.com \
  --set-secrets="ANTHROPIC_API_KEY=ANTHROPIC_API_KEY:latest" \
  --set-env-vars="APP_SHARED_KEY=1d4352cf93b7287a65b072c68c47490839e844c9" \
  --quiet
echo "✓ Done."
