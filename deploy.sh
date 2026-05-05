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
  --quiet
echo "✓ Done."
