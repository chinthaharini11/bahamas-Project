#!/bin/bash

# Write Google credentials from env var to file if provided
if [ -n "$GOOGLE_APPLICATION_CREDENTIALS_JSON" ]; then
    echo "$GOOGLE_APPLICATION_CREDENTIALS_JSON" > /app/gcp-credentials.json
    export GOOGLE_APPLICATION_CREDENTIALS=/app/gcp-credentials.json
    echo "Google credentials configured"
fi

exec "$@"
