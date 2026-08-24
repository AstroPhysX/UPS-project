#!/usr/bin/env bash
set -Eeuo pipefail

APP_ID="io.github.aleluc.BidAnalyzer"

if ! flatpak info --user "$APP_ID" >/dev/null 2>&1; then
    echo "$APP_ID is not installed for this user."
    exit 0
fi

# --delete-data removes the app-specific Flatpak configuration and cache too.
flatpak uninstall --user --delete-data -y "$APP_ID"
echo "Removed $APP_ID and its application data."
