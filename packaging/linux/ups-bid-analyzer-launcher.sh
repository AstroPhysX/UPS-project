#!/bin/sh
set -eu

# Flatpak supplies an application-specific XDG config directory. Running from
# it keeps bid_config.json out of the user's home-directory root while leaving
# the current relative CONFIG_PATH behavior unchanged.
CONFIG_DIR="${XDG_CONFIG_HOME:-${HOME}/.config}/ups-bid-analyzer"
mkdir -p "$CONFIG_DIR"
cd "$CONFIG_DIR"

export PYTHONPATH="/app/lib/ups-bid-analyzer/src${PYTHONPATH:+:${PYTHONPATH}}"
exec python3 -m bid_analyzer "$@"
