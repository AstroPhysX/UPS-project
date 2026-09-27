#!/bin/sh
set -eu

CONFIG_DIR="${XDG_CONFIG_HOME:-${HOME}/.config}/ups-bid-analyzer"
mkdir -p "$CONFIG_DIR"
cd "$CONFIG_DIR"

export PYTHONPATH="/app/lib/ups-bid-analyzer/src${PYTHONPATH:+:${PYTHONPATH}}"
exec python3 -m bid_analyzer "$@"
