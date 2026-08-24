#!/usr/bin/env bash
set -Eeuo pipefail
exec flatpak run io.github.aleluc.BidAnalyzer "$@"
