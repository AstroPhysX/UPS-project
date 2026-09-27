#!/usr/bin/env bash
set -euo pipefail

APP_ID="com.upsbidanalyzer.BidAnalyzer"
APP_VERSION="1.0.0"
BRANCH="master"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"

MANIFEST="$SCRIPT_DIR/installer_linux.yml"
PARTS_DIR="$SCRIPT_DIR/Bid_Analyzer"
INSTALLER_DIR="$SCRIPT_DIR/installer"

STATE_DIR="$PARTS_DIR/.flatpak-builder"
BUILD_DIR="$PARTS_DIR/build-dir"
REPO_DIR="$PARTS_DIR/repo"

ARCH="$(flatpak --default-arch 2>/dev/null || uname -m)"
BUNDLE="$INSTALLER_DIR/UPS_Bid_Analyzer_${APP_VERSION}_${ARCH}.flatpak"

die() {
    echo
    echo "ERROR: $*" >&2
    echo
    exit 1
}

echo
echo "============================================================"
echo " UPS Bid Analyzer - Linux Flatpak Builder"
echo "============================================================"
echo
echo "Application ID: $APP_ID"
echo "Version:        $APP_VERSION"
echo "Project:        $PROJECT_ROOT"
echo "Final output:   $BUNDLE"
echo

command -v flatpak >/dev/null 2>&1 || die "Flatpak is not installed."
[[ -f "$MANIFEST" ]] || die "Manifest not found: $MANIFEST"
[[ -d "$PROJECT_ROOT/src" ]] || die "Project src directory not found: $PROJECT_ROOT/src"
[[ -f "$PARTS_DIR/pypi-dependencies.json" ]] || die "Missing $PARTS_DIR/pypi-dependencies.json"
[[ -f "$PARTS_DIR/com.upsbidanalyzer.BidAnalyzer.desktop" ]] || die "Missing desktop file"
[[ -f "$PARTS_DIR/com.upsbidanalyzer.BidAnalyzer.metainfo.xml" ]] || die "Missing metainfo file"
[[ -f "$PARTS_DIR/ups-bid-analyzer-launcher.sh" ]] || die "Missing launcher"
[[ -f "$PROJECT_ROOT/src/bid_analyzer/resources/logo.png" ]] || die "Missing logo.png"

mkdir -p "$PARTS_DIR" "$INSTALLER_DIR"

if ! flatpak remotes --columns=name 2>/dev/null | grep -qx "flathub"; then
    flatpak remote-add --user --if-not-exists flathub \
        https://dl.flathub.org/repo/flathub.flatpakrepo
fi

echo "Checking Flatpak build dependencies..."
flatpak install --user -y flathub \
    org.kde.Sdk//6.11 \
    org.kde.Platform//6.11 \
    io.qt.PySide.BaseApp//6.11

if command -v flatpak-builder >/dev/null 2>&1; then
    BUILDER=(flatpak-builder)
else
    echo "Host flatpak-builder not found; using org.flatpak.Builder..."
    if ! flatpak info org.flatpak.Builder >/dev/null 2>&1; then
        flatpak install --user -y flathub org.flatpak.Builder
    fi
    BUILDER=(flatpak run org.flatpak.Builder)
fi

rm -rf "$BUILD_DIR" "$REPO_DIR"
rm -f "$BUNDLE"

cd "$PROJECT_ROOT"

"${BUILDER[@]}" \
    --force-clean \
    --state-dir="$STATE_DIR" \
    --repo="$REPO_DIR" \
    "$BUILD_DIR" \
    "$MANIFEST"

flatpak build-bundle \
    --runtime-repo=https://dl.flathub.org/repo/flathub.flatpakrepo \
    "$REPO_DIR" \
    "$BUNDLE" \
    "$APP_ID" \
    "$BRANCH"

[[ -f "$BUNDLE" ]] || die "The .flatpak bundle was not created."

rm -rf "$BUILD_DIR" "$REPO_DIR"

echo
echo "============================================================"
echo " SUCCESS"
echo "============================================================"
echo "Installer:"
echo "  $BUNDLE"
echo
echo "Install/reinstall with:"
echo "  flatpak install --user --reinstall \"$BUNDLE\""
echo
echo "Run with:"
echo "  flatpak run $APP_ID"
