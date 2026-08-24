#!/usr/bin/env bash
set -Eeuo pipefail

LINUX_PACKAGING_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd -- "$LINUX_PACKAGING_DIR/../.." && pwd)"

APP_ID="io.github.aleluc.BidAnalyzer"
APP_NAME="UPS_Bid_Analyzer"
APP_VERSION="1.0.0"
BRANCH="stable"
MANIFEST="$LINUX_PACKAGING_DIR/$APP_ID.yml"
BUILD_DIR="$LINUX_PACKAGING_DIR/build-dir"
REPO_DIR="$LINUX_PACKAGING_DIR/repo"
DIST_DIR="$LINUX_PACKAGING_DIR/dist"
ARCH="$(flatpak --default-arch 2>/dev/null || uname -m)"
BUNDLE="$DIST_DIR/${APP_NAME}_${APP_VERSION}_${ARCH}.flatpak"

fail() {
    printf '\n============================================================\n' >&2
    printf 'BUILD FAILED\n' >&2
    printf '============================================================\n' >&2
    printf '%s\n\n' "$*" >&2
    exit 1
}

command -v flatpak >/dev/null 2>&1 || fail "Flatpak is not installed."
[[ -f "$MANIFEST" ]] || fail "Manifest not found: $MANIFEST"
[[ -f "$PROJECT_ROOT/src/bid_analyzer/__main__.py" ]] || fail "Entry module not found under src/bid_analyzer."
[[ -f "$PROJECT_ROOT/src/bid_analyzer/resources/app_icon.png" ]] || fail "PNG icon not found: src/bid_analyzer/resources/app_icon.png"

if ! flatpak remotes --user --columns=name 2>/dev/null | grep -Fxq flathub; then
    echo "Adding the Flathub user remote..."
    flatpak remote-add --user --if-not-exists flathub \
        https://dl.flathub.org/repo/flathub.flatpakrepo
fi

if [[ ! -f "$LINUX_PACKAGING_DIR/pypi-dependencies.json" ]]; then
    echo "Python dependency manifest is missing; generating it now..."
    "$LINUX_PACKAGING_DIR/refresh_python_dependencies.sh"
fi

if command -v flatpak-builder >/dev/null 2>&1; then
    BUILDER=(flatpak-builder)
else
    echo "The host flatpak-builder command is unavailable."
    echo "Installing the Flatpak Builder application instead..."
    flatpak install --user -y flathub org.flatpak.Builder
    BUILDER=(flatpak run org.flatpak.Builder)
fi

printf '\n============================================================\n'
printf 'Building %s %s\n' "$APP_NAME" "$APP_VERSION"
printf '============================================================\n'
printf 'Project:\n  %s\n\n' "$PROJECT_ROOT"
printf 'Manifest:\n  %s\n\n' "$MANIFEST"
printf 'Bundle output:\n  %s\n\n' "$BUNDLE"

rm -rf "$BUILD_DIR" "$REPO_DIR"
mkdir -p "$DIST_DIR"
rm -f "$BUNDLE"

cd "$LINUX_PACKAGING_DIR"

"${BUILDER[@]}" \
    --user \
    --force-clean \
    --install-deps-from=flathub \
    --repo="$REPO_DIR" \
    --install \
    "$BUILD_DIR" \
    "$MANIFEST"

flatpak build-bundle \
    "$REPO_DIR" \
    "$BUNDLE" \
    "$APP_ID" \
    "$BRANCH" \
    --runtime-repo=https://dl.flathub.org/repo/flathub.flatpakrepo

[[ -f "$BUNDLE" ]] || fail "Flatpak build completed without creating the expected bundle."

printf '\n============================================================\n'
printf 'BUILD SUCCESSFUL\n'
printf '============================================================\n'
printf 'Installed test application:\n  flatpak run %s\n\n' "$APP_ID"
printf 'Distributable bundle:\n  %s\n\n' "$BUNDLE"
printf 'Install the bundle on another Linux computer with:\n'
printf '  flatpak install --user %q\n\n' "$BUNDLE"

if command -v xdg-open >/dev/null 2>&1; then
    xdg-open "$DIST_DIR" >/dev/null 2>&1 || true
fi
