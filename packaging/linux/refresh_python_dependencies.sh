#!/usr/bin/env bash
set -Eeuo pipefail

LINUX_PACKAGING_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd -- "$LINUX_PACKAGING_DIR/../.." && pwd)"
TOOLS_ROOT="$PROJECT_ROOT/build/flatpak-tools"
TOOLS_REPO="$TOOLS_ROOT/flatpak-builder-tools"
TOOLS_VENV="$TOOLS_ROOT/venv"
GENERATOR="$TOOLS_REPO/pip/flatpak-pip-generator.py"
RUNTIME="org.kde.Sdk//6.11"

fail() {
    printf '\nERROR: %s\n\n' "$*" >&2
    exit 1
}

command -v flatpak >/dev/null 2>&1 || fail "Flatpak is not installed."
command -v git >/dev/null 2>&1 || fail "Git is not installed."
command -v python3 >/dev/null 2>&1 || fail "Python 3 is not installed."

if ! flatpak remotes --user --columns=name 2>/dev/null | grep -Fxq flathub; then
    echo "Adding the Flathub user remote..."
    flatpak remote-add --user --if-not-exists flathub \
        https://dl.flathub.org/repo/flathub.flatpakrepo
fi

# The generator uses the SDK to determine valid wheel tags.
flatpak install --user -y flathub "$RUNTIME"

mkdir -p "$TOOLS_ROOT"
if [[ ! -d "$TOOLS_REPO/.git" ]]; then
    echo "Downloading flatpak-builder-tools..."
    git clone --depth 1 \
        https://github.com/flatpak/flatpak-builder-tools.git \
        "$TOOLS_REPO"
else
    echo "Updating flatpak-builder-tools..."
    git -C "$TOOLS_REPO" pull --ff-only
fi

if [[ ! -x "$TOOLS_VENV/bin/python" ]]; then
    python3 -m venv "$TOOLS_VENV"
fi

"$TOOLS_VENV/bin/python" -m pip install --quiet --upgrade pip
"$TOOLS_VENV/bin/python" -m pip install --quiet requirements-parser packaging

cd "$LINUX_PACKAGING_DIR"
rm -f pypi-dependencies.json

# Prefer platform wheels for packages that otherwise need substantial native
# compilation. The generated JSON still pins every URL and SHA-256 hash.
"$TOOLS_VENV/bin/python" "$GENERATOR" \
    --runtime "$RUNTIME" \
    --requirements-file requirements-flatpak.txt \
    --output pypi-dependencies \
    --prefer-wheels=numpy,pandas,pillow,pypdfium2,cryptography,cffi

[[ -f pypi-dependencies.json ]] || fail "Dependency generation did not create pypi-dependencies.json."

echo
echo "Python dependency manifest refreshed:"
echo "  $LINUX_PACKAGING_DIR/pypi-dependencies.json"
echo
