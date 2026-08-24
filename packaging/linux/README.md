# UPS Bid Analyzer Flatpak packaging

This folder is the Linux counterpart to `packaging/windows`.

## Output

Running `build_linux_flatpak.sh`:

1. Validates the project and PNG icon.
2. Adds the Flathub user remote when needed.
3. Generates a locked Python dependency manifest when it is missing.
4. Builds against KDE/Qt and the PySide6 BaseApp.
5. Installs the new build locally for testing.
6. Produces a distributable file under `packaging/linux/dist/`.

Expected output:

```text
packaging/linux/dist/UPS_Bid_Analyzer_1.0.0_x86_64.flatpak
```

## First build

From the project root:

```bash
chmod +x packaging/linux/*.sh
./packaging/linux/build_linux_flatpak.sh
```

Run the installed build:

```bash
flatpak run io.github.aleluc.BidAnalyzer
```

Install the generated bundle on another Linux computer:

```bash
flatpak install --user ./UPS_Bid_Analyzer_1.0.0_x86_64.flatpak
```

## Updating Python dependencies

The PySide6 runtime is supplied by `io.qt.PySide.BaseApp`; do not add PySide6
or shiboken6 to `requirements-flatpak.txt`.

When pandas, pdfplumber, openpyxl, or another runtime dependency changes:

```bash
./packaging/linux/refresh_python_dependencies.sh
./packaging/linux/build_linux_flatpak.sh
```

Commit `pypi-dependencies.json` after it is generated. It locks dependency URLs
and hashes, making subsequent builds repeatable.

## Version updates

For a new release, update:

- `APP_VERSION` in `build_linux_flatpak.sh`
- the `<release>` entry in `io.github.aleluc.BidAnalyzer.metainfo.xml`

Flatpak uses the `stable` branch while the filename carries the visible app
version.

## Application ID

The starter uses:

```text
io.github.aleluc.BidAnalyzer
```

Keep an application ID permanent after distributing the first release. If
`aleluc` is not the correct GitHub account name, change the ID before the first
public build in the manifest filename/content, desktop file, metadata file,
launcher filename, and shell scripts.

## Permissions

The initial manifest grants access to the user's home folder and removable
media so the current PDF browse and Excel export behavior works without code
changes. After testing, these permissions can be narrowed if all file access
works correctly through Qt's desktop portals.

## Removing the test installation

Keep app settings:

```bash
flatpak uninstall --user io.github.aleluc.BidAnalyzer
```

Remove the app and its Flatpak data:

```bash
./packaging/linux/uninstall_flatpak.sh
```
