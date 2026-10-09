# Color Tools

A compact color picker and palette manager with an optional always-on-top mode for Windows, built with [DearPyGui](https://github.com/hoffstadt/DearPyGui). Pick colors from the screen, explore color harmonies and color theory, check contrast and color-blindness simulations, and convert palettes between industry-standard formats — all from a single small window. The default language is **English**. Finnish can be selected in the Info tab.

![Platform](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/tested-Python_3.14-blue)
![License](https://img.shields.io/badge/license-MIT-green)

---

## Features

### Color Picker

- **Color wheel and sliders:** Switch between a color wheel and precision gradient sliders. Slider modes: RGB, HSL, CMYK, LAB and Grayscale. All modes stay in sync, and each value can also be typed in directly.
- **Screen pipette:** Press `Alt + P` (registered as a global shortcut when available) or click the picker button to sample a screen pixel, with a live color preview next to the cursor. Left-click picks, right-click or `Esc` cancels.
- **Hex input:** Type or paste a hex code (`#RRGGBB` or shorthand `#RGB`).
- **Quick copy:** Copy the current HEX color or a harmony color in the selected display format.

### Harmony

- **9 harmony modes:** Complementary, Split Complementary, Analogous, Triadic, Tetradic, Rectangle, Tints, Shades and Tones.
- **6 display formats:** HEX, RGB, HSL, HSV, CSS Name (nearest entry in the built-in CSS color table, matched in CIELAB) and CMYK.
- **Save as palette:** Store the generated harmony as a new palette with one click.
- **HTML report:** Export the harmony as a styled HTML page with the six listed color formats, contrast ratios against white and black, color theory text and HEX copy buttons. Clipboard access in the exported page depends on browser permissions.

### Color Theory

The **Theory** tab explains the current color:

- **Key associations:** the color's main associations as a one-line list, e.g. *softness · comfort · quiet elegance*, followed by a short explanation of why.
- **Colour temperature:** a labelled cool-to-warm meter with a marker, plus a note on whether the color tends to advance (draw attention) or recede (feel calm). Vivid colors on the warm/cool border, such as green and magenta, are marked *in between*; *neutral* is reserved for near-greys.
- **Meaning:** nearest CSS name, what the color is generally associated with and why, its cultural and historical background, and its visual tone and saturation character.
- **In a UI:** a hint about the role the color typically plays in an interface — error, success, primary brand color, soft background, muted accent, and so on.
- **Fine-grained categories:** beyond the main hues, sand/beige, cream/ivory, browns, olive, and warm and cool greys (taupe, slate) have their own descriptions.
- Theory descriptions are available in English and Finnish. The theory text is also included in HTML reports.

### Contrast Checking

- **WCAG contrast:** See the contrast ratio of the current color against white, black and a custom background, with AAA / AA / AA* (large text only) / Fail ratings and a live sample text.
- Set the custom background from the current color with one click.

### Color Blindness Simulation

- The **Simulation** tab shows how the current color (as text on your custom background) appears with:
  - Protanopia (red-blind)
  - Deuteranopia (green-blind)
  - Tritanopia (blue-blind)
  - …compared with normal vision.
- The simulation applies color-transformation matrices in linear RGB to both the text and its custom background, then converts back to display RGB.

### History

- Keeps up to 60 recent color entries, including color edits and imported palette colors. Consecutive identical history additions are skipped.
- Right-click any swatch to remove it.
- Select multiple colors to save them as a new palette, or export the whole history as HTML.

### Palettes & Conversion

- **Palette converter:** Import a palette from one app and export it for another, e.g. Procreate `.swatches` → Affinity `.afpalette`, or Photoshop `.aco` → GIMP `.gpl`.
- **Palette formats with import and export support (subject to the limitations below):**

  | Format | Used by |
  |--------|---------|
  | `.afpalette` | Affinity Designer / Photo / Publisher |
  | `.ase` (Adobe Swatch Exchange) | Illustrator, InDesign, Affinity |
  | `.aco` (Adobe Color) | Photoshop |
  | `.gpl` (GIMP Palette) | GIMP, Inkscape, Krita |
  | `.swatches` | Procreate |

  `.ase` import reads RGB, CMYK, Gray and LAB colors; `.aco` import (v1 and v2) reads RGB, HSB and CMYK. Unsupported ASE/ACO color spaces are skipped. **HTML report** is export-only.

  Imported colors are converted to opaque 8-bit RGB. Original color names, profiles, spot-color definitions and other palette metadata are not preserved. Procreate import reads HSV-based swatch entries. Affinity support is described below; support for a file extension does not imply support for every variant produced by those applications.

- **Drag & drop:** Drop one palette file or image onto the window. A preview lets you confirm before it is added. Only one import can be pending at a time.
- **Import from image:** Choose a target of 2–20 colors from supported PNG, JPEG, BMP, GIF, WebP or TIFF files (requires Pillow). The preview may contain fewer colors if fewer are present. Images are reduced to at most 200 × 200 pixels before color extraction; transparency is composited onto white. Animated or multi-page images use the initial frame.
- **Palette editor:** Browse 96 colors per page, select and edit colors, reorder by drag-and-drop within a page, or use **Move to position…** across pages. Right-click removes a color. Up to 30 palette-edit snapshots are retained across all palettes during the session; Undo applies to the currently open palette. Changing pages clears the selection. Deleting an entire palette is not undoable.
- **Palette list:** 20 palettes per page, with compact color previews.
- Creating or importing a new palette is blocked when 100 palettes already exist. Imports retain at most 1,000 colors, and adding a color is blocked once a palette has 1,000 colors. Larger palettes and collections already present in valid settings remain accessible without truncation.

### General

- **7 themes:** Dark, Light, Midnight, Mocha, Nord, Solarized and Black.
- **English / Finnish UI:** English is the default on first launch. Switch language in the Info tab; an existing saved language choice is preserved.
- **Always on top:** Toggleable pin-to-top.
- **Compact layout:** Collapsible sections and a custom borderless title bar; system-DPI-aware layout. Windows scales the window on monitors with a different scale; it may appear less sharp there.
- History, palettes and their order, window position, theme, language, always-on-top mode and collapsed sections are saved automatically to `%APPDATA%\Color Tools\config.json`. The active color, custom contrast background, selected modes and undo history are not restored on restart.
- If the settings file cannot be read, the app attempts to preserve a `.bak` copy before using defaults. If the backup fails, settings writes are disabled for that session. Files exceeding the settings size limit are also protected from replacement with defaults.

---

## Run the Windows application

Download `ColorTools.exe` from this repository's **Releases** when a release is available.
The executable bundles Python and the application dependencies.
Close an older Color Tools instance before starting a new version.

## Run from source

Built and regression-tested on Windows 11 x64 with Python 3.14.3. Automated checks used Dear PyGui without a visible viewport; they do not constitute a complete interactive EXE test. The application uses Win32 APIs;
it does not run on Linux or macOS. Windows 10 has not been verified for this build.

From the repository root in PowerShell:

```powershell
python -m venv .venv
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
& .\.venv\Scripts\python.exe -B run_app.py
```

Pinned runtime dependencies: Dear PyGui 2.2 and Pillow 12.3.0.
Image import is disabled if Pillow is missing or older than 12.3.0.

## Build an EXE

```powershell
& .\.venv\Scripts\python.exe -m pip install -r requirements-build.txt
& .\.venv\Scripts\python.exe -B -m PyInstaller --noconfirm ColorTools.spec
```

Output: `dist/ColorTools.exe`. Both logo images and `icon.ico` must remain at
the repository root. Publish the EXE as a GitHub Release asset; build output
and the local Python environment are excluded from source control.

## Affinity compatibility

The reader supports the observed uncompressed palette layouts `0x20` and `0x21`.
Import has been verified with an Affinity-exported palette containing 153 HSL colors. Other layouts, compressed palettes and
special color types have not been comprehensively verified. Unsupported
structures are rejected instead of scanning arbitrary bytes for colors.

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Alt + P` | Activate the screen eyedropper; global registration is attempted at startup |
| Left-click | Pick the color under the cursor (while the eyedropper is active) |
| Right-click / `Esc` | Cancel the eyedropper |

---

## Import Limits

- Palette files: up to 20 MiB; the first 1,000 supported colors are retained, with a truncation message if more are found.
- Procreate archives: the selected JSON member may be at most 5 MiB after decompression.
- Images: up to 100 MiB on disk and 60 million pixels after JPEG draft reduction when supported. Decoder-level checks may also reject a file.
- Settings: up to 20 MiB. Existing oversized settings are not overwritten with defaults.

---

## Troubleshooting

Errors and crash information are written to `%APPDATA%\Color Tools\error.log` and `crash.log` when those files are writable. If `%APPDATA%` is unavailable, the app uses a `Color Tools` folder under the home directory.

If `Alt + P` cannot be registered, the app falls back to checking the key state while its update loop is running. The picker button remains available.

If the interface starts in Finnish, an earlier language choice was saved. Select English in the Info tab. New users start in English.

---

## License

MIT — see [LICENSE](LICENSE).
