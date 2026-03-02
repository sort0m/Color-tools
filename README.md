# Color Tools

A compact, always-on-top color picker for Windows built with [DearPyGui](https://github.com/hoffstadt/DearPyGui). Pick colors from anywhere on screen, explore color harmonies, manage palettes, and export in multiple formats — all from a single window.

Built with multiple AI assistants

![Platform](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

---

## Features

### Color Picker
- **Color wheel** or **slider mode** — switch between a HSV wheel and precision sliders
- **Slider modes:** RGB, HSL, CMYK, LAB, Grayscale — all stay in sync with each other
- **Screen pipette** — click anywhere on screen to sample a color
- **Hex input** — type or paste a hex code directly (`#RRGGBB` or shorthand `#RGB`)
- **Copy to clipboard** — one click copies the current hex value

### Harmony Tab
- **9 harmony modes:** Complementary, Split Complementary, Analogous, Triadic, Tetradic, Rectangle, Tints, Shades, Tones
- **7 display formats per color:** HEX, RGB, HSL, HSV, CSS Name, CMYK, Contrast
- **WCAG contrast checker** — shows contrast ratio against white and black with AAA / AA / AA* / Fail rating
- Export the full harmony as an **HTML report** or save it as a **palette**

### History Tab
- Automatically saves the last 60 picked colors
- **Right-click** a swatch to remove it
- Select multiple colors and save them as a new palette with **New Palette**
- Export the full history as an **HTML report**

### Palettes Tab
- Create and manage unlimited named palettes
- **Import from image** — extracts a color palette from any image file using quantization (requires Pillow)
- **Import ASE** — loads Adobe Swatch Exchange files
- **Palette editor** — reorder colors by drag-and-drop, add the current color, undo changes
- Export any palette as an **HTML report** or **ASE file**

### Other
- **6 themes:** Dark, Light, Midnight, Mocha, Nord, Solarized
- **Always on top** toggle
- Config (history, palettes, window position, theme) saved automatically to `%APPDATA%\Color Tools\config.json`

---

## Requirements

- Windows 10 or 11
- Python 3.9 or newer

---

## Installation

```bash
pip install dearpygui mss pynput pyperclip Pillow
```

> Pillow is optional. Without it, the **Import from Image** feature will show an error message but everything else will work normally.

---

## Running

```bash
python color_tools.py
```

---

## Project Structure

```
color_tools.py      # Main application (single file)
colortoolsd.png     # Logo — dark theme variant
colortoolsw.png     # Logo — light theme variant
requirements.txt    # Python dependencies
README.md
```

---

## Data & Privacy

All data (color history, palettes, window position, theme) is stored locally on your machine at:

```
%APPDATA%\Color Tools\config.json
```

Nothing is sent over the network.

---

## Platform Support

Color Tools is **Windows only**. It uses Win32 API calls (`ctypes.windll`) for:
- DPI awareness
- Custom borderless title bar
- Dark mode title bar (DWM)
- Window minimize / move

These calls are wrapped in `try/except` so the app won't crash on other platforms, but the UI will not render or behave correctly outside of Windows.

---

## License

MIT License — see [LICENSE](LICENSE) for details.
