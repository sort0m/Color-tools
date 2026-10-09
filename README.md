# Color Tools

A compact color picker, color theory guide and palette manager for Windows. Pick colors from the screen, build palettes, explore harmonies, check contrast and convert palettes between formats — for example, from Procreate `.swatches` to Affinity `.afpalette`. The interface starts in **English**, with Finnish available in the Info tab.

## Download and run

The Windows executable will be available in this repository's **Releases** section once published. Close any older running version, then launch **ColorTools.exe**. No separate Python installation is needed.

Built for Windows x64. The application does not run on Linux or macOS.

## Screenshots

![Harmony tab](screenshots/ct1.jpg)

![History tab](screenshots/ct2.jpg)

![Palettes tab](screenshots/ct3.jpg)

## Features

- **Color picking and editing:** Screen eyedropper with an **Alt + P** shortcut, HEX input, quick copying, and RGB, HSL, CMYK, LAB and grayscale controls.
- **Color theory:** Explore a color's associations, historical context, mood, warm/cool character and suggested uses in interfaces. Includes a color-temperature meter and the nearest name from the built-in CSS color table.
- **Color harmonies:** Nine modes, including complementary, analogous, triadic, tints, shades and tones. Save a generated harmony as a palette.
- **Palette management:** Create palettes from selected history colors, harmonies or imported files. Edit colors, reorder them by dragging, move them between pages and undo edits. The editor displays 96 colors per page; history keeps up to 60 entries.
- **Palette conversion:** Import and export **ASE, ACO, GPL, Procreate `.swatches` and Affinity `.afpalette`**. Preview imported colors before adding them.
- **Palettes from images:** Extract colors from PNG, JPEG, BMP, GIF, WebP and TIFF images. Choose a target of 2–20 colors, preview the result and save it as a palette.
- **WCAG contrast:** Check contrast ratios against white, black or a custom background, with AA/AAA ratings and a text preview.
- **Color-blindness simulation:** Preview text and background colors with protanopia, deuteranopia and tritanopia simulations.
- **HTML reports:** Export palettes, harmonies or history as a color reference with HEX, RGB, HSL, HSV and CMYK values, CSS names, contrast information and color theory text.
- **Personalization:** Seven themes, collapsible sections, an optional always-on-top mode, and automatic saving of palettes, history and preferences.

## Palette conversion examples

Import a supported palette, confirm the preview, open it in the palette editor and choose **Export** with the desired format.

| Starting point | What you can do |
| --- | --- |
| Procreate `.swatches` palette | Export as Affinity `.afpalette`. |
| Affinity `.afpalette` palette | Export as Procreate `.swatches` or Adobe `.ase`. |
| Adobe `.aco` palette | Export as GIMP `.gpl`. |
| A photo or illustration | Extract a palette, then export it in any supported palette format. |

Palette imports retain up to **1,000 colors**. Not all format variants are supported. Imported colors become opaque RGB; original color names, profiles and other metadata are not preserved.

## Settings

Settings and logs are stored in `%APPDATA%\Color Tools`. An existing language choice is remembered; change it in the Info tab. The current color and undo history are not restored after restarting.

<details>
<summary>Run from source / Build</summary>

With Python installed (build environment: Python 3.14), run these commands from the project folder in PowerShell:

```powershell
python -m venv .venv
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
& .\.venv\Scripts\python.exe -B run_app.py
```

Dependencies are pinned to Dear PyGui 2.2 and Pillow 12.3.0.

To build the EXE:

```powershell
& .\.venv\Scripts\python.exe -m pip install -r requirements-build.txt
& .\.venv\Scripts\python.exe -B -m PyInstaller --noconfirm ColorTools.spec
```

The result is `dist/ColorTools.exe`.

</details>

## License

[MIT](LICENSE) — Copyright (c) 2026 sort0m.
