# Local File Converter Chrome Extension

This is a local Chrome extension plus a small helper app. Files stay on this computer. The extension sends selected files to `http://127.0.0.1:8765`, and the helper returns the converted result.

## What It Supports

- PPT/PPTX to PDF using LibreOffice when installed, or Microsoft PowerPoint automation
- PPT/PPTX to HTML by rendering slides into a single standalone HTML file
- PPT/PPTX to PNG or JPG by exporting to PDF first, then rendering pages
- PDF to HTML by rendering pages into a single standalone HTML file
- PDF to PNG or JPG
- PNG, JPG, WEBP, BMP, TIFF, and GIF image conversion
- Image to PDF
- Image to HTML
- TXT, MD, CSV, JSON, XML, CSS, and JS to HTML
- DOC/DOCX/XLS/XLSX to PDF using LibreOffice when installed, or Microsoft Word/Excel automation
- DOC/DOCX/XLS/XLSX to HTML by rendering exported PDF pages into a single HTML file

Chrome extensions use `manifest.json`, not XML. The required extension manifest is already included at `extension/manifest.json`.

## Start The Local Helper

Double-click:

```text
start_converter.bat
```

Leave that window open while using the extension.

If Office export says unavailable even though PowerPoint is installed, close the helper window and start `start_converter.bat` by double-clicking it from your normal Windows desktop session. Some sandboxed terminals cannot start Microsoft Office automation.

## Add The Extension To Chrome

1. Open Chrome.
2. Go to `chrome://extensions`.
3. Turn on `Developer mode`.
4. Click `Load unpacked`.
5. Select the `extension` folder inside this project.
6. Pin `Local File Converter` from the Chrome extensions menu.

## Use It

1. Start `start_converter.bat`.
2. Click the extension icon in Chrome.
3. Choose or drop files.
4. Pick the output format.
5. Click `Convert`.

If several output files are created, for example PDF pages to PNG images, the extension downloads a ZIP file.
