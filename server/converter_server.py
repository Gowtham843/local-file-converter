from __future__ import annotations

import base64
import csv
import html
import io
import json
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from PIL import Image, ImageSequence, UnidentifiedImageError


HOST = "127.0.0.1"
PORT = int(os.environ.get("LOCAL_CONVERTER_PORT", "8765"))
SERVER_DIR = Path(__file__).resolve().parent
OFFICE_TO_PDF_SCRIPT = SERVER_DIR / "office_to_pdf.ps1"

IMAGE_EXTENSIONS = {
    ".bmp",
    ".gif",
    ".jpeg",
    ".jpg",
    ".png",
    ".tif",
    ".tiff",
    ".webp",
}
OFFICE_EXTENSIONS = {
    ".doc",
    ".docx",
    ".ppt",
    ".pptx",
    ".xls",
    ".xlsx",
}
IMAGE_TARGETS = {"bmp", "jpg", "jpeg", "png", "tiff", "webp"}
TEXT_EXTENSIONS = {
    ".css",
    ".csv",
    ".htm",
    ".html",
    ".js",
    ".json",
    ".log",
    ".md",
    ".txt",
    ".xml",
}
POWERPOINT_COM_OK: bool | None = None


class ConversionError(Exception):
    pass


def json_bytes(payload: dict) -> bytes:
    return json.dumps(payload, indent=2).encode("utf-8")


def safe_stem(filename: str) -> str:
    stem = Path(filename).stem or "converted"
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._")
    return stem or "converted"


def unique_output_path(output_dir: Path, stem: str, suffix: str) -> Path:
    suffix = suffix if suffix.startswith(".") else f".{suffix}"
    output = output_dir / f"{stem}{suffix}"
    counter = 2
    while output.exists():
        output = output_dir / f"{stem}_{counter}{suffix}"
        counter += 1
    return output


def content_type_for(path: Path) -> str:
    if path.suffix.lower() == ".zip":
        return "application/zip"
    return mimetypes.guess_type(path.name)[0] or "application/octet-stream"


def find_binary(name: str) -> str | None:
    explicit = os.environ.get(f"{name.upper()}_EXE")
    if explicit and Path(explicit).exists():
        return explicit

    bundled = (
        Path.home()
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
    )
    bundled_bins = [
        bundled / "native" / "poppler" / "Library" / "bin",
        bundled / "bin",
    ]
    for bundled_bin in bundled_bins:
        for suffix in (".exe", ".cmd", ".bat", ""):
            candidate = bundled_bin / f"{name}{suffix}"
            if candidate.exists():
                return str(candidate)

    found = shutil.which(name)
    if found:
        return found

    return None


def has_powerpoint() -> bool:
    candidates = [
        Path(os.environ.get("POWERPNT_EXE", "")),
        Path(r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"),
        Path(r"C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE"),
        Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"),
    ]
    return any(path.exists() for path in candidates if str(path))


def find_soffice() -> str | None:
    explicit = os.environ.get("SOFFICE_EXE")
    if explicit and Path(explicit).exists():
        return explicit

    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found

    candidates = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)

    return None


def can_start_powerpoint_com() -> bool:
    global POWERPOINT_COM_OK
    if POWERPOINT_COM_OK is not None:
        return POWERPOINT_COM_OK
    if not has_powerpoint():
        POWERPOINT_COM_OK = False
        return POWERPOINT_COM_OK

    command = (
        "$ErrorActionPreference='Stop'; "
        "$app=New-Object -ComObject PowerPoint.Application; "
        "$app.Quit(); "
        "[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)"
    )
    try:
        completed = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
            capture_output=True,
            text=True,
            timeout=12,
        )
        POWERPOINT_COM_OK = completed.returncode == 0
    except Exception:
        POWERPOINT_COM_OK = False

    return POWERPOINT_COM_OK


def office_pdf_available() -> bool:
    return bool(find_soffice() or can_start_powerpoint_com())


def normalize_target(target: str) -> str:
    normalized = target.lower().strip().lstrip(".")
    if normalized == "jpeg":
        return "jpg"
    return normalized


def save_uploaded_file(payload: bytes, filename: str, workdir: Path) -> Path:
    clean_name = Path(filename or "upload.bin").name
    stem = safe_stem(clean_name)
    suffix = Path(clean_name).suffix.lower()
    output = workdir / f"{stem}{suffix}"
    counter = 2
    while output.exists():
        output = workdir / f"{stem}_{counter}{suffix}"
        counter += 1
    output.write_bytes(payload)
    return output


def white_background(image: Image.Image) -> Image.Image:
    if image.mode in ("RGBA", "LA") or (
        image.mode == "P" and "transparency" in image.info
    ):
        rgba = image.convert("RGBA")
        background = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        background.alpha_composite(rgba)
        return background.convert("RGB")
    if image.mode != "RGB":
        return image.convert("RGB")
    return image


def save_image(image: Image.Image, output_path: Path, target: str) -> None:
    if target == "jpg":
        image = white_background(image)
        image.save(output_path, "JPEG", quality=92, optimize=True)
    elif target == "png":
        image.save(output_path, "PNG")
    elif target == "webp":
        image.save(output_path, "WEBP", quality=92)
    elif target == "bmp":
        image.convert("RGB").save(output_path, "BMP")
    elif target == "tiff":
        image.save(output_path, "TIFF")
    else:
        raise ConversionError(f"Unsupported image target: {target}")


def image_to_pdf(input_path: Path, output_path: Path) -> None:
    try:
        with Image.open(input_path) as image:
            frames = []
            for frame in ImageSequence.Iterator(image):
                frames.append(white_background(frame.copy()))

        if not frames:
            raise ConversionError("No image frames found")

        first, *rest = frames
        first.save(output_path, "PDF", save_all=bool(rest), append_images=rest)
    except UnidentifiedImageError as exc:
        raise ConversionError(f"Cannot read image: {input_path.name}") from exc


def image_to_image(input_path: Path, target: str, output_dir: Path) -> list[Path]:
    try:
        with Image.open(input_path) as image:
            frames = [frame.copy() for frame in ImageSequence.Iterator(image)]
    except UnidentifiedImageError as exc:
        raise ConversionError(f"Cannot read image: {input_path.name}") from exc

    if not frames:
        raise ConversionError("No image frames found")

    stem = safe_stem(input_path.name)
    extension = "jpg" if target == "jpg" else target
    outputs: list[Path] = []

    for index, frame in enumerate(frames, start=1):
        output_stem = f"{stem}_frame_{index:03d}" if len(frames) > 1 else stem
        output_path = unique_output_path(output_dir, output_stem, extension)
        save_image(frame, output_path, target)
        outputs.append(output_path)

    return outputs


def data_uri_for(path: Path) -> str:
    media_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{media_type};base64,{encoded}"


def html_document(title: str, body: str) -> str:
    escaped_title = html.escape(title)
    return f"""<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{escaped_title}</title>
    <style>
      :root {{
        font-family: Arial, Helvetica, sans-serif;
        color: #1f2937;
        background: #eef1f5;
      }}
      body {{
        margin: 0;
        padding: 24px;
      }}
      main {{
        max-width: 1100px;
        margin: 0 auto;
      }}
      h1 {{
        margin: 0 0 18px;
        font-size: 24px;
        letter-spacing: 0;
      }}
      .page,
      .content {{
        margin: 0 0 18px;
        background: #ffffff;
        border: 1px solid #d9dee8;
        border-radius: 8px;
        box-shadow: 0 1px 4px rgba(15, 23, 42, 0.08);
      }}
      .page img {{
        display: block;
        width: 100%;
        height: auto;
        border-radius: 8px;
      }}
      .caption {{
        padding: 8px 12px;
        color: #64748b;
        font-size: 13px;
      }}
      pre {{
        margin: 0;
        padding: 18px;
        overflow: auto;
        white-space: pre-wrap;
        overflow-wrap: anywhere;
        font-family: Consolas, "Courier New", monospace;
        font-size: 14px;
        line-height: 1.5;
      }}
      table {{
        width: 100%;
        border-collapse: collapse;
        background: #ffffff;
      }}
      th,
      td {{
        border: 1px solid #d9dee8;
        padding: 8px 10px;
        text-align: left;
        vertical-align: top;
      }}
      th {{
        background: #f8fafc;
      }}
    </style>
  </head>
  <body>
    <main>
      <h1>{escaped_title}</h1>
{body}
    </main>
  </body>
</html>
"""


def images_to_html(image_paths: list[Path], title: str, output_path: Path) -> None:
    parts = []
    for index, image_path in enumerate(image_paths, start=1):
        caption = html.escape(f"Page {index}" if len(image_paths) > 1 else Path(title).name)
        alt = html.escape(f"{title} page {index}")
        parts.append(
            f"""      <section class="page">
        <img src="{data_uri_for(image_path)}" alt="{alt}">
        <div class="caption">{caption}</div>
      </section>"""
        )
    output_path.write_text(html_document(title, "\n".join(parts)), encoding="utf-8")


def text_to_html(input_path: Path, title: str, output_path: Path) -> None:
    source_ext = input_path.suffix.lower()
    text = input_path.read_text(encoding="utf-8-sig", errors="replace")

    if source_ext == ".json":
        try:
            text = json.dumps(json.loads(text), indent=2, ensure_ascii=False)
        except json.JSONDecodeError:
            pass

    if source_ext == ".csv":
        rows = list(csv.reader(io.StringIO(text)))
        if rows:
            table_rows = []
            for row_index, row in enumerate(rows):
                tag = "th" if row_index == 0 else "td"
                cells = "".join(f"<{tag}>{html.escape(cell)}</{tag}>" for cell in row)
                table_rows.append(f"        <tr>{cells}</tr>")
            body = (
                '      <section class="content">\n'
                "        <table>\n"
                + "\n".join(table_rows)
                + "\n        </table>\n"
                "      </section>"
            )
            output_path.write_text(html_document(title, body), encoding="utf-8")
            return

    body = f"""      <section class="content">
        <pre>{html.escape(text)}</pre>
      </section>"""
    output_path.write_text(html_document(title, body), encoding="utf-8")


def pdf_to_images(input_path: Path, target: str, output_dir: Path) -> list[Path]:
    pdftoppm = find_binary("pdftoppm")
    if not pdftoppm:
        raise ConversionError("PDF rendering tool was not found. Start with start_converter.bat so the bundled PDF tool is on PATH.")

    stem = safe_stem(input_path.name)
    prefix_name = f"{stem}_page"
    command = [pdftoppm, "-r", "180"]
    expected_extension = "jpg" if target == "jpg" else target

    prefix = output_dir / prefix_name
    counter = 2
    while list(output_dir.glob(f"{prefix.name}-*.{expected_extension}")):
        prefix = output_dir / f"{prefix_name}_{counter}"
        counter += 1

    if target == "jpg":
        command.extend(["-jpeg", "-jpegopt", "quality=92"])
    elif target == "png":
        command.append("-png")
    elif target == "tiff":
        command.append("-tiff")
    else:
        raise ConversionError("PDF pages can be rendered to PNG, JPG, or TIFF")

    command.extend([str(input_path), str(prefix)])

    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=180,
        )
    except Exception as exc:
        raise ConversionError(f"Could not render PDF: {input_path.name}") from exc

    if completed.returncode != 0:
        detail = completed.stderr.strip() or completed.stdout.strip() or "PDF render failed"
        raise ConversionError(detail)

    outputs = sorted(output_dir.glob(f"{prefix.name}-*.{expected_extension}"))
    if not outputs:
        raise ConversionError("The PDF did not contain renderable pages")

    return outputs


def office_to_pdf(input_path: Path, output_path: Path) -> None:
    if not OFFICE_TO_PDF_SCRIPT.exists():
        raise ConversionError("Office conversion script is missing")

    soffice = find_soffice()
    if soffice:
        with tempfile.TemporaryDirectory(dir=output_path.parent, prefix="office_export_") as office_temp:
            office_dir = Path(office_temp)
            completed = subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(office_dir),
                    str(input_path),
                ],
                capture_output=True,
                text=True,
                timeout=180,
            )
            libreoffice_output = office_dir / f"{input_path.stem}.pdf"
            if completed.returncode == 0 and libreoffice_output.exists():
                libreoffice_output.replace(output_path)
                return

    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(OFFICE_TO_PDF_SCRIPT),
        "-InputPath",
        str(input_path),
        "-OutputPath",
        str(output_path),
    ]
    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        timeout=180,
    )
    if completed.returncode != 0 or not output_path.exists():
        detail = completed.stderr.strip() or completed.stdout.strip() or "Office export failed"
        if "80070520" in detail:
            detail = (
                "PowerPoint is installed, but Windows blocked Office automation in this session. "
                "Start the helper by double-clicking start_converter.bat from your normal desktop session, "
                "or install LibreOffice and try again."
            )
        raise ConversionError(detail)


def convert_one(input_path: Path, original_name: str, target: str, output_dir: Path) -> list[Path]:
    source_ext = input_path.suffix.lower()
    target = normalize_target(target)
    stem = safe_stem(original_name)

    if target == "pdf":
        output_path = unique_output_path(output_dir, stem, "pdf")
        if source_ext == ".pdf":
            shutil.copy2(input_path, output_path)
        elif source_ext in IMAGE_EXTENSIONS:
            image_to_pdf(input_path, output_path)
        elif source_ext in OFFICE_EXTENSIONS:
            office_to_pdf(input_path, output_path)
        else:
            raise ConversionError(f"{source_ext or 'This file'} cannot be converted to PDF yet")
        return [output_path]

    if target == "html":
        output_path = unique_output_path(output_dir, stem, "html")
        if source_ext in {".html", ".htm"}:
            shutil.copy2(input_path, output_path)
        elif source_ext == ".pdf":
            page_images = pdf_to_images(input_path, "png", output_dir)
            images_to_html(page_images, original_name, output_path)
        elif source_ext in IMAGE_EXTENSIONS:
            preview_images = image_to_image(input_path, "png", output_dir)
            images_to_html(preview_images, original_name, output_path)
        elif source_ext in OFFICE_EXTENSIONS:
            intermediate_pdf = unique_output_path(output_dir, f"{stem}_document", "pdf")
            office_to_pdf(input_path, intermediate_pdf)
            page_images = pdf_to_images(intermediate_pdf, "png", output_dir)
            images_to_html(page_images, original_name, output_path)
        elif source_ext in TEXT_EXTENSIONS:
            text_to_html(input_path, original_name, output_path)
        else:
            raise ConversionError(f"{source_ext or 'This file'} cannot be converted to HTML yet")
        return [output_path]

    if target in IMAGE_TARGETS:
        if source_ext == ".pdf":
            return pdf_to_images(input_path, target, output_dir)
        if source_ext in IMAGE_EXTENSIONS:
            return image_to_image(input_path, target, output_dir)
        if source_ext in {".ppt", ".pptx"}:
            intermediate_pdf = unique_output_path(output_dir, f"{stem}_slides", "pdf")
            office_to_pdf(input_path, intermediate_pdf)
            return pdf_to_images(intermediate_pdf, target, output_dir)
        raise ConversionError(f"{source_ext or 'This file'} cannot be converted to {target.upper()} yet")

    raise ConversionError(f"Unsupported target format: {target}")


def make_zip(paths: list[Path], output_path: Path) -> Path:
    used_names: set[str] = set()
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in paths:
            arcname = path.name
            if arcname in used_names:
                arcname = f"{path.stem}_{len(used_names) + 1}{path.suffix}"
            used_names.add(arcname)
            archive.write(path, arcname)
    return output_path


def parse_multipart(headers, body: bytes) -> tuple[dict[str, str], list[tuple[str, bytes]]]:
    content_type = headers.get("Content-Type", "")
    message = BytesParser(policy=default).parsebytes(
        b"Content-Type: "
        + content_type.encode("utf-8")
        + b"\r\nMIME-Version: 1.0\r\n\r\n"
        + body
    )

    fields: dict[str, str] = {}
    files: list[tuple[str, bytes]] = []

    if not message.is_multipart():
        raise ConversionError("Expected multipart form data")

    for part in message.iter_parts():
        if part.get_content_disposition() != "form-data":
            continue
        name = part.get_param("name", header="content-disposition")
        filename = part.get_filename()
        payload = part.get_payload(decode=True) or b""

        if filename:
            files.append((filename, payload))
        elif name:
            fields[name] = payload.decode(part.get_content_charset() or "utf-8", errors="replace")

    return fields, files


class ConverterHandler(BaseHTTPRequestHandler):
    server_version = "LocalFileConverter/1.0"

    def end_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Expose-Headers", "Content-Disposition")
        super().end_headers()

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.end_headers()

    def do_GET(self) -> None:
        if urlparse(self.path).path != "/health":
            self.send_error(404)
            return

        payload = json_bytes(
            {
                "ok": True,
                "server": "Local File Converter",
                "powerpoint": can_start_powerpoint_com(),
                "officePdf": office_pdf_available(),
                "libreoffice": bool(find_soffice()),
                "poppler": bool(find_binary("pdftoppm") and find_binary("pdfinfo")),
                "supports": [
                    "pptx to pdf",
                    "pptx to html",
                    "pptx to png/jpg",
                    "pdf to html",
                    "pdf to png/jpg",
                    "png/jpg/webp/bmp/tiff/gif image conversion",
                    "image to pdf",
                    "image/text/csv/json/xml to html",
                ],
            }
        )
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_POST(self) -> None:
        if urlparse(self.path).path != "/convert":
            self.send_error(404)
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            if content_length <= 0:
                raise ConversionError("No files were sent")
            body = self.rfile.read(content_length)
            fields, uploaded_files = parse_multipart(self.headers, body)
            target = normalize_target(fields.get("target", ""))

            if not target:
                raise ConversionError("Choose a target format")
            if not uploaded_files:
                raise ConversionError("Choose at least one file")

            with tempfile.TemporaryDirectory(prefix="local_converter_") as temp_name:
                workdir = Path(temp_name)
                input_dir = workdir / "input"
                output_dir = workdir / "output"
                input_dir.mkdir()
                output_dir.mkdir()

                converted: list[Path] = []
                for filename, payload in uploaded_files:
                    input_path = save_uploaded_file(payload, filename, input_dir)
                    converted.extend(convert_one(input_path, filename, target, output_dir))

                if len(converted) == 1:
                    response_path = converted[0]
                    download_name = response_path.name
                    response_body = response_path.read_bytes()
                else:
                    response_path = make_zip(converted, output_dir / "converted_files.zip")
                    download_name = response_path.name
                    response_body = response_path.read_bytes()

            self.send_response(200)
            self.send_header("Content-Type", content_type_for(Path(download_name)))
            self.send_header("Content-Disposition", f'attachment; filename="{download_name}"')
            self.send_header("Content-Length", str(len(response_body)))
            self.end_headers()
            self.wfile.write(response_body)
        except ConversionError as exc:
            self.send_json_error(str(exc), 400)
        except subprocess.TimeoutExpired:
            self.send_json_error("Conversion timed out", 500)
        except Exception as exc:
            self.send_json_error(f"Unexpected error: {exc}", 500)

    def log_message(self, format: str, *args) -> None:
        sys.stderr.write("%s - %s\n" % (self.address_string(), format % args))

    def send_json_error(self, message: str, status_code: int) -> None:
        payload = json_bytes({"ok": False, "error": message})
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)


def main() -> None:
    server = ThreadingHTTPServer((HOST, PORT), ConverterHandler)
    print(f"Local File Converter is running at http://{HOST}:{PORT}")
    print("Leave this window open while using the Chrome extension.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
