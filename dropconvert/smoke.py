"""Package verification used by the build process."""

from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

from docx import Document
from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader

from .core import detect_capabilities, rearrange_pdf, run_action


def main(report_path: str) -> int:
    report: dict[str, object] = {"checks": {}}
    checks: dict[str, str] = report["checks"]  # type: ignore[assignment]
    capabilities = detect_capabilities()
    report["word_installed"] = capabilities.word
    report["tesseract_installed"] = bool(capabilities.tesseract)
    report["ocr_languages"] = capabilities.ocr_languages
    if getattr(sys, "frozen", False):
        bundled = Path(getattr(sys, "_MEIPASS")) / "tesseract" / "tesseract.exe"
        checks["portable_ocr"] = "ok" if capabilities.tesseract == str(bundled) and bundled.is_file() else "bundled engine missing"
        checks["ocr_orientation_data"] = "ok" if capabilities.tessdata_dir and (capabilities.tessdata_dir / "osd.traineddata").is_file() else "orientation data missing"
    try:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            outputs = root / "output"
            image = Image.new("RGB", (600, 110), "white")
            draw = ImageDraw.Draw(image)
            font = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", 50)
            draw.text((20, 20), "HELLO 123", font=font, fill="black")
            source_image = root / "image.png"
            image.save(source_image)

            from PySide6.QtWidgets import QApplication
            from .queue import QueueDialog
            from .ui import FilePanel

            app = QApplication.instance() or QApplication([])
            text_source = root / "note.txt"
            text_source.write_text("hello", encoding="utf-8")
            panel = FilePanel(capabilities)
            panel.set_files((source_image, text_source))
            queue = QueueDialog((source_image, text_source), outputs, capabilities, panel)
            checks["mixed_queue_ui"] = "ok" if queue.table.rowCount() == 2 and panel.queue_button.isEnabled() else "queue unavailable"
            queue.close()
            panel.close()

            pdf = run_action("image_pdf", [source_image], outputs, capabilities).files[0]
            compressed = run_action("compress_pdf", [pdf], outputs, capabilities, compression="small").files[0]
            checks["pdf_compression"] = "ok" if len(PdfReader(str(compressed)).pages) == 1 else "compressed PDF invalid"
            png = run_action("pdf_png", [pdf], outputs, capabilities).files[0]
            checks["pdf_render"] = "ok" if png.is_file() else "missing output"
            rearranged = rearrange_pdf(pdf, ((0, 90),), outputs)
            checks["page_organizer"] = "ok" if PdfReader(str(rearranged)).pages[0].get("/Rotate") == 90 else "wrong rotation"

            if capabilities.tesseract:
                ocr = run_action("image_ocr_txt", [source_image], outputs, capabilities).files[0]
                checks["ocr"] = "ok" if "HELLO" in ocr.read_text(encoding="utf-8").upper() else "text not recognized"
                searchable = run_action("pdf_ocr_searchable", [pdf], outputs, capabilities).files[0]
                checks["searchable_pdf"] = "ok" if "HELLO" in PdfReader(str(searchable)).pages[0].extract_text().upper() else "text not searchable"

            if capabilities.word:
                document = Document()
                document.add_paragraph("Xin chào Việt Nam")
                source_docx = root / "word.docx"
                document.save(source_docx)
                word_pdf = run_action("docx_pdf", [source_docx], outputs, capabilities).files[0]
                extracted = run_action("pdf_txt", [word_pdf], outputs, capabilities).files[0]
                checks["word_to_pdf"] = "ok" if "Việt Nam" in extracted.read_text(encoding="utf-8") else "text not found"
    except Exception as error:
        checks["exception"] = f"{error.__class__.__name__}: {error}"
    Path(report_path).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if checks and all(value == "ok" for value in checks.values()) else 1
