"""Optional PDF support: LibreOffice headless conversion and pypdf merging."""

from __future__ import annotations

import contextlib
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from pypdf import PdfReader, PdfWriter

INSTALL_HELP = (
    "PDF output needs LibreOffice, which was not found.\n"
    "  Ubuntu/Debian: sudo apt install libreoffice-writer\n"
    "  Fedora:        sudo dnf install libreoffice-writer\n"
    "  macOS:         brew install --cask libreoffice\n"
    "  Windows:       install from https://www.libreoffice.org/download/\n"
    "Or point ROWDOCS_SOFFICE at the soffice executable. The .docx files were still written."
)

_KNOWN_PATHS = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)
BATCH_SIZE = 200  # files per soffice call; keeps the command line short on every OS


class PdfUnavailable(RuntimeError):
    """LibreOffice is not installed (or ROWDOCS_SOFFICE points nowhere)."""


def find_soffice() -> str | None:
    """Locate the LibreOffice binary. ``ROWDOCS_SOFFICE`` overrides (empty string = disabled)."""
    override = os.environ.get("ROWDOCS_SOFFICE")
    if override is not None:
        return override if override and Path(override).is_file() and os.access(override, os.X_OK) else None
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    return next((p for p in _KNOWN_PATHS if Path(p).is_file()), None)


def convert_to_pdf(docx_files: list[Path], outdir: Path, soffice: str | None = None) -> dict[Path, Path | None]:
    """Convert .docx files to PDF in ``outdir`` with as few soffice calls as possible.

    Returns a map from each input file to its PDF, or None if that file failed to convert.
    """
    soffice = soffice or find_soffice()
    if not soffice:
        raise PdfUnavailable(INSTALL_HELP)
    results: dict[Path, Path | None] = {}
    with tempfile.TemporaryDirectory(prefix="rowdocs-lo-") as profile:
        # A private profile avoids clashing with a LibreOffice window the user already has open.
        for start in range(0, len(docx_files), BATCH_SIZE):
            batch = docx_files[start : start + BATCH_SIZE]
            cmd = [
                soffice,
                f"-env:UserInstallation={Path(profile).as_uri()}",
                "--headless",
                "--norestore",
                "--convert-to",
                "pdf",
                "--outdir",
                str(outdir),
                *[str(p) for p in batch],
            ]
            # Files without a PDF afterwards are reported as failed, so errors here need no handling.
            with contextlib.suppress(OSError, subprocess.TimeoutExpired):
                subprocess.run(cmd, capture_output=True, timeout=120 + 15 * len(batch), check=False)
            for docx in batch:
                pdf = outdir / f"{docx.stem}.pdf"
                results[docx] = pdf if pdf.is_file() and pdf.stat().st_size > 0 else None
    return results


def merge_pdfs(parts: list[Path], target: Path) -> int:
    """Concatenate PDFs in order into ``target``. Returns the total page count."""
    writer = PdfWriter()
    for part in parts:
        writer.append(str(part))
    with open(target, "wb") as handle:
        writer.write(handle)
    return len(PdfReader(str(target)).pages)
