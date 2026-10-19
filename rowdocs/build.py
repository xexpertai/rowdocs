"""Generation (``rowdocs build``): one .docx (and optionally .pdf) per selected row."""

from __future__ import annotations

import csv
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from .data import Dataset
from .naming import DEFAULT_PATTERN, plan_names
from .pdf import INSTALL_HELP, convert_to_pdf, find_soffice, merge_pdfs
from .template import WordTemplate
from .validate import Report, validate

MANIFEST_NAME = "manifest.csv"
MERGED_NAME = "ALL.pdf"
MANIFEST_FIELDS = ("row", "output_docx", "output_pdf", "status", "message")

EXIT_OK = 0
EXIT_ROW_ERRORS = 1
EXIT_VALIDATION = 2
EXIT_NO_SOFFICE = 3


class ValidationFailed(Exception):
    """Pre-flight validation found errors and ``force`` was not set."""

    def __init__(self, report: Report):
        super().__init__(f"validation failed with {len(report.errors)} error(s)")
        self.report = report


@dataclass
class RowResult:
    row: int
    output_docx: str = ""
    output_pdf: str = ""
    status: str = "ok"  # ok / skipped / error
    message: str = ""


@dataclass
class BuildResult:
    outdir: Path
    report: Report
    rows: list[RowResult] = field(default_factory=list)
    manifest: Path | None = None
    merged_pdf: Path | None = None
    zip_path: Path | None = None
    pdf_unavailable: bool = False

    @property
    def generated(self) -> list[RowResult]:
        """Rows that produced a .docx (even if a later PDF step failed)."""
        return [r for r in self.rows if r.output_docx]

    @property
    def failed(self) -> list[RowResult]:
        return [r for r in self.rows if r.status == "error"]

    @property
    def exit_code(self) -> int:
        if self.pdf_unavailable:
            return EXIT_NO_SOFFICE
        return EXIT_ROW_ERRORS if self.failed else EXIT_OK

    def files(self) -> list[Path]:
        """Every file this build wrote, in row order (manifest and ALL.pdf last)."""
        paths: list[Path] = []
        for row in self.rows:
            paths.extend(self.outdir / name for name in (row.output_docx, row.output_pdf) if name)
        if self.merged_pdf:
            paths.append(self.merged_pdf)
        if self.manifest:
            paths.append(self.manifest)
        return paths


def write_manifest(rows: list[RowResult], path: Path) -> None:
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=MANIFEST_FIELDS)
        writer.writeheader()
        for row in sorted(rows, key=lambda r: r.row):
            writer.writerow({f: getattr(row, f) for f in MANIFEST_FIELDS})


def build(
    dataset: Dataset,
    template: WordTemplate,
    outdir: Path | str,
    name_pattern: str | None = None,
    pdf: bool = False,
    merge_pdf: bool = False,
    make_zip: bool = False,
    force: bool = False,
    soffice: str | None = None,
) -> BuildResult:
    """Validate, then render every selected row. One failing row never stops the others."""
    pattern = name_pattern or DEFAULT_PATTERN
    report = validate(dataset, template, pattern)
    if report.errors and not force:
        raise ValidationFailed(report)

    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    real_outdir = outdir.resolve()
    pdf = pdf or merge_pdf
    soffice = (soffice or find_soffice()) if pdf else None
    result = BuildResult(outdir=outdir, report=report, pdf_unavailable=pdf and not soffice)

    selected = set(dataset.selected)
    results: dict[int, RowResult] = {
        n: RowResult(row=n, status="skipped", message="not selected by --rows/--where")
        for n in range(1, len(dataset.rows) + 1)
        if n not in selected
    }
    planned = plan_names(pattern, [(n, dataset.context(n)) for n in dataset.selected])
    docx_paths: dict[int, Path] = {}
    for item in planned:
        row = RowResult(row=item.row, message=item.error)
        results[item.row] = row
        target = outdir / f"{item.final}.docx"
        if target.resolve().parent != real_outdir:  # defence in depth; sanitize() already prevents this
            row.status, row.message = "error", f"refusing to write outside the output folder: {item.final!r}"
            continue
        try:
            target.write_bytes(template.render(dataset.context(item.row)))
        except Exception as exc:  # any template/filter failure is reported per row
            row.status, row.message = "error", f"{exc.__class__.__name__}: {exc}"
            continue
        row.output_docx = target.name
        docx_paths[item.row] = target

    if pdf and soffice and docx_paths:
        converted = convert_to_pdf(list(docx_paths.values()), outdir, soffice)
        failure = "PDF conversion failed"
        if not any(converted.values()):
            failure += " (LibreOffice could not open any document: is LibreOffice Writer installed?)"
        for number, docx in docx_paths.items():
            pdf_path = converted.get(docx)
            if pdf_path:
                results[number].output_pdf = pdf_path.name
            else:
                results[number].status, results[number].message = "error", failure
        if merge_pdf:
            parts = [outdir / results[n].output_pdf for n in sorted(docx_paths) if results[n].output_pdf]
            if parts:
                result.merged_pdf = outdir / MERGED_NAME
                merge_pdfs(parts, result.merged_pdf)
    elif result.pdf_unavailable:
        for number in docx_paths:
            results[number].message = "PDF not created: LibreOffice not found"

    result.rows = sorted(results.values(), key=lambda r: r.row)
    result.manifest = outdir / MANIFEST_NAME
    write_manifest(result.rows, result.manifest)

    if make_zip and not result.pdf_unavailable:
        result.zip_path = outdir.parent / f"{outdir.name}.zip"
        with zipfile.ZipFile(result.zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
            for path in result.files():
                archive.write(path, arcname=path.name)
    return result


__all__ = ["build", "BuildResult", "RowResult", "ValidationFailed", "INSTALL_HELP", "MANIFEST_NAME", "MERGED_NAME"]
