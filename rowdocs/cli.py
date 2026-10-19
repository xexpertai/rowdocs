"""Command-line interface: ``rowdocs check | build | serve``."""

from __future__ import annotations

import argparse
import json
import sys

from . import __version__
from .build import EXIT_VALIDATION, INSTALL_HELP, ValidationFailed, build
from .data import DataError, load_data, select_rows
from .naming import DEFAULT_PATTERN
from .template import TemplateFileError, WordTemplate
from .validate import format_report, validate


def _add_input_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("data", help="data file (.xlsx or .csv)")
    parser.add_argument("template", help="Word template (.docx) with {{ placeholders }}")
    parser.add_argument("--sheet", help="worksheet name for .xlsx files (default: first sheet)")
    parser.add_argument("--name", default=DEFAULT_PATTERN, help=f'file-name pattern (default: "{DEFAULT_PATTERN}")')
    parser.add_argument("--rows", help="only these data rows, e.g. 2-10,15 (row 1 = first row under the header)")
    parser.add_argument(
        "--where",
        action="append",
        default=[],
        metavar="COL=VALUE",
        help="only rows where a column equals (=) or differs from (!=) a value; repeatable, all must match",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="rowdocs",
        description="Generate one Word/PDF document per spreadsheet row, after checking the data.",
    )
    parser.add_argument("--version", action="version", version=f"rowdocs {__version__}")
    sub = parser.add_subparsers(dest="command", required=True, metavar="{check,build,serve}")

    check = sub.add_parser("check", help="validate data against a template without generating anything")
    _add_input_args(check)
    check.add_argument("--json", action="store_true", help="print a machine-readable JSON report")
    check.add_argument("--strict", action="store_true", help="treat warnings as errors (exit code 2)")

    bld = sub.add_parser("build", help="generate one document per row")
    _add_input_args(bld)
    bld.add_argument("-o", "--outdir", required=True, help="output folder")
    bld.add_argument("--pdf", action="store_true", help="also convert each document to PDF (needs LibreOffice)")
    bld.add_argument("--merge-pdf", action="store_true", help="also write ALL.pdf with every PDF (implies --pdf)")
    bld.add_argument("--zip", action="store_true", help="also write OUTDIR.zip with all outputs and the manifest")
    bld.add_argument("--force", action="store_true", help="generate even if validation finds errors")

    serve = sub.add_parser("serve", help="open the local web page (127.0.0.1 only)")
    serve.add_argument("--port", type=int, default=8765, help="port (default: 8765)")
    return parser


def _load(args: argparse.Namespace):
    dataset = select_rows(load_data(args.data, args.sheet), args.rows, args.where)
    return dataset, WordTemplate(args.template)


def cmd_check(args: argparse.Namespace) -> int:
    dataset, template = _load(args)
    report = validate(dataset, template, args.name)
    if args.json:
        print(json.dumps(report.to_dict(), indent=2, ensure_ascii=False))
    else:
        print(format_report(report))
    return report.exit_code(args.strict)


def cmd_build(args: argparse.Namespace) -> int:
    dataset, template = _load(args)
    try:
        result = build(
            dataset,
            template,
            args.outdir,
            name_pattern=args.name,
            pdf=args.pdf,
            merge_pdf=args.merge_pdf,
            make_zip=args.zip,
            force=args.force,
        )
    except ValidationFailed as failed:
        print(format_report(failed.report), file=sys.stderr)
        print("\nNothing was generated. Fix the errors above, or pass --force.", file=sys.stderr)
        return EXIT_VALIDATION

    report = result.report
    if report.issues:
        print(
            f"check: {len(report.errors)} error(s), {len(report.warnings)} warning(s) "
            "(run `rowdocs check` for details)",
            file=sys.stderr,
        )
    for row in result.failed:
        print(f"row {row.row}: ERROR {row.message}", file=sys.stderr)
    pdfs = sum(1 for r in result.rows if r.output_pdf)
    summary = f"Generated {len(result.generated)} document(s)"
    if pdfs:
        summary += f" and {pdfs} PDF(s)"
    summary += f" in {result.outdir}/ ({len(result.failed)} failed). Manifest: {result.manifest}"
    print(summary)
    if result.merged_pdf:
        print(f"Merged PDF: {result.merged_pdf}")
    if result.zip_path:
        print(f"ZIP: {result.zip_path}")
    if result.pdf_unavailable:
        print(INSTALL_HELP, file=sys.stderr)
    return result.exit_code


def cmd_serve(args: argparse.Namespace) -> int:
    from .web import serve

    serve(args.port)
    return 0


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    handler = {"check": cmd_check, "build": cmd_build, "serve": cmd_serve}[args.command]
    try:
        return handler(args)
    except (DataError, TemplateFileError, OSError) as exc:
        print(f"rowdocs: error: {exc}", file=sys.stderr)
        return EXIT_VALIDATION


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
