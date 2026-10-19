"""Pre-flight validation (``rowdocs check``): compare template placeholders with the data."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any

from jinja2 import TemplateError

from .data import Dataset, is_empty
from .naming import DEFAULT_PATTERN, RESERVED, pattern_variables, plan_names
from .template import BUILTIN_VARIABLES, WordTemplate

ERROR = "ERROR"
WARN = "WARN"


@dataclass
class Issue:
    level: str  # ERROR or WARN
    code: str  # missing_column, unused_column, empty_value, duplicate_name, template_syntax, ...
    message: str
    row: int | None = None
    variable: str | None = None


@dataclass
class Report:
    data: str
    template: str
    name_pattern: str
    rows_total: int
    rows_selected: list[int]
    columns: list[dict[str, Any]]  # {"header", "key", "used"}
    variables: dict[str, list[str]]  # name -> where it is used
    issues: list[Issue] = field(default_factory=list)

    @property
    def errors(self) -> list[Issue]:
        return [i for i in self.issues if i.level == ERROR]

    @property
    def warnings(self) -> list[Issue]:
        return [i for i in self.issues if i.level == WARN]

    def exit_code(self, strict: bool = False) -> int:
        if self.errors or (strict and self.warnings):
            return 2
        return 0

    def to_dict(self) -> dict[str, Any]:
        return {
            "data": self.data,
            "template": self.template,
            "name_pattern": self.name_pattern,
            "rows_total": self.rows_total,
            "rows_selected": self.rows_selected,
            "columns": self.columns,
            "variables": self.variables,
            "issues": [asdict(i) for i in self.issues],
            "errors": len(self.errors),
            "warnings": len(self.warnings),
            "ok": not self.errors,
        }


def validate(dataset: Dataset, template: WordTemplate, name_pattern: str | None = None) -> Report:
    """Run every pre-flight check and return a :class:`Report`."""
    pattern = name_pattern or DEFAULT_PATTERN
    info = template.inspect()
    variables = {k: list(v) for k, v in info.variables.items()}
    issues: list[Issue] = [Issue(ERROR, "template_syntax", f"template {msg}") for msg in info.errors]

    pattern_ok = True
    try:
        for name in sorted(pattern_variables(pattern)):
            variables.setdefault(name, []).append("name pattern")
    except TemplateError as exc:
        pattern_ok = False
        issues.append(Issue(ERROR, "name_pattern_syntax", f"--name pattern is invalid: {exc}"))

    keys = set(dataset.keys)
    data_vars = sorted(v for v in variables if v not in BUILTIN_VARIABLES)

    # (a) placeholders without a column
    for name in data_vars:
        if name not in keys:
            where = ", ".join(variables[name])
            message = f"'{name}' is used in the {where} but no column matches it"
            issues.append(Issue(ERROR, "missing_column", message, variable=name))

    # (b) columns the template never uses
    used = set(variables)
    for col in dataset.columns:
        if col.key not in used:
            issues.append(
                Issue(
                    WARN,
                    "unused_column",
                    f"column '{col.header}' ({col.key}) is not used by the template",
                    variable=col.key,
                )
            )

    # (c) empty values of used variables, per row
    present = [v for v in data_vars if v in keys]
    for number in dataset.selected:
        row = dataset.row(number)
        empty = [v for v in present if is_empty(row.get(v))]
        for name in empty:
            issues.append(
                Issue(
                    WARN,
                    "empty_value",
                    f"row {number} (sheet line {number + 1}): '{name}' is empty",
                    row=number,
                    variable=name,
                )
            )

    # (d) duplicate output names
    if pattern_ok:
        planned = plan_names(pattern, [(n, dataset.context(n)) for n in dataset.selected])
        groups: dict[str, list[int]] = {}
        for item in planned:
            if item.error:
                issues.append(Issue(WARN, "name_error", f"row {item.row}: {item.error}", row=item.row))
            groups.setdefault(item.base.casefold(), []).append(item.row)
        reported: set[str] = set()
        for item in planned:
            key = item.base.casefold()
            if key in reported:
                continue
            reported.add(key)
            rows = groups[key]
            if len(rows) > 1:
                listed = ", ".join(str(r) for r in rows)
                message = f"rows {listed} would all be named '{item.base}'; extra files get _2, _3... suffixes"
                issues.append(Issue(WARN, "duplicate_name", message, row=rows[0]))
            elif key in RESERVED:
                message = f"row {item.row}: '{item.base}' is reserved by rowdocs; it becomes '{item.final}'"
                issues.append(Issue(WARN, "duplicate_name", message, row=item.row))

    columns = [{"header": c.header, "key": c.key, "used": c.key in used} for c in dataset.columns]
    return Report(
        data=dataset.source,
        template=template.name,
        name_pattern=pattern,
        rows_total=len(dataset.rows),
        rows_selected=list(dataset.selected),
        columns=columns,
        variables=dict(sorted(variables.items())),
        issues=issues,
    )


def format_report(report: Report) -> str:
    """Human-readable report for the terminal."""
    lines = [
        f"rowdocs check: {report.data} + {report.template}",
        f"Rows: {report.rows_total} in file, {len(report.rows_selected)} selected",
        "",
        "Column -> key",
    ]
    width = max([len(c["header"]) for c in report.columns] + [6])
    key_width = max([len(c["key"]) for c in report.columns] + [3])
    for col in report.columns:
        mark = "used" if col["used"] else "unused"
        lines.append(f"  {col['header']:<{width}}  ->  {col['key']:<{key_width}}  {mark}")
    lines.append("")
    lines.append(f"Template variables ({len(report.variables)}):")
    for name, where in report.variables.items():
        lines.append(f"  {name:<{key_width}}  {', '.join(where)}")
    lines.append("")
    if report.issues:
        lines.append(f"{'LEVEL':<6} {'WHERE':<10} MESSAGE")
        for issue in sorted(report.issues, key=lambda i: (i.level != ERROR, i.row or 0)):
            where = f"row {issue.row}" if issue.row else ("data" if issue.code == "unused_column" else "template")
            lines.append(f"{issue.level:<6} {where:<10} {issue.message}")
        lines.append("")
    lines.append(f"Result: {len(report.errors)} error(s), {len(report.warnings)} warning(s)")
    return "\n".join(lines)
