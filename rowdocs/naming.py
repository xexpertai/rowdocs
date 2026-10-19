"""Output file names: render the --name pattern, sanitise it and resolve collisions."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from typing import Any

from jinja2 import TemplateError, meta

from .filters import make_env

DEFAULT_PATTERN = "row_{{ _row }}"
MAX_NAME_LENGTH = 120
# Names rowdocs writes itself; a row rendering to one of these gets a suffix instead.
RESERVED = frozenset({"all", "manifest"})
_WINDOWS_RESERVED = frozenset(
    {"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)), *(f"lpt{i}" for i in range(1, 10))}
)


@dataclass
class PlannedName:
    row: int
    base: str  # sanitised name before collision handling
    final: str  # unique name (without extension)
    error: str = ""


def pattern_variables(pattern: str) -> set[str]:
    """Undeclared variables used by a naming pattern. Raises TemplateError on bad syntax."""
    env = make_env()
    variables = meta.find_undeclared_variables(env.parse(pattern))
    env.from_string(pattern)
    return variables


def sanitize(name: str, fallback: str) -> str:
    """Make ``name`` safe as a single file name: no separators, no traversal, <= 120 chars."""
    name = unicodedata.normalize("NFC", name)
    name = re.sub(r"[\x00-\x1f\x7f]", "", name)
    name = re.sub(r'[/\\<>:"|?*]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    name = name.strip(". ")  # no leading dots (hidden files, "..") and no trailing dots/spaces
    name = name[:MAX_NAME_LENGTH].rstrip(". ")
    if not name:
        name = fallback
    if name.split(".")[0].casefold() in _WINDOWS_RESERVED:
        name = f"_{name}"
    return name


def plan_names(pattern: str, contexts: list[tuple[int, dict[str, Any]]]) -> list[PlannedName]:
    """Render the pattern for each (row number, context) and make every name unique.

    The first occurrence keeps the plain name; later ones get ``_2``, ``_3``...
    Comparison is case-insensitive so output is safe on Windows and macOS file systems.
    """
    env = make_env()
    template = env.from_string(pattern or DEFAULT_PATTERN)
    used: set[str] = set(RESERVED)
    planned: list[PlannedName] = []
    for row, context in contexts:
        fallback = f"row_{row}"
        error = ""
        try:
            base = sanitize(template.render(context), fallback)
        except (TemplateError, ValueError, TypeError) as exc:
            base, error = fallback, f"cannot render file name: {exc}"
        final = base
        counter = 2
        while final.casefold() in used:
            suffix = f"_{counter}"
            final = base[: MAX_NAME_LENGTH - len(suffix)] + suffix
            counter += 1
        used.add(final.casefold())
        planned.append(PlannedName(row=row, base=base, final=final, error=error))
    return planned
