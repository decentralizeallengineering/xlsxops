# SPDX-License-Identifier: Apache-2.0
"""xlsxops: reliable Excel (.xlsx) operations for LLM tool layers.

Every public operation is a plain function returning a JSON-serializable report::

    {"ok": bool, "changed": bool, "warnings": [str], "details": {...}}

User errors never raise; they come back as ``ok: false`` with ``details.error``.
``SCHEMAS[name]`` holds the JSON Schema (<= 600 chars) for each operation's arguments.
"""
from ._core import OPERATIONS, SCHEMAS, jsonable
from .formula import SUPPORTED_FUNCTIONS
from .ops_basic import (add_column, add_sheet, calculated_column, read_summary, set_formula,
                        total_row, verify_formulas, write_cells)
from .ops_data import (dedupe, filter_rows, from_csv, lookup, multisheet_report, pivot_summary,
                       sort, split_column)
from .ops_format import chart, conditional_format, data_validation, format_range

__version__ = "0.1.0"


def call(name: str, args: dict | None = None) -> dict:
    """Dispatch a tool call by name with a JSON arguments object."""
    fn = OPERATIONS.get(name)
    if fn is None:
        return {"ok": False, "changed": False, "warnings": [],
                "details": {"error": f"unknown operation {name!r}", "error_type": "unknown_operation",
                            "operations": sorted(OPERATIONS)}}
    if args is not None and not isinstance(args, dict):
        return {"ok": False, "changed": False, "warnings": [],
                "details": {"error": "arguments must be a JSON object", "error_type": "invalid_arguments"}}
    return fn(**(args or {}))


def tool_definitions() -> list[dict]:
    """Tool specs (name, description, input_schema) ready for an LLM tool-use API."""
    out = []
    for name, fn in sorted(OPERATIONS.items()):
        doc = (fn.__doc__ or "").strip().split("\n\n")[0].replace("\n", " ")
        out.append({"name": name, "description": " ".join(doc.split()), "input_schema": SCHEMAS[name]})
    return out


__all__ = [
    "read_summary", "write_cells", "add_sheet", "add_column", "calculated_column", "set_formula",
    "total_row", "sort", "filter_rows", "dedupe", "split_column", "lookup", "pivot_summary", "chart",
    "conditional_format", "data_validation", "format_range", "from_csv", "multisheet_report",
    "verify_formulas", "OPERATIONS", "SCHEMAS", "SUPPORTED_FUNCTIONS", "call", "tool_definitions",
    "jsonable",
]
