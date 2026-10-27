# SPDX-License-Identifier: Apache-2.0
"""The tool-layer contract: schemas, report shape, and never raising."""
from __future__ import annotations

import ast
import json
from pathlib import Path

import jsonschema
import pytest

import xlsxops as X

OPS = ["read_summary", "write_cells", "add_sheet", "add_column", "calculated_column", "set_formula",
       "total_row", "sort", "filter_rows", "dedupe", "split_column", "lookup", "pivot_summary", "chart",
       "conditional_format", "data_validation", "format_range", "from_csv", "multisheet_report",
       "verify_formulas"]


def test_every_operation_has_a_small_valid_schema():
    assert sorted(X.SCHEMAS) == sorted(OPS) == sorted(X.OPERATIONS)
    for name, schema in X.SCHEMAS.items():
        assert len(json.dumps(schema)) <= 600, name
        jsonschema.Draft202012Validator.check_schema(schema)
        assert schema["type"] == "object" and "path" in schema["required"]


def test_tool_definitions():
    tools = X.tool_definitions()
    assert {t["name"] for t in tools} == set(OPS)
    assert all(t["description"] and t["input_schema"] for t in tools)


@pytest.mark.parametrize("name", OPS)
def test_missing_file_reports_instead_of_raising(name, tmp_path):
    args = {"path": str(tmp_path / "nope.xlsx")}
    required = X.SCHEMAS[name]["required"]
    filler = {"name": "S", "header": "H", "formula": "=1", "cell": "A1", "keys": [{"column": "A"}],
              "conditions": [{"column": "A", "op": "blank"}], "new_sheet": "N", "column": "A",
              "key": "A", "source_sheet": "S", "columns": ["A"], "group_by": ["A"],
              "values": [{"column": "A", "agg": "sum"}], "type": "bar", "rule": "data_bar",
              "csv_path": str(tmp_path / "missing.csv")}
    if name == "multisheet_report":
        filler["group_by"] = "A"
    if name == "data_validation":
        filler["type"] = "list"
    for k in required:
        if k != "path":
            args[k] = filler[k]
    args.update({"write_cells": {"cells": {"A1": 1}}, "format_range": {"autofit": True},
                 "split_column": {"delimiter": ","}}.get(name, {}))
    r = X.call(name, args)
    json.dumps(r)
    if name == "add_sheet":  # add_sheet creates a missing workbook on purpose
        assert r["ok"] is True and r["details"]["created_workbook"] is True
        return
    assert r["ok"] is False and r["changed"] is False
    assert r["details"]["error_type"] == "file_not_found"


def test_bad_arguments_are_reported(sales):
    r = X.call("sort", {"path": sales, "keys": "Qty"})
    assert r["ok"] is False and r["details"]["error_type"] == "invalid_arguments"
    r = X.call("sort", {"path": sales, "keys": [{"column": "Qty"}], "bogus": 1})
    assert r["ok"] is False and r["details"]["error_type"] == "invalid_arguments"
    r = X.call("no_such_op", {})
    assert r["ok"] is False
    r = X.sort(sales, keys=[{"column": "Nope"}], sheet="Sales")
    assert r["details"]["error_type"] == "column_not_found" and "Qty" in r["details"]["error"]
    r = X.read_summary(sales, sheet="Nope")
    assert r["details"]["error_type"] == "sheet_not_found"


def test_not_a_workbook(tmp_path):
    p = tmp_path / "fake.xlsx"
    p.write_text("not a zip")
    r = X.read_summary(str(p))
    assert r["ok"] is False and r["details"]["error_type"] == "invalid_file"
    r = X.read_summary(str(tmp_path / "x.csv"))
    assert r["ok"] is False


def test_failed_write_leaves_original_untouched(sales):
    before = Path(sales).read_bytes()
    r = X.filter_rows(sales, sheet="Sales", new_sheet="Notes",
                      conditions=[{"column": "Qty", "op": ">", "value": 1}])
    assert r["ok"] is False
    assert Path(sales).read_bytes() == before
    assert not [p for p in Path(sales).parent.iterdir() if p.name.startswith(".xlsxops-")]


def test_out_path_leaves_source_untouched(sales, tmp_path):
    before = Path(sales).read_bytes()
    out = tmp_path / "copy.xlsx"
    r = X.add_sheet(sales, name="New", out_path=str(out))
    assert r["ok"] and Path(sales).read_bytes() == before
    assert "New" in X.read_summary(str(out))["details"]["sheet_names"]


def test_no_forbidden_imports():
    allowed = {"openpyxl", "jsonschema", "pandas", "numpy", "docx", "docxcompose", "pptx", "pypdf",
               "pdfplumber", "fpdf"}
    banned = {"subprocess", "socket", "urllib", "http", "requests", "ctypes", "win32com"}
    stdlib = {"__future__", "contextlib", "csv", "datetime", "functools", "inspect", "io", "json",
              "math", "os", "re", "statistics", "sys", "tempfile", "threading", "time", "warnings", "zipfile",
              "copy", "dataclasses", "decimal", "pathlib", "xml"}
    for f in (Path(X.__file__).parent).glob("*.py"):
        tree = ast.parse(f.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            names = []
            if isinstance(node, ast.Import):
                names = [a.name for a in node.names]
            elif isinstance(node, ast.ImportFrom) and node.level == 0:
                names = [node.module]
            for n in names:
                top = n.split(".")[0]
                assert top not in banned, (f.name, n)
                assert top in allowed | stdlib, (f.name, n)
