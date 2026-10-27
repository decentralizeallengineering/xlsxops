# SPDX-License-Identifier: Apache-2.0
"""Every write operation must leave styles, charts, other sheets and names intact."""
from __future__ import annotations

import base64
import zipfile

import openpyxl
import pytest
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.table import Table, TableStyleInfo

import xlsxops as X
from conftest import snapshot

# Each case: operation name, kwargs, and snapshot keys the operation is allowed to change.
CASES = [
    ("write_cells", {"sheet": "Sales", "cells": {"H1": "x"}}, set()),
    ("add_sheet", {"name": "Extra"}, {"sheets", "charts"}),
    ("add_column", {"sheet": "Sales", "header": "Extra", "fill": 1}, set()),
    ("calculated_column", {"sheet": "Sales", "header": "Calc", "formula": "=D2*2"}, set()),
    ("set_formula", {"sheet": "Sales", "cell": "J20", "formula": "=SUM(D2:D11)"}, set()),
    ("total_row", {"sheet": "Sales", "columns": ["Qty"]}, set()),
    ("sort", {"sheet": "Sales", "keys": [{"column": "Qty"}]}, set()),
    ("filter_rows", {"sheet": "Sales", "new_sheet": "F",
                     "conditions": [{"column": "Qty", "op": ">", "value": 2}]}, {"sheets", "charts"}),
    ("dedupe", {"sheet": "Sales"}, {"notes_formulas"}),  # SUM(Sales!D2:D11) shrinks with the table
    ("split_column", {"sheet": "Sales", "column": "Rep", "regex": "^(.)(.*)$"}, set()),
    ("lookup", {"sheet": "Sales", "key": "Product", "source_sheet": "Products",
                "columns": ["Cost"]}, set()),
    ("pivot_summary", {"sheet": "Sales", "group_by": ["Region"], "new_sheet": "P", "formulas": True,
                       "values": [{"column": "Qty", "agg": "sum"}]}, {"sheets", "charts"}),
    ("chart", {"sheet": "Sales", "type": "bar", "x": "Rep", "y": ["Total"]}, {"charts", "chart_titles"}),
    ("conditional_format", {"sheet": "Sales", "column": "Qty", "rule": "outliers"}, {"cf"}),
    ("data_validation", {"sheet": "Sales", "column": "Qty", "type": "whole", "min": 0}, {"dv"}),
    ("format_range", {"sheet": "Sales", "range": "D2:D11", "number_format": "0.0",
                      "freeze": "A2"}, set()),
    ("multisheet_report", {"sheet": "Sales", "group_by": "Region"}, {"sheets", "charts"}),
    ("verify_formulas", {"write_cache": True}, set()),
]


@pytest.mark.parametrize("name,kwargs,may_change", CASES, ids=[c[0] for c in CASES])
def test_roundtrip_preserves_workbook(sales, name, kwargs, may_change):
    before = snapshot(sales)
    r = X.call(name, {"path": sales, **kwargs})
    assert r["ok"], r["details"]
    assert r["changed"] is True
    after = snapshot(sales)
    for key in before:
        if key in may_change:
            continue
        assert after[key] == before[key], f"{name} changed {key}"
    if "sheets" in may_change:  # sheets may be added, never removed or reordered
        assert [s for s in after["sheets"] if s in before["sheets"]] == before["sheets"]
        assert after["charts"]["Sales"] == before["charts"]["Sales"]
    if "charts" in may_change and "sheets" not in may_change:
        assert after["charts"]["Sales"] == before["charts"]["Sales"] + 1


def test_all_formulas_have_cached_values_after_write(sales):
    X.calculated_column(sales, sheet="Sales", header="Calc", formula="=TEXT(F2,\"mmm d\")&\": \"&D2")
    wb = openpyxl.load_workbook(sales, data_only=True)
    wf = openpyxl.load_workbook(sales)
    missing = [f"{ws.title}!{c.coordinate}" for ws in wf.worksheets for row in ws.iter_rows()
               for c in row if c.data_type == "f" and wb[ws.title][c.coordinate].value is None]
    assert missing == []
    assert wb["Sales"]["H2"].value == "Jan 5: 10"
    with zipfile.ZipFile(sales) as z:
        xml = z.read("xl/worksheets/sheet1.xml").decode()
    assert '<c r="H2" t="str"><f>' in xml or 't="str"' in xml
    assert 'fullCalcOnLoad="1"' in zipfile.ZipFile(sales).read("xl/workbook.xml").decode()


def test_tables_and_images_survive(tmp_path):
    png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP8z8DAwMDAxMDAwMDAAAANHQEDasKb6QAAAABJRU5ErkJggg==")
    (tmp_path / "logo.png").write_bytes(png)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Name", "Score"])
    for i, (n, s) in enumerate([("a", 1), ("b", 5), ("c", 3)]):
        ws.append([n, s])
    t = Table(displayName="Scores", ref="A1:B4")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(t)
    path = tmp_path / "t.xlsx"
    try:
        ws.add_image(XLImage(str(tmp_path / "logo.png")), "E2")
        has_image = True
    except ImportError:  # Pillow not installed: build the image part by hand below
        has_image = False
    wb.save(path)
    r = X.add_column(str(path), sheet="Data", header="Double", fill="=B2*2")
    assert r["ok"], r["details"]
    wb2 = openpyxl.load_workbook(path)
    tbl = wb2["Data"].tables["Scores"]
    assert tbl.ref == "A1:C4" and [c.name for c in tbl.tableColumns] == ["Name", "Score", "Double"]
    if has_image:
        with zipfile.ZipFile(path) as z:
            assert any(n.startswith("xl/media/") for n in z.namelist())
    r = X.sort(str(path), sheet="Data", keys=[{"column": "Score", "descending": True}])
    assert r["ok"]
    assert [c.value for c in openpyxl.load_workbook(path)["Data"]["A"]][1:] == ["b", "c", "a"]


def test_lossy_content_requires_force(sales, tmp_path):
    # Fake a sparkline extension in Sales so the preflight trips.
    src = zipfile.ZipFile(sales)
    out = tmp_path / "spark.xlsx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = data.replace(b"</worksheet>", b'<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">'
                                    b'<x14:sparklineGroups xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/>'
                                    b"</ext></extLst></worksheet>")
            z.writestr(item, data)
    src.close()
    r = X.write_cells(str(out), sheet="Notes", cells={"D1": 1})
    assert r["ok"] is False and r["details"]["error_type"] == "would_lose_content"
    assert "sparklines" in r["details"]["lossy_features"]
    summary = X.read_summary(str(out))
    assert summary["ok"] and "sparklines" in summary["details"]["content_lost_on_write"]
    r = X.write_cells(str(out), sheet="Notes", cells={"D1": 1}, force=True,
                      out_path=str(tmp_path / "copy.xlsx"))
    assert r["ok"] and any("sparklines" in w for w in r["warnings"])


def test_images_survive_without_pillow(tmp_path, monkeypatch):
    """openpyxl drops pictures when Pillow is missing; xlsxops keeps the raw bytes."""
    import io
    import openpyxl.reader.drawings as rd
    from xlsxops import _core
    png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP8z8DAwMDAxMDAwMDAAAANHQEDasKb6QAAAABJRU5ErkJggg==")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "logo below"
    ws.add_image(_core._RawImage(io.BytesIO(png)), "B3")
    path = tmp_path / "img.xlsx"
    wb.save(path)
    monkeypatch.setattr(rd, "PILImage", False)        # behave as if Pillow is not installed
    monkeypatch.setattr(_core, "_PIL_AVAILABLE", False)
    r = X.write_cells(str(path), cells={"A2": 1})
    assert r["ok"], r["details"]
    with zipfile.ZipFile(path) as z:
        media = [n for n in z.namelist() if n.startswith("xl/media/")]
        assert len(media) == 1 and z.read(media[0]) == png
    assert rd.PILImage is False  # the temporary patch was undone
