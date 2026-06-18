from __future__ import annotations

import gzip

import pandas as pd
import pytest

from excel2csv.cli import Excel2CsvError, main, merge_frames
from excel2csv.cli import SheetFrame


def write_workbook(path, sheets):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, index=False, sheet_name=sheet_name)


def test_merges_workbooks_and_sheets_to_utf8_sig_csv(tmp_path):
    workbook_a = tmp_path / "a.xlsx"
    workbook_b = tmp_path / "b.xlsx"
    output = tmp_path / "merged.csv"

    write_workbook(
        workbook_a,
        {
            "January": pd.DataFrame({"id": [1], "name": ["Alice"]}),
            "February": pd.DataFrame({"id": [2], "name": ["Bob"]}),
        },
    )
    write_workbook(workbook_b, {"March": pd.DataFrame({"id": [3], "name": ["Carol"]})})

    assert main([str(workbook_a), str(workbook_b), "-o", str(output)]) == 0

    raw = output.read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf")
    assert output.read_text(encoding="utf-8-sig").splitlines() == [
        "id,name",
        "1,Alice",
        "2,Bob",
        "3,Carol",
    ]


def test_writes_gzip_without_bom(tmp_path):
    workbook = tmp_path / "book.xlsx"
    output = tmp_path / "merged.csv.gz"
    write_workbook(workbook, {"Sheet1": pd.DataFrame({"id": [1], "name": ["Alice"]})})

    assert main([str(workbook), "-o", str(output)]) == 0

    with gzip.open(output, "rb") as file_obj:
        raw = file_obj.read()

    assert not raw.startswith(b"\xef\xbb\xbf")
    assert raw.decode("utf-8").splitlines() == ["id,name", "1,Alice"]


def test_rejects_mismatched_columns(tmp_path):
    first = SheetFrame(
        workbook=tmp_path / "a.xlsx",
        sheet_name="Sheet1",
        frame=pd.DataFrame({"id": [1], "name": ["Alice"]}),
    )
    second = SheetFrame(
        workbook=tmp_path / "b.xlsx",
        sheet_name="Sheet1",
        frame=pd.DataFrame({"id": [2], "email": ["alice@example.com"]}),
    )

    with pytest.raises(Excel2CsvError, match="column mismatch"):
        merge_frames([first, second])
