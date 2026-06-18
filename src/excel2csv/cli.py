from __future__ import annotations

import argparse
import gzip
import io
import os
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO, Iterable, Sequence

import msoffcrypto
import pandas as pd
from msoffcrypto.exceptions import DecryptionError, FileFormatError, InvalidKeyError


SUPPORTED_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
    ".xls",
    ".xlsb",
}


class Excel2CsvError(Exception):
    """User-facing conversion error."""


@dataclass(frozen=True)
class SheetFrame:
    workbook: Path
    sheet_name: str
    frame: pd.DataFrame


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel2csv",
        description=(
            "Merge same-column Excel workbooks/sheets into one CSV "
            "(UTF-8 with BOM) or CSV.GZ (UTF-8 without BOM)."
        ),
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        type=Path,
        help="Excel files or directories containing Excel files.",
    )
    parser.add_argument(
        "-o",
        "--output",
        required=True,
        type=Path,
        help="Output path. Use .csv for UTF-8 with BOM, or .csv.gz for gzip UTF-8.",
    )
    parser.add_argument(
        "--gzip",
        action="store_true",
        help="Write gzip output even when the output path does not end with .gz.",
    )
    parser.add_argument(
        "--password",
        help=(
            "Password for encrypted workbooks. EXCEL2CSV_PASSWORD can also be used. "
            "All encrypted inputs must currently share this password."
        ),
    )
    parser.add_argument(
        "--password-file",
        type=Path,
        help="Read the workbook password from the first line of this file.",
    )
    parser.add_argument(
        "--sheet",
        action="append",
        dest="sheets",
        help="Sheet name to include. Repeat to include multiple sheets. Defaults to all sheets.",
    )
    parser.add_argument(
        "--include-empty-sheets",
        action="store_true",
        help="Include sheets with no rows and no columns. By default they are skipped.",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="When an input is a directory, search it recursively for supported Excel files.",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        password = resolve_password(args.password, args.password_file)
        input_files = expand_inputs(args.inputs, recursive=args.recursive)
        sheets = load_sheets(
            input_files,
            requested_sheets=args.sheets,
            password=password,
            include_empty_sheets=args.include_empty_sheets,
        )
        merged = merge_frames(sheets)
        write_output(merged, args.output, force_gzip=args.gzip)
    except Excel2CsvError as exc:
        parser.exit(2, f"excel2csv: error: {exc}\n")

    print(
        f"wrote {len(merged):,} rows from {len(sheets):,} sheet(s) "
        f"across {len(input_files):,} workbook(s) to {args.output}",
        file=sys.stderr,
    )
    return 0


def resolve_password(password: str | None, password_file: Path | None) -> str | None:
    if password and password_file:
        raise Excel2CsvError("use either --password or --password-file, not both")

    if password_file:
        try:
            return password_file.read_text(encoding="utf-8").splitlines()[0]
        except IndexError as exc:
            raise Excel2CsvError(f"password file is empty: {password_file}") from exc
        except OSError as exc:
            raise Excel2CsvError(f"cannot read password file {password_file}: {exc}") from exc

    return password or os.environ.get("EXCEL2CSV_PASSWORD") or None


def expand_inputs(inputs: Iterable[Path], *, recursive: bool) -> list[Path]:
    files: list[Path] = []
    for input_path in inputs:
        path = input_path.expanduser()
        if not path.exists():
            raise Excel2CsvError(f"input does not exist: {input_path}")
        if path.is_dir():
            pattern = "**/*" if recursive else "*"
            files.extend(
                sorted(
                    candidate
                    for candidate in path.glob(pattern)
                    if candidate.is_file() and is_supported_excel_file(candidate)
                )
            )
        elif is_supported_excel_file(path):
            files.append(path)
        else:
            raise Excel2CsvError(f"unsupported Excel file extension: {path}")

    if not files:
        raise Excel2CsvError("no supported Excel files were found")
    return files


def is_supported_excel_file(path: Path) -> bool:
    return not path.name.startswith("~$") and path.suffix.lower() in SUPPORTED_EXTENSIONS


def load_sheets(
    input_files: Iterable[Path],
    *,
    requested_sheets: Sequence[str] | None,
    password: str | None,
    include_empty_sheets: bool,
) -> list[SheetFrame]:
    sheet_frames: list[SheetFrame] = []
    for workbook in input_files:
        source = open_workbook_source(workbook, password=password)
        engine = engine_for(workbook)
        try:
            with pd.ExcelFile(source, engine=engine) as excel_file:
                sheet_names = select_sheets(
                    workbook,
                    excel_file.sheet_names,
                    requested_sheets=requested_sheets,
                )
                for sheet_name in sheet_names:
                    frame = pd.read_excel(
                        excel_file,
                        sheet_name=sheet_name,
                        dtype=object,
                        keep_default_na=False,
                    )
                    if frame.empty and len(frame.columns) == 0 and not include_empty_sheets:
                        continue
                    validate_unique_columns(workbook, sheet_name, frame.columns)
                    sheet_frames.append(
                        SheetFrame(
                            workbook=workbook,
                            sheet_name=sheet_name,
                            frame=frame,
                        )
                    )
        except Excel2CsvError:
            raise
        except Exception as exc:  # pandas/openpyxl/xlrd expose several exception types.
            raise Excel2CsvError(f"failed to read {workbook}: {exc}") from exc

    if not sheet_frames:
        raise Excel2CsvError("no readable sheets were found")
    return sheet_frames


def open_workbook_source(path: Path, *, password: str | None) -> Path | BinaryIO:
    try:
        with path.open("rb") as file_obj:
            office_file = msoffcrypto.OfficeFile(file_obj)
            if not office_file.is_encrypted():
                return path
            if not password:
                raise Excel2CsvError(
                    f"{path} is encrypted; pass --password, --password-file, "
                    "or EXCEL2CSV_PASSWORD"
                )
            decrypted = io.BytesIO()
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            return decrypted
    except FileFormatError:
        return path
    except InvalidKeyError as exc:
        raise Excel2CsvError(f"invalid password for {path}") from exc
    except DecryptionError as exc:
        raise Excel2CsvError(f"could not decrypt {path}: {exc}") from exc
    except OSError as exc:
        raise Excel2CsvError(f"cannot open {path}: {exc}") from exc


def engine_for(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return "openpyxl"
    if suffix == ".xls":
        return "xlrd"
    if suffix == ".xlsb":
        return "pyxlsb"
    raise Excel2CsvError(f"unsupported Excel file extension: {path}")


def select_sheets(
    workbook: Path,
    available_sheets: Sequence[str],
    *,
    requested_sheets: Sequence[str] | None,
) -> Sequence[str]:
    if not requested_sheets:
        return available_sheets

    missing = [sheet for sheet in requested_sheets if sheet not in available_sheets]
    if missing:
        available = ", ".join(available_sheets)
        requested = ", ".join(missing)
        raise Excel2CsvError(
            f"{workbook} does not contain requested sheet(s): {requested}. "
            f"Available sheets: {available}"
        )
    return requested_sheets


def validate_unique_columns(workbook: Path, sheet_name: str, columns: Sequence[object]) -> None:
    counts = Counter(str(column) for column in columns)
    duplicates = sorted(column for column, count in counts.items() if count > 1)
    if duplicates:
        duplicate_text = ", ".join(duplicates)
        raise Excel2CsvError(
            f"{workbook} sheet {sheet_name!r} has duplicate column names: {duplicate_text}"
        )


def merge_frames(sheets: Sequence[SheetFrame]) -> pd.DataFrame:
    expected_columns = list(sheets[0].frame.columns)
    for sheet in sheets[1:]:
        columns = list(sheet.frame.columns)
        if columns != expected_columns:
            raise Excel2CsvError(
                "column mismatch in "
                f"{sheet.workbook} sheet {sheet.sheet_name!r}. "
                f"Expected {expected_columns!r}, got {columns!r}"
            )

    return pd.concat([sheet.frame for sheet in sheets], ignore_index=True)


def write_output(frame: pd.DataFrame, output: Path, *, force_gzip: bool) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    gzip_output = force_gzip or output.name.lower().endswith(".gz")
    try:
        if gzip_output:
            with gzip.open(output, "wt", encoding="utf-8", newline="") as file_obj:
                frame.to_csv(file_obj, index=False, lineterminator="\n")
        else:
            with output.open("w", encoding="utf-8-sig", newline="") as file_obj:
                frame.to_csv(file_obj, index=False, lineterminator="\n")
    except OSError as exc:
        raise Excel2CsvError(f"cannot write output {output}: {exc}") from exc
