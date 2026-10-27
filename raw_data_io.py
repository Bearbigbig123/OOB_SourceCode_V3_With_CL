"""Shared helpers for locating and reading raw measurement data files."""

from __future__ import annotations

import os
import re
from collections.abc import Iterator

import pandas as pd


RAW_DATA_EXTENSIONS = (".csv", ".xlsx", ".xls", ".xlsm", ".xlsb")
_EXTENSION_PRIORITY = {extension: index for index, extension in enumerate(RAW_DATA_EXTENSIONS)}


def is_raw_data_file(path: str) -> bool:
    return os.path.splitext(path)[1].lower() in _EXTENSION_PRIORITY


def iter_raw_data_files(directory: str) -> Iterator[str]:
    """Yield supported raw-data files in a stable, format-aware order."""
    if not os.path.isdir(directory):
        return
    paths = (
        os.path.join(directory, filename)
        for filename in os.listdir(directory)
        if is_raw_data_file(filename)
    )
    yield from sorted(
        paths,
        key=lambda path: (
            _EXTENSION_PRIORITY[os.path.splitext(path)[1].lower()],
            os.path.basename(path).casefold(),
        ),
    )


def find_matching_raw_data_file(directory: str, group_name, chart_name) -> str | None:
    """Find a chart's CSV or Excel raw-data file using the established naming rule."""
    base_name = f"{str(group_name).strip()}_{str(chart_name).strip()}"
    pattern = re.compile(
        rf"^{re.escape(base_name)}(?:_\d+_\d+)?(?:\.csv|\.xlsx|\.xls|\.xlsm|\.xlsb)$",
        re.IGNORECASE,
    )
    return next(
        (path for path in iter_raw_data_files(directory) if pattern.match(os.path.basename(path))),
        None,
    )


def find_raw_data_file_by_prefix(directory: str, prefix: str) -> str | None:
    """Find the first supported file whose stem starts with ``prefix``."""
    normalized_prefix = str(prefix).casefold()
    return next(
        (
            path
            for path in iter_raw_data_files(directory)
            if os.path.splitext(os.path.basename(path))[0].casefold().startswith(normalized_prefix)
        ),
        None,
    )


def read_raw_data(path: str, **kwargs) -> pd.DataFrame:
    """Read CSV or Excel raw data; Excel data comes from the first worksheet."""
    extension = os.path.splitext(path)[1].lower()
    if extension == ".csv":
        return pd.read_csv(path, **kwargs)
    if extension not in _EXTENSION_PRIORITY:
        raise ValueError(f"Unsupported raw data format: {extension or '(no extension)'}")

    excel_kwargs = {
        key: value
        for key, value in kwargs.items()
        if key not in {"encoding", "float_precision"}
    }
    excel_kwargs.setdefault("sheet_name", 0)
    if extension in {".xlsx", ".xlsm"}:
        excel_kwargs.setdefault("engine", "openpyxl")
    elif extension == ".xls":
        excel_kwargs.setdefault("engine", "xlrd")
    elif extension == ".xlsb":
        excel_kwargs.setdefault("engine", "pyxlsb")
    return pd.read_excel(path, **excel_kwargs)
