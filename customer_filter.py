"""Shared Customer discovery and DataFrame filtering helpers."""

from __future__ import annotations

import glob
import os
from dataclasses import dataclass
from typing import Iterable

import pandas as pd


CUSTOMER_COLUMN = "Customer"


@dataclass(frozen=True)
class CustomerFilterResult:
    data: pd.DataFrame
    missing_column: bool = False


def normalize_customer(value) -> str | None:
    if pd.isna(value):
        return None
    normalized = str(value).strip()
    return normalized or None


def discover_customers(raw_data_directory: str) -> tuple[list[str], list[str]]:
    """Return sorted unique customers and files that could not be inspected."""
    customers: set[str] = set()
    errors: list[str] = []
    pattern = os.path.join(raw_data_directory, "*.csv")
    for path in glob.glob(pattern):
        try:
            header = pd.read_csv(path, nrows=0)
            if CUSTOMER_COLUMN not in header.columns:
                continue
            values = pd.read_csv(path, usecols=[CUSTOMER_COLUMN])[CUSTOMER_COLUMN]
            customers.update(v for v in map(normalize_customer, values) if v is not None)
        except Exception:
            errors.append(path)
    return sorted(customers, key=lambda value: value.casefold()), errors


def apply_customer_filter(
    dataframe: pd.DataFrame,
    selected_customers: Iterable[str] | None,
) -> CustomerFilterResult:
    """Filter a copy of dataframe. An empty selection means all customers."""
    data = dataframe.copy()
    selected = {value for value in map(normalize_customer, selected_customers or []) if value}
    if not selected:
        return CustomerFilterResult(data)
    if CUSTOMER_COLUMN not in data.columns:
        return CustomerFilterResult(data.iloc[0:0].copy(), missing_column=True)
    normalized = data[CUSTOMER_COLUMN].map(normalize_customer)
    return CustomerFilterResult(data.loc[normalized.isin(selected)].copy())
