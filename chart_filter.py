"""Shared helpers for limiting calculations to selected charts."""

from __future__ import annotations

from collections.abc import Iterable

import pandas as pd


def chart_key(group_name, chart_name) -> tuple[str, str]:
    return str(group_name).strip(), str(chart_name).strip()


def filter_chart_information(
    chart_info: pd.DataFrame,
    selected_chart_keys: Iterable[tuple[str, str]] | None,
) -> pd.DataFrame:
    """Return a filtered copy; None means unrestricted and an empty set means none."""
    if selected_chart_keys is None:
        return chart_info.copy()
    selected = {chart_key(group, chart) for group, chart in selected_chart_keys}
    if not selected or 'GroupName' not in chart_info.columns or 'ChartName' not in chart_info.columns:
        return chart_info.iloc[0:0].copy()
    keys = [chart_key(group, chart) for group, chart in zip(chart_info['GroupName'], chart_info['ChartName'])]
    return chart_info.loc[[key in selected for key in keys]].copy()
