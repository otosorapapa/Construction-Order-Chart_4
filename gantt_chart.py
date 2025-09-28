"""Utility for drawing a project level Gantt chart.

The module exposes :func:`create_project_gantt_chart` which receives a
``pandas.DataFrame`` having the following columns:

* ``案件名`` – project name shown on the Y axis and in the legend.
* ``開始日`` – task start date.
* ``終了日`` – task end date (inclusive).

The function returns a ``plotly.graph_objects.Figure`` instance that renders a
Gantt style visualisation fulfilling the requirements from the task
description.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, List, Sequence

import pandas as pd
import plotly.graph_objects as go
from dateutil.relativedelta import relativedelta


@dataclass(frozen=True)
class _AxisMarks:
    """Container for tick positions, labels and drawing domain."""

    tick_positions: Sequence[pd.Timestamp]
    tick_labels: Sequence[str]
    major_marks: Sequence[pd.Timestamp]
    minor_marks: Sequence[pd.Timestamp]
    domain_start: pd.Timestamp
    domain_end: pd.Timestamp


def _ensure_datetime(series: pd.Series, label: str) -> pd.Series:
    """Convert a column to ``datetime64`` while raising on complete failure."""

    converted = pd.to_datetime(series, errors="coerce")
    if converted.isna().all():
        raise ValueError(f"列 '{label}' の有効な日付が見つかりません。")
    return converted


def _build_axis_marks(starts: pd.Series, ends: pd.Series) -> _AxisMarks:
    """Create tick marks and grid positions for the chart."""

    start_value = starts.min()
    end_value = ends.max()
    if pd.isna(start_value) or pd.isna(end_value):
        raise ValueError("開始日または終了日に有効な値が必要です。")

    domain_start = pd.Timestamp(start_value.year, start_value.month, 1)
    end_month_start = pd.Timestamp(end_value.year, end_value.month, 1)
    domain_end = (
        end_month_start + relativedelta(months=1) - pd.Timedelta(days=1)
    )

    months = pd.date_range(domain_start, domain_end, freq="MS")
    major_marks: List[pd.Timestamp] = []
    major_labels: List[str] = []
    minor_marks: List[pd.Timestamp] = []

    for month_start in months:
        month_end = month_start + relativedelta(months=1) - pd.Timedelta(days=1)
        major_marks.append(month_start)
        major_labels.append(month_start.strftime("%Y/%m"))

        for day in (6, 12, 18, 24):
            candidate = month_start + pd.Timedelta(days=day - 1)
            if candidate <= month_end:
                minor_marks.append(candidate)

        if month_end not in major_marks:
            minor_marks.append(month_end)

    # Remove duplicates while keeping chronological order.
    def _deduplicate(values: Iterable[pd.Timestamp]) -> List[pd.Timestamp]:
        seen: Dict[pd.Timestamp, None] = {}
        for value in values:
            if value not in seen:
                seen[value] = None
        return list(seen.keys())

    minor_marks = _deduplicate(sorted(minor_marks))
    tick_positions = _deduplicate(sorted(list(major_marks) + list(minor_marks)))
    label_map: Dict[pd.Timestamp, str] = {
        mark: label for mark, label in zip(major_marks, major_labels)
    }
    tick_labels = [label_map.get(position, "") for position in tick_positions]

    return _AxisMarks(
        tick_positions=tick_positions,
        tick_labels=tick_labels,
        major_marks=major_marks,
        minor_marks=minor_marks,
        domain_start=domain_start,
        domain_end=domain_end,
    )


def create_project_gantt_chart(df: pd.DataFrame) -> go.Figure:
    """Create a Plotly Gantt chart from the provided dataframe.

    Parameters
    ----------
    df:
        DataFrame that must contain the columns ``案件名`` (project name),
        ``開始日`` (start date) and ``終了日`` (end date). The start/end dates
        are treated as inclusive.

    Returns
    -------
    plotly.graph_objects.Figure
        A figure instance containing the configured Gantt chart.
    """

    required_columns = {"案件名", "開始日", "終了日"}
    missing = required_columns - set(df.columns)
    if missing:
        missing_str = ", ".join(sorted(missing))
        raise ValueError(f"DataFrame に必要な列がありません: {missing_str}")

    working = df.copy()
    working["開始日"] = _ensure_datetime(working["開始日"], "開始日")
    working["終了日"] = _ensure_datetime(working["終了日"], "終了日")

    valid_mask = (~working["開始日"].isna()) & (~working["終了日"].isna())
    filtered = working.loc[valid_mask].copy()
    filtered = filtered.loc[filtered["終了日"] >= filtered["開始日"]]
    if filtered.empty:
        raise ValueError("開始日と終了日が正しく設定された行が存在しません。")

    axis_marks = _build_axis_marks(filtered["開始日"], filtered["終了日"])

    unique_projects = list(dict.fromkeys(filtered["案件名"].astype(str)))
    color_sequence = go.Figure().layout.template.layout.colorway or []
    if not color_sequence:
        color_sequence = [
            "#1f77b4",
            "#ff7f0e",
            "#2ca02c",
            "#d62728",
            "#9467bd",
            "#8c564b",
            "#e377c2",
            "#7f7f7f",
            "#bcbd22",
            "#17becf",
        ]
    color_map = {
        project: color_sequence[i % len(color_sequence)]
        for i, project in enumerate(unique_projects)
    }

    fig = go.Figure()
    legend_drawn: Dict[str, bool] = {}
    for _, row in filtered.iterrows():
        project = str(row["案件名"])
        start = row["開始日"]
        end = row["終了日"]
        duration = (end - start).days + 1
        if duration <= 0:
            continue

        fig.add_trace(
            go.Bar(
                x=[duration],
                y=[project],
                base=start,
                orientation="h",
                marker=dict(color=color_map[project]),
                name=project,
                legendgroup=project,
                showlegend=not legend_drawn.get(project, False),
                customdata=[(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))],
                hovertemplate=(
                    "案件名: %{y}<br>開始日: %{customdata[0]}<br>"
                    "終了日: %{customdata[1]}<extra></extra>"
                ),
            )
        )
        legend_drawn[project] = True

    y_categories: List[str] = []
    for trace in fig.data:
        if trace.y:
            value = str(trace.y[0])
            if value not in y_categories:
                y_categories.append(value)
    chart_height = max(360, 40 * len(y_categories) + 120)

    fig.update_layout(
        barmode="overlay",
        title="案件別ガントチャート",
        height=chart_height,
        template="plotly_white",
        xaxis=dict(
            range=[axis_marks.domain_start, axis_marks.domain_end + pd.Timedelta(days=1)],
            tickmode="array",
            tickvals=list(axis_marks.tick_positions),
            ticktext=list(axis_marks.tick_labels),
            showgrid=False,
            title="日付",
        ),
        yaxis=dict(autorange="reversed", title="案件名"),
        legend=dict(title="案件名"),
        margin=dict(t=80, b=40, l=80, r=20),
    )

    for mark in axis_marks.major_marks:
        fig.add_vline(x=mark, line_color="#8899aa", line_width=1.2)

    major_set = set(axis_marks.major_marks)
    for mark in axis_marks.minor_marks:
        if mark in major_set:
            continue
        fig.add_vline(
            x=mark,
            line_color="#ccd2d9",
            line_width=0.8,
            line_dash="dot",
        )

    return fig


__all__ = ["create_project_gantt_chart"]

