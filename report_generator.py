"""Generate recurring HTML and Markdown operating reports for sold dedicated lines.

Excel and CSV sources are treated as equivalent systems of record. Derived contract
values are recalculated when source totals are unavailable, so the report remains
usable after a workbook refresh or a CSV export.
"""

from __future__ import annotations

import argparse
import csv
import html
import io
import math
import secrets
import tempfile
import warnings
from contextlib import contextmanager
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable, Mapping
from zoneinfo import ZoneInfo
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import load_workbook


REPORT_TITLE = "专线在售线路运营分析报告"
REPORT_FILENAME_PREFIX = "专线运营分析报告"
TIMEZONE = ZoneInfo("Asia/Shanghai")
AMOUNT_LABEL = "人民币（元）"
REQUIRED_FIELDS = (
    "专线ID",
    "专线名称",
    "线路平台",
    "地域",
    "带宽（Mbps）",
    "创建时间",
    "开始时间",
    "结束时间",
    "付款周期",
    "时长（月）",
    "销售单价（Mb/月）",
    "成本价（Mb/月）",
)


@dataclass(frozen=True)
class Line:
    row_number: int
    line_id: str
    name: str
    platform: str
    region: str
    bandwidth: float
    create_date: date
    start_date: date
    end_date: date
    payment_cycle: str
    term_months: float
    sale_unit_price: float
    cost_unit_price: float
    settlement_unit_price: float
    sales_total: float
    settlement_total: float
    cost_total: float

    @property
    def gross_profit(self) -> float:
        return self.sales_total - self.cost_total

    @property
    def gross_margin(self) -> float | None:
        if self.sales_total == 0:
            return None
        return self.gross_profit / self.sales_total

    @property
    def bandwidth_months(self) -> float:
        return self.bandwidth * self.term_months


def as_number(value: Any, *, default: float | None = None) -> float | None:
    if value is None or value == "":
        return default
    if isinstance(value, bool):
        return default
    try:
        number = float(value)
    except (TypeError, ValueError):
        return default
    if not math.isfinite(number):
        return default
    return number


def parse_date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%Y.%m.%d", "%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"无法解析日期：{value!r}")


def parse_as_of(value: str | date | None, now: datetime) -> date:
    if value is None:
        return now.date()
    if isinstance(value, date):
        return value
    return parse_date(value)


def text_value(value: Any) -> str:
    return "" if value is None else str(value).strip()


def numeric_from_row(row: Mapping[str, Any], field: str, fallback: float) -> float:
    value = as_number(row.get(field))
    return fallback if value is None else value


@contextmanager
def open_readable_workbook(source_path: Path):
    """Open a workbook, repairing only known malformed style nodes in a temp copy."""

    try:
        workbook = load_workbook(source_path, data_only=True, read_only=True)
    except TypeError:
        with tempfile.TemporaryDirectory(prefix="line-report-") as temp_dir:
            normalized_path = Path(temp_dir) / source_path.name
            with ZipFile(source_path, "r") as source_zip, ZipFile(
                normalized_path, "w", compression=ZIP_DEFLATED
            ) as normalized_zip:
                for item in source_zip.infolist():
                    data = source_zip.read(item.filename)
                    if item.filename == "xl/styles.xml":
                        styles = data.decode("utf-8")
                        styles = styles.replace(
                            "<fill/>",
                            '<fill><patternFill patternType="none"/></fill>',
                        )
                        data = styles.encode("utf-8")
                    normalized_zip.writestr(item, data)
            with warnings.catch_warnings():
                warnings.filterwarnings(
                    "ignore", message="Workbook contains no default style"
                )
                workbook = load_workbook(normalized_path, data_only=True, read_only=True)
            try:
                yield workbook
            finally:
                workbook.close()
        return
    try:
        yield workbook
    finally:
        workbook.close()


def read_source_rows(source_path: Path) -> tuple[list[str], list[tuple[Any, ...]]]:
    """Read headers and data rows from a supported Excel or CSV source."""

    suffix = source_path.suffix.casefold()
    if suffix == ".xlsx":
        with open_readable_workbook(source_path) as workbook:
            if not workbook.sheetnames:
                raise ValueError("Excel 文件不包含工作表")
            worksheet = workbook[workbook.sheetnames[0]]
            rows = worksheet.iter_rows(values_only=True)
            try:
                headers = [text_value(value) for value in next(rows)]
            except StopIteration as exc:
                raise ValueError("Excel 工作表为空") from exc
            return headers, [tuple(row) for row in rows]

    if suffix == ".csv":
        raw = source_path.read_bytes()
        text: str | None = None
        for encoding in ("utf-8-sig", "gb18030"):
            try:
                text = raw.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError("CSV 文件编码无法识别，请使用 UTF-8 或 GB18030")
        rows = csv.reader(io.StringIO(text, newline=""))
        try:
            headers = [text_value(value) for value in next(rows)]
        except StopIteration as exc:
            raise ValueError("CSV 文件为空") from exc
        while headers and not headers[-1]:
            headers.pop()
        return headers, [tuple(row[: len(headers)]) for row in rows]

    raise ValueError(f"不支持的输入格式：{source_path.suffix or '无扩展名'}；仅支持 .xlsx 和 .csv")


def load_lines(source_path: Path) -> tuple[list[Line], list[str]]:
    """Load valid line records and return data-quality findings separately."""

    headers, rows = read_source_rows(source_path)
    missing_headers = [field for field in REQUIRED_FIELDS if field not in headers]
    if missing_headers:
        raise ValueError(f"缺少必要字段：{', '.join(missing_headers)}")

    lines: list[Line] = []
    issues: list[str] = []
    for row_number, values in enumerate(rows, start=2):
        if not any(value not in (None, "") for value in values):
            continue
        row = dict(zip(headers, values))
        missing = [field for field in REQUIRED_FIELDS if row.get(field) in (None, "")]
        if missing:
            issues.append(f"第 {row_number} 行缺少字段：{', '.join(missing)}")
            continue
        try:
            create_date = parse_date(row["创建时间"])
            start_date = parse_date(row["开始时间"])
            end_date = parse_date(row["结束时间"])
        except ValueError as exc:
            issues.append(f"第 {row_number} 行日期错误：{exc}")
            continue

        bandwidth = as_number(row.get("带宽（Mbps）"))
        term_months = as_number(row.get("时长（月）"))
        sale_unit_price = as_number(row.get("销售单价（Mb/月）"))
        cost_unit_price = as_number(row.get("成本价（Mb/月）"))
        if None in (bandwidth, term_months, sale_unit_price, cost_unit_price):
            issues.append(f"第 {row_number} 行数值字段无法解析")
            continue
        assert bandwidth is not None
        assert term_months is not None
        assert sale_unit_price is not None
        assert cost_unit_price is not None
        if bandwidth <= 0:
            issues.append(f"第 {row_number} 行带宽不是正数：{bandwidth}")
        if term_months <= 0:
            issues.append(f"第 {row_number} 行时长不是正数：{term_months}")
        if end_date < start_date:
            issues.append(f"第 {row_number} 行结束时间早于开始时间")

        settlement_unit = numeric_from_row(
            row,
            "结算单价（Mb/月）",
            (sale_unit_price + cost_unit_price) / 2,
        )
        sales_total = numeric_from_row(
            row,
            "销售价格",
            bandwidth * term_months * sale_unit_price,
        )
        settlement_total = numeric_from_row(
            row,
            "结算价格",
            bandwidth * term_months * settlement_unit,
        )
        cost_total = numeric_from_row(
            row,
            "成本价合计",
            bandwidth * term_months * cost_unit_price,
        )
        lines.append(
            Line(
                row_number=row_number,
                line_id=text_value(row["专线ID"]),
                name=text_value(row["专线名称"]),
                platform=text_value(row["线路平台"]),
                region=text_value(row["地域"]),
                bandwidth=bandwidth,
                create_date=create_date,
                start_date=start_date,
                end_date=end_date,
                payment_cycle=text_value(row["付款周期"]),
                term_months=term_months,
                sale_unit_price=sale_unit_price,
                cost_unit_price=cost_unit_price,
                settlement_unit_price=settlement_unit,
                sales_total=sales_total,
                settlement_total=settlement_total,
                cost_total=cost_total,
            )
        )
    return lines, issues


def grouped(lines: Iterable[Line], key: str) -> list[tuple[str, list[Line]]]:
    groups: dict[str, list[Line]] = defaultdict(list)
    for line in lines:
        groups[getattr(line, key)].append(line)
    return sorted(groups.items(), key=lambda item: (-sum(x.bandwidth for x in item[1]), item[0]))


def total(lines: Iterable[Line], attr: str) -> float:
    return sum(float(getattr(line, attr)) for line in lines)


def ratio(value: float, denominator: float) -> float | None:
    return value / denominator if denominator else None


def fmt_number(value: float, decimals: int = 1) -> str:
    if abs(value - round(value)) < 1e-9:
        return f"{value:,.0f}"
    return f"{value:,.{decimals}f}"


def fmt_money(value: float | None, decimals: int = 2) -> str:
    if value is None:
        return "—"
    return f"¥{value:,.{decimals}f}"


def fmt_percent(value: float | None, decimals: int = 1) -> str:
    if value is None:
        return "—"
    return f"{value * 100:.{decimals}f}%"


def fmt_date(value: date) -> str:
    return value.strftime("%Y-%m-%d")


ACTIVE_STATUS = "在售"
RECLAIMED_STATUS = "已回收"


def line_status(line: Line, as_of: date) -> str:
    return RECLAIMED_STATUS if line.end_date < as_of else ACTIVE_STATUS


def status_metrics(lines: list[Line], as_of: date, summary: Mapping[str, Any]) -> list[dict[str, Any]]:
    total_count = len(lines)
    total_bandwidth = float(summary["total_bandwidth"])
    total_sales = float(summary["total_sales"])
    total_settlement = float(summary["total_settlement"])
    result: list[dict[str, Any]] = []
    for status in (ACTIVE_STATUS, RECLAIMED_STATUS):
        items = [line for line in lines if line_status(line, as_of) == status]
        bandwidth = total(items, "bandwidth")
        sales = total(items, "sales_total")
        cost = total(items, "cost_total")
        settlement = total(items, "settlement_total")
        result.append(
            {
                "status": status,
                "line_count": len(items),
                "count_share": ratio(len(items), total_count),
                "bandwidth": bandwidth,
                "bandwidth_share": ratio(bandwidth, total_bandwidth),
                "sales": sales,
                "sales_share": ratio(sales, total_sales),
                "settlement": settlement,
                "settlement_share": ratio(settlement, total_settlement),
                "cost": cost,
                "profit": sales - cost,
                "margin": ratio(sales - cost, sales),
                "settlement_surplus": settlement - cost,
                "settlement_margin": ratio(settlement - cost, settlement),
            }
        )
    return result


def summarize(lines: list[Line], as_of: date, issues: list[str]) -> dict[str, Any]:
    total_bandwidth = total(lines, "bandwidth")
    total_sales = total(lines, "sales_total")
    total_settlement = total(lines, "settlement_total")
    total_cost = total(lines, "cost_total")
    total_bandwidth_months = total(lines, "bandwidth_months")
    total_profit = total_sales - total_cost
    settlement_surplus = total_settlement - total_cost
    weighted_sale_unit = ratio(
        sum(line.bandwidth * line.sale_unit_price for line in lines), total_bandwidth
    )
    weighted_cost_unit = ratio(
        sum(line.bandwidth * line.cost_unit_price for line in lines), total_bandwidth
    )
    contract_monthly_unit = ratio(total_sales, total_bandwidth_months)
    expiring = sorted(
        [
            (line, (line.end_date - as_of).days)
            for line in lines
            if 0 <= (line.end_date - as_of).days <= 7
        ],
        key=lambda item: (item[1], item[0].end_date, item[0].line_id),
    )
    expired = sorted(
        [line for line in lines if line.end_date < as_of],
        key=lambda line: (line.end_date, line.line_id),
    )
    summary = {
        "line_count": len(lines),
        "total_bandwidth": total_bandwidth,
        "total_sales": total_sales,
        "total_settlement": total_settlement,
        "total_cost": total_cost,
        "total_profit": total_profit,
        "gross_margin": ratio(total_profit, total_sales),
        "settlement_surplus": settlement_surplus,
        "settlement_margin": ratio(settlement_surplus, total_settlement),
        "weighted_sale_unit": weighted_sale_unit,
        "weighted_cost_unit": weighted_cost_unit,
        "weighted_margin_unit": None
        if weighted_sale_unit is None or weighted_cost_unit is None
        else weighted_sale_unit - weighted_cost_unit,
        "total_bandwidth_months": total_bandwidth_months,
        "contract_monthly_unit": contract_monthly_unit,
        "expiring": expiring,
        "expired": expired,
        "margin_risks": sorted(
            [line for line in lines if line.gross_profit <= 0],
            key=lambda line: (line.gross_profit, line.line_id),
        ),
        "issues": issues,
    }
    summary["status_rows"] = status_metrics(lines, as_of, summary)
    summary["active_line_count"] = next(
        row["line_count"] for row in summary["status_rows"] if row["status"] == ACTIVE_STATUS
    )
    summary["reclaimed_line_count"] = next(
        row["line_count"] for row in summary["status_rows"] if row["status"] == RECLAIMED_STATUS
    )
    summary["active_bandwidth"] = next(
        row["bandwidth"] for row in summary["status_rows"] if row["status"] == ACTIVE_STATUS
    )
    summary["reclaimed_bandwidth"] = next(
        row["bandwidth"] for row in summary["status_rows"] if row["status"] == RECLAIMED_STATUS
    )
    return summary


def group_metrics(
    lines: list[Line], key: str, summary: Mapping[str, Any], as_of: date
) -> list[dict[str, Any]]:
    total_bandwidth = float(summary["total_bandwidth"])
    total_sales = float(summary["total_sales"])
    total_settlement = float(summary["total_settlement"])
    result: list[dict[str, Any]] = []
    for name, items in grouped(lines, key):
        bandwidth = total(items, "bandwidth")
        sales = total(items, "sales_total")
        cost = total(items, "cost_total")
        settlement = total(items, "settlement_total")
        active_items = [line for line in items if line_status(line, as_of) == ACTIVE_STATUS]
        reclaimed_items = [line for line in items if line_status(line, as_of) == RECLAIMED_STATUS]
        result.append(
            {
                "name": name or "未填写",
                "line_count": len(items),
                "bandwidth": bandwidth,
                "bandwidth_share": ratio(bandwidth, total_bandwidth),
                "active_line_count": len(active_items),
                "active_bandwidth": total(active_items, "bandwidth"),
                "reclaimed_line_count": len(reclaimed_items),
                "reclaimed_bandwidth": total(reclaimed_items, "bandwidth"),
                "sales": sales,
                "sales_share": ratio(sales, total_sales),
                "settlement": settlement,
                "settlement_share": ratio(settlement, total_settlement),
                "profit": sales - cost,
                "margin": ratio(sales - cost, sales),
                "settlement_surplus": settlement - cost,
                "settlement_margin": ratio(settlement - cost, settlement),
                "weighted_sale_unit": ratio(
                    sum(line.bandwidth * line.sale_unit_price for line in items), bandwidth
                ),
            }
        )
    return result


def platform_region_metrics(
    lines: list[Line], platform: str, as_of: date
) -> list[dict[str, Any]]:
    """Aggregate region metrics within one platform.

    Shares are calculated against the selected platform, not against all
    platforms, so the UC regional mix can be read independently.
    """

    platform_lines = [line for line in lines if line.platform.casefold() == platform.casefold()]
    total_count = len(platform_lines)
    total_bandwidth = total(platform_lines, "bandwidth")
    total_sales = total(platform_lines, "sales_total")
    total_settlement = total(platform_lines, "settlement_total")
    result: list[dict[str, Any]] = []
    for name, items in grouped(platform_lines, "region"):
        bandwidth = total(items, "bandwidth")
        sales = total(items, "sales_total")
        settlement = total(items, "settlement_total")
        cost = total(items, "cost_total")
        active_items = [line for line in items if line_status(line, as_of) == ACTIVE_STATUS]
        reclaimed_items = [line for line in items if line_status(line, as_of) == RECLAIMED_STATUS]
        result.append(
            {
                "name": name or "未填写",
                "line_count": len(items),
                "count_share": ratio(len(items), total_count),
                "bandwidth": bandwidth,
                "bandwidth_share": ratio(bandwidth, total_bandwidth),
                "active_line_count": len(active_items),
                "active_bandwidth": total(active_items, "bandwidth"),
                "reclaimed_line_count": len(reclaimed_items),
                "reclaimed_bandwidth": total(reclaimed_items, "bandwidth"),
                "sales": sales,
                "sales_share": ratio(sales, total_sales),
                "settlement": settlement,
                "settlement_share": ratio(settlement, total_settlement),
                "profit": sales - cost,
                "margin": ratio(sales - cost, sales),
            }
        )
    return result


def creation_metrics(lines: list[Line], as_of: date) -> list[dict[str, Any]]:
    groups: dict[str, dict[str, Any]] = defaultdict(
        lambda: {
            "month": "",
            "line_count": 0,
            "bandwidth": 0.0,
            "settlement": 0.0,
            "active_line_count": 0,
            "active_bandwidth": 0.0,
            "active_settlement": 0.0,
            "reclaimed_line_count": 0,
            "reclaimed_bandwidth": 0.0,
            "reclaimed_settlement": 0.0,
        }
    )
    for line in lines:
        month = line.create_date.strftime("%Y-%m")
        groups[month]["month"] = month
        groups[month]["line_count"] += 1
        groups[month]["bandwidth"] += line.bandwidth
        groups[month]["settlement"] += line.settlement_total
        if line_status(line, as_of) == ACTIVE_STATUS:
            groups[month]["active_line_count"] += 1
            groups[month]["active_bandwidth"] += line.bandwidth
            groups[month]["active_settlement"] += line.settlement_total
        else:
            groups[month]["reclaimed_line_count"] += 1
            groups[month]["reclaimed_bandwidth"] += line.bandwidth
            groups[month]["reclaimed_settlement"] += line.settlement_total
    return [groups[month] for month in sorted(groups)]


def add_months(source_date: date, offset: int) -> date:
    """Return the first day of the month ``offset`` months after ``source_date``."""

    month_index = source_date.year * 12 + source_date.month - 1 + offset
    year, month_index = divmod(month_index, 12)
    return date(year, month_index + 1, 1)


def monthly_active_bandwidth_metrics(
    lines: list[Line], as_of: date
) -> list[dict[str, Any]]:
    """Return active bandwidth snapshots for each month through ``as_of``.

    Completed months use their final calendar day as the snapshot date. The
    current month uses ``as_of`` so future starts and already ended lines are not
    counted in that month's total.
    """

    started_lines = [line for line in lines if line.start_date <= as_of]
    if not started_lines:
        return []

    first_month = min(line.start_date for line in started_lines).replace(day=1)
    last_month = as_of.replace(day=1)
    rows: list[dict[str, Any]] = []
    month_start = first_month
    while month_start <= last_month:
        natural_month_end = add_months(month_start, 1) - timedelta(days=1)
        snapshot_date = min(natural_month_end, as_of)
        active_lines = [
            line
            for line in started_lines
            if line.start_date <= snapshot_date <= line.end_date
        ]
        rows.append(
            {
                "month": month_start.strftime("%Y-%m"),
                "active_line_count": len(active_lines),
                "active_bandwidth": total(active_lines, "bandwidth"),
            }
        )
        month_start = add_months(month_start, 1)
    return rows


def monthly_reclaimed_metrics(lines: list[Line], as_of: date) -> list[dict[str, Any]]:
    """Aggregate reclaimed line count and bandwidth by end-date month."""

    reclaimed_lines = [line for line in lines if line.end_date < as_of]
    if not reclaimed_lines:
        return []

    first_month = min(line.end_date for line in reclaimed_lines).replace(day=1)
    last_month = as_of.replace(day=1)
    rows: list[dict[str, Any]] = []
    month_start = first_month
    while month_start <= last_month:
        next_month = add_months(month_start, 1)
        items = [
            line
            for line in reclaimed_lines
            if month_start <= line.end_date < next_month
        ]
        rows.append(
            {
                "month": month_start.strftime("%Y-%m"),
                "reclaimed_line_count": len(items),
                "reclaimed_bandwidth": total(items, "bandwidth"),
            }
        )
        month_start = next_month
    return rows


def service_monthly_sales_metrics(
    lines: list[Line], through_date: date | None = None
) -> list[dict[str, Any]]:
    """Allocate monthly contract and settlement sales across service months.

    A line starts contributing in the month containing ``start_date`` and
    contributes its monthly equivalent for every full contracted month. If a
    line spans only part of a calendar month, the final contribution is
    prorated so that all source revenue is allocated exactly once. This keeps
    quarterly, half-year and annual lines visible in each month they cover.
    When ``through_date`` is supplied, only months through that date's
    calendar month are returned for report display.
    """

    groups: dict[str, dict[str, Any]] = defaultdict(
        lambda: {
            "month": "",
            "contract_monthly_sales": 0.0,
            "settlement_monthly_sales": 0.0,
            "line_count": 0,
        }
    )
    for line in lines:
        contract_monthly = line.sales_total / line.term_months
        settlement_monthly = line.settlement_total / line.term_months
        month_count = max(1, math.ceil(line.term_months))
        for offset in range(month_count):
            month = add_months(line.start_date, offset).strftime("%Y-%m")
            if through_date is not None and month > through_date.strftime("%Y-%m"):
                continue
            months_in_period = 1.0
            if offset == month_count - 1 and line.term_months < month_count:
                months_in_period = line.term_months - (month_count - 1)
            groups[month]["month"] = month
            groups[month]["contract_monthly_sales"] += contract_monthly * months_in_period
            groups[month]["settlement_monthly_sales"] += settlement_monthly * months_in_period
            groups[month]["line_count"] += 1
    return [groups[month] for month in sorted(groups)]


PAYMENT_CYCLE_MONTHS = {"月": 1, "季度": 3, "半年": 6, "年": 12}


def payment_cycle_months(line: Line) -> int:
    """Return the renewal duration in months for a line's payment cycle."""

    return PAYMENT_CYCLE_MONTHS.get(line.payment_cycle, max(1, math.ceil(line.term_months)))


def forecast_monthly_sales_metrics(
    lines: list[Line], as_of: date, horizon: int = 6
) -> list[dict[str, Any]]:
    """Forecast monthly sales for active lines with automatic renewals.

    The forecast begins in the calendar month after ``as_of``. Active lines
    keep their current monthly contract and settlement equivalents. When the
    current term ends, the line is extended repeatedly by its payment-cycle
    duration, so monthly sales continue without a gap.
    """

    if horizon <= 0:
        return []
    active_lines = [line for line in lines if line.end_date >= as_of]
    groups: dict[str, dict[str, Any]] = defaultdict(
        lambda: {
            "month": "",
            "line_count": 0,
            "bandwidth": 0.0,
            "contract_monthly_sales": 0.0,
            "settlement_monthly_sales": 0.0,
        }
    )
    forecast_months = [
        add_months(date(as_of.year, as_of.month, 1), offset)
        for offset in range(1, horizon + 1)
    ]
    for line in active_lines:
        renewal_months = payment_cycle_months(line)
        coverage_end = line.end_date
        for month_start in forecast_months:
            month_end = add_months(month_start, 1) - timedelta(days=1)
            while coverage_end < month_end:
                renewal_start = coverage_end + timedelta(days=1)
                coverage_end = add_months(renewal_start, renewal_months) - timedelta(days=1)
            month = month_start.strftime("%Y-%m")
            groups[month]["month"] = month
            groups[month]["line_count"] += 1
            groups[month]["bandwidth"] += line.bandwidth
            groups[month]["contract_monthly_sales"] += line.sales_total / line.term_months
            groups[month]["settlement_monthly_sales"] += line.settlement_total / line.term_months
    return [groups[month] for month in sorted(groups)]


def payment_metrics(lines: list[Line], summary: Mapping[str, Any], as_of: date) -> list[dict[str, Any]]:
    total_bandwidth = float(summary["total_bandwidth"])
    total_sales = float(summary["total_sales"])
    total_settlement = float(summary["total_settlement"])
    total_count = len(lines)
    result: list[dict[str, Any]] = []
    for name, items in sorted(grouped(lines, "payment_cycle"), key=lambda item: item[0]):
        bandwidth = total(items, "bandwidth")
        sales = total(items, "sales_total")
        settlement = total(items, "settlement_total")
        active_items = [line for line in items if line_status(line, as_of) == ACTIVE_STATUS]
        reclaimed_items = [line for line in items if line_status(line, as_of) == RECLAIMED_STATUS]
        result.append(
            {
                "name": name or "未填写",
                "line_count": len(items),
                "count_share": ratio(len(items), total_count),
                "bandwidth": bandwidth,
                "bandwidth_share": ratio(bandwidth, total_bandwidth),
                "active_line_count": len(active_items),
                "active_bandwidth": total(active_items, "bandwidth"),
                "reclaimed_line_count": len(reclaimed_items),
                "reclaimed_bandwidth": total(reclaimed_items, "bandwidth"),
                "sales": sales,
                "sales_share": ratio(sales, total_sales),
                "settlement": settlement,
                "settlement_share": ratio(settlement, total_settlement),
                "avg_term": ratio(total(items, "term_months"), len(items)),
            }
        )
    return result


def html_table(headers: list[str], rows: list[list[str]], class_name: str = "") -> str:
    header_html = "".join(f"<th>{html.escape(header)}</th>" for header in headers)
    body = []
    for row in rows:
        cells = "".join(f"<td>{html.escape(str(cell))}</td>" for cell in row)
        body.append(f"<tr>{cells}</tr>")
    cls = f' class="{html.escape(class_name)}"' if class_name else ""
    return f'<table{cls}><thead><tr>{header_html}</tr></thead><tbody>{"".join(body)}</tbody></table>'


def markdown_table(headers: list[str], rows: list[list[str]]) -> str:
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    for row in rows:
        safe = [str(cell).replace("|", "\\|").replace("\n", " ") for cell in row]
        lines.append("| " + " | ".join(safe) + " |")
    return "\n".join(lines)


def platform_pie_svg(platform_rows: list[dict[str, Any]]) -> str:
    """Render an inline SVG pie chart using absolute platform bandwidth totals."""

    total_bandwidth = sum(float(row["bandwidth"]) for row in platform_rows)
    if not platform_rows or total_bandwidth <= 0:
        return '<p class="muted">暂无平台带宽数据。</p>'

    width, height = 520, 290
    cx, cy, radius = 145, 145, 108
    colors = ["#2b75b3", "#f59e0b", "#22a06b", "#8b5cf6", "#ef4444", "#06b6d4"]
    paths: list[str] = []
    legend: list[str] = []
    start_angle = -math.pi / 2
    for index, row in enumerate(platform_rows):
        bandwidth = float(row["bandwidth"])
        share = bandwidth / total_bandwidth
        end_angle = start_angle + share * 2 * math.pi
        color = colors[index % len(colors)]
        if share >= 0.999999:
            paths.append(f'<circle cx="{cx}" cy="{cy}" r="{radius}" fill="{color}"/>')
        else:
            start_x = cx + radius * math.cos(start_angle)
            start_y = cy + radius * math.sin(start_angle)
            end_x = cx + radius * math.cos(end_angle)
            end_y = cy + radius * math.sin(end_angle)
            large_arc = 1 if share > 0.5 else 0
            paths.append(
                f'<path d="M {cx} {cy} L {start_x:.2f} {start_y:.2f} '
                f'A {radius} {radius} 0 {large_arc} 1 {end_x:.2f} {end_y:.2f} Z" '
                f'fill="{color}"/>'
            )
        name = html.escape(str(row["name"]))
        bandwidth_text = html.escape(fmt_number(bandwidth))
        share_text = html.escape(fmt_percent(row["bandwidth_share"]))
        legend_y = 72 + index * 34
        legend.append(
            f'<rect x="315" y="{legend_y - 13}" width="14" height="14" rx="3" fill="{color}"/>'
            f'<text x="340" y="{legend_y}" class="pie-label">{name}：{bandwidth_text} Mbps（{share_text}）</text>'
        )
        start_angle = end_angle

    return (
        f'<svg class="platform-pie" viewBox="0 0 {width} {height}" role="img" '
        f'aria-label="各线路平台售出带宽占比饼图">'
        f'<title>各线路平台售出带宽占比</title>{"".join(paths)}'
        f'<text x="145" y="140" text-anchor="middle" class="pie-total">{fmt_number(total_bandwidth)}</text>'
        f'<text x="145" y="163" text-anchor="middle" class="pie-unit">Mbps 总带宽</text>'
        f'{"".join(legend)}</svg>'
    )


def creation_bar_svg(monthly_rows: list[dict[str, Any]]) -> str:
    """Render monthly created-line count and bandwidth as grouped bars."""

    if not monthly_rows:
        return '<p class="muted">暂无创建时间数据。</p>'

    width, height = 860, 390
    left, right, top, bottom = 62, 66, 58, 70
    plot_width = width - left - right
    plot_height = height - top - bottom
    max_count = max(1, max(row["line_count"] for row in monthly_rows))
    max_bandwidth = max(1, max(row["bandwidth"] for row in monthly_rows))
    count_axis_max = max(1, math.ceil(max_count / 5) * 5)
    bandwidth_axis_max = max(1, math.ceil(max_bandwidth / 10) * 10)
    group_width = plot_width / len(monthly_rows)
    bar_width = min(24, max(12, group_width * 0.22))
    count_color = "#2b75b3"
    bandwidth_color = "#f59e0b"
    svg: list[str] = [
        f'<svg class="creation-bars" viewBox="0 0 {width} {height}" role="img" '
        f'aria-label="按创建时间统计每月新增线路数量和带宽量柱状图">',
        "<title>每月新增线路数量和带宽量</title>",
        f'<text x="{left}" y="24" class="chart-axis-title">线路数（左轴）</text>',
        f'<text x="{width - right}" y="24" text-anchor="end" class="chart-axis-title">带宽 Mbps（右轴）</text>',
    ]
    for tick in range(0, 5):
        fraction = tick / 4
        y = top + plot_height - fraction * plot_height
        count_value = count_axis_max * fraction
        bandwidth_value = bandwidth_axis_max * fraction
        svg.append(
            f'<line x1="{left}" y1="{y:.2f}" x2="{width - right}" y2="{y:.2f}" class="chart-gridline"/>'
            f'<text x="{left - 10}" y="{y + 4:.2f}" text-anchor="end" class="chart-tick">{fmt_number(count_value)}</text>'
            f'<text x="{width - right + 10}" y="{y + 4:.2f}" class="chart-tick">{fmt_number(bandwidth_value)}</text>'
        )
    for index, row in enumerate(monthly_rows):
        center_x = left + group_width * (index + 0.5)
        count_height = row["line_count"] / count_axis_max * plot_height
        bandwidth_height = row["bandwidth"] / bandwidth_axis_max * plot_height
        count_x = center_x - bar_width - 3
        bandwidth_x = center_x + 3
        count_y = top + plot_height - count_height
        bandwidth_y = top + plot_height - bandwidth_height
        svg.extend(
            [
                f'<rect x="{count_x:.2f}" y="{count_y:.2f}" width="{bar_width:.2f}" height="{count_height:.2f}" rx="3" fill="{count_color}"/>',
                f'<rect x="{bandwidth_x:.2f}" y="{bandwidth_y:.2f}" width="{bar_width:.2f}" height="{bandwidth_height:.2f}" rx="3" fill="{bandwidth_color}"/>',
                f'<text x="{count_x + bar_width / 2:.2f}" y="{max(top + 14, count_y - 7):.2f}" text-anchor="middle" class="chart-value">{row["line_count"]}</text>',
                f'<text x="{bandwidth_x + bar_width / 2:.2f}" y="{max(top + 14, bandwidth_y - 7):.2f}" text-anchor="middle" class="chart-value">{fmt_number(row["bandwidth"])}</text>',
                f'<text x="{center_x:.2f}" y="{height - bottom + 25}" text-anchor="middle" class="chart-month">{html.escape(row["month"])}</text>',
            ]
        )
    legend_y = height - 20
    svg.extend(
        [
            f'<rect x="{left}" y="{legend_y - 12}" width="14" height="14" rx="3" fill="{count_color}"/><text x="{left + 22}" y="{legend_y}" class="chart-legend">新增线路数</text>',
            f'<rect x="{left + 150}" y="{legend_y - 12}" width="14" height="14" rx="3" fill="{bandwidth_color}"/><text x="{left + 172}" y="{legend_y}" class="chart-legend">新增带宽（Mbps）</text>',
            "</svg>",
        ]
    )
    return "".join(svg)


def active_bandwidth_bar_svg(
    monthly_rows: list[dict[str, Any]],
    *,
    count_key: str = "active_line_count",
    bandwidth_key: str = "active_bandwidth",
    chart_class: str = "active-statistics-bars",
    chart_title: str = "每月在售线路数量和总带宽",
    aria_label: str = "每月在售线路数量和总带宽柱状图",
    empty_text: str = "暂无月度在售数据。",
    count_legend: str = "在售线路数",
    bandwidth_legend: str = "在售总带宽（Mbps）",
) -> str:
    """Render monthly line count and bandwidth as grouped bars."""

    if not monthly_rows:
        return f'<p class="muted">{html.escape(empty_text)}</p>'

    width, height = 860, 390
    left, right, top, bottom = 62, 66, 58, 70
    plot_width = width - left - right
    plot_height = height - top - bottom
    max_count = max(1, max(row[count_key] for row in monthly_rows))
    max_bandwidth = max(1, max(row[bandwidth_key] for row in monthly_rows))
    count_axis_max = max(1, math.ceil(max_count / 5) * 5)
    bandwidth_axis_max = max(10, math.ceil(max_bandwidth / 10) * 10)
    group_width = plot_width / len(monthly_rows)
    bar_width = min(24, max(12, group_width * 0.22))
    count_color = "#2b75b3"
    bandwidth_color = "#f59e0b"
    svg: list[str] = [
        f'<svg class="{html.escape(chart_class)}" viewBox="0 0 {width} {height}" role="img" '
        f'aria-label="{html.escape(aria_label)}">',
        f"<title>{html.escape(chart_title)}</title>",
        f'<text x="{left}" y="24" class="chart-axis-title">线路数（左轴）</text>',
        f'<text x="{width - right}" y="24" text-anchor="end" class="chart-axis-title">带宽 Mbps（右轴）</text>',
    ]
    for tick in range(0, 5):
        fraction = tick / 4
        y = top + plot_height - fraction * plot_height
        count_value = count_axis_max * fraction
        bandwidth_value = bandwidth_axis_max * fraction
        svg.append(
            f'<line x1="{left}" y1="{y:.2f}" x2="{width - right}" y2="{y:.2f}" class="chart-gridline"/>'
            f'<text x="{left - 10}" y="{y + 4:.2f}" text-anchor="end" class="chart-tick">{fmt_number(count_value)}</text>'
            f'<text x="{width - right + 10}" y="{y + 4:.2f}" class="chart-tick">{fmt_number(bandwidth_value)}</text>'
        )
    for index, row in enumerate(monthly_rows):
        center_x = left + group_width * (index + 0.5)
        count_height = row[count_key] / count_axis_max * plot_height
        bandwidth_height = row[bandwidth_key] / bandwidth_axis_max * plot_height
        count_x = center_x - bar_width - 3
        bandwidth_x = center_x + 3
        count_y = top + plot_height - count_height
        bandwidth_y = top + plot_height - bandwidth_height
        svg.extend(
            [
                f'<rect x="{count_x:.2f}" y="{count_y:.2f}" width="{bar_width:.2f}" height="{count_height:.2f}" rx="3" fill="{count_color}"/>',
                f'<rect x="{bandwidth_x:.2f}" y="{bandwidth_y:.2f}" width="{bar_width:.2f}" height="{bandwidth_height:.2f}" rx="3" fill="{bandwidth_color}"/>',
                f'<text x="{count_x + bar_width / 2:.2f}" y="{max(top + 14, count_y - 7):.2f}" text-anchor="middle" class="chart-value">{row[count_key]}</text>',
                f'<text x="{bandwidth_x + bar_width / 2:.2f}" y="{max(top + 14, bandwidth_y - 7):.2f}" text-anchor="middle" class="chart-value">{fmt_number(row[bandwidth_key])}</text>',
                f'<text x="{center_x:.2f}" y="{height - bottom + 25}" text-anchor="middle" class="chart-month">{html.escape(row["month"])}</text>',
            ]
        )
    legend_y = height - 20
    svg.extend(
        [
            f'<rect x="{left}" y="{legend_y - 12}" width="14" height="14" rx="3" fill="{count_color}"/><text x="{left + 22}" y="{legend_y}" class="chart-legend">{html.escape(count_legend)}</text>',
            f'<rect x="{left + 150}" y="{legend_y - 12}" width="14" height="14" rx="3" fill="{bandwidth_color}"/><text x="{left + 172}" y="{legend_y}" class="chart-legend">{html.escape(bandwidth_legend)}</text>',
            "</svg>",
        ]
    )
    return "".join(svg)


def monthly_sales_bar_svg(
    monthly_rows: list[dict[str, Any]],
    *,
    chart_class: str = "monthly-sales-bars",
    chart_title: str = "每月合同月销与结算月销",
    aria_label: str = "按服务月份统计合同月销和结算月销柱状图",
) -> str:
    """Render monthly contract and settlement sales as grouped bars."""

    if not monthly_rows:
        return '<p class="muted">暂无月销金额数据。</p>'

    width, height = 860, 390
    left, right, top, bottom = 72, 36, 58, 70
    plot_width = width - left - right
    plot_height = height - top - bottom
    max_value = max(
        1,
        max(
            max(row["contract_monthly_sales"], row["settlement_monthly_sales"])
            for row in monthly_rows
        ),
    )
    axis_max = max(100, math.ceil(max_value / 1000) * 1000)
    group_width = plot_width / len(monthly_rows)
    bar_width = min(28, max(12, group_width * 0.22))
    contract_color = "#2b75b3"
    settlement_color = "#f59e0b"
    svg: list[str] = [
        f'<svg class="{html.escape(chart_class)}" viewBox="0 0 {width} {height}" role="img" '
        f'aria-label="{html.escape(aria_label)}">',
        f"<title>{html.escape(chart_title)}</title>",
        f'<text x="{left}" y="24" class="chart-axis-title">金额（人民币元）</text>',
    ]
    for tick in range(0, 5):
        fraction = tick / 4
        y = top + plot_height - fraction * plot_height
        value = axis_max * fraction
        svg.append(
            f'<line x1="{left}" y1="{y:.2f}" x2="{width - right}" y2="{y:.2f}" class="chart-gridline"/>'
            f'<text x="{left - 10}" y="{y + 4:.2f}" text-anchor="end" class="chart-tick">{fmt_money(value, 0)}</text>'
        )
    for index, row in enumerate(monthly_rows):
        center_x = left + group_width * (index + 0.5)
        contract_height = row["contract_monthly_sales"] / axis_max * plot_height
        settlement_height = row["settlement_monthly_sales"] / axis_max * plot_height
        contract_x = center_x - bar_width - 3
        settlement_x = center_x + 3
        contract_y = top + plot_height - contract_height
        settlement_y = top + plot_height - settlement_height
        contract_label_y = contract_y - 7
        settlement_label_y = settlement_y - 7
        if abs(contract_label_y - settlement_label_y) < 18:
            if contract_label_y <= settlement_label_y:
                contract_label_y -= 18
            else:
                settlement_label_y -= 18
        contract_label_y = max(top + 14, contract_label_y)
        settlement_label_y = max(top + 14, settlement_label_y)
        svg.extend(
            [
                f'<rect x="{contract_x:.2f}" y="{contract_y:.2f}" width="{bar_width:.2f}" height="{contract_height:.2f}" rx="3" fill="{contract_color}"/>',
                f'<rect x="{settlement_x:.2f}" y="{settlement_y:.2f}" width="{bar_width:.2f}" height="{settlement_height:.2f}" rx="3" fill="{settlement_color}"/>',
                f'<text x="{contract_x + bar_width / 2:.2f}" y="{contract_label_y:.2f}" text-anchor="middle" class="chart-value">{fmt_money(row["contract_monthly_sales"], 0)}</text>',
                f'<text x="{settlement_x + bar_width / 2:.2f}" y="{settlement_label_y:.2f}" text-anchor="middle" class="chart-value">{fmt_money(row["settlement_monthly_sales"], 0)}</text>',
                f'<text x="{center_x:.2f}" y="{height - bottom + 25}" text-anchor="middle" class="chart-month">{html.escape(row["month"])}</text>',
            ]
        )
    legend_y = height - 20
    svg.extend(
        [
            f'<rect x="{left}" y="{legend_y - 12}" width="14" height="14" rx="3" fill="{contract_color}"/><text x="{left + 22}" y="{legend_y}" class="chart-legend">合同月销</text>',
            f'<rect x="{left + 130}" y="{legend_y - 12}" width="14" height="14" rx="3" fill="{settlement_color}"/><text x="{left + 152}" y="{legend_y}" class="chart-legend">结算月销（实际收入）</text>',
            "</svg>",
        ]
    )
    return "".join(svg)


def render_report(
    lines: list[Line],
    summary: Mapping[str, Any],
    platform_rows: list[dict[str, Any]],
    monthly_rows: list[dict[str, Any]],
    active_bandwidth_rows: list[dict[str, Any]],
    reclaimed_rows: list[dict[str, Any]],
    monthly_sales_rows: list[dict[str, Any]],
    forecast_rows: list[dict[str, Any]],
    region_rows: list[dict[str, Any]],
    uc_region_rows: list[dict[str, Any]],
    payment_rows: list[dict[str, Any]],
    *,
    source_name: str,
    generated_at: datetime,
    as_of: date,
    report_filename: str,
) -> tuple[str, str]:
    expiring: list[tuple[Line, int]] = list(summary["expiring"])
    margin_risks: list[Line] = list(summary["margin_risks"])
    expired: list[Line] = list(summary["expired"])
    issues: list[str] = list(summary["issues"])
    status_rows: list[dict[str, Any]] = list(summary["status_rows"])
    expiring_bandwidth = total((line for line, _ in expiring), "bandwidth") if expiring else 0
    expiring_sales = total((line for line, _ in expiring), "sales_total") if expiring else 0

    urgent_count = sum(1 for _, days in expiring if days <= 3)
    route_leader = region_rows[0] if region_rows else None
    platform_leader = platform_rows[0] if platform_rows else None
    observations = [
        f"平台结构以 {platform_leader['name']} 为主，占总带宽 {fmt_number(platform_leader['bandwidth'])} Mbps（{fmt_percent(platform_leader['bandwidth_share'])}）。"
        if platform_leader
        else "暂无平台结构数据。",
        f"地域结构中 {route_leader['name']} 带宽最高，占比 {fmt_percent(route_leader['bandwidth_share'])}。"
        if route_leader
        else "暂无地域结构数据。",
        f"未来 7 天有 {len(expiring)} 条线路到期，其中 {urgent_count} 条在 3 天内到期。",
        f"共有 {len(margin_risks)} 条线路毛利小于或等于 0，需复核定价或成本。",
    ]

    kpis = [
        ("全量记录线路", f"{summary['line_count']:,}", "条"),
        ("当前在售线路", f"{summary['active_line_count']:,}", "条"),
        ("已回收线路", f"{summary['reclaimed_line_count']:,}", "条"),
        ("全量记录带宽", f"{fmt_number(summary['total_bandwidth'])}", "Mbps"),
        ("在售总带宽", f"{fmt_number(summary['active_bandwidth'])}", "Mbps"),
        ("合同销售额", fmt_money(summary["total_sales"], 0), AMOUNT_LABEL),
        ("实际收入（结算价格）", fmt_money(summary["total_settlement"], 0), AMOUNT_LABEL),
        ("合同毛利", fmt_money(summary["total_profit"], 0), AMOUNT_LABEL),
        ("毛利率", fmt_percent(summary["gross_margin"]), ""),
        ("加权平均销售单价", fmt_money(summary["weighted_sale_unit"]), "/ Mbps / 月"),
    ]

    expiry_rows = [
        [
            line.line_id,
            line.name,
            f"{days} 天",
            fmt_date(line.end_date),
            line.platform,
            line.region,
            fmt_number(line.bandwidth),
            fmt_money(line.sales_total, 0),
            fmt_money(line.settlement_total, 0),
        ]
        for line, days in expiring
    ]
    risk_rows = [
        [
            line.line_id,
            line.name,
            line_status(line, as_of),
            line.platform,
            fmt_number(line.bandwidth),
            fmt_money(line.sale_unit_price),
            fmt_money(line.cost_unit_price),
            fmt_money(line.settlement_total, 0),
            fmt_money(line.gross_profit, 0),
            fmt_percent(line.gross_margin),
        ]
        for line in margin_risks
    ]
    platform_table_rows = [
        [
            row["name"],
            str(row["line_count"]),
            fmt_number(row["bandwidth"]),
            fmt_percent(row["bandwidth_share"]),
            str(row["active_line_count"]),
            fmt_number(row["active_bandwidth"]),
            str(row["reclaimed_line_count"]),
            fmt_number(row["reclaimed_bandwidth"]),
            fmt_money(row["sales"], 0),
            fmt_percent(row["sales_share"]),
            fmt_money(row["settlement"], 0),
            fmt_percent(row["settlement_share"]),
            fmt_money(row["weighted_sale_unit"]),
            fmt_percent(row["margin"]),
        ]
        for row in platform_rows
    ]
    region_table_rows = [
        [
            row["name"],
            str(row["line_count"]),
            fmt_number(row["bandwidth"]),
            fmt_percent(row["bandwidth_share"]),
            str(row["active_line_count"]),
            fmt_number(row["active_bandwidth"]),
            str(row["reclaimed_line_count"]),
            fmt_number(row["reclaimed_bandwidth"]),
            fmt_money(row["sales"], 0),
            fmt_percent(row["sales_share"]),
            fmt_money(row["settlement"], 0),
            fmt_percent(row["settlement_share"]),
            fmt_percent(row["margin"]),
        ]
        for row in region_rows
    ]
    payment_table_rows = [
        [
            row["name"],
            str(row["line_count"]),
            fmt_percent(row["count_share"]),
            fmt_number(row["bandwidth"]),
            fmt_percent(row["bandwidth_share"]),
            str(row["active_line_count"]),
            fmt_number(row["active_bandwidth"]),
            str(row["reclaimed_line_count"]),
            fmt_number(row["reclaimed_bandwidth"]),
            fmt_money(row["sales"], 0),
            fmt_percent(row["sales_share"]),
            fmt_money(row["settlement"], 0),
            fmt_percent(row["settlement_share"]),
            fmt_number(row["avg_term"], 1),
        ]
        for row in payment_rows
    ]
    monthly_table_rows = [
        [
            row["month"],
            str(row["line_count"]),
            fmt_number(row["bandwidth"]),
            fmt_money(row["settlement"], 0),
            str(row["active_line_count"]),
            fmt_number(row["active_bandwidth"]),
            fmt_money(row["active_settlement"], 0),
            str(row["reclaimed_line_count"]),
            fmt_number(row["reclaimed_bandwidth"]),
            fmt_money(row["reclaimed_settlement"], 0),
        ]
        for row in monthly_rows
    ]
    reclaimed_creation_detail_rows = [
        [
            line.create_date.strftime("%Y-%m"),
            line.line_id,
            line.name,
            line.platform,
            line.region,
            fmt_number(line.bandwidth),
            fmt_date(line.create_date),
            fmt_date(line.start_date),
            fmt_date(line.end_date),
            fmt_money(line.sales_total, 0),
            fmt_money(line.settlement_total, 0),
        ]
        for line in sorted(
            (line for line in lines if line.end_date < as_of),
            key=lambda line: (line.create_date, line.end_date, line.line_id),
        )
    ]
    active_bandwidth_table_rows = [
        [
            row["month"],
            str(row["active_line_count"]),
            fmt_number(row["active_bandwidth"]),
        ]
        for row in active_bandwidth_rows
    ]
    reclaimed_table_rows = [
        [
            row["month"],
            str(row["reclaimed_line_count"]),
            fmt_number(row["reclaimed_bandwidth"]),
        ]
        for row in reclaimed_rows
    ]
    monthly_sales_table_rows = [
        [
            row["month"],
            fmt_money(row["contract_monthly_sales"], 0),
            fmt_money(row["settlement_monthly_sales"], 0),
        ]
        for row in monthly_sales_rows
    ]
    uc_region_table_rows = [
        [
            row["name"],
            str(row["line_count"]),
            fmt_percent(row["count_share"]),
            fmt_number(row["bandwidth"]),
            fmt_percent(row["bandwidth_share"]),
            str(row["active_line_count"]),
            fmt_number(row["active_bandwidth"]),
            str(row["reclaimed_line_count"]),
            fmt_number(row["reclaimed_bandwidth"]),
            fmt_money(row["sales"], 0),
            fmt_percent(row["sales_share"]),
            fmt_money(row["settlement"], 0),
            fmt_percent(row["settlement_share"]),
        ]
        for row in uc_region_rows
    ]
    forecast_table_rows = [
        [
            row["month"],
            fmt_money(row["contract_monthly_sales"], 0),
            fmt_money(row["settlement_monthly_sales"], 0),
        ]
        for row in forecast_rows
    ]
    status_table_rows = [
        [
            row["status"],
            str(row["line_count"]),
            fmt_percent(row["count_share"]),
            fmt_number(row["bandwidth"]),
            fmt_percent(row["bandwidth_share"]),
            fmt_money(row["sales"], 0),
            fmt_percent(row["sales_share"]),
            fmt_money(row["settlement"], 0),
            fmt_percent(row["settlement_share"]),
            fmt_money(row["profit"], 0),
            fmt_percent(row["margin"]),
        ]
        for row in status_rows
    ]
    platform_pie = platform_pie_svg(platform_rows)
    active_bandwidth_bar = active_bandwidth_bar_svg(active_bandwidth_rows)
    reclaimed_bar = active_bandwidth_bar_svg(
        reclaimed_rows,
        count_key="reclaimed_line_count",
        bandwidth_key="reclaimed_bandwidth",
        chart_class="reclaimed-statistics-bars",
        chart_title="每月回收线路数量和回收带宽",
        aria_label="每月回收线路数量和回收带宽柱状图",
        empty_text="暂无月度回收数据。",
        count_legend="回收线路数",
        bandwidth_legend="回收带宽（Mbps）",
    )
    creation_bar = creation_bar_svg(monthly_rows)
    monthly_sales_bar = monthly_sales_bar_svg(monthly_sales_rows)
    forecast_sales_bar = monthly_sales_bar_svg(
        forecast_rows,
        chart_class="forecast-sales-bars",
        chart_title="未来6个月合同月销与结算月销预测",
        aria_label="未来6个月合同月销与结算月销预测柱状图",
    )
    active_bandwidth_months = ", ".join(
        f'"{row["month"]}"' for row in active_bandwidth_rows
    )
    mermaid_active_statistics_charts = [
        "```mermaid",
        "xychart-beta",
        "    title \"每月在售线路数量\"",
        f"    x-axis [{active_bandwidth_months}]",
        f"    y-axis \"线路数\" 0 --> {max((row['active_line_count'] for row in active_bandwidth_rows), default=0)}",
        f"    bar [{', '.join(str(row['active_line_count']) for row in active_bandwidth_rows)}]",
        "```",
        "",
        "```mermaid",
        "xychart-beta",
        "    title \"每月在售总带宽\"",
        f"    x-axis [{active_bandwidth_months}]",
        f"    y-axis \"Mbps\" 0 --> {max((row['active_bandwidth'] for row in active_bandwidth_rows), default=0):g}",
        f"    bar [{', '.join(format(row['active_bandwidth'], 'g') for row in active_bandwidth_rows)}]",
        "```",
    ]
    reclaimed_months = ", ".join(f'"{row["month"]}"' for row in reclaimed_rows)
    mermaid_reclaimed_charts = [
        "```mermaid",
        "xychart-beta",
        "    title \"每月回收线路数量\"",
        f"    x-axis [{reclaimed_months}]",
        f"    y-axis \"线路数\" 0 --> {max((row['reclaimed_line_count'] for row in reclaimed_rows), default=0)}",
        f"    bar [{', '.join(str(row['reclaimed_line_count']) for row in reclaimed_rows)}]",
        "```",
        "",
        "```mermaid",
        "xychart-beta",
        "    title \"每月回收带宽\"",
        f"    x-axis [{reclaimed_months}]",
        f"    y-axis \"Mbps\" 0 --> {max((row['reclaimed_bandwidth'] for row in reclaimed_rows), default=0):g}",
        f"    bar [{', '.join(format(row['reclaimed_bandwidth'], 'g') for row in reclaimed_rows)}]",
        "```",
    ]
    mermaid_months = ", ".join(f'"{row["month"]}"' for row in monthly_rows)
    mermaid_creation_charts = [
        "```mermaid",
        "xychart-beta",
        "    title \"每月新增线路数\"",
        f"    x-axis [{mermaid_months}]",
        f"    y-axis \"线路数\" 0 --> {max((row['line_count'] for row in monthly_rows), default=0)}",
        f"    bar [{', '.join(str(row['line_count']) for row in monthly_rows)}]",
        "```",
        "",
        "```mermaid",
        "xychart-beta",
        "    title \"每月新增带宽量\"",
        f"    x-axis [{mermaid_months}]",
        f"    y-axis \"Mbps\" 0 --> {max((row['bandwidth'] for row in monthly_rows), default=0)}",
        f"    bar [{', '.join(fmt_number(row['bandwidth']) for row in monthly_rows)}]",
        "```",
    ]
    monthly_sales_months = ", ".join(f'"{row["month"]}"' for row in monthly_sales_rows)
    contract_monthly_values = ", ".join(
        f"{row['contract_monthly_sales']:.0f}" for row in monthly_sales_rows
    )
    settlement_monthly_values = ", ".join(
        f"{row['settlement_monthly_sales']:.0f}" for row in monthly_sales_rows
    )
    mermaid_monthly_sales_chart = [
        "```mermaid",
        "xychart-beta",
        "    title \"每月合同月销与结算月销\"",
        f"    x-axis [{monthly_sales_months}]",
        f"    y-axis \"人民币元\" 0 --> {max((max(row['contract_monthly_sales'], row['settlement_monthly_sales']) for row in monthly_sales_rows), default=0):.0f}",
        f"    bar [{contract_monthly_values}]",
        f"    bar [{settlement_monthly_values}]",
        "```",
    ]
    forecast_months = ", ".join(f'"{row["month"]}"' for row in forecast_rows)
    forecast_contract_values = ", ".join(
        f"{row['contract_monthly_sales']:.0f}" for row in forecast_rows
    )
    forecast_settlement_values = ", ".join(
        f"{row['settlement_monthly_sales']:.0f}" for row in forecast_rows
    )
    mermaid_forecast_sales_chart = [
        "```mermaid",
        "xychart-beta",
        "    title \"未来6个月合同月销与结算月销预测\"",
        f"    x-axis [{forecast_months}]",
        f"    y-axis \"人民币元\" 0 --> {max((max(row['contract_monthly_sales'], row['settlement_monthly_sales']) for row in forecast_rows), default=0):.0f}",
        f"    bar [{forecast_contract_values}]",
        f"    bar [{forecast_settlement_values}]",
        "```",
    ]
    mermaid_platform_pie = [
        "```mermaid",
        "pie title 各线路平台售出带宽（Mbps）",
        *[
            f'    "{str(row["name"]).replace(chr(34), chr(92) + chr(34))}": {fmt_number(row["bandwidth"])}'
            for row in platform_rows
        ],
        "```",
    ]

    md_parts = [
        f"# {REPORT_TITLE}",
        "",
        f"> 生成时间：{generated_at.strftime('%Y-%m-%d %H:%M:%S %Z')}  ",
        f"> 统计基准日：{fmt_date(as_of)}  ",
        f"> 数据源：`{source_name}`  ",
        f"> 报告文件：`{report_filename}`",
        "",
        "## 一、管理摘要",
        "",
        f"- 在售线路：**{summary['line_count']:,} 条**；总售出带宽：**{fmt_number(summary['total_bandwidth'])} Mbps**。",
        f"- 当前在售线路：**{summary['active_line_count']:,} 条**；在售总带宽：**{fmt_number(summary['active_bandwidth'])} Mbps**；已回收线路：**{summary['reclaimed_line_count']:,} 条**，已回收带宽：**{fmt_number(summary['reclaimed_bandwidth'])} Mbps**。",
        f"- 合同销售额：**{fmt_money(summary['total_sales'], 0)}**；成本：**{fmt_money(summary['total_cost'], 0)}**；合同毛利：**{fmt_money(summary['total_profit'], 0)}**；毛利率：**{fmt_percent(summary['gross_margin'])}**。",
        f"- 公司实际收入（结算价格）：**{fmt_money(summary['total_settlement'], 0)}**；实际收入减成本：**{fmt_money(summary['settlement_surplus'], 0)}**；实际收入口径差额率：**{fmt_percent(summary['settlement_margin'])}**。",
        f"- 加权平均销售单价：**{fmt_money(summary['weighted_sale_unit'])} / Mbps / 月**；合同口径每 Mbps 月均销售价格：**{fmt_money(summary['contract_monthly_unit'])}**。",
        "",
        "### 运营观察",
        "",
        *[f"- {observation}" for observation in observations],
        "",
        "## 二、线路状态汇总",
        "",
        f"按统计基准日 {fmt_date(as_of)} 判断：结束时间早于统计基准日的线路为已回收，其余线路为在售。全量财务和历史新增统计均保留已回收线路数据。",
        "",
        markdown_table(
            ["线路状态", "线路数", "线路数占比", "带宽 Mbps", "带宽占比", "销售额", "销售额占比", "实际收入（结算价格）", "实际收入占比", "合同毛利", "毛利率"],
            status_table_rows,
        ),
        "",
        "## 三、7 天内到期线路提醒",
        "",
        f"共 **{len(expiring)} 条**线路将在 7 天内到期，涉及 **{fmt_number(expiring_bandwidth)} Mbps**、合同销售额 **{fmt_money(expiring_sales, 0)}**、实际收入 **{fmt_money(total((line for line, _ in expiring), 'settlement_total') if expiring else 0, 0)}**。",
        "",
        markdown_table(
            ["专线ID", "线路名称", "剩余时间", "结束日期", "平台", "地域", "带宽 Mbps", "合同销售额", "实际收入"],
            expiry_rows,
        )
        if expiry_rows
        else "暂无 7 天内到期线路。",
        "",
        "## 四、每月在售统计",
        "",
        "按月末快照统计每个自然月的在售线路数量和在售总带宽，当前月以统计基准日作为快照日期。快照日处于线路开始至结束日期（含边界）时计入。",
        "",
        *mermaid_active_statistics_charts,
        "",
        markdown_table(
            ["月份", "在售线路数", "在售总带宽（Mbps）"],
            active_bandwidth_table_rows,
        ),
        "",
        "## 五、每月回收线路统计",
        "",
        "按线路结束日期所在月份统计每月回收线路数量和回收带宽；仅统计结束日期早于统计基准日的线路。",
        "",
        *(mermaid_reclaimed_charts if reclaimed_rows else []),
        "",
        markdown_table(
            ["月份", "回收线路数", "回收带宽（Mbps）"],
            reclaimed_table_rows,
        )
        if reclaimed_table_rows
        else "暂无已回收线路。",
        "",
        "## 六、按创建时间的月度新增",
        "",
        "按线路创建时间统计每月新创建的线路数量和新增带宽量；新增带宽为该月创建线路带宽之和。",
        "",
        *mermaid_creation_charts,
        "",
        markdown_table(
            ["创建月份", "新增线路数", "新增带宽（Mbps）", "新增实际收入", "其中在售线路", "其中在售带宽（Mbps）", "其中在售实际收入", "其中已回收线路", "其中已回收带宽（Mbps）", "其中已回收实际收入"],
            monthly_table_rows,
        ),
        "",
        "### 按创建月份的已回收线路明细",
        "",
        "用于对应上表“其中已回收线路”数据，按创建月份、结束日期和线路 ID 排序。",
        "",
        markdown_table(
            ["创建月份", "专线ID", "线路名称", "平台", "地域", "带宽 Mbps", "创建日期", "开始日期", "结束日期", "合同销售额", "实际收入"],
            reclaimed_creation_detail_rows,
        )
        if reclaimed_creation_detail_rows
        else "暂无已回收线路。",
        "",
        "## 七、月销金额分析",
        "",
        "按线路开始时间所在月份作为第一个服务月，按合同月数逐月展开并汇总全量历史线路；合同月销 = 销售价格 ÷ 时长（月），结算月销 = 结算价格 ÷ 时长（月），不足整月按实际月数比例计入。",
        "",
        *mermaid_monthly_sales_chart,
        "",
        markdown_table(
            ["月份", "合同月销", "结算月销（实际收入）"],
            monthly_sales_table_rows,
        ),
        "",
        "## 八、未来6个月月销预测",
        "",
        f"预测区间为统计基准日 {fmt_date(as_of)} 后的 6 个自然月，基于当前在用的 {summary['active_line_count']} 条线路；假设线路到期后均按当前付款周期续约，合同金额和结算金额均按当前月销水平延续。",
        "",
        *mermaid_forecast_sales_chart,
        "",
        markdown_table(
            ["预测月份", "合同月销", "结算月销（实际收入）"],
            forecast_table_rows,
        ),
        "",
        "## 九、线路平台分析",
        "",
        *mermaid_platform_pie,
        "",
        markdown_table(
            ["平台", "全量线路数", "全量带宽（Mbps）", "带宽占比", "在售线路数", "在售带宽（Mbps）", "已回收线路数", "已回收带宽（Mbps）", "销售额", "销售额占比", "实际收入（结算价格）", "实际收入占比", "加权销售单价", "毛利率"],
            platform_table_rows,
        ),
        "",
        "### UC 平台地域分析",
        "",
        "以下占比均以 UC 平台自身汇总为分母，用于识别 UC 平台的地域线路数量、带宽和收入结构。",
        "",
        markdown_table(
            ["地域", "线路数", "线路数占比", "带宽（Mbps）", "带宽占比", "在售线路数", "在售带宽（Mbps）", "已回收线路数", "已回收带宽（Mbps）", "合同销售额", "销售额占比", "实际收入（结算价格）", "实际收入占比"],
            uc_region_table_rows,
        )
        if uc_region_table_rows
        else "暂无 UC 平台地域数据。",
        "",
        "## 十、地域/产品结构分析",
        "",
        markdown_table(
            ["地域", "全量线路数", "全量带宽（Mbps）", "带宽占比", "在售线路数", "在售带宽（Mbps）", "已回收线路数", "已回收带宽（Mbps）", "销售额", "销售额占比", "实际收入（结算价格）", "实际收入占比", "毛利率"],
            region_table_rows,
        ),
        "",
        "## 十一、付款周期分析",
        "",
        markdown_table(
            ["付款周期", "全量线路数", "线路数占比", "全量带宽（Mbps）", "带宽占比", "在售线路数", "在售带宽（Mbps）", "已回收线路数", "已回收带宽（Mbps）", "销售额", "销售额占比", "实际收入（结算价格）", "实际收入占比", "平均时长（月）"],
            payment_table_rows,
        ),
        "",
        "## 十二、定价与利润风险",
        "",
        markdown_table(
            ["专线ID", "线路名称", "状态", "平台", "带宽 Mbps", "销售单价", "成本单价", "实际收入（结算价格）", "合同毛利", "毛利率"],
            risk_rows,
        )
        if risk_rows
        else "暂无毛利小于或等于 0 的线路。",
        "",
        "## 十三、数据质量检查",
        "",
        *([f"- {issue}" for issue in issues] if issues else ["- 未发现必填字段、日期或基础数值异常。"]),
        f"- 已过期线路：{len(expired)} 条。",
        "",
        "## 十四、口径说明",
        "",
        "- 金额按源表约定展示为人民币元；结算价格为公司实际收入，销售价格为合同销售额。",
        "- 总售出带宽为所有有效线路带宽之和；平台和地域占比均以总带宽为分母。",
        "- 加权平均销售单价 = Σ（带宽 × 销售单价）÷ Σ带宽。",
        "- 合同口径每 Mbps 月均销售价格 = 合同销售额 ÷ Σ（带宽 × 合同月数）。",
        "- 每月在售线路数和在售总带宽采用月末快照口径，当前月以统计基准日作为快照日期；快照日尚未开始或已经结束的线路不计入。",
        "- 每月回收线路统计按结束日期所在自然月聚合；仅统计结束日期早于统计基准日的线路，回收带宽为这些线路带宽之和。",
        "- 月销统计按服务月份展开：以开始时间所在月份作为第一个服务月，按时长（月）逐月计入；因此季度、半年、年付线路会在覆盖到的每个月计入合同月销和结算月销，不足整月按实际月数比例计入。",
        "- 未来 6 个月预测仅纳入统计基准日在用线路；线路结束后按付款周期（月、季度、半年、年）自动续约，合同月销和结算月销按当前月度折算金额延续。",
        "- 7 天内到期口径为结束日期落在统计基准日到基准日后 7 个自然日之间，包含边界日期。",
    ]
    markdown = "\n".join(md_parts)

    html_kpis = "".join(
        f'<div class="kpi"><div class="kpi-label">{html.escape(label)}</div><div class="kpi-value">{html.escape(value)}</div><div class="kpi-unit">{html.escape(unit)}</div></div>'
        for label, value, unit in kpis
    )
    html_observations = "".join(f"<li>{html.escape(observation)}</li>" for observation in observations)
    html_expiry = (
        html_table(
            ["专线ID", "线路名称", "剩余时间", "结束日期", "平台", "地域", "带宽 Mbps", "合同销售额", "实际收入"],
            expiry_rows,
        )
        if expiry_rows
        else '<p class="muted">暂无 7 天内到期线路。</p>'
    )
    html_risks = (
        html_table(
            ["专线ID", "线路名称", "状态", "平台", "带宽 Mbps", "销售单价", "成本单价", "实际收入（结算价格）", "合同毛利", "毛利率"],
            risk_rows,
        )
        if risk_rows
        else '<p class="muted">暂无毛利小于或等于 0 的线路。</p>'
    )
    html_issue = (
        "".join(f"<li>{html.escape(issue)}</li>" for issue in issues)
        if issues
        else "<li>未发现必填字段、日期或基础数值异常。</li>"
    )
    html_doc = f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(REPORT_TITLE)}</title>
<style>
:root {{ color-scheme: light; font-family: "Microsoft YaHei", "PingFang SC", Arial, sans-serif; }}
body {{ margin: 0; background: #f4f7fb; color: #1f2937; line-height: 1.6; }}
.page {{ max-width: 1360px; margin: 0 auto; padding: 32px 24px 64px; }}
h1 {{ margin: 0 0 8px; color: #0f2747; font-size: 30px; }}
h2 {{ margin-top: 34px; color: #12395e; border-left: 4px solid #2b75b3; padding-left: 10px; }}
h3 {{ margin-top: 24px; color: #244b6f; }}
.meta {{ color: #64748b; font-size: 14px; margin-bottom: 24px; }}
.kpis {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin: 20px 0 28px; }}
.kpi {{ background: white; border: 1px solid #dbe5ef; border-radius: 10px; padding: 16px; box-shadow: 0 2px 8px rgba(15,39,71,.05); }}
.kpi-label {{ color: #64748b; font-size: 13px; }}
.kpi-value {{ color: #0f2747; font-size: 25px; font-weight: 700; margin-top: 4px; }}
.kpi-unit {{ color: #64748b; font-size: 12px; }}
.panel {{ background: white; border: 1px solid #dbe5ef; border-radius: 10px; padding: 18px 20px; margin: 14px 0; overflow-x: hidden; }}
table {{ width: 100%; table-layout: fixed; border-collapse: collapse; font-size: 13px; }}
th {{ text-align: left; background: #eaf2f9; color: #244b6f; font-weight: 700; white-space: normal; overflow-wrap: anywhere; word-break: break-word; }}
th, td {{ padding: 9px 8px; border-bottom: 1px solid #e5edf4; vertical-align: top; overflow-wrap: anywhere; word-break: break-word; }}
tr:last-child td {{ border-bottom: 0; }}
tr:hover td {{ background: #f8fbfe; }}
ul {{ margin-top: 8px; }}
.muted {{ color: #64748b; }}
.warning {{ color: #9a3412; background: #fff7ed; border: 1px solid #fed7aa; border-radius: 8px; padding: 12px 14px; }}
.platform-pie {{ width: 100%; max-width: 520px; min-width: 420px; height: auto; }}
.pie-label {{ fill: #173b63; font-size: 14px; font-weight: 600; }}
.pie-total {{ fill: #ffffff; font-size: 24px; font-weight: 800; paint-order: stroke; stroke: #173b63; stroke-width: 4px; stroke-linejoin: round; }}
.pie-unit {{ fill: #ffffff; font-size: 13px; font-weight: 700; paint-order: stroke; stroke: #173b63; stroke-width: 2px; stroke-linejoin: round; }}
.chart-wrap {{ display: grid; grid-template-columns: minmax(420px, 1fr) minmax(620px, 1.6fr); gap: 20px; align-items: center; }}
.chart-card {{ min-width: 0; }}
.active-statistics-bars, .reclaimed-statistics-bars, .creation-bars {{ width: 100%; min-width: 760px; height: auto; }}
.monthly-sales-bars, .forecast-sales-bars {{ width: 100%; min-width: 760px; height: auto; }}
.uc-region-table {{ min-width: 0; table-layout: fixed; font-size: 11px; }}
.uc-region-table th, .uc-region-table td {{ padding: 7px 5px; white-space: normal; overflow-wrap: anywhere; word-break: break-word; }}
.reclaimed-creation-detail {{ min-width: 0; table-layout: fixed; font-size: 11px; }}
.reclaimed-creation-detail th, .reclaimed-creation-detail td {{ padding: 7px 5px; }}
.chart-gridline {{ stroke: #dbe5ef; stroke-width: 1; }}
.chart-axis-title, .chart-tick, .chart-value, .chart-month, .chart-legend {{ fill: #173b63; font-size: 13px; }}
.chart-axis-title, .chart-legend {{ font-weight: 700; }}
.chart-value {{ font-weight: 700; }}
.footer {{ margin-top: 38px; color: #94a3b8; font-size: 12px; }}
@media (max-width: 700px) {{ .page {{ padding: 20px 12px 40px; }} h1 {{ font-size: 24px; }} }}
</style>
</head>
<body>
<main class="page">
<h1>{html.escape(REPORT_TITLE)}</h1>
<div class="meta">生成时间：{html.escape(generated_at.strftime('%Y-%m-%d %H:%M:%S %Z'))} ｜ 统计基准日：{html.escape(fmt_date(as_of))} ｜ 数据源：{html.escape(source_name)}<br>金额单位：{html.escape(AMOUNT_LABEL)} ｜ 文件名：{html.escape(report_filename)}</div>
<section class="kpis">{html_kpis}</section>
<section class="panel"><h2>一、管理摘要</h2><ul>{html_observations}</ul></section>
<section class="panel"><h2>二、线路状态汇总</h2><p>按统计基准日 {fmt_date(as_of)} 判断：结束时间早于统计基准日的线路为已回收，其余线路为在售。全量财务和历史新增统计均保留已回收线路数据。</p>{html_table(['线路状态','线路数','线路数占比','带宽 Mbps','带宽占比','销售额','销售额占比','实际收入（结算价格）','实际收入占比','合同毛利','毛利率'], status_table_rows)}</section>
<section class="panel"><h2>三、7 天内到期线路提醒</h2><p>共 <strong>{len(expiring)} 条</strong>线路将在 7 天内到期，涉及 <strong>{fmt_number(total((line for line, _ in expiring), 'bandwidth') if expiring else 0)} Mbps</strong>、合同销售额 <strong>{fmt_money(total((line for line, _ in expiring), 'sales_total') if expiring else 0, 0)}</strong>、实际收入 <strong>{fmt_money(total((line for line, _ in expiring), 'settlement_total') if expiring else 0, 0)}</strong>。</p>{html_expiry}</section>
<section class="panel"><h2>四、每月在售统计</h2><p>按月末快照统计每个自然月的在售线路数量和在售总带宽，当前月以统计基准日作为快照日期。快照日处于线路开始至结束日期（含边界）时计入。</p>{active_bandwidth_bar}{html_table(['月份','在售线路数','在售总带宽（Mbps）'], active_bandwidth_table_rows)}</section>
<section class="panel"><h2>五、每月回收线路统计</h2><p>按线路结束日期所在月份统计每月回收线路数量和回收带宽；仅统计结束日期早于统计基准日的线路。</p>{reclaimed_bar}{html_table(['月份','回收线路数','回收带宽（Mbps）'], reclaimed_table_rows) if reclaimed_table_rows else '<p class="muted">暂无已回收线路。</p>'}</section>
<section class="panel"><h2>六、按创建时间的月度新增</h2><p>按线路创建时间统计每月新创建的线路数量和新增带宽量；新增带宽为该月创建线路带宽之和，并拆分展示当前在售和已回收数据。</p>{creation_bar}{html_table(['创建月份','新增线路数','新增带宽（Mbps）','新增实际收入','其中在售线路','其中在售带宽（Mbps）','其中在售实际收入','其中已回收线路','其中已回收带宽（Mbps）','其中已回收实际收入'], monthly_table_rows)}<h3>按创建月份的已回收线路明细</h3><p>用于对应上表“其中已回收线路”数据，按创建月份、结束日期和线路 ID 排序。</p>{html_table(['创建月份','专线ID','线路名称','平台','地域','带宽 Mbps','创建日期','开始日期','结束日期','合同销售额','实际收入'], reclaimed_creation_detail_rows, 'reclaimed-creation-detail') if reclaimed_creation_detail_rows else '<p class="muted">暂无已回收线路。</p>'}</section>
<section class="panel"><h2>七、月销金额分析</h2><p>按线路开始时间所在月份作为第一个服务月，按合同月数逐月展开并汇总全量历史线路；合同月销 = 销售价格 ÷ 时长（月），结算月销 = 结算价格 ÷ 时长（月），不足整月按实际月数比例计入。</p>{monthly_sales_bar}{html_table(['月份','合同月销','结算月销（实际收入）'], monthly_sales_table_rows)}</section>
<section class="panel"><h2>八、未来6个月月销预测</h2><p>预测区间为统计基准日 {fmt_date(as_of)} 后的 6 个自然月，基于当前在用的 {summary['active_line_count']} 条线路；假设线路到期后均按当前付款周期续约，合同金额和结算金额均按当前月销水平延续。</p>{forecast_sales_bar}{html_table(['预测月份','合同月销','结算月销（实际收入）'], forecast_table_rows)}</section>
<section class="panel"><h2>九、线路平台分析</h2><div class="chart-wrap"><div class="chart-card"><h3>平台售出带宽饼图</h3>{platform_pie}</div><div class="chart-card">{html_table(['平台','全量线路数','全量带宽（Mbps）','带宽占比','在售线路数','在售带宽（Mbps）','已回收线路数','已回收带宽（Mbps）','销售额','销售额占比','实际收入（结算价格）','实际收入占比','加权销售单价','毛利率'], platform_table_rows)}</div></div><h3>UC 平台地域分析</h3><p>以下占比均以 UC 平台自身汇总为分母，用于识别 UC 平台的地域线路数量、带宽和收入结构。</p>{html_table(['地域','线路数','线路数占比','带宽（Mbps）','带宽占比','在售线路数','在售带宽（Mbps）','已回收线路数','已回收带宽（Mbps）','合同销售额','销售额占比','实际收入（结算价格）','实际收入占比'], uc_region_table_rows, 'uc-region-table') if uc_region_table_rows else '<p class="muted">暂无 UC 平台地域数据。</p>'}</section>
<section class="panel"><h2>十、地域/产品结构分析</h2>{html_table(['地域','全量线路数','全量带宽（Mbps）','带宽占比','在售线路数','在售带宽（Mbps）','已回收线路数','已回收带宽（Mbps）','销售额','销售额占比','实际收入（结算价格）','实际收入占比','毛利率'], region_table_rows)}</section>
<section class="panel"><h2>十一、付款周期分析</h2>{html_table(['付款周期','全量线路数','线路数占比','全量带宽（Mbps）','带宽占比','在售线路数','在售带宽（Mbps）','已回收线路数','已回收带宽（Mbps）','销售额','销售额占比','实际收入（结算价格）','实际收入占比','平均时长（月）'], payment_table_rows)}</section>
<section class="panel"><h2>十二、定价与利润风险</h2>{html_risks}</section>
<section class="panel"><h2>十三、数据质量检查</h2><ul>{html_issue}<li>已回收线路：{len(expired)} 条。</li></ul></section>
<section class="panel"><h2>十四、口径说明</h2><ul><li>金额按源表约定展示为人民币元。</li><li>总售出带宽为所有有效线路带宽之和；平台和地域占比均以全量带宽为分母。</li><li>已回收线路定义为结束时间早于统计基准日的线路；全量财务和历史新增统计保留已回收线路数据。</li><li>加权平均销售单价 = Σ（带宽 × 销售单价）÷ Σ带宽。</li><li>合同口径每 Mbps 月均销售价格 = 合同销售额 ÷ Σ（带宽 × 合同月数）。</li><li>每月在售线路数和在售总带宽采用月末快照口径，当前月以统计基准日作为快照日期；快照日尚未开始或已经结束的线路不计入。</li><li>每月回收线路统计按结束日期所在自然月聚合；仅统计结束日期早于统计基准日的线路，回收带宽为这些线路带宽之和。</li><li>按线路创建时间按自然月聚合；新增带宽为该月新创建线路带宽之和。</li><li>月销统计按服务月份展开：以开始时间所在月份作为第一个服务月，按时长（月）逐月计入；因此季度、半年、年付线路会在覆盖到的每个月计入合同月销和结算月销，不足整月按实际月数比例计入。</li><li>未来 6 个月预测仅纳入统计基准日在用线路；线路结束后按付款周期（月、季度、半年、年）自动续约，合同月销和结算月销按当前月度折算金额延续。</li><li>7 天内到期口径为结束日期落在统计基准日到基准日后 7 个自然日之间，包含边界日期。</li></ul></section>
<div class="footer">由 report_generator.py 生成。</div>
</main>
</body>
</html>"""
    return html_doc, markdown


def generate_report(
    input_path: Path,
    output_root: Path,
    *,
    as_of: str | date | None = None,
    generated_at: datetime | None = None,
    unique_id: str | None = None,
) -> Path:
    generated_at = generated_at or datetime.now(TIMEZONE)
    if generated_at.tzinfo is None:
        generated_at = generated_at.replace(tzinfo=TIMEZONE)
    report_date = parse_as_of(as_of, generated_at)
    unique_id = unique_id or secrets.token_hex(4)
    if len(unique_id) != 8 or any(char not in "0123456789abcdefABCDEF" for char in unique_id):
        raise ValueError("unique_id 必须是 8 位十六进制字符串")

    lines, issues = load_lines(input_path)
    summary = summarize(lines, report_date, issues)
    platform_rows = group_metrics(lines, "platform", summary, report_date)
    monthly_rows = creation_metrics(lines, report_date)
    active_bandwidth_rows = monthly_active_bandwidth_metrics(lines, report_date)
    reclaimed_rows = monthly_reclaimed_metrics(lines, report_date)
    monthly_sales_rows = service_monthly_sales_metrics(lines, through_date=report_date)
    forecast_rows = forecast_monthly_sales_metrics(lines, report_date, horizon=6)
    region_rows = group_metrics(lines, "region", summary, report_date)
    uc_region_rows = platform_region_metrics(lines, "uc", report_date)
    payment_rows = payment_metrics(lines, summary, report_date)

    output_root.mkdir(parents=True, exist_ok=True)
    run_dir = output_root / f"{generated_at:%Y%m%d_%H%M%S}_{unique_id.lower()}"
    if run_dir.exists():
        raise FileExistsError(f"报告目录已存在，为避免覆盖而停止：{run_dir}")
    run_dir.mkdir()
    filename_stem = f"{REPORT_FILENAME_PREFIX}_{generated_at:%Y%m%d}_{unique_id.lower()}"
    html_path = run_dir / f"{filename_stem}.html"
    markdown_path = run_dir / f"{filename_stem}.md"
    html_doc, markdown = render_report(
        lines,
        summary,
        platform_rows,
        monthly_rows,
        active_bandwidth_rows,
        reclaimed_rows,
        monthly_sales_rows,
        forecast_rows,
        region_rows,
        uc_region_rows,
        payment_rows,
        source_name=input_path.name,
        generated_at=generated_at,
        as_of=report_date,
        report_filename=html_path.name,
    )
    html_path.write_text(html_doc, encoding="utf-8")
    markdown_path.write_text(markdown, encoding="utf-8")
    return run_dir


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=REPORT_TITLE)
    parser.add_argument(
        "--input",
        type=Path,
        default=Path("otherlive-线路汇总.xlsx"),
        help="输入 Excel 或 CSV 文件路径，默认：otherlive-线路汇总.xlsx",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=Path("reports"),
        help="报告输出根目录，默认：reports",
    )
    parser.add_argument(
        "--as-of",
        type=str,
        default=None,
        help="统计基准日期，支持 YYYY-MM-DD；默认使用当前 Asia/Shanghai 日期",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    input_path = args.input.resolve()
    if not input_path.exists():
        raise SystemExit(f"找不到输入文件：{input_path}")
    run_dir = generate_report(input_path, args.output_root.resolve(), as_of=args.as_of)
    print(f"报告已生成：{run_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
