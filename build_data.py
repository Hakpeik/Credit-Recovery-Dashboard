"""
Credit Recovery dashboard - data automation.

Scans the Excel folders in this workspace and writes ../data.js, which the
dashboard (index.html) loads directly. Run once, or use --watch so the
dashboard refreshes automatically whenever a workbook is added or replaced.

Sources
-------
Data Dashboard.xlsx ................ Loan Listing Report, columns BE:CQ
NPL Movement/<year>/Recovery*.xlsx .. NPL Ratio sheet (monthly totals)
Credit Recovery Workflow/Daily Overdue Data/** .... Term Loan, Overdraft, Credit Card
Credit Recovery Workflow/Daily Consolidate for Overdue Report/** ... Consolidate sheet

Usage
-----
    python automation/build_data.py
    python automation/build_data.py --watch
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import os
import re
import sys
import time
import warnings
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

ROOT = Path(__file__).resolve().parent.parent
AUTOMATION = Path(__file__).resolve().parent
CACHE_PATH = AUTOMATION / ".cache.json"
DATA_JS = ROOT / "data.js"

# Bump this whenever a parser changes shape. Old cache entries are discarded.
CACHE_VERSION = 6

RATE_KHR_PER_USD = 4055
MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

OVERDUE_ROOT = ROOT / "Credit Recovery Workflow" / "Daily Overdue Data"
# The workspace copy may only contain the newest reports. The complete archive
# lives in the sibling "Loan Recovery" folder, so scan it as a read-only
# fallback and let the workspace copy win when both contain the same date.
ARCHIVE_OVERDUE_ROOT = (
    ROOT.parent / "Loan Recovery" / "Credit Recovery Workflow" / "Daily Overdue Data"
)
OVERDUE_ROOTS = [OVERDUE_ROOT]
if ARCHIVE_OVERDUE_ROOT.exists() and ARCHIVE_OVERDUE_ROOT.resolve() != OVERDUE_ROOT.resolve():
    OVERDUE_ROOTS.append(ARCHIVE_OVERDUE_ROOT)
CONSOLIDATE_ROOT = ROOT / "Credit Recovery Workflow" / "Daily Consolidate for Overdue Report"
NPL_ROOT = ROOT / "NPL Movement"
DASHBOARD_FILE = ROOT / "Data Dashboard.xlsx"


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def norm(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def num(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return 0.0
        return float(value)
    text = str(value).strip().replace(",", "").replace(" ", "")
    if text in ("", "-", "N/A", "n/a", "None", "#REF!", "#DIV/0!"):
        return 0.0
    text = text.replace("%", "")
    try:
        return float(text)
    except ValueError:
        return 0.0


def text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.10g}"
    return str(value).strip()


def iso_date(value) -> str:
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    raw = str(value).strip()
    if not raw or raw in ("N/A", "-"):
        return ""
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", raw)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})", raw)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2})$", raw)
    if m:
        return f"20{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    return raw[:10]


def month_from_name(name: str) -> int | None:
    lowered = name.lower()
    for index, month in enumerate(MONTH_NAMES, start=1):
        if month.lower() in lowered:
            return index
    m = re.match(r"^(\d{1,2})[.\s_-]", name)
    if m and 1 <= int(m.group(1)) <= 12:
        return int(m.group(1))
    return None


def parse_date_from_path(path: Path) -> str:
    name = path.name
    patterns = [
        (r"(20\d{2})[-_.](\d{1,2})[-_.](\d{1,2})", ("y", "m", "d")),
        (r"(\d{1,2})[-_.](\d{1,2})[-_.](20\d{2})", ("d", "m", "y")),
        (r"(\d{2})[-_.](\d{1,2})[-_.](\d{2})(?!\d)", ("y2", "m", "d")),
    ]
    for pattern, order in patterns:
        match = re.search(pattern, name)
        if not match:
            continue
        parts = dict(zip(order, match.groups()))
        year = int(parts.get("y") or ("20" + parts.get("y2", "")))
        month = int(parts["m"])
        day = int(parts["d"])
        if 2000 <= year <= 2100 and 1 <= month <= 12 and 1 <= day <= 31:
            return f"{year}-{month:02d}-{day:02d}"

    year = None
    month = None
    for part in path.parts:
        if re.fullmatch(r"20\d{2}", part):
            year = int(part)
        found = month_from_name(part)
        if found:
            month = found
    if year and month:
        return f"{year}-{month:02d}-01"
    try:
        return dt.date.fromtimestamp(path.stat().st_mtime).isoformat()
    except OSError:
        return ""


def sig(path: Path) -> str:
    try:
        stat = path.stat()
        return f"{stat.st_mtime_ns}:{stat.st_size}"
    except OSError:
        return "missing"


def load_cache() -> dict:
    try:
        payload = json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"version": CACHE_VERSION, "entries": {}}
    if not isinstance(payload, dict) or payload.get("version") != CACHE_VERSION:
        return {"version": CACHE_VERSION, "entries": {}}
    if not isinstance(payload.get("entries"), dict):
        payload["entries"] = {}
    return payload


def save_cache(cache: dict) -> None:
    cache["version"] = CACHE_VERSION
    cache.setdefault("entries", {})
    try:
        CACHE_PATH.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    except OSError as exc:
        log(f"Cache write skipped: {exc}")


def log(message: str) -> None:
    print(f"[{dt.datetime.now():%H:%M:%S}] {message}", flush=True)


def workbook_rows(path: Path, sheet: str | None = None, max_col: int | None = None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb[wb.sheetnames[0]]
        rows = ws.iter_rows(
            min_row=1,
            max_row=ws.max_row,
            min_col=1,
            max_col=max_col or ws.max_column,
            values_only=True,
        )
        for row in rows:
            yield row
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Loan Listing Report (columns BE:CQ)
# ---------------------------------------------------------------------------

def build_loan_listing() -> dict:
    if not DASHBOARD_FILE.exists():
        return {"asOf": "", "rate": RATE_KHR_PER_USD, "columns": [], "rows": []}

    wb = openpyxl.load_workbook(DASHBOARD_FILE, read_only=True, data_only=True)
    try:
        ws = wb["Loan Listing Report"] if "Loan Listing Report" in wb.sheetnames else wb[wb.sheetnames[0]]
        first_col, last_col = 57, 95  # BE:CQ
        header_row = 10
        headers = []
        for col in range(first_col, last_col + 1):
            header = text(ws.cell(row=header_row, column=col).value)
            headers.append({
                "letter": get_column_letter(col),
                "header": header or f"Column {get_column_letter(col)}",
            })

        rows = []
        as_of = ""
        for col in range(last_col, first_col - 1, -1):
            candidate = iso_date(ws.cell(row=7, column=col).value)
            if candidate:
                as_of = candidate
                break
        for row in ws.iter_rows(
            min_row=header_row + 1,
            max_row=ws.max_row,
            min_col=first_col,
            max_col=last_col,
            values_only=True,
        ):
            if not row or row[0] in (None, ""):
                continue
            values = []
            for value in row:
                if isinstance(value, dt.datetime):
                    values.append(value.date().isoformat())
                elif isinstance(value, dt.date):
                    values.append(value.isoformat())
                elif isinstance(value, float):
                    values.append(round(value, 6) if not (math.isnan(value) or math.isinf(value)) else None)
                elif isinstance(value, int):
                    values.append(value)
                else:
                    cleaned = text(value)
                    values.append(cleaned if cleaned != "" else None)
            rows.append(values)

        return {
            "asOf": as_of,
            "rate": RATE_KHR_PER_USD,
            "totalRows": len(rows),
            "columns": headers,
            "rows": rows,
            "source": f"{DASHBOARD_FILE.name} · Loan Listing Report · BE:CQ",
        }
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# NPL Movement (NPL Ratio sheet)
# ---------------------------------------------------------------------------

def parse_npl_branches(ws) -> list[dict]:
    header_row = None
    for row_index in range(1, min(ws.max_row, 90) + 1):
        left = norm(ws.cell(row=row_index, column=17).value)
        right = norm(ws.cell(row=row_index, column=18).value)
        if left == "branch" and right.startswith("totalaccount"):
            header_row = row_index
            break
    if header_row is None:
        return []

    branches = []
    for row_index in range(header_row + 1, ws.max_row + 1):
        name = text(ws.cell(row=row_index, column=17).value)
        if not name:
            break
        if norm(name) == "total":
            break
        ratio = num(ws.cell(row=row_index, column=26).value)
        if 0 < ratio < 1:
            ratio *= 100
        branches.append({
            "branch": name,
            "accounts": int(num(ws.cell(row=row_index, column=18).value)),
            "totalOS": round(num(ws.cell(row=row_index, column=19).value), 6),
            "normal": round(num(ws.cell(row=row_index, column=20).value), 6),
            "specialMention": round(num(ws.cell(row=row_index, column=21).value), 6),
            "substandard": round(num(ws.cell(row=row_index, column=22).value), 6),
            "doubtful": round(num(ws.cell(row=row_index, column=23).value), 6),
            "loss": round(num(ws.cell(row=row_index, column=24).value), 6),
            "nplAmount": round(num(ws.cell(row=row_index, column=25).value), 6),
            "ratio": round(ratio, 6),
        })
    branches.sort(key=lambda item: item["ratio"], reverse=True)
    return branches


def build_npl_months() -> dict:
    months = []
    series = []
    sources = []
    if not NPL_ROOT.exists():
        return {"months": [], "series": [], "latest": None}

    files = sorted(
        path for path in NPL_ROOT.rglob("*.xlsx")
        if path.is_file() and not path.name.startswith("~$")
    )
    for path in files:
        match = re.search(r"Recovery\s*(\d{2})(\d{4})", path.name, re.IGNORECASE)
        if not match:
            continue
        month = int(match.group(1))
        year = int(match.group(2))
        if not (1 <= month <= 12):
            continue

        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            if "NPL Ratio" not in wb.sheetnames:
                wb.close()
                continue
            ws = wb["NPL Ratio"]
            labels = {}
            for row in ws.iter_rows(min_row=1, max_row=15, min_col=17, max_col=26, values_only=True):
                label = text(row[0])
                if label:
                    labels[norm(label)] = row
            total_row = labels.get("total")
            npl_row = labels.get("totalnpl")
            ratio_row = labels.get("totalnplratio")
            if not total_row or not npl_row or not ratio_row:
                wb.close()
                continue
            total_accounts = num(total_row[3])       # column T (Acc)
            total_os = num(total_row[4])             # column U (OS)
            npl_accounts = num(npl_row[3])
            npl_amount = num(npl_row[4])
            ratio = num(ratio_row[4])
            if ratio and ratio < 1:
                ratio *= 100

            # Optional long history block in the same sheet (header row varies).
            history = []
            history_rows = list(ws.iter_rows(
                min_row=95, max_row=ws.max_row, min_col=1, max_col=6, values_only=True
            ))
            history_start = 0
            for index, row in enumerate(history_rows):
                if norm(row[0]) == "date":
                    history_start = index + 1
                    break
            for row in history_rows[history_start:]:
                date_value = row[0]
                if not isinstance(date_value, (dt.date, dt.datetime)):
                    continue
                os_value = num(row[2])
                npl_value = num(row[4])
                ratio_value = num(row[5])
                if ratio_value and ratio_value < 1:
                    ratio_value *= 100
                if os_value or npl_value:
                    history.append({
                        "date": iso_date(date_value),
                        "totalAccounts": int(num(row[1])) if num(row[1]) else None,
                        "totalOS": round(os_value, 6),
                        "nplAccounts": int(num(row[3])) if num(row[3]) else None,
                        "nplAmount": round(npl_value, 6),
                        "ratio": round(ratio_value, 6),
                    })

            # Branch table lives to the right of the monthly summary.
            branches = parse_npl_branches(ws)
            wb.close()
        except Exception as exc:  # keep the pipeline running on a corrupt workbook
            log(f"NPL skip {path.name}: {exc}")
            continue

        if not total_accounts and not total_os:
            continue
        date = f"{year}-{month:02d}-{_month_end(year, month):02d}"
        entry = {
            "key": f"{year}-{month:02d}",
            "year": year,
            "month": MONTH_NAMES[month - 1],
            "monthShort": MONTH_NAMES[month - 1][:3],
            "date": date,
            "totalAccounts": int(total_accounts),
            "totalOS": round(total_os, 6),
            "nplAccounts": int(npl_accounts),
            "nplAmount": round(npl_amount, 6),
            "ratio": round(ratio, 6),
            "branches": branches,
            "source": path.name,
        }
        months.append(entry)
        sources.append(path.name)
        if history:
            series = history  # the most recent workbook carries the full history

    months.sort(key=lambda item: item["key"])
    series.sort(key=lambda item: item["date"])
    # Keep the raw history values: the March monthly rows repeat February in
    # some source workbooks, while the history block is accurate.
    series = [point for point in series if point["date"] >= "2023-01-01"]

    return {
        "months": months,
        "series": series,
        "latest": months[-1] if months else None,
        "sources": sources,
    }


def _month_end(year: int, month: int) -> int:
    if month == 12:
        return 31
    return (dt.date(year, month + 1, 1) - dt.timedelta(days=1)).day


# ---------------------------------------------------------------------------
# Daily Overdue Data
# ---------------------------------------------------------------------------

LOAN_TYPES = {
    "Term Loan": ("longterm", "termloan"),
    "Overdraft": ("overdraft",),
    "Credit Card": ("creditcard",),
}


def overdue_type_for(path: Path) -> str | None:
    joined = norm(str(path))
    for loan_type, needles in LOAN_TYPES.items():
        if any(needle in joined for needle in needles):
            return loan_type
    return None


def overdue_file_allowed(path: Path, loan_type: str) -> bool:
    name = norm(path.name)
    if "reportbiendofmonth" in norm(str(path.parent)) or norm(path.parent.name).startswith("report"):
        return False
    if loan_type == "Term Loan":
        return name.startswith("overduereport") and "onoverdraft" not in name
    if loan_type == "Overdraft":
        return "overduereportonoverdraft" in name
    if loan_type == "Credit Card":
        return "creditcardoveduereport" in name or "creditcardoverduereport" in name
    return False


def _find_header(rows, candidates, max_scan=12):
    for index, row in enumerate(rows[:max_scan]):
        normalised = [norm(value) for value in row]
        for candidate in candidates:
            if norm(candidate) in normalised:
                return index
    return None


def _col_index(headers, *names):
    normalised = [norm(value) for value in headers]
    for name in names:
        target = norm(name)
        if target in normalised:
            return normalised.index(target)
    return -1


def _value(row, index):
    if index is None or index < 0 or index >= len(row):
        return None
    return row[index]


def _aging_bucket(days) -> str:
    days = num(days)
    if days <= 0:
        return "Current / 0"
    if days <= 30:
        return "1-30"
    if days <= 60:
        return "31-60"
    if days <= 90:
        return "61-90"
    if days <= 180:
        return "91-180"
    if days <= 365:
        return "181-365"
    return "Over 365"


def parse_overdue(path: Path, loan_type: str) -> dict | None:
    try:
        rows = list(workbook_rows(path))
    except Exception as exc:
        log(f"Overdue skip {path.name}: {exc}")
        return None

    if loan_type == "Credit Card":
        header_index = _find_header(rows, ["Financial Contract"])
    else:
        header_index = _find_header(rows, ["Account No", "Currecnt Account"])
    if header_index is None:
        return None

    headers = rows[header_index]
    if loan_type == "Credit Card":
        i_account = _col_index(headers, "Financial Contract", "Autopay Account Number")
        i_customer = _col_index(headers, "Customer Name")
        i_branch = _col_index(headers, "Default Card Branch Name")
        i_ccy = _col_index(headers, "Default Account CCY")
        i_product = _col_index(headers, "Card Product")
        i_os = _col_index(headers, "Principal Outstanding Balance")
        i_overdue = _col_index(headers, "Total Payment", "Min Payment", "Full Payment")
        i_aging = _col_index(headers, "Aging")
        i_class = _col_index(headers, "Classification")
    elif loan_type == "Overdraft":
        i_account = _col_index(headers, "Currecnt Account", "Current Account", "Account No")
        i_customer = _col_index(headers, "Account Name", "Customer Name")
        i_branch = _col_index(headers, "Branch Code")
        i_ccy = _col_index(headers, "CCY")
        i_product = _col_index(headers, "Loan Product")
        i_os = _col_index(headers, "Outstanding Balance")
        i_overdue = _col_index(headers, "Total Balance", "Overdue Amount")
        i_aging = _col_index(headers, "Aging")
        i_class = _col_index(headers, "Classification")
    else:
        i_account = _col_index(headers, "Account No")
        i_customer = _col_index(headers, "Customer Name")
        i_branch = _col_index(headers, "Brancha Code", "Branch Code")
        i_ccy = _col_index(headers, "CCY")
        i_product = _col_index(headers, "Loan Product")
        i_os = _col_index(headers, "Outstanding Balance")
        i_overdue = _col_index(headers, "Total Balance", "OV-Prin")
        i_aging = _col_index(headers, "AGING", "Aging")
        i_class = _col_index(headers, "Classification")

    if i_account < 0:
        return None

    records = []
    classifications = {}
    aging = {}
    branches = {}
    total_os = total_overdue = usd_os = khr_os = 0.0
    os_usd_total = overdue_usd_total = 0.0

    for row in rows[header_index + 1:]:
        account = text(_value(row, i_account))
        if not account:
            continue
        os_value = num(_value(row, i_os))
        overdue_value = num(_value(row, i_overdue))
        if not os_value and not overdue_value:
            # Credit card files contain many closed / zero-balance rows.
            continue
        ccy = text(_value(row, i_ccy)).upper() or "USD"
        fx = 1.0 if ccy == "USD" else (RATE_KHR_PER_USD if ccy == "KHR" else 1.0)
        os_usd = os_value / fx
        overdue_usd = overdue_value / fx
        os_usd_total += os_usd
        overdue_usd_total += overdue_usd
        branch = text(_value(row, i_branch))
        classification = canonical_classification(_value(row, i_class))
        record = {
            "account": account,
            "customer": text(_value(row, i_customer)),
            "branch": branch,
            "ccy": ccy,
            "product": text(_value(row, i_product)),
            "os": round(os_value, 4),
            "overdue": round(overdue_value, 4),
            "aging": int(num(_value(row, i_aging))),
            "classification": classification,
            "loanType": loan_type,
        }
        records.append(record)

        total_os += os_value
        total_overdue += overdue_value
        if ccy == "KHR":
            khr_os += os_value
        elif ccy == "USD":
            usd_os += os_value

        item = classifications.setdefault(
            classification,
            {"accounts": 0, "os": 0.0, "overdue": 0.0, "osUSD": 0.0, "overdueUSD": 0.0},
        )
        item["accounts"] += 1
        item["os"] += os_value
        item["overdue"] += overdue_value
        item["osUSD"] += os_usd
        item["overdueUSD"] += overdue_usd
        bucket = _aging_bucket(record["aging"])
        aging[bucket] = aging.get(bucket, 0) + 1
        if branch:
            b_item = branches.setdefault(
                branch,
                {"accounts": 0, "osUSD": 0.0, "overdueUSD": 0.0, "classes": {}},
            )
            b_item["accounts"] += 1
            b_item["osUSD"] += os_usd
            b_item["overdueUSD"] += overdue_usd
            c_item = b_item["classes"].setdefault(
                classification, {"accounts": 0, "osUSD": 0.0}
            )
            c_item["accounts"] += 1
            c_item["osUSD"] += os_usd

    if not records:
        return None

    return {
        "date": parse_date_from_path(path),
        "loanType": loan_type,
        "accounts": len(records),
        "totalOS": round(total_os, 4),
        "overdue": round(total_overdue, 4),
        "totalOSUSD": round(os_usd_total, 4),
        "overdueUSD": round(overdue_usd_total, 4),
        "usdOS": round(usd_os, 4),
        "khrOS": round(khr_os, 4),
        "classifications": {
            key: {
                "accounts": value["accounts"],
                "os": round(value["os"], 4),
                "overdue": round(value["overdue"], 4),
                "osUSD": round(value["osUSD"], 4),
                "overdueUSD": round(value["overdueUSD"], 4),
            }
            for key, value in classifications.items()
        },
        "aging": aging,
        "branches": {
            key: {
                "accounts": value["accounts"],
                "osUSD": round(value["osUSD"], 4),
                "overdueUSD": round(value["overdueUSD"], 4),
                "classes": {
                    class_key: {
                        "accounts": class_value["accounts"],
                        "osUSD": round(class_value["osUSD"], 4),
                    }
                    for class_key, class_value in value["classes"].items()
                },
            }
            for key, value in sorted(branches.items(), key=lambda item: item[1]["osUSD"], reverse=True)
        },
        "source": path.name,
        "_records": records,
    }


def build_overdue(cache: dict) -> dict:
    summaries = []
    entries = cache.setdefault("entries", {})
    candidates = []
    seen = set()
    for root in OVERDUE_ROOTS:
        if not root.exists():
            continue
        files = sorted(
            path for path in root.rglob("*.xlsx")
            if path.is_file() and not path.name.startswith("~$")
        )
        for path in files:
            loan_type = overdue_type_for(path)
            if not loan_type or not overdue_file_allowed(path, loan_type):
                continue
            try:
                key = "overdue:" + str(path.relative_to(ROOT))
            except ValueError:
                key = "overdue-archive:" + str(path.relative_to(root))
            signature = sig(path)
            cached = entries.get(key)
            if cached and cached.get("sig") == signature:
                detail = cached.get("data")
            else:
                result = parse_overdue(path, loan_type)
                if not result:
                    continue
                detail = {k: v for k, v in result.items() if k != "_records"}
                entries[key] = {"sig": signature, "data": detail}
            if not detail:
                continue
            if not detail.get("date"):
                detail["date"] = parse_date_from_path(path)
            stamp = (detail.get("date"), detail.get("loanType"))
            if stamp in seen:
                continue
            seen.add(stamp)
            summaries.append(detail)

    summaries.sort(key=lambda item: (item["date"], item["loanType"]))
    return {
        "types": ["Term Loan", "Overdraft", "Credit Card"],
        "summaries": summaries,
        "files": len(summaries),
    }


# ---------------------------------------------------------------------------
# Daily Consolidate for Overdue Report
# ---------------------------------------------------------------------------

RESULT_KEYS = ["Promise to pay", "Repromise", "No Result", "N/A", "Paid Already"]


def canonical_classification(value) -> str:
    raw = text(value)
    if not raw:
        return "Unclassified"
    key = norm(raw)
    mapping = {
        "normal": "Normal",
        "specialmention": "Special Mention",
        "substandard": "Substandard",
        "doubtful": "Doubtful",
        "loss": "Loss",
    }
    return mapping.get(key, raw)


def canonical_result(value) -> str:
    raw = text(value)
    if not raw:
        return "N/A"
    key = norm(raw)
    if key.startswith("promise"):
        return "Promise to pay"
    if key.startswith("repromise") or key.startswith("re-promise"):
        return "Repromise"
    if key.startswith("noresult"):
        return "No Result"
    if key in ("na", "n/a"):
        return "N/A"
    if key.startswith("paid"):
        return "Paid Already"
    return raw


def parse_consolidate(path: Path) -> dict | None:
    try:
        rows = list(workbook_rows(path, sheet="Consolidate"))
    except Exception as exc:
        log(f"Consolidate skip {path.name}: {exc}")
        return None

    header_index = None
    for index, row in enumerate(rows[:12]):
        if "accountno" in [norm(value) for value in row]:
            header_index = index
            break
    if header_index is None:
        return None

    headers = rows[header_index]
    i_account = _col_index(headers, "ACCOUNT NO")
    i_customer = _col_index(headers, "CUSTOEMR NAME", "Customer Name")
    i_branch = _col_index(headers, "Branch Code: All", "Branch Code")
    i_ccy = _col_index(headers, "CCY")
    i_disbursed = _col_index(headers, "Disbursed Date")
    i_maturity = _col_index(headers, "Maturity Date")
    i_limit = _col_index(headers, "Credit Limit")
    i_os = _col_index(headers, "Outstanding Balance")
    i_rate = _col_index(headers, "INT RATE", "INT Rate")
    i_due_principal = _col_index(headers, "DUE-Prin", "DUE Prin")
    i_due_interest = _col_index(headers, "Due Int", "DUE Int")
    i_due_date = _col_index(headers, "Due-Date", "Due Date")
    i_ov_principal = _col_index(headers, "OV-Prin", "OV Prin")
    i_ov_interest = _col_index(headers, "OV-Int", "OV Int")
    i_penalty = _col_index(headers, "Penalty")
    i_total = _col_index(headers, "Total-Balance", "Total Balance")
    i_aging = _col_index(headers, "AGING", "Aging")
    i_class = _col_index(headers, "Classification")
    i_product = _col_index(headers, "Loan Product")
    i_rm = _col_index(headers, "RM Incharge")
    i_biz = _col_index(headers, "Biz Type")
    i_income = _col_index(headers, "Current Income/month", "Current Income")
    i_result = _col_index(headers, "Result")
    i_promise = _col_index(headers, "Promise\u200b to pay Date", "Promise to pay Date", "Promise to pay  Date")
    i_letter = _col_index(headers, "Issue Letter")
    i_capacity = _col_index(headers, "Capacity Payment")
    i_restructure = _col_index(headers, "Need to Restructuring Loan (Yes/No)")
    i_remark = _col_index(headers, "Remark")
    i_next = _col_index(headers, "Next Action")
    i_classinfo = _col_index(headers, "Class. Info.")

    if i_account < 0:
        return None

    records = []
    results = {key: 0 for key in RESULT_KEYS}
    classifications = {}
    branches = {}
    aging = {}
    total_os_usd = 0.0
    total_overdue_usd = 0.0
    restructure_yes = 0

    for row in rows[header_index + 1:]:
        account = text(_value(row, i_account))
        if not account:
            continue
        ccy = text(_value(row, i_ccy)).upper() or "USD"
        os_value = num(_value(row, i_os))
        total_value = num(_value(row, i_total))
        fx = 1.0 if ccy == "USD" else (RATE_KHR_PER_USD if ccy == "KHR" else 1.0)
        os_usd = os_value / fx
        overdue_usd = total_value / fx
        total_os_usd += os_usd
        total_overdue_usd += overdue_usd

        result_key = canonical_result(_value(row, i_result))
        results[result_key] = results.get(result_key, 0) + 1
        classification = canonical_classification(_value(row, i_class))
        branch = text(_value(row, i_branch)) or "Unassigned"
        restructure = text(_value(row, i_restructure))
        if norm(restructure).startswith("yes"):
            restructure_yes += 1

        c_item = classifications.setdefault(classification, {"accounts": 0, "osUSD": 0.0, "results": {}})
        c_item["accounts"] += 1
        c_item["osUSD"] += os_usd
        c_item["results"][result_key] = c_item["results"].get(result_key, 0) + 1

        b_item = branches.setdefault(branch, {"accounts": 0, "osUSD": 0.0, "results": {}})
        b_item["accounts"] += 1
        b_item["osUSD"] += os_usd
        b_item["results"][result_key] = b_item["results"].get(result_key, 0) + 1

        bucket = _aging_bucket(num(_value(row, i_aging)))
        aging[bucket] = aging.get(bucket, 0) + 1

        records.append({
            "account": account,
            "customer": text(_value(row, i_customer)),
            "branch": branch,
            "ccy": ccy,
            "product": text(_value(row, i_product)),
            "rm": text(_value(row, i_rm)),
            "bizType": text(_value(row, i_biz)),
            "os": round(os_value, 4),
            "osUSD": round(os_usd, 4),
            "overdue": round(total_value, 4),
            "aging": int(num(_value(row, i_aging))),
            "classification": classification,
            "result": result_key,
            "promiseDate": iso_date(_value(row, i_promise)),
            "letter": text(_value(row, i_letter)),
            "capacity": text(_value(row, i_capacity)),
            "restructure": restructure,
            "remark": text(_value(row, i_remark))[:500],
            "nextAction": text(_value(row, i_next))[:300],
            "classInfo": text(_value(row, i_classinfo))[:300],
            "currentIncome": text(_value(row, i_income)),
            "dueDate": iso_date(_value(row, i_due_date)),
        })

    if not records:
        return None

    summary = {
        "date": parse_date_from_path(path),
        "accounts": len(records),
        "totalOSUSD": round(total_os_usd, 4),
        "totalOverdueUSD": round(total_overdue_usd, 4),
        "results": results,
        "promises": results.get("Promise to pay", 0),
        "repromise": results.get("Repromise", 0),
        "noResult": results.get("No Result", 0),
        "na": results.get("N/A", 0),
        "paid": results.get("Paid Already", 0),
        "restructure": restructure_yes,
        "classifications": {
            key: {
                "accounts": value["accounts"],
                "osUSD": round(value["osUSD"], 4),
                "results": value["results"],
            }
            for key, value in classifications.items()
        },
        "branches": {
            key: {
                "accounts": value["accounts"],
                "osUSD": round(value["osUSD"], 4),
                "results": value["results"],
            }
            for key, value in sorted(branches.items(), key=lambda item: item[1]["accounts"], reverse=True)
        },
        "aging": aging,
        "source": path.name,
    }
    return {"summary": summary, "records": records}


def compact_consolidate_record(record: dict) -> dict:
    fx = RATE_KHR_PER_USD if record.get("ccy") == "KHR" else 1.0
    return {
        "account": record.get("account", ""),
        "customer": record.get("customer", ""),
        "branch": record.get("branch", ""),
        "osUSD": record.get("osUSD", 0),
        "overdueUSD": round((record.get("overdue") or 0) / fx, 4),
        "aging": record.get("aging", 0),
        "classification": record.get("classification", "Unclassified"),
        "result": record.get("result", "N/A"),
        "promiseDate": record.get("promiseDate", ""),
        "rm": record.get("rm", ""),
        "nextAction": record.get("nextAction", ""),
    }


def build_consolidate(cache: dict) -> dict:
    summaries = []
    details = {}
    entries = cache.setdefault("entries", {})
    if CONSOLIDATE_ROOT.exists():
        files = sorted(
            path for path in CONSOLIDATE_ROOT.rglob("*.xlsx")
            if path.is_file() and not path.name.startswith("~$") and "consolidateof" in norm(path.name)
        )
    else:
        files = []

    for path in files:
        key = "consolidate:" + str(path.relative_to(ROOT))
        signature = sig(path)
        cached = entries.get(key)
        if cached and cached.get("sig") == signature:
            summary = cached.get("data")
            records = cached.get("records", [])
        else:
            parsed = parse_consolidate(path)
            if not parsed:
                continue
            summary = parsed["summary"]
            records = [compact_consolidate_record(item) for item in parsed["records"]]
            entries[key] = {"sig": signature, "data": summary, "records": records}
        if not summary:
            continue
        summaries.append(summary)
        if records:
            details[summary["date"]] = {
                "date": summary["date"],
                "source": summary["source"],
                "records": records,
            }

    summaries.sort(key=lambda item: item["date"])
    latest = summaries[-1] if summaries else None
    return {
        "summaries": summaries,
        "details": details,
        "detail": details.get(latest["date"]) if latest else None,
        "files": len(summaries),
        "resultKeys": RESULT_KEYS,
    }


# ---------------------------------------------------------------------------
# Internal Report Department (Consolidate + Large Exposure + Related Party)
# ---------------------------------------------------------------------------

INTERNAL_ROOT = ROOT / "Internal Report Department"


def _cell(rows, row_index, col_index):
    if 1 <= row_index <= len(rows):
        row = rows[row_index - 1]
        if 1 <= col_index <= len(row):
            return row[col_index - 1]
    return None


def _series_block(rows, start, end):
    items = []
    total = None
    for row_index in range(start, end + 1):
        name = text(_cell(rows, row_index, 3))
        if not name:
            continue
        item = {
            "name": name,
            "prevAccounts": int(num(_cell(rows, row_index, 4))),
            "prevAmount": round(num(_cell(rows, row_index, 5)), 6),
            "accounts": int(num(_cell(rows, row_index, 6))),
            "amount": round(num(_cell(rows, row_index, 7)), 6),
            "change": round(num(_cell(rows, row_index, 8)), 6),
        }
        if norm(name) == "total":
            total = item
        else:
            items.append(item)
    return items, total


def _ytd_values(rows, row_index):
    return {
        "business": {
            "accounts": int(num(_cell(rows, row_index, 4))),
            "limit": round(num(_cell(rows, row_index, 5)), 6),
            "approved": round(num(_cell(rows, row_index, 6)), 6),
        },
        "retail": {
            "accounts": int(num(_cell(rows, row_index, 7))),
            "limit": round(num(_cell(rows, row_index, 8)), 6),
            "approved": round(num(_cell(rows, row_index, 9)), 6),
        },
    }


def parse_internal(path: Path) -> dict | None:
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        log(f"Internal skip {path.name}: {exc}")
        return None

    try:
        if "Consolidate" not in wb.sheetnames:
            return None
        ws = wb["Consolidate"]
        rows = list(ws.iter_rows(min_row=1, max_row=175, min_col=1, max_col=25, values_only=True))

        current_date = iso_date(_cell(rows, 7, 6))
        previous_date = iso_date(_cell(rows, 7, 4))
        if not current_date:
            return None
        year = int(current_date[:4])
        month = int(current_date[5:7])
        key = f"{year}-{month:02d}"

        biz_units, biz_total = _series_block(rows, 9, 13)
        branches, branch_total = _series_block(rows, 25, 30)
        commitment, commitment_total = _series_block(rows, 104, 109)
        products, product_total = _series_block(rows, 119, 126)
        disbursement, disbursement_total = _series_block(rows, 142, 146)
        paid_off, paid_off_total = _series_block(rows, 162, 166)

        # Loan received · progressive branch matrix.
        progressive_headers = [text(_cell(rows, 36, col)) for col in range(14, 23)]
        progressive_previous_headers = [text(_cell(rows, 36, col)) for col in range(5, 14)]
        progressive = []
        row_index = 37
        while row_index <= 52:
            stage = text(_cell(rows, row_index, 3))
            if not stage:
                break
            progressive.append({
                "stage": stage,
                "accounts": [round(num(_cell(rows, row_index, col)), 4) for col in range(14, 23)],
                "amounts": [round(num(_cell(rows, row_index + 1, col)), 6) for col in range(14, 23)],
                "prevAccounts": [round(num(_cell(rows, row_index, col)), 4) for col in range(5, 14)],
                "prevAmounts": [round(num(_cell(rows, row_index + 1, col)), 6) for col in range(5, 14)],
            })
            row_index += 2

        # Loan received · year to date.
        ytd_totals = {
            "received": _ytd_values(rows, 57),
            "approved": _ytd_values(rows, 58),
            "rejected": _ytd_values(rows, 59),
        }
        ytd_months = []
        current_month = ""
        month_lookup = {}
        for row_index in range(60, 96):
            label = text(_cell(rows, row_index, 2))
            status = text(_cell(rows, row_index, 3))
            if label:
                current_month = label
            if not current_month or not status:
                continue
            status_key = norm(status)
            if status_key not in ("received", "approved", "rejected"):
                continue
            entry = month_lookup.setdefault(current_month, {"month": current_month})
            entry[status_key] = _ytd_values(rows, row_index)
        for month_name in MONTH_NAMES:
            if month_name in month_lookup:
                ytd_months.append(month_lookup[month_name])
            elif any(name.strip().lower() == month_name.lower() for name in month_lookup):
                match = next(name for name in month_lookup if name.strip().lower() == month_name.lower())
                ytd_months.append(month_lookup[match])

        # Product type summary on the right of LOAN BY PRODUCTS.
        product_types = []
        for row_index in range(117, 129):
            name = text(_cell(rows, row_index, 11))
            accounts = int(num(_cell(rows, row_index, 12)))
            amount = round(num(_cell(rows, row_index, 13)), 6)
            if not name and not accounts and not amount:
                continue
            product_types.append({"name": name or "Total", "accounts": accounts, "amount": amount})

        large_exposure = None
        if "Large Exposure" in wb.sheetnames:
            le_ws = wb["Large Exposure"]
            le_rows = list(le_ws.iter_rows(min_row=1, max_row=24, min_col=1, max_col=16, values_only=True))
            le_items = []
            for row_index in range(13, 25):
                borrower = text(_cell(le_rows, row_index, 2))
                if not borrower:
                    continue
                le_items.append({
                    "no": text(_cell(le_rows, row_index, 1)),
                    "borrower": borrower,
                    "maturity": iso_date(_cell(le_rows, row_index, 4)),
                    "type": text(_cell(le_rows, row_index, 5)),
                    "approved": round(num(_cell(le_rows, row_index, 6)), 6),
                    "outstanding": round(num(_cell(le_rows, row_index, 7)), 6),
                    "offBalance": round(num(_cell(le_rows, row_index, 8)), 6),
                    "totalExposure": round(num(_cell(le_rows, row_index, 9)), 6),
                    "riskWeight": round(num(_cell(le_rows, row_index, 10)), 6),
                    "weighted": round(num(_cell(le_rows, row_index, 11)), 6),
                    "ratio": round(num(_cell(le_rows, row_index, 12)), 8),
                    "maxPct": round(num(_cell(le_rows, row_index, 13)), 8),
                    "excess": round(num(_cell(le_rows, row_index, 14)), 8),
                    "purpose": text(_cell(le_rows, row_index, 15)),
                })
            large_exposure = {
                "asOf": iso_date(_cell(le_rows, 4, 2)),
                "rate": round(num(_cell(le_rows, 6, 2)), 4),
                "netWorth": round(num(_cell(le_rows, 9, 2)), 6),
                "unit": "KHR million",
                "totals": {
                    "approved": round(num(_cell(le_rows, 12, 6)), 6),
                    "outstanding": round(num(_cell(le_rows, 12, 7)), 6),
                    "offBalance": round(num(_cell(le_rows, 12, 8)), 6),
                    "totalExposure": round(num(_cell(le_rows, 12, 9)), 6),
                    "weighted": round(num(_cell(le_rows, 12, 11)), 6),
                    "ratio": round(num(_cell(le_rows, 12, 12)), 8),
                    "maxPct": round(num(_cell(le_rows, 12, 13)), 8),
                    "excess": round(num(_cell(le_rows, 12, 14)), 8),
                },
                "rows": le_items,
            }

        related_party = None
        rp_sheet = next(
            (name for name in ("Related Party", "Loan to Related Party") if name in wb.sheetnames),
            None,
        )
        if rp_sheet:
            rp_ws = wb[rp_sheet]
            rp_rows = list(rp_ws.iter_rows(min_row=1, max_row=40, min_col=1, max_col=14, values_only=True))
            rp_items = []
            for row_index in range(5, 41):
                number = _cell(rp_rows, row_index, 3)
                if not isinstance(number, (int, float)) or isinstance(number, bool):
                    continue
                borrower = text(_cell(rp_rows, row_index, 4))
                if not borrower:
                    continue
                rp_items.append({
                    "no": int(number),
                    "borrower": borrower,
                    "creditLimit": round(num(_cell(rp_rows, row_index, 5)), 6),
                    "outstanding": round(num(_cell(rp_rows, row_index, 6)), 6),
                })
            related_party = {
                "asOf": text(_cell(rp_rows, 2, 6)),
                "rate": round(num(_cell(rp_rows, 1, 6)), 4),
                "netWorth": round(num(_cell(rp_rows, 2, 12)), 6),
                "count": len(rp_items),
                "creditLimit": round(sum(item["creditLimit"] for item in rp_items), 6),
                "outstanding": round(sum(item["outstanding"] for item in rp_items), 6),
                "rows": rp_items,
            }

        return {
            "key": key,
            "year": year,
            "month": MONTH_NAMES[month - 1],
            "monthShort": MONTH_NAMES[month - 1][:3],
            "date": current_date,
            "previousDate": previous_date,
            "source": path.name,
            "bizUnits": biz_units,
            "bizTotal": biz_total,
            "branches": branches,
            "branchTotal": branch_total,
            "progressive": {
                "headers": progressive_headers,
                "previousHeaders": progressive_previous_headers,
                "rows": progressive,
            },
            "ytd": {"totals": ytd_totals, "months": ytd_months},
            "commitment": commitment,
            "commitmentTotal": commitment_total,
            "products": products,
            "productTotal": product_total,
            "productTypes": product_types,
            "disbursement": disbursement,
            "disbursementTotal": disbursement_total,
            "paidOff": paid_off,
            "paidOffTotal": paid_off_total,
            "largeExposure": large_exposure,
            "relatedParty": related_party,
        }
    except Exception as exc:
        log(f"Internal skip {path.name}: {exc}")
        return None
    finally:
        wb.close()


def build_internal(cache: dict) -> dict:
    entries = cache.setdefault("entries", {})
    candidates = []
    if INTERNAL_ROOT.exists():
        for path in sorted(INTERNAL_ROOT.rglob("*.xlsx")):
            if not path.is_file() or path.name.startswith("~$"):
                continue
            key = "internal:" + str(path.relative_to(ROOT))
            signature = sig(path)
            cached = entries.get(key)
            if cached and cached.get("sig") == signature:
                data = cached.get("data")
            else:
                data = parse_internal(path)
                if not data:
                    continue
                entries[key] = {"sig": signature, "data": data}
            if not data or not data.get("key"):
                continue
            match = re.match(r"^(\d+)", path.name)
            leading = int(match.group(1)) if match else 0
            try:
                modified = path.stat().st_mtime
            except OSError:
                modified = 0
            candidates.append((data["key"], leading, modified, data))

    chosen: dict[str, tuple[int, float, dict]] = {}
    for key, leading, modified, data in candidates:
        current = chosen.get(key)
        if current is None or (leading, modified) > (current[0], current[1]):
            chosen[key] = (leading, modified, data)

    months = [item[2] for item in sorted(chosen.values(), key=lambda item: item[2]["key"])]
    return {
        "months": months,
        "latest": months[-1]["key"] if months else None,
        "files": len(months),
    }


# ---------------------------------------------------------------------------
# assembly
# ---------------------------------------------------------------------------

def relevant_files() -> list[Path]:
    files = []
    for root in (OVERDUE_ROOT, CONSOLIDATE_ROOT, NPL_ROOT, INTERNAL_ROOT):
        if root.exists():
            files.extend(
                path for path in root.rglob("*.xlsx")
                if path.is_file() and not path.name.startswith("~$")
            )
    if DASHBOARD_FILE.exists():
        files.append(DASHBOARD_FILE)
    return files


def build() -> dict:
    started = time.time()
    cache = load_cache()
    loan = build_loan_listing()
    npl = build_npl_months()
    overdue = build_overdue(cache)
    consolidate = build_consolidate(cache)
    internal = build_internal(cache)
    save_cache(cache)

    payload = {
        "generatedAt": dt.datetime.now().isoformat(timespec="seconds"),
        "generatedAtLabel": dt.datetime.now().strftime("%d %b %Y, %H:%M"),
        "rateKhrPerUsd": RATE_KHR_PER_USD,
        "loan": loan,
        "npl": npl,
        "overdue": overdue,
        "consolidate": consolidate,
        "internal": internal,
        "sources": {
            "loan": "Data Dashboard.xlsx",
            "npl": "NPL Movement",
            "overdue": "Credit Recovery Workflow / Daily Overdue Data",
            "consolidate": "Credit Recovery Workflow / Daily Consolidate for Overdue Report",
            "internal": "Internal Report Department",
        },
    }

    serialised = json.dumps(payload, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    serialised = serialised.replace("</", "<\\/")
    DATA_JS.write_text(
        "/* Generated automatically by automation/build_data.py. Do not edit. */\n"
        f"window.DASHBOARD_DATA = {serialised};\n",
        encoding="utf-8",
    )
    log(
        "data.js written "
        f"({DATA_JS.stat().st_size / 1024 / 1024:.2f} MB, {time.time() - started:.1f}s) · "
        f"{len(loan.get('rows', []))} loan rows · {len(npl.get('months', []))} NPL months · "
        f"{overdue.get('files', 0)} overdue snapshots · {consolidate.get('files', 0)} consolidate days · "
        f"{internal.get('files', 0)} internal months"
    )
    return payload


def file_signature() -> str:
    parts = []
    for path in sorted(relevant_files()):
        parts.append(f"{path}:{sig(path)}")
    return "|".join(parts)


def watch(interval: int) -> None:
    log("Watching folders. Leave this window open while using the dashboard.")
    last = None
    while True:
        current = file_signature()
        if current != last:
            try:
                build()
            except Exception as exc:  # keep watching after a bad workbook
                log(f"Build failed: {exc}")
            last = current
        time.sleep(interval)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build Credit Recovery dashboard data")
    parser.add_argument("--watch", action="store_true", help="keep watching folders for changes")
    parser.add_argument("--interval", type=int, default=15, help="seconds between folder checks")
    args = parser.parse_args()

    if args.watch:
        watch(max(5, args.interval))
    else:
        build()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(0)
