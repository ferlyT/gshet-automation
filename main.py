# =========================================================
# GOOGLE SHEETS PROFESSIONAL DASHBOARD SYNC
# v4: Data sheets (Udara/Laut) + Dashboard sheet per moda
# =========================================================

from __future__ import annotations

import os
import re
import time
import random
import logging
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any
from urllib.parse import quote_plus

import pandas as pd
import sqlalchemy as sa
import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError

# =========================================================
# BOOTSTRAP
# =========================================================

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger(__name__)


# =========================================================
# CONFIG
# =========================================================

@dataclass(frozen=True)
class Config:
    db_driver:  str = field(default_factory=lambda: os.getenv("DB_DRIVER", "{ODBC Driver 17 for SQL Server}"))
    db_server:  str = field(default_factory=lambda: os.getenv("DB_SERVER", ""))
    db_name:    str = field(default_factory=lambda: os.getenv("DB_NAME", ""))
    db_user:    str = field(default_factory=lambda: os.getenv("DB_USER", ""))
    db_pass:    str = field(default_factory=lambda: os.getenv("DB_PASS", ""))

    sheet_input:      str = field(default_factory=lambda: os.getenv("GOOGLE_SHEET_ID", ""))
    credentials_file: str = field(default_factory=lambda: os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json"))

    default_rows: int = 5000
    default_cols: int = 50
    chunk_size:   int = 5000

    status_columns: tuple[str, ...] = ("Status", "StatusKirim", "PaymentStatus")

    @property
    def sheet_id(self) -> str:
        raw = self.sheet_input
        if "/d/" in raw:
            m = re.search(r"/d/([a-zA-Z0-9-_]+)", raw)
            if m:
                return m.group(1)
        return raw

    @property
    def db_url(self) -> str:
        conn = (
            f"DRIVER={self.db_driver};"
            f"SERVER={self.db_server};"
            f"DATABASE={self.db_name};"
            f"UID={self.db_user};"
            f"PWD={self.db_pass};"
        )
        return "mssql+pyodbc:///?odbc_connect=" + quote_plus(conn)


CFG = Config()


# =========================================================
# RETRY
# =========================================================

class RetryHandler:
    MAX_RETRIES = 6
    RETRYABLE   = ("429", "500")

    @classmethod
    def run(cls, func, *args, **kwargs):
        for attempt in range(cls.MAX_RETRIES):
            try:
                return func(*args, **kwargs)
            except APIError as exc:
                if not any(code in str(exc) for code in cls.RETRYABLE):
                    raise
                wait = min((2 ** attempt) + random.uniform(1, 4), 60)
                logger.warning("API error – retry in %.1fs (%d/%d)", wait, attempt + 1, cls.MAX_RETRIES)
                time.sleep(wait)
        raise RuntimeError("Maximum retries exceeded")


# =========================================================
# STATUS FORMATTING
# =========================================================

_STATUS_ICONS: dict[str, str] = {
    "NO STATUS": "⚪",
    "WARNING":   "⚠",
    "URGENT":    "🚨",
    "COD":       "💰",
    "OK":        "✅",
    "PAID":      "🟢",
    "PROCESS":   "🟡",
    "PENDING":   "🟡",
    "CANCEL":    "🔴",
    "HOLD":      "🟠",
}


def _format_status(value) -> str:
    if pd.isna(value):
        return "⚪ NO STATUS"
    text = str(value).strip()
    if not text:
        return "⚪ NO STATUS"
    icon = _STATUS_ICONS.get(text.upper(), "⚪")
    return f"{icon} {text}"


def apply_status_icons(df: pd.DataFrame) -> pd.DataFrame:
    for col in CFG.status_columns:
        if col in df.columns:
            df[col] = df[col].apply(_format_status)
    return df


# =========================================================
# DATAFRAME HELPERS
# =========================================================

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

    if "StatusKirim" in df.columns:
        df["StatusKirim"] = (
            df["StatusKirim"]
            .apply(lambda x: str(x).strip() if pd.notna(x) else x)
            .replace(["nan", "None"], "")
        )

    for col in df.select_dtypes(include=["datetime", "datetimetz"]).columns:
        df[col] = df[col].dt.strftime("%d-%m-%Y")

    df = apply_status_icons(df)
    return df


def build_sheet_values(df: pd.DataFrame) -> list[list[Any]]:
    col_count  = len(df.columns)
    updated_at = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    rows: list[list[Any]] = [
        [f"Last Update: {updated_at}"] + [""] * (col_count - 1),
        [""] * col_count,
        df.columns.tolist(),
    ]
    rows.extend(df.fillna("").values.tolist())
    return rows


# =========================================================
# DASHBOARD BUILDER
# =========================================================

def _simplify_status_kirim(s: str) -> str:
    s = str(s)
    if "BELUM DITERIMA" in s: return "Belum Diterima GDG"
    if "BELUM DIKIRIM"  in s: return "Belum Dikirim"
    if "TERKIRIM"       in s: return "Terkirim"
    if "TERIMA SJ"      in s: return "Terima SJ"
    if "SJ BATAL"       in s: return "SJ Batal"
    if "PROSES"         in s: return "Proses Kirim"
    return "Lainnya"


def _strip_icon(s: str) -> str:
    """Remove leading emoji + space from status strings."""
    return re.sub(r"^[\U00010000-\U0010ffff\u2600-\u27BF\u2B50\u2B55\u231A-\u231B]\s*", "", str(s)).strip()


def build_dashboard_values(df: pd.DataFrame, moda: str) -> list[list[Any]]:
    """
    Build a fully self-contained dashboard layout for one moda (Udara/Laut).

    Layout (row numbers, 1-based):
      1        – Title bar
      2        – blank
      3-4      – KPI headers + values  (5 KPI columns side by side)
      5        – blank
      6        – Section: Status Pengiriman  |  Section: Status Kirim
      7        – Header row for both tables
      8..N     – Data rows (longest of the two sections)
      N+1      – blank
      N+2      – Section: Top Cabang by Berat
      N+3      – Header
      N+4..M   – Data
      M+1      – blank
      M+2      – Section: Tipe Barang
      M+3      – Header
      M+4..P   – Data
      P+1      – blank
      P+2      – Section: Antrian Terlama (SEMUA STATUS)
      P+3      – Header
      P+4..    – Data rows (top 15 by hari desc, all statuses)
    """
    updated_at = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    rows: list[list[Any]] = []

    total_shipment = len(df)
    total_berat    = df["Berat"].sum() if "Berat" in df.columns else 0
    total_pack     = int(df["Jml_pack"].sum()) if "Jml_pack" in df.columns else 0
    urgent_count   = int((df["Status"] == "🚨 URGENT").sum()) if "Status" in df.columns else 0
    avg_hari       = df["hari"].mean() if "hari" in df.columns else 0

    # ── Row 1: Title ──────────────────────────────────────────────────
    rows.append([f"📊 DASHBOARD BILLING PENDING — {moda.upper()}   |   Update: {updated_at}"] + [""] * 9)

    # ── Row 2: blank ──────────────────────────────────────────────────
    rows.append([""] * 10)

    # ── Row 3: KPI labels ─────────────────────────────────────────────
    rows.append([
        "TOTAL SHIPMENT", "", "TOTAL BERAT (kg)", "", "TOTAL KOLI", "", "URGENT", "", "RATA-RATA HARI", ""
    ])

    # ── Row 4: KPI values ─────────────────────────────────────────────
    rows.append([
        total_shipment, "",
        round(total_berat, 1), "",
        total_pack, "",
        urgent_count, "",
        round(avg_hari, 1), ""
    ])

    # ── Row 5: blank ──────────────────────────────────────────────────
    rows.append([""] * 10)

    # ── Status Pengiriman ─────────────────────────────────────────────
    status_counts = {}
    if "Status" in df.columns:
        for raw, cnt in df["Status"].value_counts().items():
            status_counts[_strip_icon(raw)] = int(cnt)

    # ── Status Kirim ──────────────────────────────────────────────────
    sk_counts = {}
    if "StatusKirim" in df.columns:
        df["_sk_group"] = df["StatusKirim"].apply(_simplify_status_kirim)
        for k, v in df["_sk_group"].value_counts().items():
            sk_counts[str(k)] = int(v)

    # ── Section headers ───────────────────────────────────────────────
    rows.append(["📦 STATUS PENGIRIMAN", "", "", "", "🚚 STATUS KIRIM", "", "", "", "", ""])
    rows.append(["Status", "Jumlah", "Persentase (%)", "", "Status Kirim", "Jumlah", "Persentase (%)", "", "", ""])

    max_rows = max(len(status_counts), len(sk_counts))
    status_items = list(status_counts.items())
    sk_items     = list(sk_counts.items())

    for i in range(max_rows):
        col_a = col_b = col_c = ""
        col_e = col_f = col_g = ""

        if i < len(status_items):
            k, v = status_items[i]
            pct  = round(v / total_shipment * 100, 1) if total_shipment else 0
            col_a, col_b, col_c = k, v, pct

        if i < len(sk_items):
            k, v = sk_items[i]
            pct  = round(v / total_shipment * 100, 1) if total_shipment else 0
            col_e, col_f, col_g = k, v, pct

        rows.append([col_a, col_b, col_c, "", col_e, col_f, col_g, "", "", ""])

    # ── blank ─────────────────────────────────────────────────────────
    rows.append([""] * 10)

    # ── Top Cabang ────────────────────────────────────────────────────
    rows.append(["🏢 TOP CABANG BY BERAT (kg)", "", "", "", "", "", "", "", "", ""])
    rows.append(["Cabang", "Jumlah Shipment", "Total Berat (kg)", "Persentase (%)"] + [""] * 6)

    if "Branch" in df.columns:
        branch_berat = df.groupby("Branch")["Berat"].sum().sort_values(ascending=False)
        branch_count = df["Branch"].value_counts()
        for branch, berat in branch_berat.items():
            cnt = int(branch_count.get(branch, 0))
            pct = round(berat / total_berat * 100, 1) if total_berat else 0
            rows.append([str(branch), cnt, round(berat, 1), pct] + [""] * 6)

    # ── blank ─────────────────────────────────────────────────────────
    rows.append([""] * 10)

    # ── Tipe Barang ───────────────────────────────────────────────────
    rows.append(["📋 TIPE BARANG", "", "", "", "", "", "", "", "", ""])
    rows.append(["Tipe", "Jumlah", "Persentase (%)"] + [""] * 7)

    if "Type" in df.columns:
        for tipe, cnt in df["Type"].value_counts().items():
            pct = round(cnt / total_shipment * 100, 1) if total_shipment else 0
            rows.append([str(tipe), int(cnt), pct] + [""] * 7)

    # ── blank ─────────────────────────────────────────────────────────
    rows.append([""] * 10)

    # ── Antrian Terlama — SEMUA STATUS ────────────────────────────────
    # Sebelumnya hanya filter URGENT; sekarang tampilkan semua status,
    # diurutkan by hari descending, top 15.
    rows.append(["⏳ ANTRIAN TERLAMA (SEMUA STATUS)", "", "", "", "", "", "", "", "", ""])
    rows.append(["Customer", "Cabang", "Hari Pending", "Berat (kg)", "Status", "Status Kirim", "", "", "", ""])

    if "hari" in df.columns:
        cols_needed = ["Customer", "Branch", "hari", "Berat", "StatusKirim"]
        if "Status" in df.columns:
            cols_needed = ["Customer", "Branch", "hari", "Berat", "Status", "StatusKirim"]

        antrian_df = (
            df[cols_needed]
            .sort_values("hari", ascending=False)
            .head(15)
        )
        for _, r in antrian_df.iterrows():
            status_val = _strip_icon(str(r["Status"])) if "Status" in antrian_df.columns else ""
            rows.append([
                str(r["Customer"]),
                str(r["Branch"]),
                int(r["hari"])              if pd.notna(r["hari"])  else 0,
                round(float(r["Berat"]), 1) if pd.notna(r["Berat"]) else 0,
                status_val,
                _strip_icon(str(r["StatusKirim"])),
                "", "", "", ""
            ])

    return rows


# =========================================================
# SHEETS REQUEST BUILDERS
# =========================================================

def col_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


_STATUS_CF_RULES: list[tuple[str, dict, bool]] = [
    ("NO STATUS", {"red": 0.90, "green": 0.90, "blue": 0.90}, False),
    ("WARNING",   {"red": 0.96, "green": 0.80, "blue": 0.80}, False),
    ("URGENT",    {"red": 0.88, "green": 0.40, "blue": 0.40}, True),
    ("COD",       {"red": 1.0,  "green": 0.88, "blue": 0.72}, False),
    ("OK",        {"red": 0.82, "green": 0.93, "blue": 0.84}, False),
    ("PAID",      {"red": 0.80, "green": 0.94, "blue": 0.84}, False),
    ("PROCESS",   {"red": 1.0,  "green": 0.96, "blue": 0.75}, False),
    ("PENDING",   {"red": 1.0,  "green": 0.96, "blue": 0.75}, False),
    ("HOLD",      {"red": 1.0,  "green": 0.87, "blue": 0.70}, False),
    ("CANCEL",    {"red": 0.95, "green": 0.75, "blue": 0.75}, False),
]

_LIGHT_TEXT = {"red": 0.15, "green": 0.15, "blue": 0.15}
_WHITE_TEXT  = {"red": 1.0,  "green": 1.0,  "blue": 1.0}


def _cf_range(sheet_id: int, col_idx: int, end_row: int = None) -> dict:
    return {
        "sheetId":          sheet_id,
        "startRowIndex":    3,
        "endRowIndex":      end_row or CFG.default_rows,
        "startColumnIndex": col_idx,
        "endColumnIndex":   col_idx + 1,
    }


def build_status_cf_requests(sheet_id: int, col_idx: int) -> list[dict]:
    letter   = col_letter(col_idx + 1)
    requests = []
    for keyword, bg, dark in _STATUS_CF_RULES:
        formula    = f'=REGEXMATCH(${letter}4,"{keyword}")'
        text_color = _WHITE_TEXT if dark else _LIGHT_TEXT
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [_cf_range(sheet_id, col_idx)],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": formula}]},
                        "format": {
                            "backgroundColor": bg,
                            "textFormat": {"bold": True, "foregroundColor": text_color},
                        },
                    },
                },
                "index": 0,
            }
        })
    return requests


def build_row_highlight_request(sheet_id: int, col_count: int) -> dict:
    return {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{"sheetId": sheet_id, "startRowIndex": 3, "endRowIndex": CFG.default_rows,
                            "startColumnIndex": 0, "endColumnIndex": col_count}],
                "booleanRule": {
                    "condition": {"type": "CUSTOM_FORMULA",
                                  "values": [{"userEnteredValue": '=ISNUMBER(SEARCH("BELUM DITERIMA GDG",$P4))'}]},
                    "format": {
                        "backgroundColor": {"red": 1.0, "green": 0.85, "blue": 0.85},
                        "textFormat": {"bold": True, "foregroundColor": {"red": 0.5, "green": 0.0, "blue": 0.0}},
                    },
                },
            },
            "index": 0,
        }
    }


def build_gradient_request(sheet_id: int, col_idx: int) -> dict:
    return {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [_cf_range(sheet_id, col_idx)],
                "gradientRule": {
                    "minpoint": {"color": {"red": 0.85, "green": 0.93, "blue": 0.83}, "type": "NUMBER", "value": "0"},
                    "midpoint": {"color": {"red": 1.0,  "green": 1.0,  "blue": 0.70}, "type": "NUMBER", "value": "4"},
                    "maxpoint": {"color": {"red": 0.95, "green": 0.40, "blue": 0.40}, "type": "NUMBER", "value": "7"},
                },
            },
            "index": 0,
        }
    }


def build_freeze_request(sheet_id: int, frozen_rows: int = 2) -> dict:
    return {
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": frozen_rows}},
            "fields": "gridProperties.frozenRowCount",
        }
    }


def build_header_style_request(sheet_id: int) -> dict:
    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                    "textFormat": {"bold": True, "fontSize": 9,
                                   "foregroundColor": {"red": 0.3, "green": 0.3, "blue": 0.3}},
                    "verticalAlignment": "MIDDLE",
                }
            },
            "fields": "userEnteredFormat",
        }
    }


def build_add_table_request(sheet_id: int, col_count: int, row_count: int, table_name: str) -> dict:
    end_row = max(3 + row_count, 4)
    return {
        "addTable": {
            "table": {
                "name":  table_name,
                "range": {
                    "sheetId":          sheet_id,
                    "startRowIndex":    2,
                    "endRowIndex":      end_row,
                    "startColumnIndex": 0,
                    "endColumnIndex":   col_count,
                },
            }
        }
    }


def build_auto_resize_request(sheet_id: int, col_count: int) -> dict:
    return {
        "autoResizeDimensions": {
            "dimensions": {"sheetId": sheet_id, "dimension": "COLUMNS",
                           "startIndex": 0, "endIndex": col_count}
        }
    }


def build_set_col_width_request(sheet_id: int, col_idx: int, width_px: int) -> dict:
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": col_idx, "endIndex": col_idx + 1},
            "properties": {"pixelSize": width_px},
            "fields": "pixelSize",
        }
    }


# ──────────────────────────────────────────────────────────
# Dashboard-specific request builders
# ──────────────────────────────────────────────────────────

_DARK_BG    = {"red": 0.13, "green": 0.20, "blue": 0.35}   # navy title bar
_WHITE_FG   = {"red": 1.0,  "green": 1.0,  "blue": 1.0}
_SECTION_BG = {"red": 0.23, "green": 0.35, "blue": 0.58}   # section header blue
_KPI_BG     = {"red": 0.91, "green": 0.95, "blue": 1.0}    # light blue KPI bg
_KPI_VAL_FG = {"red": 0.10, "green": 0.20, "blue": 0.50}   # dark blue KPI value
_URGENT_FG  = {"red": 0.55, "green": 0.05, "blue": 0.05}
_HEADER_BG  = {"red": 0.87, "green": 0.91, "blue": 0.97}
_HEADER_FG  = {"red": 0.10, "green": 0.20, "blue": 0.45}


def _rng(sheet_id, r1, r2, c1, c2):
    return {"sheetId": sheet_id, "startRowIndex": r1, "endRowIndex": r2,
            "startColumnIndex": c1, "endColumnIndex": c2}


def _fmt(bg=None, fg=None, bold=False, size=10, halign=None, valign=None, wrap=None):
    f: dict[str, Any] = {"textFormat": {"bold": bold, "fontSize": size}}
    if fg:
        f["textFormat"]["foregroundColor"] = fg
    if bg:
        f["backgroundColor"] = bg
    if halign:
        f["horizontalAlignment"] = halign
    if valign:
        f["verticalAlignment"] = valign
    if wrap:
        f["wrapStrategy"] = wrap
    return f


def _repeat(sheet_id, r1, r2, c1, c2, fmt_dict, fields="userEnteredFormat"):
    return {
        "repeatCell": {
            "range": _rng(sheet_id, r1, r2, c1, c2),
            "cell": {"userEnteredFormat": fmt_dict},
            "fields": fields,
        }
    }


def _merge(sheet_id, r1, r2, c1, c2):
    return {"mergeCells": {"range": _rng(sheet_id, r1, r2, c1, c2), "mergeType": "MERGE_ALL"}}


# Warna background per keyword status untuk tabel antrian
_STATUS_CF_ANTRIAN: list[tuple[str, dict]] = [
    ("URGENT",    {"red": 0.95, "green": 0.82, "blue": 0.82}),
    ("WARNING",   {"red": 0.96, "green": 0.80, "blue": 0.80}),
    ("HOLD",      {"red": 1.0,  "green": 0.87, "blue": 0.70}),
    ("COD",       {"red": 1.0,  "green": 0.88, "blue": 0.72}),
    ("PENDING",   {"red": 1.0,  "green": 0.96, "blue": 0.75}),
    ("PROCESS",   {"red": 1.0,  "green": 0.96, "blue": 0.75}),
    ("OK",        {"red": 0.82, "green": 0.93, "blue": 0.84}),
    ("PAID",      {"red": 0.80, "green": 0.94, "blue": 0.84}),
    ("CANCEL",    {"red": 0.95, "green": 0.75, "blue": 0.75}),
    ("NO STATUS", {"red": 0.93, "green": 0.93, "blue": 0.93}),
]


def build_dashboard_style_requests(sheet_id: int, total_rows: int,
                                   section_rows: dict) -> list[dict]:
    """
    section_rows keys (all 0-based row indices):
      kpi_label_row, kpi_val_row,
      status_section_row, status_header_row, status_data_start, status_data_end,
      branch_section_row, branch_header_row, branch_data_start, branch_data_end,
      type_section_row, type_header_row, type_data_start, type_data_end,
      antrian_section_row, antrian_header_row, antrian_data_start, antrian_data_end,
    """
    reqs = []
    sr = section_rows

    # ── Title bar (row 0) ─────────────────────────────────────────────
    reqs.append(_repeat(sheet_id, 0, 1, 0, 10,
                        _fmt(bg=_DARK_BG, fg=_WHITE_FG, bold=True, size=11, valign="MIDDLE")))
    reqs.append(_merge(sheet_id, 0, 1, 0, 10))

    # ── KPI label row ─────────────────────────────────────────────────
    for c in range(0, 10, 2):
        reqs.append(_repeat(sheet_id, sr["kpi_label_row"], sr["kpi_label_row"] + 1, c, c + 1,
                            _fmt(bg=_KPI_BG, fg=_HEADER_FG, bold=True, size=9,
                                 halign="CENTER", valign="MIDDLE")))

    # ── KPI value row ─────────────────────────────────────────────────
    for c in range(0, 10, 2):
        reqs.append(_repeat(sheet_id, sr["kpi_val_row"], sr["kpi_val_row"] + 1, c, c + 1,
                            _fmt(bg=_KPI_BG, fg=_KPI_VAL_FG, bold=True, size=16,
                                 halign="CENTER", valign="MIDDLE")))
    # Urgent KPI value → merah
    reqs.append(_repeat(sheet_id, sr["kpi_val_row"], sr["kpi_val_row"] + 1, 6, 7,
                        _fmt(bg={"red": 1.0, "green": 0.92, "blue": 0.92},
                             fg=_URGENT_FG, bold=True, size=16,
                             halign="CENTER", valign="MIDDLE")))

    # ── Row heights for KPI ───────────────────────────────────────────
    for r in [sr["kpi_label_row"], sr["kpi_val_row"]]:
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "ROWS",
                          "startIndex": r, "endIndex": r + 1},
                "properties": {"pixelSize": 36},
                "fields": "pixelSize",
            }
        })

    def _section_header(row, col_end=10):
        reqs.append(_repeat(sheet_id, row, row + 1, 0, col_end,
                            _fmt(bg=_SECTION_BG, fg=_WHITE_FG, bold=True, size=10, valign="MIDDLE")))
        reqs.append(_merge(sheet_id, row, row + 1, 0, col_end))
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "ROWS",
                          "startIndex": row, "endIndex": row + 1},
                "properties": {"pixelSize": 28},
                "fields": "pixelSize",
            }
        })

    def _table_header(row, col_end):
        reqs.append(_repeat(sheet_id, row, row + 1, 0, col_end,
                            _fmt(bg=_HEADER_BG, fg=_HEADER_FG, bold=True, size=9,
                                 halign="CENTER", valign="MIDDLE")))

    def _zebra(data_start, data_end, col_end):
        for r in range(data_start, data_end):
            if (r - data_start) % 2 == 0:
                reqs.append(_repeat(sheet_id, r, r + 1, 0, col_end,
                                    _fmt(bg={"red": 0.97, "green": 0.98, "blue": 1.0})))

    # ── Status section ────────────────────────────────────────────────
    _section_header(sr["status_section_row"])
    _table_header(sr["status_header_row"], 8)
    _zebra(sr["status_data_start"], sr["status_data_end"], 8)

    # ── Branch section ────────────────────────────────────────────────
    _section_header(sr["branch_section_row"])
    _table_header(sr["branch_header_row"], 5)
    _zebra(sr["branch_data_start"], sr["branch_data_end"], 5)

    # ── Type section ──────────────────────────────────────────────────
    _section_header(sr["type_section_row"])
    _table_header(sr["type_header_row"], 4)
    _zebra(sr["type_data_start"], sr["type_data_end"], 4)

    # ── Antrian Terlama section ───────────────────────────────────────
    _section_header(sr["antrian_section_row"])
    _table_header(sr["antrian_header_row"], 6)

    # Conditional format per baris berdasarkan nilai kolom Status (col E, 0-based=4).
    # Formula memakai referensi absolut ke kolom E dengan baris pertama data.
    antrian_range_all = _rng(sheet_id,
                             sr["antrian_data_start"],
                             sr["antrian_data_end"],
                             0, 6)
    status_col_letter = col_letter(5)   # col index 4 (0-based) = kolom E

    for keyword, bg in _STATUS_CF_ANTRIAN:
        formula = (
            f'=REGEXMATCH(UPPER(${status_col_letter}'
            f'{sr["antrian_data_start"] + 1}),"{keyword}")'
        )
        reqs.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [antrian_range_all],
                    "booleanRule": {
                        "condition": {
                            "type":   "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": formula}],
                        },
                        "format": {"backgroundColor": bg},
                    },
                },
                "index": 0,
            }
        })

    # Kolom Hari Pending (col C, 0-based=2) → bold + center
    reqs.append(_repeat(sheet_id,
                        sr["antrian_data_start"], sr["antrian_data_end"],
                        2, 3,
                        _fmt(bold=True, halign="CENTER")))

    # ── Freeze row 1 (title) ──────────────────────────────────────────
    reqs.append({
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
            "fields": "gridProperties.frozenRowCount",
        }
    })

    return reqs


# =========================================================
# GOOGLE SHEETS CLIENT
# =========================================================

_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


class GoogleSheetsClient:
    def __init__(self):
        creds            = Credentials.from_service_account_file(CFG.credentials_file, scopes=_SCOPES)
        self.client      = gspread.authorize(creds)
        self.spreadsheet = self.client.open_by_key(CFG.sheet_id)
        self._styled:    set[str] = set()

    # ------------------------------------------------------------------
    # Sheet management
    # ------------------------------------------------------------------

    def get_or_create_sheet(self, name: str, index: int | None = None) -> gspread.Worksheet:
        try:
            return self.spreadsheet.worksheet(name)
        except gspread.exceptions.WorksheetNotFound:
            logger.info("Creating worksheet: %s", name)
            kwargs: dict[str, Any] = {"title": name, "rows": CFG.default_rows, "cols": CFG.default_cols}
            if index is not None:
                kwargs["index"] = index
            return self.spreadsheet.add_worksheet(**kwargs)

    def _recreate_sheet(self, worksheet: gspread.Worksheet) -> gspread.Worksheet:
        name  = worksheet.title
        index = worksheet.index
        logger.info("Recreating worksheet '%s' to clear all tables/formatting...", name)
        self.spreadsheet.del_worksheet(worksheet)
        time.sleep(2)
        new_ws = self.spreadsheet.add_worksheet(
            title=name, rows=CFG.default_rows, cols=CFG.default_cols, index=index
        )
        logger.info("Worksheet '%s' recreated (sheet_id=%s)", name, new_ws.id)
        time.sleep(1)
        return new_ws

    # ------------------------------------------------------------------
    # Batch helper
    # ------------------------------------------------------------------

    def _batch(self, requests: list[dict]):
        if requests:
            RetryHandler.run(self.spreadsheet.batch_update, {"requests": requests})

    # ------------------------------------------------------------------
    # Data sheet style
    # ------------------------------------------------------------------

    def ensure_data_style(self, worksheet: gspread.Worksheet, df: pd.DataFrame) -> gspread.Worksheet:
        if worksheet.title in self._styled:
            return worksheet

        logger.info("Applying data style: %s", worksheet.title)
        worksheet = self._recreate_sheet(worksheet)
        sheet_id  = worksheet.id
        col_count = len(df.columns)
        row_count = len(df)

        add_requests: list[dict] = []
        add_requests.append(build_freeze_request(sheet_id))
        add_requests.append(build_header_style_request(sheet_id))
        add_requests.append(build_add_table_request(sheet_id, col_count, row_count, worksheet.title))

        if "hari" in df.columns:
            add_requests.append(build_gradient_request(sheet_id, df.columns.get_loc("hari")))

        add_requests.append(build_row_highlight_request(sheet_id, col_count))

        for status_col in CFG.status_columns:
            if status_col in df.columns:
                add_requests.extend(build_status_cf_requests(sheet_id, df.columns.get_loc(status_col)))

        self._batch(add_requests)
        self._styled.add(worksheet.title)
        return worksheet

    # ------------------------------------------------------------------
    # Dashboard sheet
    # ------------------------------------------------------------------

    def sync_dashboard(self, df: pd.DataFrame, moda: str):
        """Build and upload the dashboard sheet for one moda."""
        dash_name = f"📊 Dashboard {moda}"
        logger.info("Building dashboard: %s", dash_name)

        values = build_dashboard_values(df, moda)

        # ── Hitung section row indices (0-based) ──────────────────────
        status_counts = df["Status"].value_counts() if "Status" in df.columns else pd.Series([], dtype=int)
        sk_counts_len = df["StatusKirim"].apply(_simplify_status_kirim).value_counts().shape[0] \
                        if "StatusKirim" in df.columns else 0
        max_status_rows = max(len(status_counts), sk_counts_len)

        branch_rows = df["Branch"].nunique() if "Branch" in df.columns else 0
        type_rows   = df["Type"].nunique()   if "Type"   in df.columns else 0

        # CHANGED: semua baris (bukan hanya URGENT), max 15
        antrian_rows = min(len(df), 15) if "hari" in df.columns else 0

        r = 0
        kpi_label_row = r + 2
        kpi_val_row   = r + 3
        status_sec    = r + 5
        status_hdr    = r + 6
        status_ds     = r + 7
        status_de     = status_ds + max_status_rows
        branch_sec    = status_de + 1
        branch_hdr    = branch_sec + 1
        branch_ds     = branch_hdr + 1
        branch_de     = branch_ds + branch_rows
        type_sec      = branch_de + 1
        type_hdr      = type_sec + 1
        type_ds       = type_hdr + 1
        type_de       = type_ds + type_rows
        antrian_sec   = type_de + 1
        antrian_hdr   = antrian_sec + 1
        antrian_ds    = antrian_hdr + 1
        antrian_de    = antrian_ds + antrian_rows

        section_rows = {
            "kpi_label_row":       kpi_label_row,
            "kpi_val_row":         kpi_val_row,
            "status_section_row":  status_sec,
            "status_header_row":   status_hdr,
            "status_data_start":   status_ds,
            "status_data_end":     status_de,
            "branch_section_row":  branch_sec,
            "branch_header_row":   branch_hdr,
            "branch_data_start":   branch_ds,
            "branch_data_end":     branch_de,
            "type_section_row":    type_sec,
            "type_header_row":     type_hdr,
            "type_data_start":     type_ds,
            "type_data_end":       type_de,
            # Renamed: urgent_* → antrian_*
            "antrian_section_row": antrian_sec,
            "antrian_header_row":  antrian_hdr,
            "antrian_data_start":  antrian_ds,
            "antrian_data_end":    antrian_de,
        }

        # ── Recreate sheet (clean state) ──────────────────────────────
        try:
            ws = self.spreadsheet.worksheet(dash_name)
            ws = self._recreate_sheet(ws)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(
                title=dash_name, rows=max(len(values) + 10, 100), cols=12
            )

        # ── Upload data ───────────────────────────────────────────────
        logger.info("Uploading dashboard values (%d rows)...", len(values))
        RetryHandler.run(ws.update, values=values, range_name="A1")
        time.sleep(2)

        # ── Apply formatting ──────────────────────────────────────────
        style_reqs = build_dashboard_style_requests(ws.id, len(values), section_rows)
        style_reqs.append(build_auto_resize_request(ws.id, 10))

        # Set min width for key columns
        style_reqs.append(build_set_col_width_request(ws.id, 0, 220))   # Customer/Status
        style_reqs.append(build_set_col_width_request(ws.id, 1, 130))   # Cabang/Jumlah
        style_reqs.append(build_set_col_width_request(ws.id, 2, 110))   # Hari/Berat
        style_reqs.append(build_set_col_width_request(ws.id, 4, 160))   # Status col
        style_reqs.append(build_set_col_width_request(ws.id, 5, 200))   # Status Kirim col

        self._batch(style_reqs)
        logger.info("✅ Dashboard synced: %s", dash_name)

    # ------------------------------------------------------------------
    # Column sizing
    # ------------------------------------------------------------------

    def resize_columns(self, worksheet: gspread.Worksheet, df: pd.DataFrame):
        sheet_id  = worksheet.id
        col_count = len(df.columns)
        requests  = [build_auto_resize_request(sheet_id, col_count)]
        if "hari" in df.columns:
            requests.append(build_set_col_width_request(sheet_id, df.columns.get_loc("hari"), 30))
        logger.info("Resizing %d columns in '%s'", col_count, worksheet.title)
        self._batch(requests)

    # ------------------------------------------------------------------
    # Data upload
    # ------------------------------------------------------------------

    def upload_values(self, worksheet: gspread.Worksheet, values: list[list[Any]]):
        total = len(values)
        logger.info("Uploading %d rows → '%s'", total, worksheet.title)
        for start in range(0, total, CFG.chunk_size):
            chunk  = values[start: start + CFG.chunk_size]
            row_a1 = start + 1
            logger.info("  chunk rows %d – %d", row_a1, row_a1 + len(chunk) - 1)
            RetryHandler.run(worksheet.update, values=chunk, range_name=f"A{row_a1}")
            time.sleep(2)

    # ------------------------------------------------------------------
    # Main sync (data sheet)
    # ------------------------------------------------------------------

    def sync(self, df: pd.DataFrame, sheet_name: str):
        if df.empty:
            logger.warning("Empty dataframe, skipping: %s", sheet_name)
            return

        ws = self.get_or_create_sheet(sheet_name)
        ws = self.ensure_data_style(ws, df)
        self.upload_values(ws, build_sheet_values(df))
        self.resize_columns(ws, df)
        logger.info("✅ Data sheet synced: %s", sheet_name)
        time.sleep(3)


# =========================================================
# DATABASE CLIENT
# =========================================================

class DatabaseClient:
    def __init__(self):
        self.engine = sa.create_engine(CFG.db_url, pool_pre_ping=True)

    def fetch(self, query: str) -> pd.DataFrame:
        try:
            logger.info("Running query...")
            df = pd.read_sql(sa.text(query), self.engine)
            df = clean_dataframe(df)
            logger.info("Fetched %d rows", len(df))
            return df
        except Exception:
            logger.exception("Database fetch failed")
            return pd.DataFrame()


# =========================================================
# APP
# =========================================================

@dataclass
class Task:
    sheet_name: str   # raw data sheet name  (e.g. "Udara")
    query:      str   # SQL / stored proc


class BillingSyncApp:
    def __init__(self):
        self.db = DatabaseClient()
        self.gs = GoogleSheetsClient()

    def run(self, tasks: list[Task]):
        for task in tasks:
            logger.info("═══ Task: %s ═══", task.sheet_name)
            try:
                df = self.db.fetch(task.query)

                # 1. Sync raw data sheet
                self.gs.sync(df, task.sheet_name)

                # 2. Build & upload dashboard sheet
                self.gs.sync_dashboard(df, task.sheet_name)

            except Exception:
                logger.exception("Task failed: %s", task.sheet_name)


# =========================================================
# ENTRY POINT
# =========================================================

TASKS = [
    Task(sheet_name="Udara", query="EXEC get_data_billing_gsheet 1"),
    Task(sheet_name="Laut",  query="EXEC get_data_billing_gsheet 2"),
]


def main():
    BillingSyncApp().run(TASKS)


if __name__ == "__main__":
    main()