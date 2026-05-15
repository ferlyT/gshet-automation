# =========================================================
# GOOGLE SHEETS PROFESSIONAL DASHBOARD SYNC
# REFACTORED VERSION + Native Table (addTable API)
# =========================================================

import os
import re
import time
import random
import logging
from datetime import datetime
from typing import List, Dict, Any

import pandas as pd
import sqlalchemy as sa
import gspread
from dotenv import load_dotenv
from urllib.parse import quote_plus
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError

# =========================================================
# INIT
# =========================================================

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)
logger = logging.getLogger(__name__)


# =========================================================
# CONFIG
# =========================================================

class Config:
    DB_DRIVER  = os.getenv("DB_DRIVER", "{ODBC Driver 17 for SQL Server}")
    DB_SERVER  = os.getenv("DB_SERVER")
    DB_NAME    = os.getenv("DB_NAME")
    DB_USER    = os.getenv("DB_USER")
    DB_PASS    = os.getenv("DB_PASS")

    GOOGLE_SHEET_INPUT      = os.getenv("GOOGLE_SHEET_ID")
    GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")

    DEFAULT_ROWS = 5000
    DEFAULT_COLS = 50
    CHUNK_SIZE   = 5000

    STATUS_COLUMNS = ["Status", "StatusKirim", "PaymentStatus"]

    @classmethod
    def get_sheet_id(cls):
        if "/d/" in (cls.GOOGLE_SHEET_INPUT or ""):
            match = re.search(r"/d/([a-zA-Z0-9-_]+)", cls.GOOGLE_SHEET_INPUT)
            if match:
                return match.group(1)
        return cls.GOOGLE_SHEET_INPUT

    @classmethod
    def get_db_url(cls):
        conn = (
            f"DRIVER={cls.DB_DRIVER};"
            f"SERVER={cls.DB_SERVER};"
            f"DATABASE={cls.DB_NAME};"
            f"UID={cls.DB_USER};"
            f"PWD={cls.DB_PASS};"
        )
        return "mssql+pyodbc:///?odbc_connect=" + quote_plus(conn)


# =========================================================
# RETRY HANDLER
# =========================================================

class RetryHandler:
    MAX_RETRIES = 6

    @classmethod
    def execute(cls, func, *args, **kwargs):
        for attempt in range(cls.MAX_RETRIES):
            try:
                return func(*args, **kwargs)
            except APIError as e:
                if "429" not in str(e) and "500" not in str(e):
                    raise
                wait_time = min((2 ** attempt) + random.uniform(1, 4), 60)
                logger.warning(
                    f"Quota/API error. Retry in {wait_time:.2f}s "
                    f"(Attempt {attempt + 1}/{cls.MAX_RETRIES})"
                )
                time.sleep(wait_time)
        raise Exception("Maximum retries exceeded")


# =========================================================
# STATUS FORMATTER
# =========================================================

class StatusFormatter:
    STATUS_CONFIG = {
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

    @classmethod
    def format_status(cls, value) -> str:
        if pd.isna(value):
            return "⚪ NO STATUS"
        text = str(value).strip()
        if not text:
            return "⚪ NO STATUS"
        icon = cls.STATUS_CONFIG.get(text.upper(), "⚪")
        return f"{icon} {text}"

    @classmethod
    def apply_icons(cls, df: pd.DataFrame) -> pd.DataFrame:
        for col in Config.STATUS_COLUMNS:
            if col in df.columns:
                df[col] = df[col].apply(cls.format_status)
        return df


# =========================================================
# DATAFRAME HELPER
# =========================================================

class DataFrameHelper:
    @staticmethod
    def clean(df: pd.DataFrame) -> pd.DataFrame:
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

        df = StatusFormatter.apply_icons(df)
        return df

    @staticmethod
    def build_sheet_values(df: pd.DataFrame) -> List[List[Any]]:
        """
        Layout:
          Row 1 : Last Update timestamp
          Row 2 : kosong (spacer)
          Row 3 : header kolom  <- baris pertama tabel (startRowIndex=2)
          Row 4+: data
        """
        col_count  = len(df.columns)
        updated_at = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

        values = [
            [f"Last Update: {updated_at}"] + [""] * (col_count - 1),
            [""] * col_count,
            df.columns.tolist(),
        ]
        values.extend(df.fillna("").values.tolist())
        return values


# =========================================================
# GOOGLE SHEETS CLIENT
# =========================================================

class GoogleSheetsClient:
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    # FIX: format pada addConditionalFormatRule hanya boleh berisi:
    #   backgroundColor, textFormat.bold, textFormat.italic,
    #   textFormat.strikethrough, textFormat.foregroundColor.
    # DILARANG: horizontalAlignment, verticalAlignment, fontSize, dll.
    STATUS_RULES = [
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

    def __init__(self):
        creds = Credentials.from_service_account_file(
            Config.GOOGLE_CREDENTIALS_FILE,
            scopes=self.SCOPES
        )
        self.client      = gspread.authorize(creds)
        self.spreadsheet = self.client.open_by_key(Config.get_sheet_id())
        self.styled_sheets: set = set()

    # ------------------------------------------------------------------
    # HELPERS
    # ------------------------------------------------------------------

    @staticmethod
    def column_letter(col_num: int) -> str:
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def get_or_create_sheet(self, sheet_name: str):
        try:
            return self.spreadsheet.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            logger.info(f"Creating worksheet: {sheet_name}")
            return self.spreadsheet.add_worksheet(
                title=sheet_name,
                rows=Config.DEFAULT_ROWS,
                cols=Config.DEFAULT_COLS,
            )

    # ------------------------------------------------------------------
    # CONDITIONAL FORMAT BUILDERS
    # ------------------------------------------------------------------

    def _build_status_rules(self, sheet_id: int, col_index: int) -> List[Dict]:
        col_letter = self.column_letter(col_index + 1)
        rules = []
        for keyword, bg_color, dark_mode in self.STATUS_RULES:
            formula    = f'=REGEXMATCH(${col_letter}4,"{keyword}")'
            text_color = (
                {"red": 1.0, "green": 1.0, "blue": 1.0}
                if dark_mode
                else {"red": 0.15, "green": 0.15, "blue": 0.15}
            )
            rules.append({
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{
                            "sheetId":          sheet_id,
                            "startRowIndex":    3,
                            "endRowIndex":      Config.DEFAULT_ROWS,
                            "startColumnIndex": col_index,
                            "endColumnIndex":   col_index + 1,
                        }],
                        "booleanRule": {
                            "condition": {
                                "type":   "CUSTOM_FORMULA",
                                "values": [{"userEnteredValue": formula}],
                            },
                            "format": {
                                "backgroundColor": bg_color,
                                "textFormat": {
                                    "bold":            True,
                                    "foregroundColor": text_color,
                                },
                            },
                        },
                    },
                    "index": 0,
                }
            })
        return rules

    def _build_row_highlight_rule(self, sheet_id: int, col_count: int) -> Dict:
        return {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId":          sheet_id,
                        "startRowIndex":    3,
                        "endRowIndex":      Config.DEFAULT_ROWS,
                        "startColumnIndex": 0,
                        "endColumnIndex":   col_count,
                    }],
                    "booleanRule": {
                        "condition": {
                            "type":   "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": '=ISNUMBER(SEARCH("BELUM DITERIMA GDG",$P4))'}],
                        },
                        "format": {
                            "backgroundColor": {"red": 1.0, "green": 0.85, "blue": 0.85},
                            "textFormat": {
                                "bold":            True,
                                "foregroundColor": {"red": 0.5, "green": 0.0, "blue": 0.0},
                            },
                        },
                    },
                },
                "index": 0,
            }
        }

    def _build_gradient_rule(self, sheet_id: int, col_index: int) -> Dict:
        return {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId":          sheet_id,
                        "startRowIndex":    3,
                        "endRowIndex":      Config.DEFAULT_ROWS,
                        "startColumnIndex": col_index,
                        "endColumnIndex":   col_index + 1,
                    }],
                    "gradientRule": {
                        "minpoint": {"color": {"red": 0.85, "green": 0.93, "blue": 0.83}, "type": "NUMBER", "value": "0"},
                        "midpoint": {"color": {"red": 1.0,  "green": 1.0,  "blue": 0.70}, "type": "NUMBER", "value": "4"},
                        "maxpoint": {"color": {"red": 0.95, "green": 0.40, "blue": 0.40}, "type": "NUMBER", "value": "7"},
                    },
                },
                "index": 0,
            }
        }

    # ------------------------------------------------------------------
    # NATIVE TABLE (addTable — Sheets API since April 29, 2025)
    # ------------------------------------------------------------------

    def _get_existing_tables(self, sheet_id: int) -> List[Dict]:
        """Ambil daftar tabel yang sudah ada di sheet ini via metadata."""
        try:
            metadata = RetryHandler.execute(self.spreadsheet.fetch_sheet_metadata)
            for sheet in metadata.get("sheets", []):
                if sheet.get("properties", {}).get("sheetId") == sheet_id:
                    return sheet.get("tables", [])
        except Exception:
            pass
        return []

    def _delete_existing_tables(self, sheet_id: int) -> List[Dict]:
        """
        Hapus semua tabel lama di sheet agar addTable tidak konflik.
        Returns list of deleteTable requests.
        """
        existing = self._get_existing_tables(sheet_id)
        requests = []
        for table in existing:
            table_id = table.get("tableId")
            if table_id:
                requests.append({"deleteTable": {"tableId": table_id}})
                logger.info(f"Queuing deleteTable for tableId={table_id}")
        return requests

    def _build_add_table_request(
        self,
        sheet_id:   int,
        col_count:  int,
        row_count:  int,
        table_name: str,
    ) -> Dict:
        """
        addTable request (Sheets API v4, tersedia sejak 29 Apr 2025).

        Layout di sheet:
          Row index 0 = Last Update  -> bukan bagian tabel
          Row index 1 = kosong       -> bukan bagian tabel
          Row index 2 = header tabel <- startRowIndex tabel
          Row index 3+ = data baris

        endRowIndex = 3 + row_count, minimal 4 (setidaknya 1 baris data).
        """
        end_row = max(3 + row_count, 4)

        return {
            "addTable": {
                "table": {
                    "name":       table_name,
                    "showFooter": False,
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

    # ------------------------------------------------------------------
    # STYLING ORCHESTRATOR
    # ------------------------------------------------------------------

    def ensure_style(self, worksheet, df: pd.DataFrame):
        if worksheet.title in self.styled_sheets:
            return

        logger.info(f"Applying style: {worksheet.title}")
        sheet_id  = worksheet.id
        col_count = len(df.columns)
        row_count = len(df)

        metadata      = RetryHandler.execute(self.spreadsheet.fetch_sheet_metadata)
        current_sheet = next(
            (s for s in metadata.get("sheets", [])
             if s.get("properties", {}).get("sheetId") == sheet_id),
            None,
        )
        existing_cf_count = len(current_sheet.get("conditionalFormats", [])) if current_sheet else 0

        requests: List[Dict] = []

        # 1. Hapus tabel lama agar addTable tidak konflik
        requests.extend(self._delete_existing_tables(sheet_id))

        # 2. Hapus conditional format rules lama
        for i in reversed(range(existing_cf_count)):
            requests.append({
                "deleteConditionalFormatRule": {"sheetId": sheet_id, "index": i}
            })

        # 3. Freeze 2 baris info (Last Update + spacer) di atas tabel
        requests.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId":        sheet_id,
                    "gridProperties": {"frozenRowCount": 2},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        })

        # 4. Styling baris "Last Update" (row index 0)
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                        "textFormat": {
                            "bold":            True,
                            "fontSize":        9,
                            "foregroundColor": {"red": 0.3, "green": 0.3, "blue": 0.3},
                        },
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat",
            }
        })

        # 5. Native Table — header, alternating rows, filter otomatis
        requests.append(
            self._build_add_table_request(sheet_id, col_count, row_count, worksheet.title)
        )

        # 6. Gradient rule untuk kolom 'hari'
        if "hari" in df.columns:
            hari_idx = df.columns.get_loc("hari")
            requests.append(self._build_gradient_rule(sheet_id, hari_idx))

        # 7. Row highlight "BELUM DITERIMA GDG" di kolom P
        requests.append(self._build_row_highlight_rule(sheet_id, col_count))

        # 8. Status column conditional format
        for status_col in Config.STATUS_COLUMNS:
            if status_col in df.columns:
                status_idx = df.columns.get_loc(status_col)
                requests.extend(self._build_status_rules(sheet_id, status_idx))

        if requests:
            RetryHandler.execute(
                self.spreadsheet.batch_update, {"requests": requests}
            )

        self.styled_sheets.add(worksheet.title)

    # ------------------------------------------------------------------
    # COLUMN SIZING
    # ------------------------------------------------------------------

    def auto_fit_columns(self, worksheet, col_count: int):
        logger.info(f"Auto-fitting {col_count} columns in '{worksheet.title}'...")
        RetryHandler.execute(
            self.spreadsheet.batch_update,
            {"requests": [{
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId":    worksheet.id,
                        "dimension":  "COLUMNS",
                        "startIndex": 0,
                        "endIndex":   col_count,
                    }
                }
            }]}
        )

    def set_column_width(self, worksheet, col_index: int, width: int):
        logger.info(f"Setting width {width}px for column {col_index} in '{worksheet.title}'...")
        RetryHandler.execute(
            self.spreadsheet.batch_update,
            {"requests": [{
                "updateDimensionProperties": {
                    "range": {
                        "sheetId":    worksheet.id,
                        "dimension":  "COLUMNS",
                        "startIndex": col_index,
                        "endIndex":   col_index + 1,
                    },
                    "properties": {"pixelSize": width},
                    "fields":     "pixelSize",
                }
            }]}
        )

    # ------------------------------------------------------------------
    # DATA UPLOAD
    # ------------------------------------------------------------------

    def update_in_chunks(self, worksheet, values: List[List[Any]]):
        total_rows = len(values)
        logger.info(f"Uploading {total_rows} rows to '{worksheet.title}'")

        for start in range(0, total_rows, Config.CHUNK_SIZE):
            chunk     = values[start : start + Config.CHUNK_SIZE]
            row_start = start + 1
            logger.info(f"Chunk rows {row_start} – {row_start + len(chunk) - 1}")
            RetryHandler.execute(
                worksheet.update,
                values=chunk,
                range_name=f"A{row_start}",
            )
            time.sleep(2)

    def sync_dataframe(self, df: pd.DataFrame, sheet_name: str):
        if df.empty:
            logger.warning(f"Empty dataframe, skipping: {sheet_name}")
            return

        worksheet = self.get_or_create_sheet(sheet_name)

        logger.info(f"Clearing old data in '{sheet_name}'...")
        RetryHandler.execute(worksheet.clear)
        time.sleep(1)

        # Style + native table dipasang SEBELUM data diisi
        # agar endRowIndex akurat berdasarkan len(df)
        self.ensure_style(worksheet, df)

        values = DataFrameHelper.build_sheet_values(df)
        self.update_in_chunks(worksheet, values)

        self.auto_fit_columns(worksheet, len(df.columns))

        if "hari" in df.columns:
            self.set_column_width(worksheet, df.columns.get_loc("hari"), 30)

        logger.info(f"✅ Sync success: {sheet_name}")
        time.sleep(3)


# =========================================================
# DATABASE CLIENT
# =========================================================

class DatabaseClient:
    def __init__(self):
        self.engine = sa.create_engine(Config.get_db_url(), pool_pre_ping=True)

    def fetch(self, query: str) -> pd.DataFrame:
        try:
            logger.info("Fetching SQL data...")
            df = pd.read_sql(sa.text(query), self.engine)
            df = DataFrameHelper.clean(df)
            logger.info(f"Fetched {len(df)} rows")
            return df
        except Exception as e:
            logger.exception(f"Database error: {e}")
            return pd.DataFrame()


# =========================================================
# MAIN APP
# =========================================================

class BillingSyncApp:
    def __init__(self):
        self.db = DatabaseClient()
        self.gs = GoogleSheetsClient()

    def process_task(self, task: Dict[str, str]):
        logger.info(f"--- Processing Task: {task['sheet_name']} ---")
        df = self.db.fetch(task["query"])
        self.gs.sync_dataframe(df, task["sheet_name"])

    def run(self, tasks: List[Dict[str, str]]):
        for task in tasks:
            try:
                self.process_task(task)
            except Exception as e:
                logger.exception(f"Task failed: {e}")


# =========================================================
# ENTRY POINT
# =========================================================

def main():
    tasks = [
        {"sheet_name": "Udara", "query": "EXEC get_data_billing_gsheet 1"},
        {"sheet_name": "Laut",  "query": "EXEC get_data_billing_gsheet 2"},
    ]
    BillingSyncApp().run(tasks)


if __name__ == "__main__":
    main()