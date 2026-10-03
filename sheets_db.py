"""
Google Sheets–backed data layer.

Every "table" (course, user, document, lab_session, attendance) is a tab
inside a single Google Spreadsheet. Row 1 of each tab is the header; each
subsequent row is a record. An integer "id" column is maintained by hand
since Sheets has no autoincrement.

Setup required before running the app — see README.md "Google Sheets setup".
"""

import threading
from datetime import datetime, timezone

import gspread
from google.oauth2.service_account import Credentials
from werkzeug.security import generate_password_hash

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

TABLE_SCHEMAS = {
    "course": ["id", "name", "code"],
    "user": ["id", "name", "email", "password_hash", "role", "course_id", "created_at"],
    "document": ["id", "course_id", "title", "filename", "uploaded_at"],
    "lab_session": ["id", "course_id", "code", "created_at", "expires_at", "is_active"],
    "attendance": ["id", "student_id", "lab_session_id", "timestamp"],
}

_INT_FIELDS = {"id", "course_id", "student_id", "lab_session_id", "is_active"}

# gspread's client isn't safe for concurrent writers hitting the same sheet
# from multiple threads/requests at once, so writes are serialized.
_write_lock = threading.Lock()


def utcnow_iso():
    return datetime.now(timezone.utc).replace(tzinfo=None).isoformat()


def _cast(row_dict):
    out = dict(row_dict)
    for key in _INT_FIELDS:
        if key in out and out[key] not in (None, ""):
            try:
                out[key] = int(out[key])
            except (TypeError, ValueError):
                pass
        elif key in out and out[key] == "":
            out[key] = None
    return out


class IdCounter:
    """Tracks the next id to hand out per table in a dedicated '_meta' sheet,
    so an id is never reused even after its row has been deleted."""

    def __init__(self, meta_worksheet):
        self.ws = meta_worksheet
        with _write_lock:
            existing = {r["table"] for r in self.ws.get_all_records()}
            for name in TABLE_SCHEMAS:
                if name not in existing:
                    self.ws.append_row([name, "0"])

    def next(self, table_name):
        with _write_lock:
            records = self.ws.get_all_records()
            for i, r in enumerate(records):
                if r["table"] == table_name:
                    new_val = int(r["next_id"] or 0) + 1
                    self.ws.update_cell(i + 2, 2, str(new_val))
                    return new_val
            # Shouldn't happen (row is seeded in __init__), but stay safe.
            self.ws.append_row([table_name, "1"])
            return 1


class SheetTable:
    def __init__(self, worksheet, name, counter):
        self.ws = worksheet
        self.name = name
        self.columns = TABLE_SCHEMAS[name]
        self.counter = counter

    def _rows(self):
        records = self.ws.get_all_records()
        return [_cast(r) for r in records]

    def all(self, order_by=None, reverse=False):
        rows = self._rows()
        if order_by:
            rows.sort(key=lambda r: (r.get(order_by) is None, r.get(order_by)), reverse=reverse)
        return rows

    def find(self, **filters):
        rows = self._rows()

        def match(r):
            return all(str(r.get(k)) == str(v) for k, v in filters.items())

        return [r for r in rows if match(r)]

    def find_one(self, **filters):
        results = self.find(**filters)
        return results[0] if results else None

    def get(self, id_):
        if id_ in (None, ""):
            return None
        return self.find_one(id=id_)

    def count(self, **filters):
        return len(self.find(**filters)) if filters else len(self._rows())

    def insert(self, **fields):
        new_id = self.counter.next(self.name)
        with _write_lock:
            fields["id"] = new_id
            row = [str(fields.get(col, "")) for col in self.columns]
            self.ws.append_row(row, value_input_option="RAW")
        return new_id

    def _find_row_index(self, id_):
        """1-based sheet row (including header) for a given id, or None."""
        id_col = self.columns.index("id") + 1
        cell = self.ws.find(str(id_), in_column=id_col)
        return cell.row if cell else None

    def update(self, id_, **fields):
        with _write_lock:
            row_idx = self._find_row_index(id_)
            if not row_idx:
                return False
            for key, value in fields.items():
                if key not in self.columns:
                    continue
                col_idx = self.columns.index(key) + 1
                self.ws.update_cell(row_idx, col_idx, str(value))
        return True

    def delete(self, id_):
        with _write_lock:
            row_idx = self._find_row_index(id_)
            if not row_idx:
                return False
            self.ws.delete_rows(row_idx)
        return True

    def delete_where(self, **filters):
        matches = self.find(**filters)
        for r in matches:
            self.delete(r["id"])
        return len(matches)


class SheetDB:
    def __init__(self, spreadsheet):
        self.ss = spreadsheet
        meta_ws = self._sheet("_meta", headers=["table", "next_id"])
        self.counter = IdCounter(meta_ws)
        self.course = SheetTable(self._sheet("course"), "course", self.counter)
        self.user = SheetTable(self._sheet("user"), "user", self.counter)
        self.document = SheetTable(self._sheet("document"), "document", self.counter)
        self.lab_session = SheetTable(self._sheet("lab_session"), "lab_session", self.counter)
        self.attendance = SheetTable(self._sheet("attendance"), "attendance", self.counter)

    def _sheet(self, name, headers=None):
        headers = headers or TABLE_SCHEMAS[name]
        try:
            return self.ss.worksheet(name)
        except gspread.WorksheetNotFound:
            ws = self.ss.add_worksheet(title=name, rows=1000, cols=len(headers))
            ws.append_row(headers)
            return ws


_client = None
_client_lock = threading.Lock()


def _get_client(credentials_file):
    global _client
    if _client is None:
        with _client_lock:
            if _client is None:
                creds = Credentials.from_service_account_file(credentials_file, scopes=SCOPES)
                _client = gspread.authorize(creds)
    return _client


def get_conn(credentials_file, sheet_id):
    client = _get_client(credentials_file)
    spreadsheet = client.open_by_key(sheet_id)
    return SheetDB(spreadsheet)


def init_db(credentials_file, sheet_id):
    db = get_conn(credentials_file, sheet_id)

    legacy_admin = db.user.find_one(email="admin@example.com", role="admin")
    email_in_use = db.user.find_one(email="admin@ayu.com")
    if legacy_admin and not email_in_use:
        db.user.update(legacy_admin["id"], email="admin@ayu.com")

    existing_admin = db.user.find_one(role="admin")
    if not existing_admin:
        db.user.insert(
            name="Admin",
            email="admin@ayu.com",
            password_hash=generate_password_hash("admin123"),
            role="admin",
            course_id="",
            created_at=utcnow_iso(),
        )
        print("Seeded default admin -> email: admin@ayu.com  password: admin123 (CHANGE THIS)")

    return db
