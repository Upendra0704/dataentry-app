"""
Collaborative Data Entry Web App
- Dynamic schema from uploaded Excel
- SQLite runtime DB (not raw Excel for concurrent safety)
- Admin controls: start task, end task, export
- User auth: name + user ID, login/logout time tracking
- Idle detection via heartbeat
- 100-200 concurrent users safe
"""

import os
import json
import uuid
import threading
from datetime import datetime, timezone
from functools import wraps

from flask import (
    Flask, render_template, request, redirect,
    url_for, session, jsonify, flash, send_file
)
import openpyxl
from openpyxl import load_workbook
import sqlite3

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-this-in-production-" + uuid.uuid4().hex)

DB_PATH = os.path.join(os.path.dirname(__file__), "data.db")
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
EXPORT_FOLDER = os.path.join(os.path.dirname(__file__), "exports")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin1234")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)

db_lock = threading.Lock()

# ─────────────────────────────────────────────
# Database setup
# ─────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def init_db():
    with get_db() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            ended_at TEXT,
            status TEXT NOT NULL DEFAULT 'active',
            columns_meta TEXT NOT NULL,
            created_by TEXT NOT NULL DEFAULT 'admin'
        );

        CREATE TABLE IF NOT EXISTS entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id INTEGER NOT NULL,
            row_index INTEGER NOT NULL,
            prefill_data TEXT NOT NULL,
            user_data TEXT,
            status TEXT NOT NULL DEFAULT 'pending',
            claimed_by TEXT,
            claimed_by_name TEXT,
            claimed_at TEXT,
            submitted_at TEXT,
            time_spent_seconds INTEGER,
            FOREIGN KEY(task_id) REFERENCES tasks(id)
        );

        CREATE TABLE IF NOT EXISTS users (
            user_id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT NOT NULL,
            task_id INTEGER,
            login_at TEXT NOT NULL,
            logout_at TEXT,
            last_heartbeat TEXT,
            idle_seconds INTEGER DEFAULT 0,
            entries_completed INTEGER DEFAULT 0,
            FOREIGN KEY(user_id) REFERENCES users(user_id)
        );
        """)

init_db()

# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def now_iso():
    return datetime.now(timezone.utc).isoformat()

def get_active_task():
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM tasks WHERE status='active' ORDER BY id DESC LIMIT 1"
        ).fetchone()

def require_user(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

def require_admin(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("is_admin"):
            return redirect(url_for("admin_login"))
        return f(*args, **kwargs)
    return decorated

# ─────────────────────────────────────────────
# Auth routes
# ─────────────────────────────────────────────

@app.route("/", methods=["GET"])
def index():
    if session.get("is_admin"):
        return redirect(url_for("admin_dashboard"))
    if "user_id" in session:
        return redirect(url_for("work"))
    return redirect(url_for("login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    task = get_active_task()
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        user_id = request.form.get("user_id", "").strip().upper()
        if not name or not user_id:
            flash("Both Name and User ID are required.", "error")
            return render_template("login.html", task=task)

        with get_db() as conn:
            existing = conn.execute(
                "SELECT * FROM users WHERE user_id=?", (user_id,)
            ).fetchone()
            if existing and existing["name"].lower() != name.lower():
                flash("User ID already registered to a different name.", "error")
                return render_template("login.html", task=task)
            if not existing:
                conn.execute(
                    "INSERT INTO users (user_id, name, created_at) VALUES (?,?,?)",
                    (user_id, name, now_iso())
                )
            conn.execute(
                "INSERT INTO sessions (user_id, task_id, login_at, last_heartbeat) VALUES (?,?,?,?)",
                (user_id, task["id"] if task else None, now_iso(), now_iso())
            )
            sess_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

        session["user_id"] = user_id
        session["user_name"] = name
        session["session_db_id"] = sess_id
        return redirect(url_for("work"))

    return render_template("login.html", task=task)

@app.route("/logout")
@require_user
def logout():
    sess_id = session.get("session_db_id")
    if sess_id:
        with get_db() as conn:
            conn.execute(
                "UPDATE sessions SET logout_at=? WHERE id=?",
                (now_iso(), sess_id)
            )
    session.clear()
    return redirect(url_for("login"))

# ─────────────────────────────────────────────
# Heartbeat (idle detection)
# ─────────────────────────────────────────────

@app.route("/heartbeat", methods=["POST"])
@require_user
def heartbeat():
    sess_id = session.get("session_db_id")
    idle = request.json.get("idle_seconds", 0)
    if sess_id:
        with get_db() as conn:
            conn.execute(
                "UPDATE sessions SET last_heartbeat=?, idle_seconds=idle_seconds+? WHERE id=?",
                (now_iso(), idle, sess_id)
            )
    return jsonify({"ok": True})

# ─────────────────────────────────────────────
# Work routes
# ─────────────────────────────────────────────

@app.route("/work")
@require_user
def work():
    task = get_active_task()
    if not task:
        return render_template("no_task.html")

    columns_meta = json.loads(task["columns_meta"])
    user_id = session["user_id"]

    with db_lock:
        with get_db() as conn:
            entry = conn.execute(
                """SELECT * FROM entries
                   WHERE task_id=? AND claimed_by=? AND status='in_progress'
                   LIMIT 1""",
                (task["id"], user_id)
            ).fetchone()

            if not entry:
                entry = conn.execute(
                    """SELECT * FROM entries
                       WHERE task_id=? AND status='pending'
                       ORDER BY row_index ASC LIMIT 1""",
                    (task["id"],)
                ).fetchone()
                if entry:
                    conn.execute(
                        """UPDATE entries SET status='in_progress', claimed_by=?,
                           claimed_by_name=?, claimed_at=? WHERE id=?""",
                        (user_id, session["user_name"], now_iso(), entry["id"])
                    )

    if not entry:
        return render_template("queue_empty.html", task=task)

    prefill = json.loads(entry["prefill_data"])
    return render_template(
        "work.html",
        task=task,
        entry=entry,
        prefill=prefill,
        columns_meta=columns_meta,
        user_name=session["user_name"],
        user_id=user_id
    )

@app.route("/submit", methods=["POST"])
@require_user
def submit():
    task = get_active_task()
    if not task:
        return redirect(url_for("work"))

    entry_id = request.form.get("entry_id", type=int)
    time_spent = request.form.get("time_spent", type=int, default=0)
    columns_meta = json.loads(task["columns_meta"])

    user_data = {}
    for col in columns_meta:
        if not col["readonly"]:
            val = request.form.get(f"field_{col['name']}", "")
            user_data[col["name"]] = val

    with get_db() as conn:
        entry = conn.execute(
            "SELECT * FROM entries WHERE id=? AND claimed_by=? AND status='in_progress'",
            (entry_id, session["user_id"])
        ).fetchone()
        if not entry:
            flash("Entry not found or already submitted.", "error")
            return redirect(url_for("work"))

        conn.execute(
            """UPDATE entries SET status='done', user_data=?,
               submitted_at=?, time_spent_seconds=? WHERE id=?""",
            (json.dumps(user_data), now_iso(), time_spent, entry_id)
        )
        conn.execute(
            "UPDATE sessions SET entries_completed=entries_completed+1 WHERE id=?",
            (session["session_db_id"],)
        )

    return redirect(url_for("work"))

# ─────────────────────────────────────────────
# Admin routes
# ─────────────────────────────────────────────

@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        pw = request.form.get("password", "")
        if pw == ADMIN_PASSWORD:
            session["is_admin"] = True
            return redirect(url_for("admin_dashboard"))
        flash("Wrong password.", "error")
    return render_template("admin_login.html")

@app.route("/admin/logout")
def admin_logout():
    session.pop("is_admin", None)
    return redirect(url_for("admin_login"))

@app.route("/admin")
@require_admin
def admin_dashboard():
    with get_db() as conn:
        active_task = get_active_task()
        all_tasks = conn.execute(
            "SELECT * FROM tasks ORDER BY id DESC"
        ).fetchall()
        active_sessions = []
        if active_task:
            active_sessions = conn.execute(
                """SELECT s.*, u.name FROM sessions s
                   JOIN users u ON s.user_id = u.user_id
                   WHERE s.task_id=? AND s.logout_at IS NULL
                   ORDER BY s.login_at DESC""",
                (active_task["id"],)
            ).fetchall()
            stats = conn.execute(
                """SELECT
                    COUNT(*) as total,
                    SUM(CASE WHEN status='done' THEN 1 ELSE 0 END) as done,
                    SUM(CASE WHEN status='in_progress' THEN 1 ELSE 0 END) as in_progress,
                    SUM(CASE WHEN status='pending' THEN 1 ELSE 0 END) as pending
                   FROM entries WHERE task_id=?""",
                (active_task["id"],)
            ).fetchone()
        else:
            stats = None

    return render_template(
        "admin_dashboard.html",
        active_task=active_task,
        all_tasks=all_tasks,
        active_sessions=active_sessions,
        stats=stats
    )

@app.route("/admin/upload", methods=["POST"])
@require_admin
def admin_upload():
    active = get_active_task()
    if active:
        flash("End the current task before starting a new one.", "error")
        return redirect(url_for("admin_dashboard"))

    file = request.files.get("excel_file")
    task_name = request.form.get("task_name", "").strip()
    if not file or not task_name:
        flash("Task name and Excel file are required.", "error")
        return redirect(url_for("admin_dashboard"))

    filename = f"{uuid.uuid4().hex}_{file.filename}"
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        # Pass 1: read cell values
        wb_data = load_workbook(filepath, read_only=True, data_only=True)
        ws_data = wb_data.active
        headers = [
            str(cell.value).strip()
            for cell in next(ws_data.iter_rows(min_row=1, max_row=1))
            if cell.value
        ]

        rows_data = []
        for row in ws_data.iter_rows(min_row=2, values_only=True):
            if all(v is None for v in row):
                continue
            row_dict = {}
            for i, h in enumerate(headers):
                val = row[i] if i < len(row) else None
                row_dict[h] = str(val) if val is not None else ""
            rows_data.append(row_dict)
        wb_data.close()

        # Pass 2: extract Data Validation dropdowns
        wb_full = load_workbook(filepath, read_only=False, data_only=False)
        ws_full = wb_full.active

        from openpyxl.utils import get_column_letter, range_boundaries
        col_letter_to_name = {}
        for i, h in enumerate(headers):
            col_letter_to_name[get_column_letter(i + 1)] = h

        dropdown_options = {}
        for dv in ws_full.data_validations.dataValidation:
            if dv.type == "list" and dv.formula1:
                raw = dv.formula1.strip('"')
                if "," in raw and not raw.startswith("$") and "!" not in raw:
                    options = [o.strip() for o in raw.split(",") if o.strip()]
                elif "!" in raw or raw.startswith("$"):
                    options = []
                    try:
                        ref_ws = ws_full
                        ref_range = raw
                        if "!" in raw:
                            sheet_name, ref_range = raw.split("!", 1)
                            sheet_name = sheet_name.strip("'")
                            if sheet_name in wb_full.sheetnames:
                                ref_ws = wb_full[sheet_name]
                        for ref_row in ref_ws[ref_range]:
                            for cell in (ref_row if hasattr(ref_row, '__iter__') else [ref_row]):
                                if cell.value is not None:
                                    options.append(str(cell.value).strip())
                    except Exception:
                        options = []
                else:
                    options = []

                if options:
                    for cell_range in str(dv.sqref).split():
                        try:
                            min_col, _, max_col, _ = range_boundaries(cell_range)
                            for col_idx in range(min_col, max_col + 1):
                                col_name = col_letter_to_name.get(get_column_letter(col_idx))
                                if col_name:
                                    dropdown_options[col_name] = options
                        except Exception:
                            pass

        wb_full.close()

        columns_meta = []
        for h in headers:
            columns_meta.append({
                "name": h,
                "readonly": False,
                "dropdown": dropdown_options.get(h, [])
            })

    except Exception as e:
        flash(f"Could not read Excel file: {e}", "error")
        return redirect(url_for("admin_dashboard"))

    with get_db() as conn:
        conn.execute(
            "INSERT INTO tasks (name, created_at, status, columns_meta, created_by) VALUES (?,?,?,?,?)",
            (task_name, now_iso(), "active", json.dumps(columns_meta), "admin")
        )
        task_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        for i, row in enumerate(rows_data):
            conn.execute(
                "INSERT INTO entries (task_id, row_index, prefill_data, status) VALUES (?,?,?,?)",
                (task_id, i, json.dumps(row), "pending")
            )

    flash(f"Task '{task_name}' started with {len(rows_data)} entries.", "success")
    return redirect(url_for("admin_configure_columns", task_id=task_id))

@app.route("/admin/configure/<int:task_id>", methods=["GET", "POST"])
@require_admin
def admin_configure_columns(task_id):
    with get_db() as conn:
        task = conn.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
    if not task:
        flash("Task not found.", "error")
        return redirect(url_for("admin_dashboard"))

    columns_meta = json.loads(task["columns_meta"])

    if request.method == "POST":
        updated = []
        for col in columns_meta:
            is_readonly = request.form.get(f"readonly_{col['name']}") == "on"
            updated.append({
                "name": col["name"],
                "readonly": is_readonly,
                "dropdown": col.get("dropdown", [])
            })
        with get_db() as conn:
            conn.execute(
                "UPDATE tasks SET columns_meta=? WHERE id=?",
                (json.dumps(updated), task_id)
            )
        flash("Column configuration saved.", "success")
        return redirect(url_for("admin_dashboard"))

    return render_template("admin_configure.html", task=task, columns_meta=columns_meta)

@app.route("/admin/end_task", methods=["POST"])
@require_admin
def admin_end_task():
    task_id = request.form.get("task_id", type=int)
    confirmed = request.form.get("confirmed") == "yes"

    with get_db() as conn:
        task = conn.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not task:
            flash("Task not found.", "error")
            return redirect(url_for("admin_dashboard"))

        pending_count = conn.execute(
            "SELECT COUNT(*) FROM entries WHERE task_id=? AND status IN ('pending','in_progress')",
            (task_id,)
        ).fetchone()[0]

    if not confirmed:
        return render_template(
            "admin_end_confirm.html",
            task=task,
            pending_count=pending_count
        )

    export_path = _export_task(task_id)
    with get_db() as conn:
        conn.execute(
            "UPDATE tasks SET status='ended', ended_at=? WHERE id=?",
            (now_iso(), task_id)
        )
        conn.execute(
            "UPDATE sessions SET logout_at=? WHERE task_id=? AND logout_at IS NULL",
            (now_iso(), task_id)
        )

    flash(f"Task ended and data exported to {os.path.basename(export_path)}.", "success")
    return redirect(url_for("admin_dashboard"))

@app.route("/admin/export/<int:task_id>")
@require_admin
def admin_export(task_id):
    path = _export_task(task_id)
    return send_file(path, as_attachment=True)

def _export_task(task_id):
    with get_db() as conn:
        task = conn.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        entries = conn.execute(
            "SELECT * FROM entries WHERE task_id=? ORDER BY row_index",
            (task_id,)
        ).fetchall()
        sessions_data = conn.execute(
            """SELECT s.*, u.name as user_name FROM sessions s
               JOIN users u ON s.user_id = u.user_id
               WHERE s.task_id=? ORDER BY s.login_at""",
            (task_id,)
        ).fetchall()

    columns_meta = json.loads(task["columns_meta"])
    col_names = [c["name"] for c in columns_meta]

    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "Entries"
    headers = col_names + ["Status", "Claimed By", "Claimed By Name",
                           "Claimed At", "Submitted At", "Time Spent (s)"]
    ws1.append(headers)
    for entry in entries:
        prefill = json.loads(entry["prefill_data"])
        user_data = json.loads(entry["user_data"]) if entry["user_data"] else {}
        merged = {**prefill, **user_data}
        row = [merged.get(c, "") for c in col_names]
        row += [
            entry["status"],
            entry["claimed_by"] or "",
            entry["claimed_by_name"] or "",
            entry["claimed_at"] or "",
            entry["submitted_at"] or "",
            entry["time_spent_seconds"] or 0
        ]
        ws1.append(row)

    ws2 = wb.create_sheet("Time Tracking")
    ws2.append(["User ID", "Name", "Login", "Logout", "Idle (s)",
                "Entries Completed", "Total Session Duration"])
    for s in sessions_data:
        login = s["login_at"] or ""
        logout = s["logout_at"] or ""
        duration = ""
        if login and logout:
            try:
                d = datetime.fromisoformat(logout) - datetime.fromisoformat(login)
                duration = str(d).split(".")[0]
            except Exception:
                duration = ""
        ws2.append([
            s["user_id"], s["user_name"], login, logout,
            s["idle_seconds"] or 0, s["entries_completed"] or 0, duration
        ])

    fname = f"export_task_{task_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    path = os.path.join(EXPORT_FOLDER, fname)
    wb.save(path)
    return path

@app.route("/admin/task_entries/<int:task_id>")
@require_admin
def admin_task_entries(task_id):
    with get_db() as conn:
        task = conn.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        entries = conn.execute(
            "SELECT * FROM entries WHERE task_id=? ORDER BY row_index",
            (task_id,)
        ).fetchall()
    columns_meta = json.loads(task["columns_meta"])
    return render_template(
        "admin_entries.html",
        task=task,
        entries=entries,
        columns_meta=columns_meta
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)