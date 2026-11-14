# DataEntry App

Collaborative web-based data entry with dynamic Excel schema, admin controls,
user auth, idle tracking, and time tracking.

---

## Setup (Local — 5 minutes)

### 1. Install Python dependencies
```bash
cd dataentry_app
pip install -r requirements.txt
```

### 2. Set your admin password (optional but recommended)
Edit `app.py` line:
```python
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin1234")
```
Or set environment variable:
```bash
export ADMIN_PASSWORD=your_secure_password
```

### 3. Run the app
```bash
python app.py
```
App runs at: **http://localhost:5000**

For LAN access (all users on same Wi-Fi connect to your IP):
```bash
python app.py
# Users connect to: http://YOUR_IP:5000
```

---

## Admin Workflow

1. Go to **http://localhost:5000/admin/login** (default password: `admin1234`)
2. Upload an Excel file → name the task → configure columns (pre-filled vs editable)
3. Users log in at the main URL with Name + User ID
4. Each user gets the next unclaimed row automatically
5. When all entries are done (or you decide to end):
   - Click **End Task** → it shows a warning if entries are incomplete
   - Confirm → data is exported to Excel → task ends
6. Start a fresh task by uploading a new Excel file

---

## User Workflow

1. Go to **http://localhost:5000**
2. Enter your Name and User ID (same ID every session for audit continuity)
3. Pre-filled fields are shown in grey — fill in the editable fields
4. Click **Submit & Next** → next entry loads automatically
5. Timer runs on each entry; idle time is tracked automatically
6. Click **Logout** when done for the day

---

## Exports

- **Admin Dashboard → Export Now**: downloads current state as .xlsx
- **On End Task**: auto-exported before task closes

Export contains two sheets:
- **Entries**: all rows with pre-fill + user-filled data + status + time spent
- **Time Tracking**: every user session with login, logout, idle time, entries completed

---

## Hosting on Render (free, public internet)

1. Push this folder to a GitHub repo
2. Go to **render.com** → New Web Service → connect your repo
3. Build command: `pip install -r requirements.txt`
4. Start command: `python app.py`
5. Set environment variables: `SECRET_KEY`, `ADMIN_PASSWORD`
6. Done — Render gives you a public URL

---

## Security notes

- Change `ADMIN_PASSWORD` before deploying
- Set a strong `SECRET_KEY` in production
- For 100-200 users, the SQLite + WAL mode setup handles concurrency fine
- For 500+ users, migrate to PostgreSQL (swap `sqlite3` for `psycopg2`)
