import csv
import sqlite3
from pathlib import Path
from datetime import datetime, timedelta

DATABASE_PATH = Path(__file__).resolve().parent / "messages.db"


def get_connection():
    connection = sqlite3.connect(DATABASE_PATH, timeout=10)
    connection.row_factory = sqlite3.Row
    return connection


def initialize_database():
    """Create tables and add compatible columns without deleting existing data."""
    with get_connection() as connection:
        connection.execute("""
            CREATE TABLE IF NOT EXISTS messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                message_id TEXT UNIQUE NOT NULL,
                phone_number TEXT NOT NULL,
                message TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'queued',
                created_at TEXT NOT NULL
            )
        """)
        connection.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id TEXT UNIQUE NOT NULL,
                amount REAL NOT NULL,
                phone_number TEXT NOT NULL,
                message_id TEXT,
                created_at TEXT NOT NULL
            )
        """)
        # Safe additive migrations. Existing rows and IDs are preserved.
        message_columns = {
            row["name"] for row in connection.execute("PRAGMA table_info(messages)")
        }
        if "provider" not in message_columns:
            connection.execute(
                "ALTER TABLE messages ADD COLUMN provider TEXT NOT NULL DEFAULT 'local-simulation'"
            )
        if "updated_at" not in message_columns:
            connection.execute("ALTER TABLE messages ADD COLUMN updated_at TEXT")
            connection.execute(
                "UPDATE messages SET updated_at = created_at WHERE updated_at IS NULL"
            )


def save_message(message_data):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection() as connection:
        connection.execute("""
            INSERT INTO messages
                (message_id, phone_number, message, status, created_at, provider, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            message_data["message_id"],
            message_data["phone_number"],
            message_data["message"],
            message_data["status"],
            message_data["created_at"],
            message_data.get("provider", "local-simulation"),
            message_data.get("updated_at", now),
        ))


def get_all_messages():
    with get_connection() as connection:
        rows = connection.execute(
            "SELECT * FROM messages ORDER BY id DESC"
        ).fetchall()
        return [dict(row) for row in rows]


def get_message_by_id(message_id):
    with get_connection() as connection:
        row = connection.execute(
            "SELECT * FROM messages WHERE message_id = ?", (message_id,)
        ).fetchone()
        return dict(row) if row else None


def update_message_status(message_id, new_status):
    status = (new_status or "").strip().lower()
    if status not in {"queued", "delivered", "failed"}:
        return False
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection() as connection:
        cursor = connection.execute("""
            UPDATE messages SET status = ?, updated_at = ?
            WHERE message_id = ?
        """, (status, now, message_id))
        return cursor.rowcount > 0


def get_message_stats():
    with get_connection() as connection:
        rows = connection.execute("""
            SELECT LOWER(status) AS status, COUNT(*) AS total
            FROM messages GROUP BY LOWER(status)
        """).fetchall()
    counts = {"total": 0, "queued": 0, "delivered": 0, "failed": 0}
    for row in rows:
        status, count = row["status"], row["total"]
        counts["total"] += count
        if status in counts:
            counts[status] = count
    resolved = counts["delivered"] + counts["failed"]
    counts["delivery_rate"] = round(counts["delivered"] / resolved * 100, 1) if resolved else 0
    counts["failure_rate"] = round(counts["failed"] / resolved * 100, 1) if resolved else 0
    return counts


def get_daily_stats(days=7):
    days = max(1, min(int(days), 90))
    start_date = (datetime.now() - timedelta(days=days - 1)).strftime("%Y-%m-%d")
    with get_connection() as connection:
        rows = connection.execute("""
            SELECT substr(created_at, 1, 10) AS day,
                   COUNT(*) AS total,
                   SUM(CASE WHEN LOWER(status) = 'delivered' THEN 1 ELSE 0 END) AS delivered,
                   SUM(CASE WHEN LOWER(status) = 'failed' THEN 1 ELSE 0 END) AS failed,
                   SUM(CASE WHEN LOWER(status) = 'queued' THEN 1 ELSE 0 END) AS queued
            FROM messages
            WHERE substr(created_at, 1, 10) >= ?
            GROUP BY substr(created_at, 1, 10)
            ORDER BY day
        """, (start_date,)).fetchall()
    mapping = {row["day"]: dict(row) for row in rows}
    result = []
    for offset in range(days - 1, -1, -1):
        day = (datetime.now() - timedelta(days=offset)).strftime("%Y-%m-%d")
        result.append(mapping.get(day, {
            "day": day, "total": 0, "delivered": 0, "failed": 0, "queued": 0
        }))
    return result


def search_messages(search_text="", status="all", start_date="", end_date=""):
    search_text = (search_text or "").strip()
    query = "SELECT * FROM messages WHERE 1=1"
    params = []
    if search_text:
        query += """ AND (message_id LIKE ? OR phone_number LIKE ? OR message LIKE ?)"""
        term = f"%{search_text}%"
        params.extend([term, term, term])
    selected_status = (status or "all").strip().lower()
    if selected_status in {"queued", "delivered", "failed"}:
        query += " AND LOWER(status) = ?"
        params.append(selected_status)
    if start_date:
        query += " AND substr(created_at, 1, 10) >= ?"
        params.append(start_date)
    if end_date:
        query += " AND substr(created_at, 1, 10) <= ?"
        params.append(end_date)
    query += " ORDER BY id DESC"
    with get_connection() as connection:
        return [dict(row) for row in connection.execute(query, params).fetchall()]


def get_message_history(message_id):
    with get_connection() as connection:
        row = connection.execute(
            "SELECT * FROM messages WHERE message_id = ?", (message_id,)
        ).fetchone()
    if not row:
        return None
    item = dict(row)
    item["history"] = [
        {"time": item.get("created_at"), "status": "queued", "note": "Message created"}
    ]
    if item.get("updated_at") and item["updated_at"] != item.get("created_at"):
        item["history"].append({
            "time": item["updated_at"], "status": item.get("status"),
            "note": "Latest status recorded"
        })
    return item


def save_order(order_id, amount, phone_number, message_id, created_at):
    with get_connection() as connection:
        connection.execute("""
            INSERT INTO orders (order_id, amount, phone_number, message_id, created_at)
            VALUES (?, ?, ?, ?, ?)
        """, (order_id, amount, phone_number, message_id, created_at))


def get_all_orders(search_text=""):
    search_text = (search_text or "").strip()
    query = "SELECT * FROM orders WHERE 1=1"
    params = []
    if search_text:
        query += " AND (order_id LIKE ? OR phone_number LIKE ? OR message_id LIKE ?)"
        term = f"%{search_text}%"
        params.extend([term, term, term])
    query += " ORDER BY id DESC"
    with get_connection() as connection:
        return [dict(row) for row in connection.execute(query, params).fetchall()]


def export_messages_csv():
    """Return all messages as CSV text for browser download."""
    import io
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Message ID", "Phone Number", "Message", "Status", "Created At", "Updated At", "Provider"])
    for row in get_all_messages():
        writer.writerow([
            row.get("message_id", ""), row.get("phone_number", ""),
            row.get("message", ""), row.get("status", ""),
            row.get("created_at", ""), row.get("updated_at", ""),
            row.get("provider", "local-simulation")
        ])
    return output.getvalue()


def export_orders_csv():
    import io
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Order ID", "Amount", "Phone Number", "Message ID", "Created At"])
    for row in get_all_orders():
        writer.writerow([
            row.get("order_id", ""), row.get("amount", ""),
            row.get("phone_number", ""), row.get("message_id", ""),
            row.get("created_at", "")
        ])
    return output.getvalue()
