# -*- coding: utf-8 -*-
"""
仓库扣账 API（给生产系统调用）
端口 8510。库存仍写 warehouse.db，与 Streamlit 页面共用同一文件。
启动：python stock_api.py
"""
import json
import re
import sqlite3
import os
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "warehouse.db")
PORT = int(os.environ.get("STOCK_API_PORT", "8510"))


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def get_setting(conn, key, default=""):
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row[0] if row else default


def migrate():
    conn = get_conn()
    rec_cols = [r[1] for r in conn.execute("PRAGMA table_info(records)")]
    if "task_id" not in rec_cols:
        conn.execute("ALTER TABLE records ADD COLUMN task_id TEXT DEFAULT ''")
    if "stage_id" not in rec_cols:
        conn.execute("ALTER TABLE records ADD COLUMN stage_id TEXT DEFAULT ''")
    mat_cols = [r[1] for r in conn.execute("PRAGMA table_info(materials)")]
    if "kind" not in mat_cols:
        conn.execute("ALTER TABLE materials ADD COLUMN kind TEXT DEFAULT 'connector'")
        conn.execute("UPDATE materials SET kind='connector' WHERE kind IS NULL OR kind=''")
    # 给下料演示准备一条公用 POM 原料（已有则跳过）
    row = conn.execute(
        "SELECT id FROM materials WHERE code=? AND customer=?",
        ("POM-黑色-20mm", "公用"),
    ).fetchone()
    if not row:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur = conn.execute(
            "INSERT INTO materials (code,name,customer,color,spec,unit,location,"
            "safety_stock,note,created_at,status,kind)"
            " VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            ("POM-黑色-20mm", "POM板", "公用", "黑色", "20mm", "kg", "R-1",
             10, "生产下料扣账用", now, "在用", "pom"),
        )
        mid = cur.lastrowid
        conn.execute(
            "INSERT INTO records (type,material_id,quantity,operator,note,created_at,sub_type)"
            " VALUES ('期初',?,100,'系统','下料演示期初',?,'')",
            (mid, now),
        )
    conn.commit()
    conn.close()


def stock_of(conn, mid):
    row = conn.execute(
        "SELECT COALESCE(SUM(CASE WHEN type IN ('期初','入库','盘点') "
        "THEN quantity ELSE -quantity END), 0) FROM records WHERE material_id=?",
        (mid,),
    ).fetchone()
    return row[0] or 0


def material_json(conn, row):
    mid = row["id"]
    base = get_setting(conn, "base_url", "").rstrip("/")
    return {
        "id": mid,
        "qr": "%s/?id=%s" % (base, mid),
        "kind": row["kind"] if "kind" in row.keys() else "connector",
        "code": row["code"],
        "name": row["name"] or row["code"],
        "customer": row["customer"] or "",
        "spec": row["spec"] or "",
        "color": row["color"] or "",
        "unit": row["unit"] or "个",
        "location": row["location"] or "",
        "stock": stock_of(conn, mid),
    }


def lookup(q):
    q = (q or "").strip()
    conn = get_conn()
    url_id = re.search(r"[?&]id=(\d+)", q)
    if q.upper().startswith("MAT:") and q[4:].strip().isdigit():
        rows = conn.execute("SELECT * FROM materials WHERE id=?", (int(q[4:].strip()),)).fetchall()
    elif q.isdigit():
        rows = conn.execute("SELECT * FROM materials WHERE id=?", (int(q),)).fetchall()
    elif url_id:
        # 扫码枪/生产系统直接读到二维码里印的网址（如 http://.../?id=123），从里面取编号
        rows = conn.execute("SELECT * FROM materials WHERE id=?", (int(url_id.group(1)),)).fetchall()
    else:
        like = "%" + q + "%"
        rows = conn.execute(
            "SELECT m.* FROM materials m "
            "LEFT JOIN material_aliases a ON a.material_id=m.id "
            "WHERE m.code LIKE ? OR m.name LIKE ? OR m.customer LIKE ? OR m.spec LIKE ? "
            "OR a.code LIKE ? "
            "GROUP BY m.id ORDER BY m.code LIMIT 30",
            (like, like, like, like, like),
        ).fetchall()
    out = [material_json(conn, r) for r in rows]
    conn.close()
    return out


def by_task(task_id):
    conn = get_conn()
    rows = conn.execute(
        "SELECT r.id, r.created_at, r.quantity, r.operator, r.note, r.stage_id, r.sub_type,"
        " m.id AS material_id, m.code, m.customer, m.unit "
        "FROM records r JOIN materials m ON m.id=r.material_id "
        "WHERE r.task_id=? AND r.type='出库' ORDER BY r.id DESC",
        (task_id,),
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def issue(body):
    mid = int(body.get("material_id") or 0)
    qty = float(body.get("qty") or 0)
    task_id = str(body.get("task_id") or "").strip()
    stage_id = str(body.get("stage_id") or "").strip()
    operator = str(body.get("operator") or "").strip() or "生产报工"
    note = str(body.get("note") or "").strip()
    if mid <= 0 or qty <= 0:
        return 400, {"error": "material_id 和 qty 必填且大于 0"}
    conn = get_conn()
    try:
        conn.execute("BEGIN IMMEDIATE")
        mat = conn.execute("SELECT * FROM materials WHERE id=?", (mid,)).fetchone()
        if not mat:
            conn.rollback()
            return 404, {"error": "物料不存在"}
        current = stock_of(conn, mid)
        if qty > current:
            conn.rollback()
            return 409, {
                "error": "stock_short",
                "message": "库存不足：当前只剩 %g，不能出库 %g" % (current, qty),
                "stock": current,
                "material": material_json(conn, mat),
            }
        conn.execute(
            "INSERT INTO records (type, material_id, quantity, operator, note,"
            " sub_type, created_at, task_id, stage_id)"
            " VALUES ('出库',?,?,?,?, '工单扣料',?,?,?)",
            (mid, qty, operator, note or ("工单 " + task_id),
             datetime.now().strftime("%Y-%m-%d %H:%M:%S"), task_id, stage_id),
        )
        conn.commit()
        left = stock_of(conn, mid)
        return 200, {
            "ok": True,
            "stock": left,
            "material": material_json(conn, mat),
            "task_id": task_id,
        }
    except Exception as e:
        conn.rollback()
        return 500, {"error": str(e)}
    finally:
        conn.close()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print("[stock-api]", fmt % args)

    def _send(self, code, obj):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type,x-auth")
        self.end_headers()

    def do_GET(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)
        if u.path == "/health":
            return self._send(200, {"ok": True})
        if u.path in ("/stock/lookup", "/api/stock/lookup"):
            return self._send(200, {"items": lookup(q.get("q", [""])[0])})
        if u.path in ("/stock/by-task", "/api/stock/by-task"):
            return self._send(200, {"items": by_task(q.get("id", [""])[0])})
        return self._send(404, {"error": "not found"})

    def do_POST(self):
        u = urlparse(self.path)
        n = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(n).decode("utf-8") if n else "{}"
        try:
            body = json.loads(raw) if raw else {}
        except json.JSONDecodeError:
            return self._send(400, {"error": "invalid json"})
        if u.path in ("/stock/issue", "/api/stock/issue"):
            code, obj = issue(body)
            return self._send(code, obj)
        return self._send(404, {"error": "not found"})


if __name__ == "__main__":
    migrate()
    httpd = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print("仓库扣账 API  http://127.0.0.1:%s" % PORT)
    httpd.serve_forever()
