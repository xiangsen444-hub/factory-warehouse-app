# -*- coding: utf-8 -*-
"""
工厂仓库管理系统（演示版）
================================
物料格式对齐工厂实际清单：
    位置(序号) / 客户名称 / 料号 / 颜色 / 器件数目 / 产品照片(PDF) / 图纸照片 / 备注

功能：库存查询（搜索+低库存标红+物料详情照片）、库位视图、入库、出库、
      出入流水（可导出CSV）、AI问答（离线规则版）、物料管理（含Excel导入、照片上传）、
      每日自动备份

技术：Streamlit（网页） + SQLite（单文件数据库 warehouse.db）
启动：streamlit run app.py --server.address 0.0.0.0 --server.port 8501
局域网内手机/电脑浏览器访问： http://本机IP:8501
"""

import io
import os
import re
import shutil
import sqlite3
from datetime import datetime, date, timedelta

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image

# ---------------- 路径配置 ----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "warehouse.db")      # 数据库就是这个文件，复制它=备份
BACKUP_DIR = os.path.join(BASE_DIR, "backups")        # 自动备份存放目录
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")        # 产品照片/图纸存放目录

IMAGE_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}

FULL_ACCESS_PASSWORD = "demo2026"  # 解锁全部功能（入库/出库/物料管理/访问统计）的演示密码，公开展示用，和生产环境密码无关

# ---------------- 数据库基础 ----------------

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    """
    建表（已存在则跳过）+ 老库自动迁移。
    重要设计：料号在不同客户那里可能是不同的东西，所以唯一性 = 料号 + 客户，
    内部关联用自增 id。
    """
    conn = get_conn()
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS materials (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        code          TEXT NOT NULL,      -- 料号，如 PP5415101
        name          TEXT DEFAULT '',    -- 名称（可空，显示用料号）
        customer      TEXT DEFAULT '',    -- 客户名称，如 河泽三贤
        color         TEXT DEFAULT '',    -- 颜色
        pin_count     INTEGER,            -- 器件数目（列名沿用历史命名，实际含义是器件数目而非电气PIN数）
        spec          TEXT DEFAULT '',    -- 规格/补充描述
        unit          TEXT DEFAULT '个',  -- 单位
        location      TEXT DEFAULT '',    -- 位置（清单里的序号，如 A-1）
        safety_stock  REAL DEFAULT 0,     -- 安全库存（低于它标红）
        product_file  TEXT DEFAULT '',    -- 预留，不再使用
        drawing_file  TEXT DEFAULT '',    -- 预留，不再使用
        note          TEXT DEFAULT '',    -- 备注
        created_at    TEXT,
        UNIQUE(code, customer)            -- 同一客户内料号唯一
    );
    CREATE TABLE IF NOT EXISTS records (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        type          TEXT NOT NULL,      -- 期初 / 入库 / 出库
        material_id   INTEGER NOT NULL,   -- 关联 materials.id
        quantity      REAL NOT NULL,      -- 数量（始终为正数，方向由type决定）
        operator      TEXT DEFAULT '',    -- 经手人/领用人
        note          TEXT DEFAULT '',
        created_at    TEXT NOT NULL,
        FOREIGN KEY (material_id) REFERENCES materials(id)
    );
    CREATE TABLE IF NOT EXISTS visits (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        visited_at    TEXT NOT NULL,  -- 每次新会话打开网页的时间
        ip_address    TEXT DEFAULT '' -- 访问设备的局域网IP
    );
    """)
    visit_cols = [r[1] for r in conn.execute("PRAGMA table_info(visits)")]
    if "ip_address" not in visit_cols:
        conn.execute("ALTER TABLE visits ADD COLUMN ip_address TEXT DEFAULT ''")
    # 老库迁移：materials 没有 id 列 → 重建为「id主键 + UNIQUE(code,customer)」结构
    cols = [r[1] for r in conn.execute("PRAGMA table_info(materials)")]
    if cols and "id" not in cols:
        # v1遗留库先补列，保证下面SELECT字段齐全
        for col, ddl in {
            "customer": "customer TEXT DEFAULT ''",
            "color": "color TEXT DEFAULT ''",
            "pin_count": "pin_count INTEGER",
            "product_file": "product_file TEXT DEFAULT ''",
            "drawing_file": "drawing_file TEXT DEFAULT ''",
        }.items():
            if col not in cols:
                conn.execute(f"ALTER TABLE materials ADD COLUMN {ddl}")
        conn.executescript("""
        CREATE TABLE materials_new (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            code          TEXT NOT NULL,
            name          TEXT DEFAULT '',
            customer      TEXT DEFAULT '',
            color         TEXT DEFAULT '',
            pin_count     INTEGER,
            spec          TEXT DEFAULT '',
            unit          TEXT DEFAULT '个',
            location      TEXT DEFAULT '',
            safety_stock  REAL DEFAULT 0,
            product_file  TEXT DEFAULT '',
            drawing_file  TEXT DEFAULT '',
            note          TEXT DEFAULT '',
            created_at    TEXT,
            UNIQUE(code, customer)
        );
        INSERT INTO materials_new (code,name,customer,color,pin_count,spec,unit,location,
                                   safety_stock,product_file,drawing_file,note,created_at)
            SELECT code,name,customer,color,pin_count,spec,unit,location,
                   safety_stock,product_file,drawing_file,note,created_at FROM materials;
        CREATE TABLE records_new (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            type          TEXT NOT NULL,
            material_id   INTEGER NOT NULL,
            quantity      REAL NOT NULL,
            operator      TEXT DEFAULT '',
            note          TEXT DEFAULT '',
            created_at    TEXT NOT NULL,
            FOREIGN KEY (material_id) REFERENCES materials_new(id)
        );
        INSERT INTO records_new (type,material_id,quantity,operator,note,created_at)
            SELECT r.type, mn.id, r.quantity, r.operator, r.note, r.created_at
            FROM records r JOIN materials_new mn ON mn.code = r.material_code;
        DROP TABLE records;
        DROP TABLE materials;
        ALTER TABLE materials_new RENAME TO materials;
        ALTER TABLE records_new RENAME TO records;
        """)
    conn.commit()
    conn.close()


def get_stock(conn, material_id):
    """单个物料当前库存 = 期初 + 入库 - 出库（由流水推导，永远一致）"""
    row = conn.execute("""
        SELECT COALESCE(SUM(CASE WHEN type IN ('期初','入库','盘点') THEN quantity ELSE -quantity END), 0)
        FROM records WHERE material_id = ?
    """, (material_id,)).fetchone()
    return row[0] or 0


def log_visit():
    """每个浏览器会话只记一次访问（用 session_state 判断），避免把页面内每次点击都算成新访问"""
    if st.session_state.get("_visit_logged"):
        return
    st.session_state["_visit_logged"] = True
    try:
        ip = st.context.ip_address or ""  # 需要较新版本 Streamlit，公司电脑版本较旧时优雅降级
    except Exception:
        ip = ""
    conn = get_conn()
    conn.execute("INSERT INTO visits (visited_at, ip_address) VALUES (?, ?)",
                 (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ip))
    conn.commit()
    conn.close()


def visit_stats():
    """返回 (今日访问次数, 累计访问次数)"""
    conn = get_conn()
    today = date.today().strftime("%Y-%m-%d")
    today_count = conn.execute(
        "SELECT COUNT(*) FROM visits WHERE visited_at LIKE ?", (f"{today}%",)
    ).fetchone()[0]
    total_count = conn.execute("SELECT COUNT(*) FROM visits").fetchone()[0]
    conn.close()
    return today_count, total_count


def recent_visits(limit=20):
    """最近 limit 条访问记录：[(时间, IP), ...]，最新的在前"""
    conn = get_conn()
    rows = conn.execute(
        "SELECT visited_at, ip_address FROM visits ORDER BY id DESC LIMIT ?", (limit,)
    ).fetchall()
    conn.close()
    return rows


def check_magic_link():
    """网址带 ?key=编辑密码 时自动解锁，方便固定设备（如录物料用的手机）收藏免密链接"""
    if st.session_state.get("_full_unlocked"):
        return
    if st.query_params.get("key") == FULL_ACCESS_PASSWORD:
        st.session_state["_full_unlocked"] = True


def require_full_access():
    """入库/出库/物料管理等编辑功能的密码门：本次会话验证一次即可，返回是否已解锁"""
    if st.session_state.get("_full_unlocked"):
        return True
    st.info("查询类功能无需密码；这个功能会改动库存数据，需要密码解锁。")
    pwd = st.text_input("密码", type="password", key="_full_pwd")
    if pwd:
        if pwd == FULL_ACCESS_PASSWORD:
            st.session_state["_full_unlocked"] = True
            st.rerun()
        else:
            st.error("密码错误")
    return False


def add_record(rtype, material_id, qty, operator, note):
    """写一条出入库流水。出库时开事务并校验负库存：要么成功，要么什么都不写。"""
    conn = get_conn()
    try:
        conn.execute("BEGIN IMMEDIATE")
        if rtype == "出库":
            current = get_stock(conn, material_id)
            if qty > current:
                conn.rollback()
                return False, f"库存不足：当前只剩 {current:g}，不能出库 {qty:g}"
        conn.execute(
            "INSERT INTO records (type, material_id, quantity, operator, note, created_at)"
            " VALUES (?,?,?,?,?,?)",
            (rtype, material_id, qty, operator, note, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
        conn.commit()
        return True, "ok"
    except Exception as e:
        conn.rollback()
        return False, f"写入失败：{e}"
    finally:
        conn.close()


def inventory_df():
    """库存总表：物料档案 + 实时库存 + 状态（id 列供程序内部用，展示时去掉）"""
    conn = get_conn()
    df = pd.read_sql_query("""
        SELECT m.id       AS id,
               m.location AS 位置,
               m.customer AS 客户,
               m.code     AS 料号,
               m.color    AS 颜色,
               m.pin_count AS 器件数目,
               COALESCE(SUM(CASE WHEN r.type IN ('期初','入库','盘点') THEN r.quantity ELSE -r.quantity END), 0) AS 库存,
               m.unit     AS 单位,
               m.safety_stock AS 安全库存
        FROM materials m
        LEFT JOIN records r ON r.material_id = m.id
        GROUP BY m.id
        ORDER BY CASE WHEN instr(m.location,'-')>0 THEN substr(m.location,1,instr(m.location,'-')-1)
                      ELSE m.location END,
                 CASE WHEN instr(m.location,'-')>0 THEN CAST(substr(m.location,instr(m.location,'-')+1) AS INTEGER)
                      ELSE 0 END,
                 m.code
    """, conn)
    conn.close()
    df["状态"] = df.apply(lambda x: "需补货" if x["库存"] < x["安全库存"] else "正常", axis=1)
    return df


# ---------------- 备份 ----------------

def auto_backup():
    """每天第一次打开系统时自动备份数据库（保留最近30天）"""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    if not os.path.exists(DB_PATH):
        return
    today = date.today().strftime("%Y%m%d")
    if not any(today in f for f in os.listdir(BACKUP_DIR)):
        shutil.copy(DB_PATH, os.path.join(BACKUP_DIR, f"warehouse_{today}.db"))
    cutoff = (date.today() - timedelta(days=30)).strftime("%Y%m%d")
    for f in os.listdir(BACKUP_DIR):
        if f.startswith("warehouse_") and f[10:18] < cutoff:
            try:
                os.remove(os.path.join(BACKUP_DIR, f))
            except OSError:
                pass
    # 照片目录每天滚动备份一份（覆盖式，只留最新，防止照片误删）
    if os.path.isdir(UPLOAD_DIR):
        marker = os.path.join(BACKUP_DIR, ".uploads_backup_date")
        done = ""
        if os.path.exists(marker):
            with open(marker) as fh:
                done = fh.read().strip()
        if done != today:
            shutil.copytree(UPLOAD_DIR, os.path.join(BACKUP_DIR, "uploads_latest"),
                            dirs_exist_ok=True)
            with open(marker, "w") as fh:
                fh.write(today)


# ---------------- 文件上传 ----------------

def compress_image(data, max_side=1600, quality=82):
    """手机原图太大（3~5MB），压到最长边1600px的JPEG（约200~400KB），看料号细节足够清晰"""
    try:
        img = Image.open(io.BytesIO(data)).convert("RGB")
        img.thumbnail((max_side, max_side))
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=quality)
        return buf.getvalue()
    except Exception:
        return data   # 压缩失败就存原图，不影响上传


def safe_name(s):
    """料号里的换行/斜杠等字符不能进文件名，统一转下划线"""
    return re.sub(r'[\\/:*?"<>|\r\n]+', "_", str(s)).strip()


def file_key(mid, code):
    """照片文件名前缀：id_料号（id 保证同料号不同客户不冲突）"""
    return f"{mid}_{safe_name(code)}"


def save_uploads(uploaded_files, fkey, kind):
    """
    保存一组照片（如产品六个方向、多张图纸）到 uploads/。
    文件命名：id_料号_类型_序号.扩展名，如 153_PP5415101_产品照片_1.jpg
    图片自动压缩；同名覆盖；清旧文件失败不影响保存（沙箱可能拦截删除）。
    """
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    for f in os.listdir(UPLOAD_DIR):
        if f.startswith(f"{fkey}_{kind}_"):
            try:
                os.remove(os.path.join(UPLOAD_DIR, f))
            except OSError:
                pass
    n = 0
    for i, uf in enumerate(uploaded_files, 1):
        ext = os.path.splitext(uf.name)[1].lower() or ".jpg"
        data = uf.getbuffer()
        if ext in IMAGE_EXT:
            data, ext = compress_image(data), ".jpg"
        with open(os.path.join(UPLOAD_DIR, f"{fkey}_{kind}_{i}{ext}"), "wb") as f:
            f.write(data)
        n += 1
    return n


def render_pdf_pages(path, max_pages=12, zoom=1.5):
    """把PDF每页渲染成图片字节流（PyMuPDF），失败返回空列表"""
    try:
        import pymupdf
        doc = pymupdf.open(path)
        imgs = []
        for i, page in enumerate(doc):
            if i >= max_pages:
                break
            pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom))
            imgs.append(pix.tobytes("png"))
        doc.close()
        return imgs
    except Exception:
        return []


def show_material_files(mid, code):
    """展示某物料的全部照片/图纸：图片和PDF页面都以网格直接显示，PDF另附下载。
    新文件按 id_料号_ 前缀匹配；早期按 料号_ 命名的文件，只在当前物料是
    同料号中 id 最小者时认领，避免同料号不同客户之间张冠李戴。"""
    if not os.path.isdir(UPLOAD_DIR):
        st.caption("还没有上传照片/图纸")
        return
    fkey = file_key(mid, code)
    all_files = os.listdir(UPLOAD_DIR)
    files = [f for f in all_files if f.startswith(fkey + "_")]

    legacy_prefix = safe_name(code) + "_"
    legacy = [f for f in all_files
              if f.startswith(legacy_prefix) and not re.match(r"^\d+_", f)]
    if legacy:
        conn = get_conn()
        min_id = conn.execute(
            "SELECT MIN(id) FROM materials WHERE code=?", (code,)).fetchone()[0]
        conn.close()
        if mid == min_id:
            files += legacy

    if not files:
        st.caption("还没有上传照片/图纸")
        return

    # 收集所有要显示的图片：直接图片 + PDF渲染页
    tiles = []   # (caption, 路径或字节流)
    for fname in sorted(files):
        path = os.path.join(UPLOAD_DIR, fname)
        ext = os.path.splitext(fname)[1].lower()
        label = fname
        for prefix in (fkey + "_", legacy_prefix):
            if label.startswith(prefix):
                label = label[len(prefix):]
        label = label.rsplit(".", 1)[0]                      # 如 产品照片_1 / 产品照片
        if ext in IMAGE_EXT:
            tiles.append((label, path))
        elif ext == ".pdf":
            for pi, png in enumerate(render_pdf_pages(path), 1):
                tiles.append((f"{label} 第{pi}页", png))
            with open(path, "rb") as fh:
                st.download_button(f"下载 {fname}", fh.read(),
                                   file_name=fname, key=f"dl_{mid}_{fname}")

    # 每行3张网格显示
    for i in range(0, len(tiles), 3):
        cols = st.columns(3)
        for col, (caption, src) in zip(cols, tiles[i:i + 3]):
            col.image(src, caption=caption, use_container_width=True)


# ---------------- 演示数据（仅数据库为空时写入一次） ----------------

def seed_demo_data():
    conn = get_conn()
    if conn.execute("SELECT COUNT(*) FROM materials").fetchone()[0] > 0:
        conn.close()
        return
    # 料号, 客户, 颜色, 器件数目, 位置, 安全库存, 期初数量
    materials = [
        ("PP5415101", "鹤壁三贤", "黑色",   1,  "A-1", 50, 300),
        ("6189-7599", "河泽三贤", "黑白色", 4,  "A-2", 50, 200),
        ("PP5411001", "河泽三贤", "黑色",   2,  "A-3", 30, 80),
        ("6189-9355", "河泽三贤", "黑白色", 26, "A-4", 20, 45),
        ("PP5412002", "鹤壁三贤", "白色",   2,  "A-5", 40, 120),
        ("7283-1020", "鹤壁三贤", "黑色",   4,  "A-6", 25, 40),
        ("PP5409301", "河泽三贤", "黑色",   6,  "A-7", 30, 60),
        ("6189-0131", "河泽三贤", "黄色",   2,  "A-8", 20, 12),
    ]
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")
    mid = {}
    for code, customer, color, pin, loc, safety, init_qty in materials:
        cur = conn.execute(
            "INSERT INTO materials (code,name,customer,color,pin_count,spec,unit,location,"
            "safety_stock,note,created_at) VALUES (?,?,?,?,?, '', '个', ?,?, '', ?)",
            (code, code, customer, color, pin, loc, safety, now))
        mid[code] = cur.lastrowid
        conn.execute(
            "INSERT INTO records (type,material_id,quantity,operator,note,created_at)"
            " VALUES ('期初',?,?, '系统', '盘点录入', ?)", (mid[code], init_qty, yesterday))
    demo_records = [
        ("出库", "PP5415101", 20, "李师傅", "线束车间领用", yesterday),
        ("出库", "7283-1020", 22, "王师傅", "组装领用", yesterday),
        ("入库", "PP5411001", 50, "向立森", "供应商到货", now),
        ("出库", "6189-7599", 30, "李师傅", "线束车间领用", now),
    ]
    conn.executemany(
        "INSERT INTO records (type,material_id,quantity,operator,note,created_at)"
        " VALUES (?,?,?,?,?,?)",
        [(t, mid[c], q, op, note, ts) for t, c, q, op, note, ts in demo_records])
    conn.commit()
    conn.close()


# ---------------- 页面：库存查询 ----------------

def page_inventory():
    st.header("库存查询")
    df = inventory_df()

    low = df[df["状态"] == "需补货"]
    if len(low) > 0:
        st.error(f"有 {len(low)} 种物料低于安全库存，需要采购/补货：" +
                 "、".join(f"{r['料号']}（剩{r['库存']:g}{r['单位']}）" for _, r in low.iterrows()))

    keyword = st.text_input("搜索（料号 / 客户 / 颜色 / 位置）", placeholder="例如：6189、三贤、黑色、A-1")
    show = df
    if keyword:
        mask = (df["料号"].str.contains(keyword, case=False, na=False)
                | df["客户"].str.contains(keyword, case=False, na=False)
                | df["颜色"].str.contains(keyword, case=False, na=False)
                | df["位置"].str.contains(keyword, case=False, na=False))
        show = df[mask]

    def highlight(row):
        color = "background-color: #FCEBEB" if row["状态"] == "需补货" else ""
        return [color] * len(row)

    st.caption("提示：点击表格中的一行，可直接跳转到下方该物料的详情。")
    event = st.dataframe(show.drop(columns=["id"]).style.apply(highlight, axis=1),
                 use_container_width=True, hide_index=True,
                 on_select="rerun", selection_mode="single-row", key="inv_table")

    # 物料详情：查看产品照片/图纸（料号+客户区分，同料号不同客户是不同物料）
    pick_map = {f"{r['料号']} | {r['客户']}": r["id"] for _, r in df.iterrows()}
    pick_key = "inv_material_pick"

    selected_rows = event.selection.rows if event and event.selection else []
    if selected_rows:
        clicked_id = int(show.iloc[selected_rows[0]]["id"])
        if clicked_id != st.session_state.get("_inv_last_table_pick"):
            st.session_state["_inv_last_table_pick"] = clicked_id
            label = next((k for k, v in pick_map.items() if v == clicked_id), None)
            if label:
                st.session_state[pick_key] = label
                st.session_state["_inv_scroll_to_detail"] = True

    st.subheader("物料详情（照片 / 图纸）", anchor="material-detail")
    pick = st.selectbox("选择物料", ["（请选择）"] + list(pick_map.keys()), key=pick_key)
    if st.session_state.pop("_inv_scroll_to_detail", False):
        components.html("""
            <script>
                var el = window.parent.document.getElementById('material-detail');
                if (el) { el.scrollIntoView({behavior: 'smooth', block: 'start'}); }
            </script>
        """, height=0)
    if pick != "（请选择）":
        mid = pick_map[pick]
        conn = get_conn()
        info = conn.execute(
            "SELECT code,customer,color,pin_count,spec,location,safety_stock,note"
            " FROM materials WHERE id=?", (mid,)).fetchone()
        conn.close()
        stock = df[df["id"] == mid].iloc[0]
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(
                f"**料号**：{info[0]}  \n**客户**：{info[1] or '-'}  \n"
                f"**颜色**：{info[2] or '-'}  \n**器件数目**：{info[3] if info[3] is not None else '-'}  \n"
                f"**位置**：{info[5] or '-'}  \n**当前库存**：{stock['库存']:g}{stock['单位']}  \n"
                f"**备注**：{info[7] or '-'}")
        with c2:
            show_material_files(mid, info[0])


# ---------------- 页面：库位视图 ----------------

def page_locations():
    """按位置(序号)分组看库存：找料时用"""
    st.header("库位视图")
    df = inventory_df()
    if df.empty:
        st.warning("还没有物料数据。")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("位置数量", df["位置"].replace("", "未分配").nunique())
    c2.metric("物料种类", len(df))
    c3.metric("低于安全库存", int((df["状态"] == "需补货").sum()))

    keyword = st.text_input("搜索物料（料号 / 客户）")
    if keyword:
        df = df[df["料号"].str.contains(keyword, case=False, na=False)
                | df["客户"].str.contains(keyword, case=False, na=False)]

    df = df.copy()
    df["位置"] = df["位置"].replace("", "未分配")

    def loc_key(loc):
        """位置自然排序：A-2 排在 A-10 前面"""
        m = re.match(r"^([A-Za-z]+)-(\d+)$", loc)
        return (m.group(1), int(m.group(2))) if m else (loc, 0)

    for loc in sorted(df["位置"].unique(), key=loc_key):
        sub = df[df["位置"] == loc]
        with st.expander(f"位置 {loc}（{len(sub)} 种物料）", expanded=True):
            st.dataframe(sub[["料号", "客户", "颜色", "器件数目", "库存", "单位", "状态"]],
                         use_container_width=True, hide_index=True)


# ---------------- 页面：入库 / 出库 ----------------

def page_record(rtype):
    st.header(f"{rtype}登记")
    conn = get_conn()
    mats = conn.execute(
        "SELECT id, code, customer, color, pin_count, unit, location"
        " FROM materials ORDER BY location, code"
    ).fetchall()
    conn.close()
    if not mats:
        st.warning("还没有物料，请先到「物料管理」添加或导入。")
        return

    options = {}
    for mid, code, cust, color, pin, unit, loc in mats:
        pin_s = f"{pin}个" if pin is not None else ""
        label = f"{loc or '-'} | {code} | {cust} | {color} {pin_s}"
        options[label] = (mid, code, unit)

    label = st.selectbox("选择物料", list(options.keys()))
    mid, code, unit = options[label]

    conn = get_conn()
    current = get_stock(conn, mid)
    conn.close()
    st.info(f"当前库存：{current:g} {unit}")

    with st.form(f"form_{rtype}", clear_on_submit=True):
        qty = st.number_input(f"{rtype}数量（{unit}）", min_value=0.0, step=1.0, format="%g")
        operator_label = "经手人" if rtype == "入库" else "领用人"
        operator = st.text_input(operator_label, placeholder="谁办的这事")
        note = st.text_input("备注", placeholder="例如：供应商到货 / 线束车间领用")
        submitted = st.form_submit_button(f"确认{rtype}", type="primary")

    if submitted:
        if qty <= 0:
            st.error("数量必须大于 0")
        elif not operator.strip():
            st.error(f"请填写{operator_label}")
        else:
            ok, msg = add_record(rtype, mid, qty, operator.strip(), note.strip())
            if ok:
                st.success(f"{rtype}成功：{code} {qty:g}{unit}")
                st.rerun()
            else:
                st.error(msg)


# ---------------- 页面：出入流水 ----------------

def page_records():
    st.header("出入流水")
    col1, col2 = st.columns(2)
    with col1:
        days = st.selectbox("时间范围", ["今天", "最近7天", "最近30天", "全部"], index=1)
    with col2:
        rtype = st.selectbox("类型", ["全部", "入库", "出库", "期初", "盘点"])

    sql = """
        SELECT r.created_at AS 时间, r.type AS 类型,
               m.code AS 料号, m.customer AS 客户,
               r.quantity AS 数量, m.unit AS 单位,
               r.operator AS 经手人, r.note AS 备注
        FROM records r JOIN materials m ON m.id = r.material_id
        WHERE 1=1
    """
    params = []
    if days != "全部":
        n = {"今天": 0, "最近7天": 6, "最近30天": 29}[days]
        since = (date.today() - timedelta(days=n)).strftime("%Y-%m-%d")
        sql += " AND r.created_at >= ?"
        params.append(since)
    if rtype != "全部":
        sql += " AND r.type = ?"
        params.append(rtype)
    sql += " ORDER BY r.id DESC LIMIT 500"

    conn = get_conn()
    df = pd.read_sql_query(sql, conn, params=params)
    conn.close()

    keyword = st.text_input("搜索（料号 / 客户 / 经手人）")
    if keyword:
        df = df[df["料号"].str.contains(keyword, case=False, na=False)
                | df["客户"].str.contains(keyword, case=False, na=False)
                | df["经手人"].str.contains(keyword, case=False, na=False)]
    st.dataframe(df, use_container_width=True, hide_index=True)

    if len(df) > 0:
        csv = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("导出当前结果为 CSV（Excel可打开）", csv,
                           file_name=f"流水_{date.today()}.csv", mime="text/csv")


# ---------------- 页面：AI 问答（规则意图路由，离线可用） ----------------

def ai_answer(question):
    """
    意图路由版 AI：把问题分类到固定意图 → 调用写死的查询 → 包装成人话。
    数字全部来自数据库，AI 不编数据。
    以后接大模型时，只需替换"意图识别"部分，下面的查询逻辑不变。
    """
    q = question.strip()
    df = inventory_df()
    if df.empty:
        return "仓库还没有物料数据，请先到「物料管理」添加或导入。"

    def lcs_len(a, b):
        """最长公共子串长度：解决“7599还有吗”命中“6189-7599”这类问法"""
        dp = [0] * (len(b) + 1)
        best = 0
        for i in range(1, len(a) + 1):
            prev = 0
            for j in range(1, len(b) + 1):
                tmp = dp[j]
                dp[j] = prev + 1 if a[i - 1] == b[j - 1] else 0
                prev = tmp
                best = max(best, dp[j])
        return best

    # 物料命中：完整料号优先精确匹配；没有精确命中才用模糊（LCS>=3）
    hits = [r for _, r in df.iterrows() if str(r["料号"]).lower() in q.lower()]
    if not hits:
        hits = [r for _, r in df.iterrows() if lcs_len(q, str(r["料号"])) >= 3]

    def fmt(r):
        pin = f"{int(r['器件数目'])}个" if pd.notna(r["器件数目"]) else ""
        return (f"· {r['料号']}（{r['客户']} {r['颜色']} {pin}）："
                f"{r['库存']:g}{r['单位']}，位置 {r['位置'] or '未分配'}")

    # 意图1：低库存 / 该采购什么
    if any(k in q for k in ["预警", "补货", "采购", "不足", "该买", "低于"]):
        low = df[df["状态"] == "需补货"]
        if low.empty:
            return "目前没有低于安全库存的物料，库存都正常。"
        lines = [f"有 {len(low)} 种物料该补货了："]
        for _, r in low.iterrows():
            lines.append(fmt(r) + f"（安全库存 {r['安全库存']:g}）")
        return "\n".join(lines)

    # 意图2：今天/昨天的出入库流水
    if any(k in q for k in ["今天", "今日", "昨天", "昨日"]):
        is_today = ("今天" in q) or ("今日" in q)
        d = date.today() if is_today else date.today() - timedelta(days=1)
        label = "今天" if is_today else "昨天"
        conn = get_conn()
        rows = conn.execute("""
            SELECT r.type, m.code, r.quantity, m.unit, r.operator, r.note
            FROM records r JOIN materials m ON m.id = r.material_id
            WHERE r.created_at LIKE ? AND r.type != '期初'
            ORDER BY r.id
        """, (d.strftime("%Y-%m-%d") + "%",)).fetchall()
        conn.close()
        if not rows:
            return f"{label}（{d}）没有出入库记录。"
        n_in = sum(1 for r in rows if r[0] == "入库")
        n_out = sum(1 for r in rows if r[0] == "出库")
        lines = [f"{label}共 {len(rows)} 笔记录（入库 {n_in} 笔、出库 {n_out} 笔）："]
        for t, code, qty, unit, op, note in rows:
            lines.append(f"· 【{t}】{code} {qty:g}{unit} —— {op}" + (f"（{note}）" if note else ""))
        return "\n".join(lines)

    # 意图3：某物料的领用记录（谁领的）
    if hits and any(k in q for k in ["谁领", "领走", "领了", "领用"]):
        conn = get_conn()
        rows = conn.execute("""
            SELECT r.created_at, r.quantity, r.operator, r.note
            FROM records r WHERE r.material_id = ? AND r.type = '出库'
            ORDER BY r.id DESC LIMIT 5
        """, (hits[0]["id"],)).fetchall()
        conn.close()
        if not rows:
            return f"{hits[0]['料号']}（{hits[0]['客户']}）还没有领用记录。"
        lines = [f"{hits[0]['料号']}（{hits[0]['客户']}）最近的领用记录："]
        for ts, qty, op, note in rows:
            lines.append(f"· {ts}：{op} 领了 {qty:g}{hits[0]['单位']}"
                         + (f"（{note}）" if note else ""))
        return "\n".join(lines)

    # 意图4：某物料放在哪 / 还剩多少
    if hits and any(k in q for k in ["在哪", "哪里", "位置", "放哪", "多少", "还剩",
                                     "还有", "库存", "有没有"]):
        lines = []
        for r in hits:
            status = ("正常" if r["库存"] >= r["安全库存"]
                      else f"偏低（安全库存 {r['安全库存']:g}），建议补货")
            lines.append(fmt(r) + f"，{status}")
        return "\n".join(lines)

    # 意图5：某客户的物料
    customers = [c for c in df["客户"].unique() if c]
    hit_customer = next((c for c in customers if c in q), None)
    if hit_customer and any(k in q for k in ["哪些", "什么", "有", "料", "库存", "多少"]):
        sub = df[df["客户"] == hit_customer]
        lines = [f"{hit_customer} 共有 {len(sub)} 种物料："]
        for _, r in sub.iterrows():
            lines.append(fmt(r))
        return "\n".join(lines)

    # 意图6：某颜色的物料
    colors = [c for c in df["颜色"].unique() if c]
    hit_color = next((c for c in colors if c in q), None)
    if hit_color and any(k in q for k in ["哪些", "什么", "有", "料"]):
        sub = df[df["颜色"] == hit_color]
        lines = [f"{hit_color}的物料有 {len(sub)} 种："]
        for _, r in sub.iterrows():
            lines.append(fmt(r))
        return "\n".join(lines)

    # 意图7：库存总览
    if any(k in q for k in ["全部", "所有", "总共", "一共"]):
        low_n = int((df["状态"] == "需补货").sum())
        return (f"仓库共有 {len(df)} 种物料，其中 {low_n} 种低于安全库存。"
                f"想了解具体某种物料，直接问我「XX还剩多少」。")

    # 兜底：告诉用户能问什么
    return ("我暂时能回答这几类问题：\n"
            "· 查库存：「PP5415101还剩多少」\n"
            "· 查位置：「6189-7599放在哪」\n"
            "· 查客户：「河泽三贤有哪些料」\n"
            "· 查颜色：「黑色的料有哪些」\n"
            "· 查预警：「哪些料该补货了」\n"
            "· 查流水：「今天出了什么货」「昨天谁领了什么」\n"
            "（后续接入大模型后，能听懂更复杂的问法）")


def page_ai():
    st.header("AI 问答")
    st.caption("离线规则版：数字全部来自数据库，不需要联网、不产生费用。")

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []
    for role, text in st.session_state.chat_history:
        with st.chat_message(role):
            st.write(text)

    q = st.chat_input("试试：PP5415101还剩多少 / 河泽三贤有哪些料 / 哪些料该补货了")
    if q:
        st.session_state.chat_history.append(("user", q))
        st.session_state.chat_history.append(("assistant", ai_answer(q)))
        st.rerun()


# ---------------- 页面：物料管理 ----------------

def page_materials():
    st.header("物料管理")

    # ---- Excel 批量导入 ----
    with st.expander("从 Excel 批量导入物料（推荐）", expanded=True):
        st.caption("支持你们现有的清单格式，列名需包含：序号/位置、客户名称、物料名称、颜色、PIN数或器件数目（备注可选）")
        up = st.file_uploader("选择 Excel 或 CSV 文件", type=["xlsx", "xls", "csv"])
        if up is not None:
            try:
                if up.name.lower().endswith(".csv"):
                    imp = pd.read_csv(up)
                else:
                    imp = pd.read_excel(up)
                # 列名映射（兼容几种叫法）
                colmap = {}
                for c in imp.columns:
                    cs = str(c).strip()
                    if cs in ("序号", "位置", "库位"):
                        colmap[c] = "location"
                    elif cs in ("客户名称", "客户"):
                        colmap[c] = "customer"
                    elif cs in ("物料名称", "料号", "物料编码"):
                        colmap[c] = "code"
                    elif cs in ("颜色",):
                        colmap[c] = "color"
                    elif cs in ("PIN数", "PIN数目", "PIN", "pin数", "器件数目", "器件数量"):
                        colmap[c] = "pin_count"
                    elif cs in ("备注",):
                        colmap[c] = "note"
                    elif cs in ("安全库存",):
                        colmap[c] = "safety_stock"
                    elif cs in ("库存", "期初库存", "数量"):
                        colmap[c] = "init_qty"
                imp = imp.rename(columns=colmap)
                if "code" not in imp.columns:
                    st.error("没找到「物料名称/料号」列，请检查列名。")
                else:
                    st.write(f"识别到 {len(imp)} 行，预览：")
                    st.dataframe(imp.head(10), use_container_width=True, hide_index=True)
                    if st.button("确认导入", type="primary"):
                        conn = get_conn()
                        n_ok, n_skip = 0, 0
                        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        for _, r in imp.iterrows():
                            code = str(r.get("code", "")).strip()
                            customer = str(r.get("customer", "") or "").strip()
                            if not code or code.lower() == "nan":
                                n_skip += 1
                                continue
                            # 唯一性 = 料号 + 客户（同料号不同客户允许共存）
                            exists = conn.execute(
                                "SELECT 1 FROM materials WHERE code=? AND customer=?",
                                (code, customer)).fetchone()
                            if exists:
                                n_skip += 1
                                continue
                            pin = r.get("pin_count")
                            pin = int(pin) if pd.notna(pin) and str(pin).strip() != "" else None
                            safety = r.get("safety_stock")
                            safety = float(safety) if pd.notna(safety) else 0
                            cur = conn.execute(
                                "INSERT INTO materials (code,name,customer,color,pin_count,"
                                "spec,unit,location,safety_stock,note,created_at)"
                                " VALUES (?,?,?,?,?, '', '个', ?,?,?, ?)",
                                (code, code, customer,
                                 str(r.get("color", "") or ""),
                                 pin,
                                 str(r.get("location", "") or ""),
                                 safety,
                                 str(r.get("note", "") or ""), now))
                            init_qty = r.get("init_qty")
                            if pd.notna(init_qty) and float(init_qty) > 0:
                                conn.execute(
                                    "INSERT INTO records (type,material_id,quantity,"
                                    "operator,note,created_at) VALUES ('期初',?,?, '系统','建账期初',?)",
                                    (cur.lastrowid, float(init_qty), now))
                            n_ok += 1
                        conn.commit()
                        conn.close()
                        st.success(f"导入完成：新增 {n_ok} 种，跳过 {n_skip} 种（已存在或为空）")
                        st.rerun()
            except Exception as e:
                st.error(f"读取文件失败：{e}")

    # ---- 手动新增 ----
    with st.expander("手动新增单个物料", expanded=False):
        # 位置助手：显示各排最新位置，自动建议下一个编号（在表单外，可实时联动）
        conn = get_conn()
        all_locs = [r[0] for r in conn.execute(
            "SELECT DISTINCT location FROM materials WHERE location != ''")]
        conn.close()
        prefixes = {}
        for loc in all_locs:
            m = re.match(r"^([A-Za-z]+)-(\d+)$", loc)
            if m:
                p, n = m.group(1).upper(), int(m.group(2))
                prefixes[p] = max(prefixes.get(p, 0), n)
        if prefixes:
            tips = "　".join(f"{p}排用到 {p}-{n}" for p, n in sorted(prefixes.items()))
            st.info(f"现有位置：{tips}")

        lc1, lc2 = st.columns(2)
        with lc1:
            pre_opts = sorted(prefixes) + ["（新排）"]
            pre = st.selectbox("位置-排", pre_opts, key="loc_pre") if pre_opts else "（新排）"
            if pre == "（新排）":
                pre = st.text_input("新排名称", placeholder="如 C", key="loc_new_pre").strip().upper()
        with lc2:
            next_n = prefixes.get(pre, 0) + 1 if pre else 1
            # 编号自动跟随最新"下一个"：只要用户没手动改过（当前值还等于上次
            # 算出的建议值），就刷新为最新建议；手动改过的保留用户输入
            prev_next = st.session_state.get("_loc_prev_next")
            cur_val = st.session_state.get("loc_num")
            if cur_val is None or cur_val == prev_next:
                st.session_state["loc_num"] = next_n
            st.session_state["_loc_prev_next"] = next_n
            num = st.number_input("位置-编号（已自动填下一个）", min_value=1,
                                  step=1, key="loc_num")
        location = f"{pre}-{num}" if pre else ""
        st.caption(f"将保存为位置：**{location or '未设置'}**")

        # 客户跟随位置排：自动预选该排主力客户（少数例外可下拉改选/新增）
        conn = get_conn()
        pc_rows = conn.execute(
            "SELECT substr(location,1,1) p, customer, COUNT(*) c FROM materials"
            " WHERE location != '' AND customer != ''"
            " GROUP BY p, customer ORDER BY p, c DESC").fetchall()
        all_custs = [r[0] for r in conn.execute(
            "SELECT DISTINCT customer FROM materials WHERE customer != ''"
            " ORDER BY customer")]
        conn.close()
        main_cust = {}
        for p_, cust_, _cnt in pc_rows:
            main_cust.setdefault(p_, cust_)          # 每排数量最多的客户
        dft = main_cust.get(pre, "")
        cust_opts = ([dft] if dft else []) + [c for c in all_custs if c != dft] \
            + ["（新客户）"]
        cust_sel = st.selectbox(
            "客户（已按位置排自动选好，例外可改）", cust_opts,
            key=f"cust_sel_{pre or 'new'}")
        new_cust = ""
        if cust_sel == "（新客户）":
            new_cust = st.text_input("新客户名称", placeholder="如 河泽三贤",
                                     key="cust_new_name").strip()
        customer = new_cust if cust_sel == "（新客户）" else cust_sel

        with st.form("form_new_material", clear_on_submit=True):
            st.caption(f"位置 **{location or '未设置'}** ｜ 客户 **{customer or '未选'}**"
                       "（在表单上方选，例外情况才需要改）")
            c1, c2 = st.columns(2)
            with c1:
                code = st.text_input("料号 *", placeholder="如 PP5415101")
                color = st.text_input("颜色", placeholder="如 黑色")
                pin = st.number_input("器件数目", min_value=0, step=1, value=0)
            with c2:
                unit = st.text_input("单位", value="个")
                safety = st.number_input("安全库存", min_value=0.0, step=1.0, format="%g")
                init_qty = st.number_input("期初库存（现有多少）", min_value=0.0, step=1.0, format="%g")
            note = st.text_input("备注")
            st.caption("手机上点上传可直接调相机拍摄，产品照片建议一次选满6个方向")
            product = st.file_uploader("产品照片（可多选，最多6张）",
                                       type=["jpg", "jpeg", "png", "webp", "pdf"],
                                       accept_multiple_files=True)
            drawing = st.file_uploader("图纸照片（可多选，如有）",
                                       type=["jpg", "jpeg", "png", "webp", "pdf"],
                                       accept_multiple_files=True)
            submitted = st.form_submit_button("保存", type="primary")
        if submitted:
            if not code.strip():
                st.error("料号必填")
            elif not customer.strip():
                st.error("客户名称必填（同料号不同客户是不同的物料，靠客户区分）")
            else:
                conn = get_conn()
                # 唯一性 = 料号 + 客户
                exists = conn.execute(
                    "SELECT 1 FROM materials WHERE code=? AND customer=?",
                    (code.strip(), customer.strip())).fetchone()
                if exists:
                    conn.close()
                    st.error(f"客户「{customer}」下已存在料号 {code}。"
                             f"如果是另一个客户的同号物料，请把客户名称改对后再保存。")
                else:
                    cur = conn.execute(
                        "INSERT INTO materials (code,name,customer,color,pin_count,spec,unit,"
                        "location,safety_stock,note,created_at)"
                        " VALUES (?,?,?,?,?, '', ?,?,?,?,?)",
                        (code.strip(), code.strip(), customer.strip(), color.strip(),
                         pin if pin > 0 else None, unit.strip() or "个", location.strip(),
                         safety, note.strip(),
                         datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    new_id = cur.lastrowid
                    conn.commit()
                    conn.close()
                    fkey = file_key(new_id, code.strip())
                    if product:
                        save_uploads(product, fkey, "产品照片")
                    if drawing:
                        save_uploads(drawing, fkey, "图纸")
                    if init_qty > 0:
                        add_record("期初", new_id, init_qty, "系统", "建账期初")
                    st.session_state.pop("loc_num", None)   # 让位置编号重新自动计算
                    st.success(f"物料 {code}（{customer}）已添加，位置 {location}")
                    st.rerun()

    # ---- 给已有物料上传/更新照片 ----
    with st.expander("给已有物料上传照片/图纸（手机可直接拍照）", expanded=False):
        conn = get_conn()
        mats = conn.execute("SELECT id, code, customer FROM materials"
                            " ORDER BY code, customer").fetchall()
        conn.close()
        if mats:
            up_map = {f"{code} | {cust}": (mid, code) for mid, code, cust in mats}
            sel = st.selectbox("选择物料（料号 | 客户）", list(up_map.keys()),
                               key="upload_sel")
            st.caption("在手机上打开本页面，点「选择文件」会直接弹出相机，可现场连拍")
            pf = st.file_uploader("产品照片（可多选，最多6张，会覆盖旧照片）",
                                  type=["jpg", "jpeg", "png", "webp", "pdf"],
                                  accept_multiple_files=True, key="pf")
            df_ = st.file_uploader("图纸照片（可多选，会覆盖旧图纸）",
                                   type=["jpg", "jpeg", "png", "webp", "pdf"],
                                   accept_multiple_files=True, key="df")
            if st.button("保存照片", type="primary"):
                if not pf and not df_:
                    st.warning("请先选择文件")
                else:
                    mid, code = up_map[sel]
                    fkey = file_key(mid, code)
                    msg = []
                    if pf:
                        msg.append(f"产品照片 {save_uploads(pf, fkey, '产品照片')} 张")
                    if df_:
                        msg.append(f"图纸 {save_uploads(df_, fkey, '图纸')} 张")
                    st.success("已保存：" + "、".join(msg))
                    st.rerun()

    # ---- 导出物料清单（Excel，可打印贴箱） ----
    st.subheader("导出清单（Excel）")
    st.caption("按新增日期 + 位置排筛选生成 Excel；选多个排时会分成多个工作表，一箱一张方便打印")
    df_all = inventory_df()
    conn = get_conn()
    meta = {r[0]: (r[1], r[2]) for r in conn.execute(
        "SELECT id, created_at, note FROM materials")}
    conn.close()
    df_all["建档时间"] = df_all["id"].map(lambda i: (meta.get(i) or ("", ""))[0] or "")
    df_all["备注"] = df_all["id"].map(lambda i: (meta.get(i) or ("", ""))[1] or "")
    ec1, ec2 = st.columns([1, 2])
    with ec1:
        exp_date = st.date_input("新增日期", value=datetime.now().date())
        only_day = st.checkbox("只导出该日新增的", value=True)
    with ec2:
        all_pre = sorted({str(l)[0] for l in df_all["位置"]
                          if l and str(l)[0].isalpha()})
        exp_pre = st.multiselect("位置排（默认全部）", all_pre, default=all_pre)
    exp = df_all.copy()
    if only_day:
        exp = exp[exp["建档时间"].str.startswith(str(exp_date))]
    if exp_pre:
        exp = exp[exp["位置"].fillna("").str[:1].isin(exp_pre)]
    st.caption(f"命中 **{len(exp)}** 种物料")
    if len(exp):
        cols = ["位置", "客户", "料号", "颜色", "器件数目", "库存", "备注", "建档时间"]
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            if exp_pre and len(exp_pre) > 1:
                for p in exp_pre:
                    sub = exp[exp["位置"].fillna("").str[:1] == p]
                    if len(sub):
                        sub[cols].to_excel(w, index=False, sheet_name=f"{p}排")
            else:
                exp[cols].to_excel(w, index=False, sheet_name="物料清单")
        day_str = str(exp_date) if only_day else "全部"
        st.download_button(
            f"下载 Excel（{len(exp)} 种）", buf.getvalue(),
            file_name=f"物料清单_{day_str}_{'-'.join(exp_pre) or '全部'}排.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary")

    # ---- 现有物料清单（直接编辑全部字段 + 盘点录入） ----
    st.subheader("现有物料（点击单元格直接改；盘点时填「实盘数量」）")
    st.caption("位置/客户/料号/颜色/器件数目/安全库存都能直接改；「库存」由流水算出不能手改，"
               "盘点时在「实盘数量」填实际数，保存后自动生成盘点流水把库存调对。")
    df = inventory_df()
    edit_df = df[["位置", "客户", "料号", "颜色", "器件数目", "库存",
                  "安全库存", "状态"]].copy()
    edit_df["实盘数量"] = None
    edited = st.data_editor(
        edit_df,
        disabled=["库存", "状态"],
        use_container_width=True, hide_index=True,
        column_config={
            "实盘数量": st.column_config.NumberColumn(
                "实盘数量（盘点用）", min_value=0, step=1,
                help="盘点时填实际数量；留空表示本次不盘点该物料"),
        })
    if st.button("保存修改", type="primary"):
        conn = get_conn()
        n_cnt, errs = 0, []
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for i, row in edited.iterrows():
            mid = int(df.iloc[i]["id"])
            cust = str(row["客户"] or "").strip()
            if not cust:
                errs.append(f"第{i+1}行（料号{row['料号']}）客户为空，未保存")
                continue
            pin_v = row["器件数目"]
            pin_v = int(pin_v) if pd.notna(pin_v) else None
            try:
                conn.execute(
                    "UPDATE materials SET location=?, customer=?, code=?, color=?,"
                    " pin_count=?, safety_stock=? WHERE id=?",
                    (str(row["位置"] or "").strip(), cust,
                     str(row["料号"] or "").strip(), str(row["颜色"] or "").strip(),
                     pin_v, float(row["安全库存"] or 0), mid))
            except sqlite3.IntegrityError:
                errs.append(f"{row['料号']}（{cust}）与其他物料撞号，未保存")
                continue
            pv = row.get("实盘数量")
            if pv is not None and pd.notna(pv):
                diff = float(pv) - float(row["库存"])
                if abs(diff) > 1e-9:
                    conn.execute(
                        "INSERT INTO records (type,material_id,quantity,operator,"
                        "note,created_at) VALUES ('盘点',?,?,?,?,?)",
                        (mid, diff, "盘点", f"实盘 {float(pv):g}", now))
                    n_cnt += 1
        conn.commit()
        conn.close()
        msg = "修改已保存"
        if n_cnt:
            msg += f"，生成 {n_cnt} 条盘点流水（库存已按实盘调整）"
        if errs:
            msg += "；注意：" + "；".join(errs)
        st.success(msg)
        st.rerun()


# ---------------- 主程序 ----------------

def main():
    st.set_page_config(page_title="工厂仓库管理系统（演示版）", page_icon="W", layout="wide")
    init_db()
    auto_backup()
    seed_demo_data()
    log_visit()
    check_magic_link()

    st.sidebar.title("工厂仓库管理系统（演示版）")
    page = st.sidebar.radio("功能", ["库存查询", "库位视图", "入库登记", "出库登记",
                                     "出入流水", "AI 问答", "物料管理"])
    st.sidebar.markdown("---")
    st.sidebar.caption("数据文件：warehouse.db\n备份目录：backups/（每天自动备份）\n照片目录：uploads/")
    with st.sidebar.expander("访问统计"):
        if st.session_state.get("_full_unlocked"):
            today_visits, total_visits = visit_stats()
            st.write(f"今日访问：{today_visits} 次")
            st.write(f"累计访问：{total_visits} 次")
            st.caption("最近访问（局域网IP + 时间）：")
            for visited_at, ip in recent_visits(20):
                st.caption(f"{visited_at}　{ip or '未知IP'}")
        else:
            pwd = st.text_input("密码", type="password", key="_stats_pwd")
            if pwd:
                if pwd == FULL_ACCESS_PASSWORD:
                    st.session_state["_full_unlocked"] = True
                    st.rerun()
                else:
                    st.error("密码错误")
    if st.sidebar.button("立即手动备份"):
        os.makedirs(BACKUP_DIR, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy(DB_PATH, os.path.join(BACKUP_DIR, f"warehouse_{ts}.db"))
        st.sidebar.success("已备份到 backups/")
    if st.session_state.get("_full_unlocked"):
        st.sidebar.success("已解锁全部功能")
        if st.sidebar.button("重新锁定"):
            st.session_state["_full_unlocked"] = False
            st.rerun()

    if page == "库存查询":
        page_inventory()
    elif page == "库位视图":
        page_locations()
    elif page == "入库登记":
        if require_full_access():
            page_record("入库")
    elif page == "出库登记":
        if require_full_access():
            page_record("出库")
    elif page == "出入流水":
        page_records()
    elif page == "AI 问答":
        page_ai()
    elif page == "物料管理":
        if require_full_access():
            page_materials()


if __name__ == "__main__":
    main()
