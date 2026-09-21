# -*- coding: utf-8 -*-
"""
工厂仓库管理系统（公开演示版）
================================
物料格式对齐工厂实际清单：
    位置(序号) / 客户名称 / 料号 / 颜色 / 器件数目(导入后作为库存数量) / 产品照片(PDF) / 图纸照片 / 备注

功能：库存查询（搜索+低库存标红+物料详情照片）、库位视图（库区→格子，可清零/批量释放）、
      入库、出库（领用/报废）、
      出入流水（可导出CSV）、AI问答（离线规则版）、物料管理（含Excel导入、照片上传）、
      每日自动备份

技术：Streamlit（网页） + SQLite（单文件数据库 warehouse.db）
启动：streamlit run app.py --server.address 0.0.0.0 --server.port 8501
局域网内手机/电脑浏览器访问： http://本机IP:8501
"""

import base64
import difflib
import io
import json
import os
import re
import shutil
import sqlite3
import unicodedata
import zipfile
from datetime import datetime, date, timedelta

import pandas as pd
import qrcode
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image, ImageDraw, ImageFont

# ---------------- 路径配置 ----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "warehouse.db")      # 数据库就是这个文件，复制它=备份
BACKUP_DIR = os.path.join(BASE_DIR, "backups")        # 自动备份存放目录
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")        # 产品照片/图纸存放目录

IMAGE_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
BIG_LOC_START = 1000  # 大件料位置编号从这之后排（A-1001 起），手动新增时跟普通料分开算"下一个"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

def load_access_password():
    """从部署环境读取编辑密码，源码中不保存本机凭据。"""
    password = os.environ.get("WAREHOUSE_EDIT_PASSWORD")
    if password is not None:
        return password
    try:
        return str(st.secrets.get("WAREHOUSE_EDIT_PASSWORD", ""))
    except FileNotFoundError:
        return ""


FULL_ACCESS_PASSWORD = load_access_password()


def get_setting(key, default=""):
    conn = get_conn()
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    conn.close()
    return row[0] if row else default


def set_setting(key, value):
    conn = get_conn()
    conn.execute(
        "INSERT INTO settings (key, value) VALUES (?, ?)"
        " ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, value))
    conn.commit()
    conn.close()


def qr_payload(material_id):
    """二维码内容：网址链接（局域网访问地址 + 物料主键id），手机自带相机/扫码枪
    扫到就能直接打开这个物料的详情页。带主键id，不是料号（同料号可能对应不同客户，
    料号也可能改）。访问地址在"批量打印二维码"页设置，改了要重新打印标签。
    生产系统/条码枪那边如果只想认编号，可以从网址里解析出 id= 后面的数字
    （stock_api.py 的 lookup() 已经支持这么解析）。"""
    base = get_setting("base_url", "").rstrip("/")
    return f"{base}/?id={int(material_id)}"

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
        code          TEXT NOT NULL,      -- 料号，如 DEMO-001
        name          TEXT DEFAULT '',    -- 名称（可空，显示用料号）
        customer      TEXT DEFAULT '',    -- 客户名称，如 客户B（通用杂物可留空，不强制归到客户名下）
        category      TEXT DEFAULT '',    -- 类别，如 电线类/五金杂件/客户定制件，自由填写，用于筛选
        color         TEXT DEFAULT '',    -- 颜色
        pin_count     INTEGER,            -- [已弃用] 原"器件数目"字段，数据已一次性迁移进库存（见 init_db 里的迁移代码），
                                           -- 保留这一列只是为了兼容老库文件，代码里不再读写它
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
    CREATE TABLE IF NOT EXISTS staff (
        id   INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE   -- 领用人姓名，出库登记时下拉选
    );
    CREATE TABLE IF NOT EXISTS settings (
        key   TEXT PRIMARY KEY,
        value TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS submissions (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        kind         TEXT NOT NULL,        -- 出库提报 / 报错留言
        material_id  INTEGER NOT NULL,
        quantity     REAL,                 -- 出库提报：申请数量；报错留言不用
        submitter    TEXT DEFAULT '',      -- 提交人
        message      TEXT DEFAULT '',      -- 出库提报的备注 / 报错留言的正文
        status       TEXT NOT NULL DEFAULT '待处理',  -- 待处理/已确认/已驳回/已处理
        record_id    INTEGER,              -- 出库提报确认后对应生成的 records.id
        created_at   TEXT NOT NULL,
        handled_at   TEXT DEFAULT '',
        handled_note TEXT DEFAULT '',
        FOREIGN KEY (material_id) REFERENCES materials(id)
    );
    CREATE TABLE IF NOT EXISTS material_aliases (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        material_id   INTEGER NOT NULL,   -- 关联 materials.id
        code          TEXT NOT NULL,      -- 别名料号（同一物料，客户/别的单子上叫的别的编号）
        created_at    TEXT NOT NULL,
        UNIQUE(material_id, code),
        FOREIGN KEY (material_id) REFERENCES materials(id)
    );
    -- 辅料/配件：探针、U叉、卡扣、线材等通用耗材，跟客户认定的连接器物料(materials)是两回事，
    -- 不认客户、不挂产品照片，但同一型号常常分装在好几个盒子/库位里，所以不能像materials
    -- 那样用"料号+客户"唯一——这里每个(类别,型号,位置)组合就是独立一行，天然支持同型号跨位置。
    CREATE TABLE IF NOT EXISTS supplies (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        category      TEXT DEFAULT '',    -- 类别，如 探针/U叉/卡扣/线材，自由填写
        code          TEXT NOT NULL,      -- 型号/编号
        location      TEXT DEFAULT '',    -- 位置/盒子编号
        unit          TEXT DEFAULT 'PCS', -- 单位
        safety_stock  REAL DEFAULT 0,     -- 安全库存
        note          TEXT DEFAULT '',
        status        TEXT DEFAULT '在用',-- 在用/停用（停用=不再占用这个位置）
        created_at    TEXT,
        UNIQUE(category, code, location)
    );
    CREATE TABLE IF NOT EXISTS supply_records (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        type          TEXT NOT NULL,      -- 期初 / 入库 / 出库
        supply_id     INTEGER NOT NULL,   -- 关联 supplies.id
        quantity      REAL NOT NULL,
        operator      TEXT DEFAULT '',
        note          TEXT DEFAULT '',
        created_at    TEXT NOT NULL,
        FOREIGN KEY (supply_id) REFERENCES supplies(id)
    );
    """)
    visit_cols = [r[1] for r in conn.execute("PRAGMA table_info(visits)")]
    if "ip_address" not in visit_cols:
        conn.execute("ALTER TABLE visits ADD COLUMN ip_address TEXT DEFAULT ''")

    # materials 加库位状态字段：位置属于料号、不属于库存数量，数量归零不等于库位释放，
    # 只有走"释放库位"才会把 status 改成停用、位置清空
    mat_cols = [r[1] for r in conn.execute("PRAGMA table_info(materials)")]
    if "status" not in mat_cols:
        conn.execute("ALTER TABLE materials ADD COLUMN status TEXT DEFAULT '在用'")
    if "kind" not in mat_cols:
        conn.execute("ALTER TABLE materials ADD COLUMN kind TEXT DEFAULT 'connector'")
        conn.execute("UPDATE materials SET kind='connector' WHERE kind IS NULL OR kind=''")
    if "category" not in mat_cols:
        conn.execute("ALTER TABLE materials ADD COLUMN category TEXT DEFAULT ''")

    # records 加出库分类字段（幂等：先查列存在再 ALTER）
    rec_cols = [r[1] for r in conn.execute("PRAGMA table_info(records)")]
    if "sub_type" not in rec_cols:
        conn.execute("ALTER TABLE records ADD COLUMN sub_type TEXT DEFAULT ''")
    if "link_id" not in rec_cols:
        conn.execute("ALTER TABLE records ADD COLUMN link_id INTEGER")
    if "expect_return" not in rec_cols:
        conn.execute("ALTER TABLE records ADD COLUMN expect_return TEXT DEFAULT ''")
    if "task_id" not in rec_cols:
        conn.execute("ALTER TABLE records ADD COLUMN task_id TEXT DEFAULT ''")
    if "stage_id" not in rec_cols:
        conn.execute("ALTER TABLE records ADD COLUMN stage_id TEXT DEFAULT ''")
    # 老的出库记录没有子类型，统一回填为"领用"（幂等：只补没打过标的）
    conn.execute(
        "UPDATE records SET sub_type='领用' WHERE type='出库' AND (sub_type IS NULL OR sub_type='')")
    # 人员名单表首次建立时，从历史流水的经手人里把名字捞出来，省得从空列表开始选
    conn.execute(
        "INSERT OR IGNORE INTO staff (name)"
        " SELECT DISTINCT TRIM(operator) FROM records WHERE TRIM(operator) != ''")
    # "器件数目"字段废弃：数据一次性迁移进库存。只给还一条流水都没有、
    # 但 pin_count 有值的物料补一条期初流水，已经有流水的物料不动（避免重复叠加、
    # 也不会覆盖后续真实录入的库存），迁移过一次之后天然幂等
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for mid, pin in conn.execute(
            "SELECT id, pin_count FROM materials WHERE pin_count IS NOT NULL AND pin_count != 0"):
        has_record = conn.execute(
            "SELECT 1 FROM records WHERE material_id=? LIMIT 1", (mid,)).fetchone()
        if not has_record:
            conn.execute(
                "INSERT INTO records (type,material_id,quantity,operator,note,created_at)"
                " VALUES ('期初',?,?, '系统','由器件数目字段迁移',?)",
                (mid, pin, now_str))
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


def get_supply_stock(conn, supply_id):
    """单条辅料/配件当前库存 = 期初 + 入库 - 出库"""
    row = conn.execute("""
        SELECT COALESCE(SUM(CASE WHEN type IN ('期初','入库') THEN quantity ELSE -quantity END), 0)
        FROM supply_records WHERE supply_id = ?
    """, (supply_id,)).fetchone()
    return row[0] or 0


def get_staff_list():
    """人员名单，按姓名排序供下拉选择"""
    conn = get_conn()
    names = [r[0] for r in conn.execute("SELECT name FROM staff ORDER BY name")]
    conn.close()
    return names


def add_staff(name):
    """新增一个人员名字（重复则忽略），出库时选"新增人员"用"""
    name = name.strip()
    if not name:
        return
    conn = get_conn()
    conn.execute("INSERT OR IGNORE INTO staff (name) VALUES (?)", (name,))
    conn.commit()
    conn.close()


def get_material_aliases(material_id):
    """一个物料的别名料号列表：[(别名id, 料号), ...]，按添加顺序"""
    conn = get_conn()
    rows = conn.execute(
        "SELECT id, code FROM material_aliases WHERE material_id=? ORDER BY id",
        (material_id,)).fetchall()
    conn.close()
    return rows


def find_code_conflict(code, customer):
    """同一客户下，一个编号只能指向一个物料：跟已有物料的料号或别名撞了（不分大小写，
    停用的也算）都不行，不然扫码、搜索时不知道该认哪个。
    撞了返回 (撞上的物料id, 提示文字)，没撞返回 (None, "")"""
    conn = get_conn()
    r = conn.execute(
        "SELECT id, code, location, status, '料号' FROM materials"
        " WHERE COALESCE(customer,'') = ? AND code = ? COLLATE NOCASE"
        " UNION ALL"
        " SELECT m.id, m.code, m.location, m.status, '别名' FROM material_aliases a"
        " JOIN materials m ON m.id = a.material_id"
        " WHERE COALESCE(m.customer,'') = ? AND a.code = ? COLLATE NOCASE"
        " LIMIT 1",
        (customer, code, customer, code)).fetchone()
    conn.close()
    if not r:
        return None, ""
    mid, mcode, loc, status, via = r
    where = f"位置 {loc}" if loc else "没有位置"
    if status == "停用":
        where += "，已停用"
    return mid, (f"客户「{customer or '不填客户'}」下，「{code}」已经是物料 {mcode}"
                 f"（{where}）的{via}了")


def add_material_alias(material_id, code):
    """新增一个别名料号，同一客户下跟任何料号/别名撞了都报错，返回 (是否成功, 消息)"""
    code = code.strip()
    if not code:
        return False, "料号不能为空"
    conn = get_conn()
    customer = conn.execute("SELECT COALESCE(customer,'') FROM materials WHERE id=?",
                            (material_id,)).fetchone()[0]
    conn.close()
    hit_mid, msg = find_code_conflict(code, customer)
    if hit_mid is not None:
        return False, ("它本来就叫这个，不用再加" if hit_mid == material_id else msg)
    conn = get_conn()
    try:
        conn.execute(
            "INSERT INTO material_aliases (material_id, code, created_at) VALUES (?,?,?)",
            (material_id, code, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        return True, ""
    except sqlite3.IntegrityError:
        return False, "这个别名料号已经加过了"
    finally:
        conn.close()


def delete_material_alias(alias_id):
    conn = get_conn()
    conn.execute("DELETE FROM material_aliases WHERE id=?", (alias_id,))
    conn.commit()
    conn.close()


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
    if FULL_ACCESS_PASSWORD and st.query_params.get("key") == FULL_ACCESS_PASSWORD:
        st.session_state["_full_unlocked"] = True


def require_full_access():
    """入库/出库/物料管理等编辑功能的密码门：本次会话验证一次即可，返回是否已解锁"""
    if not FULL_ACCESS_PASSWORD:
        st.info("编辑功能尚未启用，请联系管理员配置编辑密码。")
        return False
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


def add_record(rtype, material_id, qty, operator, note, sub_type="", link_id=None,
                expect_return="", created_at=None):
    """写一条出入库流水。出库时开事务并校验负库存：要么成功，要么什么都不写。
    created_at 不传则用当前时间；出库表单允许事后补录，传入选定日期+当前时刻。
    返回 (是否成功, 消息, 新流水的id或None)"""
    conn = get_conn()
    try:
        conn.execute("BEGIN IMMEDIATE")
        if rtype == "出库":
            current = get_stock(conn, material_id)
            if qty > current:
                conn.rollback()
                return False, f"库存不足：当前只剩 {current:g}，不能出库 {qty:g}", None
        cur = conn.execute(
            "INSERT INTO records (type, material_id, quantity, operator, note,"
            " sub_type, link_id, expect_return, created_at)"
            " VALUES (?,?,?,?,?,?,?,?,?)",
            (rtype, material_id, qty, operator, note, sub_type, link_id, expect_return,
             created_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
        conn.commit()
        return True, "ok", cur.lastrowid
    except Exception as e:
        conn.rollback()
        return False, f"写入失败：{e}", None
    finally:
        conn.close()


def add_supply_record(rtype, supply_id, qty, operator, note, created_at=None):
    """辅料/配件的出入库流水，逻辑和 add_record 一样（出库校验负库存），只是没有
    借出/归还/工单那些物料专属字段——辅料不需要那么细。"""
    conn = get_conn()
    try:
        conn.execute("BEGIN IMMEDIATE")
        if rtype == "出库":
            current = get_supply_stock(conn, supply_id)
            if qty > current:
                conn.rollback()
                return False, f"库存不足：当前只剩 {current:g}，不能出库 {qty:g}", None
        cur = conn.execute(
            "INSERT INTO supply_records (type, supply_id, quantity, operator, note, created_at)"
            " VALUES (?,?,?,?,?,?)",
            (rtype, supply_id, qty, operator, note,
             created_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
        conn.commit()
        return True, "ok", cur.lastrowid
    except Exception as e:
        conn.rollback()
        return False, f"写入失败：{e}", None
    finally:
        conn.close()


def add_submission(kind, material_id, submitter, quantity=None, message=""):
    """提报入队，不改库存。出库提报要等 resolve_submission('confirm') 才真正写流水；
    报错留言本来就只是留言，不写流水"""
    conn = get_conn()
    conn.execute(
        "INSERT INTO submissions (kind,material_id,quantity,submitter,message,"
        "status,created_at) VALUES (?,?,?,?,?, '待处理', ?)",
        (kind, material_id, quantity, submitter.strip(), message.strip(),
         datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()


def pending_submissions(kind=None):
    """待处理提报列表，带物料信息，最早的排前面"""
    conn = get_conn()
    sql = ("SELECT s.id, s.kind, s.material_id, m.code, m.customer, m.unit, s.quantity,"
           " s.submitter, s.message, s.created_at"
           " FROM submissions s JOIN materials m ON m.id = s.material_id"
           " WHERE s.status='待处理'")
    params = []
    if kind:
        sql += " AND s.kind=?"
        params.append(kind)
    sql += " ORDER BY s.id"
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    cols = ["id", "kind", "material_id", "料号", "客户", "单位", "quantity",
            "submitter", "message", "created_at"]
    return [dict(zip(cols, r)) for r in rows]


def resolve_submission(sub_id, action, handled_note=""):
    """action: 'confirm'（出库提报：写入流水） / 'reject'（驳回，不写流水） / 'done'（报错留言：标记已处理）"""
    conn = get_conn()
    sub = conn.execute(
        "SELECT kind, material_id, quantity, submitter, message"
        " FROM submissions WHERE id=? AND status='待处理'", (sub_id,)).fetchone()
    conn.close()
    if not sub:
        return False, "这条提报已经被处理过了"
    kind, material_id, quantity, submitter, message = sub
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if action == "confirm" and kind == "出库提报":
        ok, msg, record_id = add_record("出库", material_id, quantity, submitter, message,
                                        sub_type="领用")
        if not ok:
            return False, msg
        conn = get_conn()
        conn.execute(
            "UPDATE submissions SET status='已确认', record_id=?, handled_at=?,"
            " handled_note=? WHERE id=?", (record_id, now, handled_note, sub_id))
        conn.commit()
        conn.close()
        return True, "ok"

    status = "已驳回" if action == "reject" else "已处理"
    conn = get_conn()
    conn.execute(
        "UPDATE submissions SET status=?, handled_at=?, handled_note=? WHERE id=?",
        (status, now, handled_note, sub_id))
    conn.commit()
    conn.close()
    return True, "ok"


def report_issue_widget(mid, key_prefix):
    """"这条不对"入口：不用密码，谁都能写几句话反馈，进待处理列表，负责人自己看着处理"""
    show_key = f"{key_prefix}_show_report_{mid}"
    if st.button("🚩 这条不对", key=f"{key_prefix}_report_btn_{mid}"):
        st.session_state[show_key] = not st.session_state.get(show_key, False)
    if st.session_state.get(show_key):
        with st.form(f"{key_prefix}_report_form_{mid}"):
            msg = st.text_area("哪里不对，随便写", key=f"{key_prefix}_report_msg_{mid}")
            who = st.text_input("你的名字", value=st.session_state.get("_last_reporter", ""),
                                key=f"{key_prefix}_report_who_{mid}")
            submitted = st.form_submit_button("提交")
        if submitted:
            if not msg.strip():
                st.error("写点内容再提交")
            elif not who.strip():
                st.error("填一下你的名字")
            else:
                add_submission("报错留言", mid, who.strip(), message=msg.strip())
                st.session_state["_last_reporter"] = who.strip()
                st.session_state[show_key] = False
                st.success("已提交，负责人会看到")
                st.rerun()


def inventory_df():
    """库存总表：物料档案 + 实时库存 + 状态（id 列供程序内部用，展示时去掉）"""
    conn = get_conn()
    df = pd.read_sql_query("""
        SELECT m.id       AS id,
               m.location AS 位置,
               m.customer AS 客户,
               m.category AS 类别,
               m.code     AS 料号,
               m.color    AS 颜色,
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


def supply_inventory_df():
    """辅料/配件库存总表，跟 inventory_df() 是同样的算法，只是数据来源换成 supplies/supply_records"""
    conn = get_conn()
    df = pd.read_sql_query("""
        SELECT s.id       AS id,
               s.category AS 类别,
               s.code     AS 型号,
               s.location AS 位置,
               COALESCE(SUM(CASE WHEN r.type IN ('期初','入库') THEN r.quantity ELSE -r.quantity END), 0) AS 库存,
               s.unit     AS 单位,
               s.safety_stock AS 安全库存,
               s.note     AS 备注
        FROM supplies s
        LEFT JOIN supply_records r ON r.supply_id = s.id
        WHERE s.status IS NULL OR s.status = '在用'
        GROUP BY s.id
        ORDER BY s.category,
                 CASE WHEN instr(s.location,'-')>0 THEN substr(s.location,1,instr(s.location,'-')-1)
                      ELSE s.location END,
                 CASE WHEN instr(s.location,'-')>0 THEN CAST(substr(s.location,instr(s.location,'-')+1) AS INTEGER)
                      ELSE 0 END,
                 s.code
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


def supply_file_key(sid, code):
    """辅料照片文件名前缀：sid_型号（供 supplies 表用，跟物料的 file_key 是同一套思路，
    但辅料没有客户/legacy历史包袱，直接用 id 就不会重名）"""
    return f"supply{sid}_{safe_name(code)}"


# 辅料/配件的位置排序：位置是"探针1-1"这种文本，直接按文本排会把 探针1-10 排到 探针1-2
# 前面（逐字符比较，"1"<"2"）；拆成"横杠前缀"+"横杠后数字"分别排序，数字部分才按大小排。
# 用在没有表别名的普通 sqlite3.execute 查询里（supply_inventory_df 里用的是 pandas + 表别名
# s.，是单独一份，写法一样但没法共用这个字符串）。
SUPPLY_LOCATION_ORDER_SQL = (
    "category,"
    " CASE WHEN instr(location,'-')>0 THEN substr(location,1,instr(location,'-')-1)"
    "      ELSE location END,"
    " CASE WHEN instr(location,'-')>0 THEN CAST(substr(location,instr(location,'-')+1) AS INTEGER)"
    "      ELSE 0 END,"
    " code"
)


def rename_material_files(mid, old_code, new_code, was_legacy_owner=False):
    """料号改名后，把该物料 uploads/ 里已有照片的文件名同步改过来（文件名前缀里存的是料号），
    不然改料号会导致原有照片"消失"（其实是文件名跟新料号对不上了，找不到而已）。
    早期照片没有 id 前缀、直接用料号命名（如 DEMO-001_产品照片.pdf）；这类文件只有在
    was_legacy_owner=True（即改名前这个物料是该料号唯一/最早的认领者）时才顺便迁移，
    改名后统一变成 id_新料号_ 前缀，避免以后再对不上。"""
    if not os.path.isdir(UPLOAD_DIR) or safe_name(old_code) == safe_name(new_code):
        return
    new_prefix = file_key(mid, new_code) + "_"
    id_prefix = file_key(mid, old_code) + "_"
    legacy_prefix = safe_name(old_code) + "_"
    for f in os.listdir(UPLOAD_DIR):
        if f.startswith(id_prefix):
            rest = f[len(id_prefix):]
        elif was_legacy_owner and f.startswith(legacy_prefix) and not re.match(r"^\d+_", f):
            rest = f[len(legacy_prefix):]
        else:
            continue
        try:
            os.rename(os.path.join(UPLOAD_DIR, f), os.path.join(UPLOAD_DIR, new_prefix + rest))
        except OSError:
            pass


def save_uploads(uploaded_files, fkey, kind):
    """
    保存一组照片（如产品六个方向、多张图纸）到 uploads/。
    文件命名：id_料号_类型_序号.扩展名，如 153_DEMO-001_产品照片_1.jpg
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


def append_uploads(uploaded_files, fkey, kind):
    """跟 save_uploads 一样存文件，但**不删旧的**：序号从已有的最大号往后接着排。
    扫码详情页「补充添加图纸」用——现场发现少了一张图纸随手补拍，
    绝不会把原来的图纸覆盖掉。"""
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    pat = re.compile(rf"^{re.escape(fkey)}_{re.escape(kind)}_(\d+)\.")
    start = 0
    for f in os.listdir(UPLOAD_DIR):
        m = pat.match(f)
        if m:
            start = max(start, int(m.group(1)))
    n = 0
    for i, uf in enumerate(uploaded_files, start + 1):
        ext = os.path.splitext(uf.name)[1].lower() or ".jpg"
        data = uf.getbuffer()
        if ext in IMAGE_EXT:
            data, ext = compress_image(data), ".jpg"
        with open(os.path.join(UPLOAD_DIR, f"{fkey}_{kind}_{i}{ext}"), "wb") as fh:
            fh.write(data)
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


def find_material_files(mid, code, kind):
    """跟 show_material_files 用同一套文件命名规则，但只挑某一类（产品照片/图纸），
    给扫码详情页的产品照片轮播、图纸区分别用"""
    if not os.path.isdir(UPLOAD_DIR):
        return []
    all_files = os.listdir(UPLOAD_DIR)
    prefix = f"{file_key(mid, code)}_{kind}_"
    files = [f for f in all_files if f.startswith(prefix)]

    legacy_prefix = f"{safe_name(code)}_{kind}"
    legacy = [f for f in all_files
              if f.startswith(legacy_prefix) and not re.match(r"^\d+_", f)]
    if legacy:
        conn = get_conn()
        min_id = conn.execute(
            "SELECT MIN(id) FROM materials WHERE code=?", (code,)).fetchone()[0]
        conn.close()
        if mid == min_id:
            files += legacy
    return [os.path.join(UPLOAD_DIR, f) for f in sorted(files)]


def material_photo_tiles(mid, code, kind):
    """把某一类文件（图片直接用、PDF渲染成页图）铺平成一串可以直接 st.image 的图片"""
    tiles = []
    for path in find_material_files(mid, code, kind):
        ext = os.path.splitext(path)[1].lower()
        if ext in IMAGE_EXT:
            tiles.append(path)
        elif ext == ".pdf":
            tiles.extend(render_pdf_pages(path))
    return tiles


def export_photo_tiles(mid, code, kind, limit=6):
    """导出Excel专用的取图：文件规则跟 material_photo_tiles 完全一样，但
    PDF 只按低倍率渲染前几页——导出几百种物料时，还按详情页那种 1.5 倍高清渲染会慢到等不起。
    最多取 limit 张。返回「文件路径 或 图片字节」混在一起的列表，excel_thumb 两种都吃。"""
    tiles = []
    for path in find_material_files(mid, code, kind):
        if len(tiles) >= limit:
            break
        ext = os.path.splitext(path)[1].lower()
        if ext in IMAGE_EXT:
            tiles.append(path)
        elif ext == ".pdf":
            tiles.extend(render_pdf_pages(path, max_pages=limit - len(tiles), zoom=0.8))
    return tiles[:limit]


def excel_thumb(src, box=170):
    """把一张图（路径或字节）压成能塞进Excel单元格的小图：
    最长边 box 像素、JPEG质量70，一张约10KB。几百张贴进去文件才不会大到打不开。
    返回 (图片流, 宽px, 高px)；读不出来的图返回 None（跳过，不影响整份导出）。"""
    try:
        img = Image.open(io.BytesIO(src) if isinstance(src, bytes) else src).convert("RGB")
        img.thumbnail((box, box))
        bio = io.BytesIO()
        img.save(bio, "JPEG", quality=70)
        bio.seek(0)
        return bio, img.width, img.height
    except Exception:
        return None


def build_photo_excel(cols, sheet_groups, thumb_px=170, max_photos=6, progress=None):
    """生成「带照片」的物料清单Excel。
    文字部分照旧交给 pandas 写，再用 openpyxl 往右边追加图片列：
    产品照片1..N、图纸1..M。N/M 取该工作表里照片最多的那一行，照片少的行右边留空，
    这样每一列的含义是固定的，看起来才对得齐。
    sheet_groups: [(工作表名, 该表的DataFrame), ...]，跟不带照片的导出共用同一套分表规则。
    """
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.utils import get_column_letter

    buf = io.BytesIO()
    keep = []      # 必须留住这些图片流的引用：openpyxl 是等到保存那一刻才回头去读图的
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, sub in sheet_groups:
            sub[cols].to_excel(w, index=False, sheet_name=name)
            ws = w.sheets[name]
            rows = []
            for r, (_, item) in enumerate(sub.iterrows(), start=2):   # 第1行是表头
                pics = {}
                for kind in ("产品照片", "图纸"):
                    got = [excel_thumb(s, thumb_px) for s in
                           export_photo_tiles(item["id"], item["料号"], kind, max_photos)]
                    pics[kind] = [t for t in got if t]
                rows.append((r, pics["产品照片"], pics["图纸"]))
                if progress:
                    progress()
            n_p = max([len(p) for _, p, _ in rows] + [0])
            n_d = max([len(d) for _, _, d in rows] + [0])
            base = len(cols)                       # 图片列接在文字列右边
            for i in range(n_p):
                ws.cell(row=1, column=base + i + 1, value=f"产品照片{i + 1}")
            for i in range(n_d):
                ws.cell(row=1, column=base + n_p + i + 1, value=f"图纸{i + 1}")
            for c in range(base + 1, base + n_p + n_d + 1):
                ws.column_dimensions[get_column_letter(c)].width = thumb_px / 7.0   # Excel列宽≈字符数
            for r, p, d in rows:
                if p or d:
                    ws.row_dimensions[r].height = thumb_px * 0.75                   # Excel行高单位是磅
                for i, (bio, iw, ih) in enumerate(p + d):
                    col = base + (i if i < len(p) else n_p + i - len(p)) + 1
                    xi = XLImage(bio)
                    xi.width, xi.height = iw, ih
                    ws.add_image(xi, f"{get_column_letter(col)}{r}")
                    keep.append(bio)
    return buf.getvalue()


def find_supply_files(sid, code):
    """某条辅料/配件的所有照片，按 supply_file_key 前缀匹配（没有物料那套legacy历史包袱）"""
    if not os.path.isdir(UPLOAD_DIR):
        return []
    prefix = supply_file_key(sid, code) + "_"
    return [os.path.join(UPLOAD_DIR, f) for f in sorted(os.listdir(UPLOAD_DIR))
            if f.startswith(prefix)]


def show_supply_files(sid, code):
    """展示某条辅料/配件的照片：网格显示，跟物料的 show_material_files 是同一套UI"""
    files = find_supply_files(sid, code)
    if not files:
        st.caption("还没有照片")
        return
    for i in range(0, len(files), 3):
        cols = st.columns(3)
        for col, path in zip(cols, files[i:i + 3]):
            col.image(path, use_container_width=True)


# ---------------- 二维码标签 ----------------

def _label_font(size):
    """找一个能显示中文的字体；实在找不到就退回默认字体（料号本身多是英文数字，不影响使用）"""
    for path in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\simhei.ttf",
                 r"C:\Windows\Fonts\simsun.ttc",
                 "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"):
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()


def make_qr_image(text, box_size=8):
    qr = qrcode.QRCode(border=2, box_size=box_size)
    qr.add_data(text)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("RGB")


def fit_text_to_width(draw, text, max_w, max_h_px, min_px=None):
    """从 max_h_px 往下试字号，找能让整行塞进 max_w 的最大字号；
    到最小字号还塞不下就从后面截断加省略号。返回 (font, 实际显示的文字)"""
    if min_px is None:
        min_px = max(6, int(max_h_px * 0.3))
    size = max(min_px, int(max_h_px))
    while size > min_px:
        font = _label_font(size)
        w = draw.textbbox((0, 0), text, font=font)[2]
        if w <= max_w:
            return font, text
        size -= 1
    font = _label_font(min_px)
    shown = text
    while shown and draw.textbbox((0, 0), shown + "…", font=font)[2] > max_w:
        shown = shown[:-1]
    return font, (shown + "…" if shown and shown != text else shown or text)


def fit_text_multiline(draw, text, max_w, max_h_px, max_lines=3, min_px=None):
    """跟 fit_text_to_width 类似，但一行塞不下时会自动换行（最多 max_lines 行），
    只有换到 max_lines 行、字号缩到下限还是装不下，才会截断最后一行加省略号——
    标签是给人认物料/位置用的名字，字段长就该换行，不该悄悄把后半段内容截没了。
    返回 (font, [要显示的每一行文字])"""
    if min_px is None:
        min_px = max(6, int(max_h_px * 0.12))

    def wrap(font):
        lines, i = [], 0
        while i < len(text):
            j = i + 1
            while j < len(text) and draw.textbbox((0, 0), text[i:j + 1], font=font)[2] <= max_w:
                j += 1
            lines.append(text[i:j])
            i = j
        return lines or [""]

    size = max(min_px, int(max_h_px))
    while size > min_px:
        font = _label_font(size)
        lines = wrap(font)
        if len(lines) <= max_lines and size * 1.15 * len(lines) <= max_h_px:
            return font, lines
        size -= 1

    font = _label_font(min_px)
    lines = wrap(font)
    if len(lines) <= max_lines:
        return font, lines
    lines = lines[:max_lines]
    last = lines[-1]
    while last and draw.textbbox((0, 0), last + "…", font=font)[2] > max_w:
        last = last[:-1]
    lines[-1] = (last or "") + "…"
    return font, lines


def make_qr_label(material_id, code, customer, location, size_mm=35, dpi=300):
    """贴纸样式：固定物理尺寸（默认35×35mm，300dpi，要跟打印机驱动里设的纸张尺寸一致），
    上面二维码（内容是网址链接，手机自带相机/扫码枪直接扫开；生产系统那边也能从网址里
    解析出 id 参数），下面两行字——
    第一行料号，第二行客户名称+位置编号合并显示（每行都自动缩小塞进一行，太长就截断加省略号）。
    按真实毫米出图，下载/打印出来就是这个实际大小"""
    def mm(v):
        return int(v / 25.4 * dpi)

    side = mm(size_mm)
    pad = mm(0.6)
    # 文字左右各留 2mm 安全边（原来跟二维码一样只留 0.6mm）：热敏标签机走纸和驱动定位
    # 总有零点几到一两毫米的误差，文字要是排到几乎贴边，稍微偏一点最长的那行（料号）
    # 右边就会有一截压出标签外面被裁掉。留够安全边，料号会自动缩小一点点，但不会缺角
    text_pad = mm(2.0)
    qr_size = int(side * 0.65)

    qr_img = make_qr_image(qr_payload(material_id), box_size=10).resize(
        (qr_size, qr_size), Image.NEAREST)

    img = Image.new("RGB", (side, side), "white")
    img.paste(qr_img, ((side - qr_size) // 2, pad))
    draw = ImageDraw.Draw(img)

    text_top = pad + qr_size + mm(0.3)
    line_h = (side - text_top - pad) / 2
    max_w = side - text_pad * 2

    cust_loc = " / ".join(t for t in (customer, location) if t)
    y = text_top
    for text in (code, cust_loc):
        if text:
            font, shown = fit_text_to_width(draw, text, max_w, line_h * 0.72)
            # anchor="mm" 按文字实际可见范围的正中心对齐，不用自己拿 textbbox 手算竖直位置——
            # 手算那版会因为字体的上下留白（bbox[1]不为0）算出偏低的坐标，两行字堆起来
            # 误差会越叠越多，第二行（客户/位置）最容易被挤到贴纸边缘外面
            draw.text((side // 2, y + line_h / 2), shown, fill="black", font=font, anchor="mm")
        y += line_h
    return img


def image_to_png_bytes(img):
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def build_label_sheet(items, cols=4):
    """把多个物料的标签贴排成一张网格图，方便一次性打印后裁开分贴。
    items: [(材料id, 料号, 客户, 位置), ...]"""
    labels = [make_qr_label(mid, code, customer, location) for mid, code, customer, location in items]
    cell_w = max(l.width for l in labels) + 20
    cell_h = max(l.height for l in labels) + 20
    rows = (len(labels) + cols - 1) // cols
    sheet = Image.new("RGB", (cell_w * cols, cell_h * rows), "white")
    for i, label in enumerate(labels):
        r, c = divmod(i, cols)
        x = c * cell_w + (cell_w - label.width) // 2
        y = r * cell_h + (cell_h - label.height) // 2
        sheet.paste(label, (x, y))
    return sheet


# ---------------- 位置二维码（贴在盒子/货架上，跟具体物料无关） ----------------

def location_qr_payload(location):
    """位置二维码内容：固定文本 LOC:位置编号，不是网址链接——这样"扫码上架"页识别的时候
    能跟物料二维码（网址）区分开，扫错码能给出明确提示，而不是随便解析出个东西来"""
    return f"LOC:{location.strip()}"


def make_location_label(location, width_mm=35, height_mm=35, show_qr=True, dpi=300):
    """位置标签贴纸：贴在盒子/货架上长期不变，跟哪个物料放在里面无关，
    所以不印料号/客户，只印二维码（可选）+ 位置编号本身（大字方便人眼直接认）。
    宽/高可以不一样（比如做成长条形）；不要二维码时整张标签都用来放大字，
    适合只是给人看、不需要扫码枪扫的场合（比如货架大牌）——但这种标签就不能再靠
    「扫码上架」页的扫码枪自动识别了，得手动输入位置编号。
    width_mm/height_mm 必须跟打印机驱动里设置的纸张/标签尺寸一致——这两个如果对不上，
    图案在实际标签纸上的位置会整体偏移（表现为一边留白一边被裁切），
    之前排查过一次就是驱动设的30mm、实际标签纸是35mm，两边对不上导致的"""
    def mm(v):
        return int(v / 25.4 * dpi)

    w, h = mm(width_mm), mm(height_mm)
    margin_top = mm(1.5)
    margin_bottom = mm(1.5)
    pad = mm(0.6)
    # 文字左右各留 2mm 安全边，理由跟 make_qr_label 里一样：打印总有一两毫米误差，
    # 字排到贴边就容易被裁掉一截，留边之后字会小一点点但不会缺角
    text_pad = mm(2.0)
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)

    if show_qr:
        usable_h = h - margin_top - margin_bottom
        qr_size = int(min(usable_h * 0.62, w - pad * 2))
        qr_img = make_qr_image(location_qr_payload(location), box_size=10).resize(
            (qr_size, qr_size), Image.NEAREST)
        img.paste(qr_img, ((w - qr_size) // 2, margin_top))
        text_top = margin_top + qr_size + mm(0.5)
        max_w = w - text_pad * 2
        max_h = h - margin_bottom - text_top
        text_h_ratio = 0.72
    else:
        text_top = margin_top
        max_w = w - text_pad * 2
        max_h = h - margin_top - margin_bottom
        text_h_ratio = 0.85   # 没有二维码，整张纸都给文字，字可以放得更大

    max_lines = 2 if show_qr else 3   # 有二维码时留给文字的高度少，行数也少留一点
    font, lines = fit_text_multiline(draw, location, max_w, max_h * text_h_ratio, max_lines=max_lines)
    # 用 anchor="mm" 直接按文字实际可见范围的正中心对齐，不用自己拿 textbbox 手算位置——
    # 手算那版会因为字体的上下留白（bbox[1]不为0）算出偏低的坐标，字容易被挤到贴纸边缘外面
    draw.multiline_text((w // 2, text_top + max_h // 2), "\n".join(lines),
                        fill="black", font=font, anchor="mm", align="center",
                        spacing=int(font.size * 0.2))
    return img


def build_location_label_sheet(locations, width_mm=35, height_mm=35, show_qr=True, cols=4):
    """位置标签排成一张网格图，用法跟 build_label_sheet 一样"""
    labels = [make_location_label(loc, width_mm=width_mm, height_mm=height_mm, show_qr=show_qr)
              for loc in locations]
    cell_w = max(l.width for l in labels) + 20
    cell_h = max(l.height for l in labels) + 20
    rows = (len(labels) + cols - 1) // cols
    sheet = Image.new("RGB", (cell_w * cols, cell_h * rows), "white")
    for i, label in enumerate(labels):
        r, c = divmod(i, cols)
        x = c * cell_w + (cell_w - label.width) // 2
        y = r * cell_h + (cell_h - label.height) // 2
        sheet.paste(label, (x, y))
    return sheet


# ---------------- 演示数据（仅数据库为空时写入一次） ----------------

def seed_demo_data():
    conn = get_conn()
    if conn.execute("SELECT COUNT(*) FROM materials").fetchone()[0] > 0:
        conn.close()
        return
    # 料号, 客户, 颜色, 位置, 安全库存, 期初数量
    materials = [
        ("DEMO-001", "客户A", "黑色",   "A-1", 50, 300),
        ("DEMO-002", "客户B", "黑白色", "A-2", 50, 200),
        ("DEMO-003", "客户B", "黑色",   "A-3", 30, 80),
        ("DEMO-004", "客户B", "黑白色", "A-4", 20, 45),
        ("DEMO-005", "客户A", "白色",   "A-5", 40, 120),
        ("DEMO-006", "客户A", "黑色",   "A-6", 25, 40),
        ("DEMO-007", "客户B", "黑色",   "A-7", 30, 60),
        ("DEMO-008", "客户B", "黄色",   "A-8", 20, 12),
    ]
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")
    mid = {}
    for code, customer, color, loc, safety, init_qty in materials:
        cur = conn.execute(
            "INSERT INTO materials (code,name,customer,color,spec,unit,location,"
            "safety_stock,note,created_at) VALUES (?,?,?,?, '', '个', ?,?, '', ?)",
            (code, code, customer, color, loc, safety, now))
        mid[code] = cur.lastrowid
        conn.execute(
            "INSERT INTO records (type,material_id,quantity,operator,note,created_at)"
            " VALUES ('期初',?,?, '系统', '盘点录入', ?)", (mid[code], init_qty, yesterday))
    demo_records = [
        ("出库", "DEMO-001", 20, "操作员1", "线束车间领用", yesterday),
        ("出库", "DEMO-006", 22, "操作员2", "组装领用", yesterday),
        ("入库", "DEMO-003", 50, "操作员3", "供应商到货", now),
        ("出库", "DEMO-002", 30, "操作员1", "线束车间领用", now),
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

    # 别名料号：查出来拼成一列插在"料号"后面，搜索能匹配到、结果表里也能看到命中的是哪个别名
    conn = get_conn()
    alias_rows_all = conn.execute(
        "SELECT material_id, code FROM material_aliases ORDER BY id").fetchall()
    conn.close()
    alias_map = {}
    for material_id, code in alias_rows_all:
        alias_map.setdefault(material_id, []).append(code)
    df = df.copy()
    df.insert(df.columns.get_loc("料号") + 1, "别名料号",
             df["id"].map(lambda i: "、".join(alias_map.get(i, []))))

    low = df[df["状态"] == "需补货"]
    if len(low) > 0:
        st.error(f"有 {len(low)} 种物料低于安全库存，需要采购/补货：" +
                 "、".join(f"{r['料号']}（剩{r['库存']:g}{r['单位']}）" for _, r in low.iterrows()))

    keyword = st.text_input("搜索（料号 / 别名料号 / 客户 / 类别 / 颜色 / 位置）",
                            placeholder="例如：6189、三贤、电线类、黑色、A-1")
    show = df
    if keyword:
        mask = (df["料号"].str.contains(keyword, case=False, na=False)
                | df["别名料号"].str.contains(keyword, case=False, na=False)
                | df["客户"].str.contains(keyword, case=False, na=False)
                | df["类别"].str.contains(keyword, case=False, na=False)
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
    pick_map = {f"{r['料号']}" + (f" | {r['客户']}" if r['客户'] else ""): r["id"] for _, r in df.iterrows()}
    pick_key = "inv_material_pick"

    selected_rows = event.selection.rows if event and event.selection else []
    if selected_rows and selected_rows[0] < len(show):
        clicked_id = int(show.iloc[selected_rows[0]]["id"])
        if clicked_id != st.session_state.get("_inv_last_table_pick"):
            st.session_state["_inv_last_table_pick"] = clicked_id
            label = next((k for k, v in pick_map.items() if v == clicked_id), None)
            if label:
                st.session_state[pick_key] = label
                st.session_state["_inv_scroll_to_detail"] = True

    st.subheader("选择物料", anchor="material-detail")
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
            "SELECT code,customer,color,spec,location,safety_stock,note,category"
            " FROM materials WHERE id=?", (mid,)).fetchone()
        conn.close()
        stock = df[df["id"] == mid].iloc[0]

        st.markdown("---")
        with st.expander("出库", expanded=False):
            act_mode = st.radio(
                "方式", ["直接出库（需要密码）", "出库提报（免密码，提交后等负责人确认）"],
                horizontal=True, key=f"qact_mode_{mid}")

            if act_mode == "直接出库（需要密码）":
                if require_full_access():
                    qsub_type = st.radio("出库类型", ["领用", "报废"], horizontal=True,
                                         key=f"qout_sub_{mid}")

                    staff_list = get_staff_list()
                    qop_options = staff_list + ["（新增人员）"]
                    qop_pick = st.selectbox("领用人 *", qop_options, key=f"qout_op_pick_{mid}")
                    if qop_pick == "（新增人员）":
                        qoperator = st.text_input("新人员姓名", key=f"qout_op_new_{mid}").strip()
                    else:
                        qoperator = qop_pick

                    with st.form(f"qout_form_{mid}", clear_on_submit=True):
                        qqty = st.number_input(f"出库数量（{stock['单位']}）", min_value=0.0,
                                               step=1.0, format="%g")
                        qnote = st.text_input("备注", placeholder="例如：线束车间领用",
                                              key=f"qout_note_{mid}")
                        qsubmitted = st.form_submit_button("确认出库", type="primary")

                    if qsubmitted:
                        if qqty <= 0:
                            st.error("数量必须大于 0")
                        elif not qoperator.strip():
                            st.error("请填写领用人")
                        else:
                            ok, msg, _rid = add_record(
                                "出库", mid, qqty, qoperator.strip(), qnote.strip(),
                                sub_type=qsub_type)
                            if ok:
                                add_staff(qoperator.strip())
                                st.session_state[f"qout_op_pick_{mid}"] = qoperator.strip()
                                st.success(f"出库成功：{info[0]} {qqty:g}{stock['单位']}")
                                st.rerun()
                            else:
                                st.error(msg)
            else:
                st.caption("不用密码，谁都能提交；提交后进「提报处理」待处理队列，"
                          "要等负责人确认才会真正扣库存。")
                qrep_who = st.text_input(
                    "你的名字 *", value=st.session_state.get("_last_reporter", ""),
                    key=f"qrep_who_{mid}")
                with st.form(f"qrep_form_{mid}", clear_on_submit=True):
                    qrep_qty = st.number_input(f"数量（{stock['单位']}）", min_value=0.0,
                                               step=1.0, format="%g")
                    qrep_note = st.text_input("备注（可选）", placeholder="例如：组装领用",
                                              key=f"qrep_note_{mid}")
                    qrep_submitted = st.form_submit_button("提交提报", type="primary")
                if qrep_submitted:
                    if qrep_qty <= 0:
                        st.error("数量必须大于 0")
                    elif not qrep_who.strip():
                        st.error("请填写你的名字")
                    else:
                        add_submission("出库提报", mid, qrep_who.strip(),
                                       quantity=qrep_qty, message=qrep_note.strip())
                        st.session_state["_last_reporter"] = qrep_who.strip()
                        st.success(f"已提交：{info[0]} {qrep_qty:g}{stock['单位']}，等负责人确认")
                        st.rerun()

        st.markdown("---")
        st.subheader("物料详情（照片 / 图纸）")
        c1, c2, c3 = st.columns([2, 2, 1])
        with c1:
            st.markdown(
                f"**料号**：{info[0]}  \n**客户**：{info[1] or '-'}  \n"
                f"**类别**：{info[7] or '-'}  \n"
                f"**颜色**：{info[2] or '-'}  \n"
                f"**位置**：{info[4] or '-'}  \n**当前库存**：{stock['库存']:g}{stock['单位']}  \n"
                f"**备注**：{info[6] or '-'}")
            report_issue_widget(mid, "inv")
        with c2:
            show_material_files(mid, info[0])
        with c3:
            st.caption("二维码标签（扫码直接打开这个物料）")
            label_img = make_qr_label(mid, info[0], info[1] or "", info[4])
            st.image(label_img, use_container_width=True)
            st.download_button(
                "下载", image_to_png_bytes(label_img),
                file_name=f"{info[0]}_{info[1] or ''}_二维码.png", mime="image/png",
                key=f"qr_dl_{mid}")

    st.markdown("---")
    page_supply_inventory()


# ---------------- 页面：库位视图 ----------------

LOC_RE = re.compile(r"^([A-Za-z]+)-(\d+)$")
OTHER_AREA = "其他位置"
CLEAR_REASONS = ["用完了", "丢失", "寄给客户", "无法识别", "其他"]

# 格子按钮的颜色：按钮 key 里带状态（lvbox_full_/lvbox_zero_/lvbox_none_），
# Streamlit 会给它套一个 st-key-<key> 的 class，这里按 class 前缀上色
LOC_VIEW_CSS = """
<style>
[class*="st-key-lvbox_"] button {min-height: 3rem;}
[class*="st-key-lvbox_"] button p {font-weight: 600;}
[class*="st-key-lvbox_full_"] button {background: #2e7d32; border-color: #2e7d32; color: #fff;}
[class*="st-key-lvbox_full_"] button:hover {background: #388e3c; color: #fff;}
[class*="st-key-lvbox_zero_"] button {background: #ffb300; border-color: #ffb300; color: #3e2700;}
[class*="st-key-lvbox_zero_"] button:hover {background: #ffc233; color: #3e2700;}
[class*="st-key-lvbox_none_"] button {border-style: dashed; opacity: .55;}
.lv-dot {display: inline-block; width: .9em; height: .9em; border-radius: 3px;
         vertical-align: -.1em; margin: 0 .3em 0 1em;}
</style>
"""


def location_view_df():
    """在用且有位置的物料 + 实时库存，位置拆成「排」(A) 和「号」(58)。
    停用的（已释放库位的）不算，它们已经不占格子了。
    位置不是「字母-数字」格式的（比如 78896），排记为「其他位置」、号为 None"""
    conn = get_conn()
    df = pd.read_sql_query("""
        SELECT m.id AS id, m.location AS 位置, m.customer AS 客户, m.code AS 料号,
               m.color AS 颜色, m.unit AS 单位,
               COALESCE(SUM(CASE WHEN r.type IN ('期初','入库','盘点') THEN r.quantity
                                 ELSE -r.quantity END), 0) AS 库存
        FROM materials m
        LEFT JOIN records r ON r.material_id = m.id
        WHERE m.location != '' AND (m.status IS NULL OR m.status = '在用')
        GROUP BY m.id
    """, conn)
    conn.close()
    areas, nums = [], []
    for loc in df["位置"]:
        m = LOC_RE.match(loc.strip())
        areas.append(m.group(1).upper() if m else OTHER_AREA)
        nums.append(int(m.group(2)) if m else None)
    df["排"] = areas
    df["号"] = nums
    return df


def area_title(sub, area):
    """库区名：这一排里的客户（按物料多少排）+ 排号，如「客户B · A 排」"""
    if area == OTHER_AREA:
        return OTHER_AREA
    custs = sub["客户"].replace("", "无客户").value_counts().index.tolist()
    name = " / ".join(custs[:3]) + (" 等" if len(custs) > 3 else "")
    return f"{name} · {area} 排"


def area_boxes(sub, area):
    """一个库区里的格子列表 [(格子名, 格子里的物料)]。
    标准排从 1 号排到最大号，中间没登记物料的号也列出来——那就是空位。
    最大号记在 settings 里（lv_top_A）只增不减：排尾几个格子释放了，盒子和标签还在货架上，
    不能因为暂时没登记物料就从图上消失"""
    if area == OTHER_AREA:
        return [(loc, sub[sub["位置"] == loc]) for loc in sorted(sub["位置"].unique())]
    top = int(sub["号"].max())
    saved = int(get_setting(f"lv_top_{area}", "0") or 0)
    if top > saved:
        set_setting(f"lv_top_{area}", str(top))
    top = max(top, saved)
    return [(f"{area}-{n}", sub[sub["号"] == n]) for n in range(1, top + 1)]


def box_state(items):
    """full=有货  zero=登记了物料但库存都是0（占着格子）  none=没登记物料（空位）"""
    if items.empty:
        return "none"
    return "full" if (items["库存"] > 0).any() else "zero"


def _lv_set(key, value):
    st.session_state[key] = value


def apply_location_action(sel, release, zero, note, operator):
    """对选中的物料：zero=按当前库存记一笔「出库/清零」流水，把库存扣到 0；
    release=停用 + 位置清空（跟物料管理里的「释放库位」一样），格子就空出来了。
    返回 (清零几个, 释放几个, 出错信息列表)"""
    n_zero = n_rel = 0
    errors = []
    for _, r in sel.iterrows():
        mid = int(r["id"])
        if zero and r["库存"] > 0:
            ok, msg, _rid = add_record("出库", mid, float(r["库存"]), operator, note,
                                       sub_type="清零")
            if not ok:
                errors.append(f"{r['料号']}：{msg}")
                continue
            n_zero += 1
        if release:
            conn = get_conn()
            conn.execute("UPDATE materials SET status='停用', location='' WHERE id=?", (mid,))
            conn.commit()
            conn.close()
            n_rel += 1
    return n_zero, n_rel, errors


def material_actions(sel, key):
    """格子弹窗和批量处理共用的操作区：标记为空 / 释放库位。sel 是 location_view_df 的子集"""
    if not st.session_state.get("_full_unlocked"):
        st.caption("清零、释放会改库存数据，要先在本页顶部输入密码解锁。")
        return
    if sel.empty:
        st.caption("先勾选要处理的物料。")
        return

    action = st.radio("要做什么", ["标记为空：库存清零，物料还留在这个格子",
                                   "释放库位：物料从格子里摘掉，格子空出来给别的料用"],
                      key=f"lv_act_{key}")
    release = action.startswith("释放")
    zero = True
    if release:
        zero = st.checkbox("同时把库存清零", value=True, key=f"lv_zero_{key}",
                           help="东西已经不在了就勾上，账才对得上；只是挪到别处放、东西还在就别勾")

    need_record = zero and bool((sel["库存"] > 0).any())
    note = operator = ""
    if need_record:
        c1, c2 = st.columns(2)
        reason = c1.selectbox("原因", CLEAR_REASONS, key=f"lv_reason_{key}")
        operator = c2.text_input("经手人 *", value=st.session_state.get("lv_operator", ""),
                                 key=f"lv_op_{key}")
        extra = st.text_input("补充说明（可选）", key=f"lv_extra_{key}",
                              placeholder="如：寄给丰顺的样品")
        note = reason + (f"：{extra.strip()}" if extra.strip() else "")
        st.caption(f"会给库存大于 0 的物料各记一笔出库（子类型「清零」，备注「{note}」），"
                   f"出入流水里查得到。")
    elif not release:
        st.info("选中的物料库存本来就是 0，不用清零。")
        return

    if st.button(f"确认处理这 {len(sel)} 个物料", type="primary", key=f"lv_go_{key}"):
        if need_record and not operator.strip():
            st.error("请填经手人")
            return
        n_zero, n_rel, errors = apply_location_action(sel, release, zero, note, operator.strip())
        if operator.strip():
            st.session_state["lv_operator"] = operator.strip()
        parts = []
        if n_zero:
            parts.append(f"清零 {n_zero} 个")
        if n_rel:
            parts.append(f"释放 {n_rel} 个")
        st.session_state["lv_flash"] = ("已处理：" + "、".join(parts)) if parts else "没有需要处理的"
        if errors:
            st.session_state["lv_flash_err"] = "；".join(errors)
        st.rerun()


@st.dialog("格子里的物料", width="large")
def box_dialog(box_name, mids):
    st.subheader(f"格子 {box_name}")
    if not mids:
        st.info("这个格子系统里没有登记物料，是个空位。\n\n"
                f"要放新料进来：去「扫码上架」，填好料号，位置扫这个格子上的位置码"
                f"（或手动填 {box_name}）就行。")
        return

    df = location_view_df()
    items = df[df["id"].isin(mids)]
    if items.empty:
        st.info("这个格子里的物料刚刚已经被处理掉了。")
        return
    for _, r in items.iterrows():
        mid = int(r["id"])
        with st.container(border=True):
            c1, c2 = st.columns([1, 3])
            photos = [p for p in find_material_files(mid, r["料号"], "产品照片")
                      if os.path.splitext(p)[1].lower() in IMAGE_EXT]
            if photos:
                c1.image(photos[0], width=120)
            else:
                c1.caption("没有照片")
            cust = f"　{r['客户']}" if r["客户"] else ""
            color = f"　{r['颜色']}" if r["颜色"] else ""
            c2.markdown(f"**{r['料号']}**{cust}{color}")
            c2.markdown(f"库存 **{r['库存']:g}** {r['单位']}")
            c2.checkbox("选中", key=f"lvsel_{mid}", value=len(items) == 1)

    sel = items[items["id"].map(lambda i: bool(st.session_state.get(f"lvsel_{int(i)}")))]
    st.markdown("---")
    material_actions(sel, key="box")


def locations_overview(df, areas):
    kw = st.text_input("找料：输入料号 / 客户，看它在哪个格子")
    if kw:
        hit = df[df["料号"].str.contains(kw, case=False, na=False)
                 | df["客户"].str.contains(kw, case=False, na=False)]
        if hit.empty:
            st.caption("没找到。")
        else:
            st.dataframe(hit[["位置", "客户", "料号", "颜色", "库存", "单位"]],
                         use_container_width=True, hide_index=True)

    st.caption("点一个库区进去，里面是一个个编了号的格子。")
    with st.container(horizontal=True, gap="medium"):
        for a in areas:
            sub = df[df["排"] == a]
            states = [box_state(items) for _, items in area_boxes(sub, a)]
            with st.container(border=True, width=320):
                st.markdown(f"**{area_title(sub, a)}**")
                st.caption(f"{len(states)} 格　有货 {states.count('full')}　"
                           f"占位但空 {states.count('zero')}　空格子 {states.count('none')}")
                st.button("打开", key=f"lv_open_{a}", on_click=_lv_set, args=("lv_area", a),
                          width="stretch")


def locations_area(sub, area):
    st.button("← 返回全部库区", on_click=_lv_set, args=("lv_area", None))
    st.subheader(area_title(sub, area))

    boxes = area_boxes(sub, area)
    states = [box_state(items) for _, items in boxes]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("格子", len(boxes))
    c2.metric("有货", states.count("full"))
    c3.metric("占位但空", states.count("zero"))
    c4.metric("空格子", states.count("none"))
    st.markdown(
        '<span class="lv-dot" style="background:#2e7d32;margin-left:0"></span>有货'
        '<span class="lv-dot" style="background:#ffb300"></span>登记了物料但库存 0（还占着格子）'
        '<span class="lv-dot" style="border:1px dashed #888"></span>没登记物料（空位，可以放新料）'
        '　格子号后面的「·2」表示里面有 2 种料',
        unsafe_allow_html=True)

    clicked = None
    with st.container(horizontal=True, gap="small"):
        for i, ((name, items), state) in enumerate(zip(boxes, states)):
            label = name if area == OTHER_AREA else name.split("-")[-1]
            if len(items) > 1:
                label += f" ·{len(items)}"
            tip = "、".join(items["料号"]) if not items.empty else "没有登记物料"
            if st.button(label, key=f"lvbox_{state}_{i}", help=tip,
                         width="content" if area == OTHER_AREA else 72):
                clicked = (name, [int(x) for x in items["id"]])
    if clicked:
        box_dialog(*clicked)

    with st.expander("批量处理这一排", expanded=False):
        only_zero = st.checkbox("只看库存为 0 的（登记着但已经空了）", key=f"lv_only0_{area}")
        view = sub[sub["库存"] <= 0] if only_zero else sub
        view = view.sort_values(["号", "位置", "料号"])
        pick_all = st.checkbox("全选", key=f"lv_all_{area}_{only_zero}")
        ed = view[["id", "位置", "料号", "颜色", "库存", "单位"]].copy()
        ed.insert(0, "选中", pick_all)
        edited = st.data_editor(
            ed, hide_index=True, use_container_width=True,
            disabled=["位置", "料号", "颜色", "库存", "单位"],
            column_config={"id": None, "选中": st.column_config.CheckboxColumn("选中")},
            key=f"lv_ed_{area}_{only_zero}_{pick_all}")
        sel_ids = set(edited.loc[edited["选中"], "id"])
        st.caption(f"已选 {len(sel_ids)} 个")
        material_actions(sub[sub["id"].isin(sel_ids)], key=f"batch_{area}")


def page_locations():
    """库位视图：库区（按排，如 客户B · A 排）→ 格子 → 格子里的物料，
    可以在格子里或整排批量把物料标记为空（库存清零）/ 释放库位"""
    st.header("库位视图")
    st.markdown(LOC_VIEW_CSS, unsafe_allow_html=True)

    flash = st.session_state.pop("lv_flash", None)
    if flash:
        st.success(flash)
    flash_err = st.session_state.pop("lv_flash_err", None)
    if flash_err:
        st.error("有几个没处理成功：" + flash_err)

    if not st.session_state.get("_full_unlocked"):
        with st.expander("查看不用密码；要清零、释放库位，先在这里解锁"):
            require_full_access()

    df = location_view_df()
    if df.empty:
        st.warning("还没有登记了位置的物料。")
        return
    areas = sorted(df["排"].unique(), key=lambda a: (a == OTHER_AREA, a))
    area = st.session_state.get("lv_area")
    if area in areas:
        locations_area(df[df["排"] == area], area)
    else:
        locations_overview(df, areas)


# ---------------- 页面：入库 / 出库 ----------------

def page_record(rtype):
    st.header(f"{rtype}登记")
    kind = st.radio("登记对象", ["物料", "辅料/配件"], horizontal=True, key=f"rec_kind_{rtype}")
    if kind == "辅料/配件":
        page_supply_record(rtype)
        return

    conn = get_conn()
    mats = conn.execute(
        "SELECT id, code, customer, color, unit, location"
        " FROM materials ORDER BY location, code"
    ).fetchall()
    conn.close()
    if not mats:
        st.warning("还没有物料，请先到「物料管理」添加或导入。")
        return

    options = {}
    for mid, code, cust, color, unit, loc in mats:
        label = f"{loc or '-'} | {code} | {cust} | {color}"
        options[label] = (mid, code, unit)

    label = st.selectbox("选择物料", list(options.keys()))
    mid, code, unit = options[label]

    conn = get_conn()
    current = get_stock(conn, mid)
    conn.close()
    st.info(f"当前库存：{current:g} {unit}")

    sub_type = ""
    operator_input = None
    if rtype == "出库":
        # 这几个字段放在 form 外面：选"新增人员"要即时冒出姓名输入框，form 里的控件不会实时联动
        sub_type = st.radio("出库类型", ["领用", "报废"], horizontal=True,
                            key="out_sub_type")

        staff_list = get_staff_list()
        op_options = staff_list + ["（新增人员）"]
        op_pick = st.selectbox("领用人 *", op_options, key="out_operator_pick")
        if op_pick == "（新增人员）":
            operator_input = st.text_input("新人员姓名", key="out_operator_new").strip()
        else:
            operator_input = op_pick

    with st.form(f"form_{rtype}", clear_on_submit=True):
        rec_date = None
        if rtype == "出库":
            rec_date = st.date_input("日期（默认今天，事后补录可以改成实际日期）",
                                     value=date.today())
            st.caption(f"领用人：**{operator_input or '（请在上面选择或填写新人员）'}**")
        qty = st.number_input(f"{rtype}数量（{unit}）", min_value=0.0, step=1.0, format="%g")
        if rtype == "入库":
            operator = st.text_input("经手人", placeholder="谁办的这事")
        else:
            operator = operator_input or ""
        note = st.text_input("备注", placeholder="例如：供应商到货 / 线束车间领用")
        submitted = st.form_submit_button(f"确认{rtype}", type="primary")

    if submitted:
        if qty <= 0:
            st.error("数量必须大于 0")
        elif not operator.strip():
            st.error("请填写领用人" if rtype == "出库" else "请填写经手人")
        else:
            created_at = None
            if rtype == "出库" and rec_date:
                created_at = datetime.combine(rec_date, datetime.now().time()) \
                    .strftime("%Y-%m-%d %H:%M:%S")
            ok, msg, _rid = add_record(
                rtype, mid, qty, operator.strip(), note.strip(),
                sub_type=sub_type,
                created_at=created_at)
            if ok:
                if rtype == "出库":
                    add_staff(operator.strip())
                    st.session_state["out_operator_pick"] = operator.strip()
                st.success(f"{rtype}成功：{code} {qty:g}{unit}")
                st.rerun()
            else:
                st.error(msg)


# ---------------- 页面：在外未还 ----------------

def outstanding_borrows():
    """所有还没还清的借出记录：借出数量 - 已归还数量(按link_id配对) > 0 的都算在外"""
    conn = get_conn()
    rows = conn.execute("""
        SELECT r.id, r.material_id, m.code, m.customer, m.unit,
               r.quantity, r.operator, r.created_at, r.expect_return,
               COALESCE((SELECT SUM(ret.quantity) FROM records ret
                         WHERE ret.link_id = r.id AND ret.type='入库' AND ret.sub_type='归还'), 0)
        FROM records r JOIN materials m ON m.id = r.material_id
        WHERE r.type='出库' AND r.sub_type='借出'
        ORDER BY r.created_at
    """).fetchall()
    conn.close()

    today = date.today()
    out = []
    for rid, mid, code, customer, unit, qty, operator, created_at, expect_return, returned in rows:
        remain = qty - returned
        if remain <= 1e-9:
            continue
        borrow_date = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S").date()
        overdue = bool(expect_return) and today > datetime.strptime(expect_return, "%Y-%m-%d").date()
        out.append({
            "流水id": rid, "料号": code, "客户": customer, "剩余数量": remain, "单位": unit,
            "领用人": operator, "借出日期": borrow_date.isoformat(),
            "已借天数": (today - borrow_date).days,
            "预计归还": expect_return or "-", "超期": "是" if overdue else "",
            "material_id": mid,
        })
    return out


def page_outstanding():
    st.header("在外未还")
    st.caption("借出但还没登记归还（或只归还了一部分）的记录，超过预计归还日期的整行标红")

    out = outstanding_borrows()
    if not out:
        st.success("目前没有在外未还的借出记录。")
        return

    out_df = pd.DataFrame(out)
    show_cols = ["料号", "客户", "剩余数量", "单位", "领用人", "借出日期",
                "已借天数", "预计归还", "超期"]

    def highlight_overdue(row):
        color = "background-color: #FCEBEB" if row["超期"] == "是" else ""
        return [color] * len(row)

    st.dataframe(out_df[show_cols].style.apply(highlight_overdue, axis=1),
                 use_container_width=True, hide_index=True)

    st.subheader("登记归还")
    pick_map = {
        f"{r['料号']} | {r['客户']} | 领用人{r['领用人']} | 借出{r['借出日期']}"
        f" | 剩{r['剩余数量']:g}{r['单位']}"
        + ("（超期）" if r["超期"] == "是" else ""): r
        for r in out
    }
    pick = st.selectbox("选择要登记归还的借出记录", ["（请选择）"] + list(pick_map.keys()))
    if pick != "（请选择）":
        rec = pick_map[pick]
        with st.form("form_return"):
            qty = st.number_input(
                f"归还数量（{rec['单位']}，最多 {rec['剩余数量']:g}）",
                min_value=0.0, max_value=float(rec["剩余数量"]),
                value=float(rec["剩余数量"]), step=1.0, format="%g")
            note = st.text_input("备注", placeholder="可选，例如：部分归还，剩下的下周还")
            submitted = st.form_submit_button("确认归还", type="primary")
        if submitted:
            if qty <= 0:
                st.error("归还数量必须大于 0")
            else:
                ok, msg, _rid = add_record(
                    "入库", rec["material_id"], qty, rec["领用人"], note.strip(),
                    sub_type="归还", link_id=rec["流水id"])
                if ok:
                    st.success(f"归还登记成功：{rec['料号']} {qty:g}{rec['单位']}")
                    st.rerun()
                else:
                    st.error(msg)


# ---------------- 页面：出库提报 / 提报处理 ----------------

def page_submit_out():
    st.header("出库提报")
    st.caption("不用密码，谁都能提交；提交后进待处理队列，要等负责人确认才会真正扣库存。")
    conn = get_conn()
    mats = conn.execute(
        "SELECT id, code, customer, color, unit, location FROM materials"
        " WHERE (status IS NULL OR status='在用') ORDER BY location, code"
    ).fetchall()
    conn.close()
    if not mats:
        st.warning("还没有物料，请先到「物料管理」添加或导入。")
        return

    options = {}
    for mid, code, cust, color, unit, loc in mats:
        label = f"{loc or '-'} | {code} | {cust} | {color}"
        options[label] = (mid, code, unit)

    keyword = st.text_input("搜料号 / 客户 / 位置（可选，先搜一下再选）")
    filtered = [k for k in options if not keyword or keyword.lower() in k.lower()]
    if not filtered:
        st.warning("没搜到匹配的物料")
        return
    label = st.selectbox("选择物料", filtered)
    mid, code, unit = options[label]

    conn = get_conn()
    stock = get_stock(conn, mid)
    conn.close()
    st.info(f"当前库存：{stock:g}{unit}")

    who = st.text_input("你的名字 *", value=st.session_state.get("_last_reporter", ""),
                        key="submit_out_who")
    with st.form("form_submit_out", clear_on_submit=True):
        qty = st.number_input(f"数量（{unit}）", min_value=0.0, step=1.0, format="%g")
        note = st.text_input("备注（可选）", placeholder="例如：组装领用")
        submitted = st.form_submit_button("提交提报", type="primary")
    if submitted:
        if qty <= 0:
            st.error("数量必须大于 0")
        elif not who.strip():
            st.error("请填写你的名字")
        else:
            add_submission("出库提报", mid, who.strip(), quantity=qty, message=note.strip())
            st.session_state["_last_reporter"] = who.strip()
            st.success(f"已提交：{code} {qty:g}{unit}，等负责人确认")
            st.rerun()


def page_report_review():
    st.header("提报处理")
    st.caption("出库提报要你逐条确认才会真正扣库存；报错留言看完点「标记已处理」就行")

    st.subheader("出库提报待确认")
    outs = pending_submissions("出库提报")
    if not outs:
        st.success("没有待确认的出库提报。")
    else:
        for s in outs:
            with st.container(border=True):
                st.markdown(
                    f"**{s['料号']}**（{s['客户']}）　数量 **{s['quantity']:g}{s['单位']}**  \n"
                    f"提报人：{s['submitter']}　时间：{s['created_at']}")
                if s["message"]:
                    st.caption(f"备注：{s['message']}")
                bc1, bc2 = st.columns(2)
                if bc1.button("✅ 确认（写入出库）", key=f"confirm_{s['id']}", type="primary",
                             use_container_width=True):
                    ok, msg = resolve_submission(s["id"], "confirm")
                    if ok:
                        st.success("已确认，库存已更新")
                        st.rerun()
                    else:
                        st.error(msg)
                if bc2.button("❌ 驳回", key=f"reject_{s['id']}", use_container_width=True):
                    resolve_submission(s["id"], "reject")
                    st.success("已驳回")
                    st.rerun()

    st.subheader("报错留言待处理")
    msgs = pending_submissions("报错留言")
    if not msgs:
        st.success("没有待处理的留言。")
    else:
        for s in msgs:
            with st.container(border=True):
                st.markdown(
                    f"**{s['料号']}**（{s['客户']}）　留言人：{s['submitter']}"
                    f"　时间：{s['created_at']}")
                st.write(s["message"])
                if st.button("标记已处理", key=f"done_{s['id']}", use_container_width=True):
                    resolve_submission(s["id"], "done")
                    st.success("已标记")
                    st.rerun()


# ---------------- 页面：出入流水 ----------------

def page_records():
    st.header("出入流水")
    col1, col2, col3 = st.columns(3)
    with col1:
        days = st.selectbox("时间范围", ["今天", "最近7天", "最近30天", "全部"], index=1)
    with col2:
        rtype = st.selectbox("类型", ["全部", "入库", "出库", "期初", "盘点"])
    with col3:
        sub_type = st.selectbox("子类型", ["全部", "领用", "借出", "归还", "报废", "清零", "退料", "工单扣料"])

    sql = """
        SELECT r.created_at AS 时间, r.type AS 类型, r.sub_type AS 子类型,
               m.code AS 料号, m.customer AS 客户,
               r.quantity AS 数量, m.unit AS 单位,
               r.task_id AS 工单, r.stage_id AS 工序,
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
    if sub_type != "全部":
        sql += " AND r.sub_type = ?"
        params.append(sub_type)
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

    st.markdown("---")
    page_supply_records()


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
        """最长公共子串长度：解决“7599还有吗”命中“DEMO-002”这类问法"""
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
        return (f"· {r['料号']}（{r['客户']} {r['颜色']}）："
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
            "· 查库存：「DEMO-001还剩多少」\n"
            "· 查位置：「DEMO-002放在哪」\n"
            "· 查客户：「客户B有哪些料」\n"
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

    q = st.chat_input("试试：DEMO-001还剩多少 / 客户B有哪些料 / 哪些料该补货了")
    if q:
        st.session_state.chat_history.append(("user", q))
        st.session_state.chat_history.append(("assistant", ai_answer(q)))
        st.rerun()


# ---------------- 页面：物料管理 ----------------

@st.dialog("释放库位确认")
def confirm_release_location(mid, code, customer, location, stock):
    """位置属于料号、不属于库存数量：数量归零不会自动腾位置，必须走这里手动确认释放，
    释放后 status 改停用、位置清空，这个位置才能重新分配给别的料"""
    st.warning(f"确定要释放 **{code}**（{customer}）当前占用的位置 **{location}** 吗？")
    if stock > 0:
        st.error(f"注意：这个物料当前库存还有 {stock:g}，释放库位不会清空库存数字，"
                 f"只是把它从这个位置摘掉、标记为停用，库存流水还在，历史可查。")
    c1, c2 = st.columns(2)
    if c1.button("确认释放", type="primary", use_container_width=True):
        conn = get_conn()
        conn.execute("UPDATE materials SET status='停用', location='' WHERE id=?", (mid,))
        conn.commit()
        conn.close()
        st.success(f"已释放位置 {location}")
        st.rerun()
    if c2.button("取消", use_container_width=True):
        st.rerun()


def page_materials():
    st.header("物料管理")

    # ---- Excel 批量导入 ----
    with st.expander("从 Excel 批量导入物料（推荐）", expanded=True):
        st.caption("支持你们现有的清单格式，列名需包含：序号/位置、客户名称、物料名称、颜色、器件数目/PIN数/库存（数量，作为期初库存导入）；分类/客户/备注可选，客户可留空（电线、五金杂件等通用物料）")
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
                    elif cs in ("分类", "类别", "类型"):
                        colmap[c] = "category"
                    elif cs in ("备注",):
                        colmap[c] = "note"
                    elif cs in ("安全库存",):
                        colmap[c] = "safety_stock"
                    elif cs in ("库存", "期初库存", "数量",
                                "PIN数", "PIN数目", "PIN", "pin数", "器件数目", "器件数量"):
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
                            safety = r.get("safety_stock")
                            safety = float(safety) if pd.notna(safety) else 0
                            cur = conn.execute(
                                "INSERT INTO materials (code,name,customer,category,color,"
                                "spec,unit,location,safety_stock,note,created_at)"
                                " VALUES (?,?,?,?,?, '', '个', ?,?,?, ?)",
                                (code, code, customer,
                                 str(r.get("category", "") or ""),
                                 str(r.get("color", "") or ""),
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
        # 每排分两段各自往后接：普通料 1~999，大件料 BIG_LOC_START 往后。
        # 分开记最大号，大件编到 1000+ 也不会把普通料的"下一个"带跑
        prefixes, big_prefixes = {}, {}
        for loc in all_locs:
            m = re.match(r"^([A-Za-z]+)-(\d+)$", loc)
            if m:
                p, n = m.group(1).upper(), int(m.group(2))
                bucket = big_prefixes if n >= BIG_LOC_START else prefixes
                bucket[p] = max(bucket.get(p, 0), n)
        all_pres = sorted(set(prefixes) | set(big_prefixes))
        if all_pres:
            def tip(p):
                s = f"{p}排用到 {p}-{prefixes[p]}" if p in prefixes else f"{p}排普通料还没有"
                if p in big_prefixes:
                    s += f"（大件 {p}-{big_prefixes[p]}）"
                return s
            st.info("现有位置：" + "　".join(tip(p) for p in all_pres))

        is_big = st.radio("物料大小", ["普通料", f"大件料（{BIG_LOC_START}往后）"],
                          horizontal=True, key="loc_size") != "普通料"
        lc1, lc2 = st.columns(2)
        with lc1:
            pre_opts = all_pres + ["（新排）"]
            pre = st.selectbox("位置-排", pre_opts, key="loc_pre") if pre_opts else "（新排）"
            if pre == "（新排）":
                pre = st.text_input("新排名称", placeholder="如 C", key="loc_new_pre").strip().upper()
        with lc2:
            if is_big:
                next_n = big_prefixes.get(pre, BIG_LOC_START) + 1 if pre else BIG_LOC_START + 1
            else:
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

        # 客户跟随位置排：自动预选该排主力客户（少数例外可下拉改选/新增）；
        # 通用杂物（电线/五金件等）没有客户，可以选"不填客户"
        conn = get_conn()
        pc_rows = conn.execute(
            "SELECT substr(location,1,1) p, customer, COUNT(*) c FROM materials"
            " WHERE location != '' AND customer != ''"
            " GROUP BY p, customer ORDER BY p, c DESC").fetchall()
        all_custs = [r[0] for r in conn.execute(
            "SELECT DISTINCT customer FROM materials WHERE customer != ''"
            " ORDER BY customer")]
        pcat_rows = conn.execute(
            "SELECT substr(location,1,1) p, category, COUNT(*) c FROM materials"
            " WHERE location != '' AND category != ''"
            " GROUP BY p, category ORDER BY p, c DESC").fetchall()
        all_cats = [r[0] for r in conn.execute(
            "SELECT DISTINCT category FROM materials WHERE category != ''"
            " ORDER BY category")]
        conn.close()
        main_cust = {}
        for p_, cust_, _cnt in pc_rows:
            main_cust.setdefault(p_, cust_)          # 每排数量最多的客户
        dft = main_cust.get(pre, "")
        cust_opts = ["（不填客户）"] + ([dft] if dft else []) \
            + [c for c in all_custs if c != dft] + ["（新客户）"]
        cust_sel = st.selectbox(
            "客户（已按位置排自动选好，例外可改；电线/五金杂件等通用物料可选「不填客户」）",
            cust_opts, index=1 if dft else 0,
            key=f"cust_sel_{pre or 'new'}")
        new_cust = ""
        if cust_sel == "（新客户）":
            new_cust = st.text_input("新客户名称", placeholder="如 客户B",
                                     key="cust_new_name").strip()
        customer = "" if cust_sel == "（不填客户）" else \
            (new_cust if cust_sel == "（新客户）" else cust_sel)

        main_cat = {}
        for p_, cat_, _cnt in pcat_rows:
            main_cat.setdefault(p_, cat_)            # 每排数量最多的类别
        dft_cat = main_cat.get(pre, "")
        cat_opts = ["（不填类别）"] + ([dft_cat] if dft_cat else []) \
            + [c for c in all_cats if c != dft_cat] + ["（新类别）"]
        cat_sel = st.selectbox(
            "类别（如：电线类/五金杂件/客户定制件，已按位置排自动选好，例外可改）",
            cat_opts, index=1 if dft_cat else 0,
            key=f"cat_sel_{pre or 'new'}")
        new_cat = ""
        if cat_sel == "（新类别）":
            new_cat = st.text_input("新类别名称", placeholder="如 电线类",
                                    key="cat_new_name").strip()
        category = "" if cat_sel == "（不填类别）" else \
            (new_cat if cat_sel == "（新类别）" else cat_sel)

        with st.form("form_new_material", clear_on_submit=True):
            st.caption(f"位置 **{location or '未设置'}** ｜ 客户 **{customer or '不填'}**"
                       f" ｜ 类别 **{category or '不填'}**"
                       "（在表单上方选，例外情况才需要改）")
            c1, c2 = st.columns(2)
            with c1:
                code = st.text_input("料号 *", placeholder="如 DEMO-001")
                alias = st.text_input("别名（可选）",
                                      placeholder="客户/别的单子上叫的另一个编号")
                color = st.text_input("颜色", placeholder="如 黑色")
            with c2:
                unit = st.text_input("单位", value="个")
                init_qty = st.number_input("库存", min_value=0.0, step=1.0, format="%g")
                safety = st.number_input("安全库存", min_value=0.0, step=1.0, format="%g")
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
            code, alias, customer = code.strip(), alias.strip(), customer.strip()
            if alias.lower() == code.lower():
                alias = ""                       # 别名跟料号一样就不用挂了，挂上也只是重复
            # 同一客户下料号、别名跟已有的料号或别名撞了都不让存（通用杂物客户留空，一样要求不重复）
            code_hit = find_code_conflict(code, customer)[1] if code else ""
            alias_hit = find_code_conflict(alias, customer)[1] if alias else ""
            if not code:
                st.error("料号必填")
            elif code_hit or alias_hit:
                if code_hit:
                    st.error(f"料号撞了：{code_hit}。如果其实是另一个客户的料，把客户改对再保存。")
                if alias_hit:
                    st.error(f"别名撞了：{alias_hit}。")
            else:
                conn = get_conn()
                cur = conn.execute(
                    "INSERT INTO materials (code,name,customer,category,color,spec,unit,"
                    "location,safety_stock,note,created_at)"
                    " VALUES (?,?,?,?,?, '', ?,?,?,?,?)",
                    (code, code, customer, category.strip(),
                     color.strip(), unit.strip() or "个", location.strip(),
                     safety, note.strip(),
                     datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                new_id = cur.lastrowid
                conn.commit()
                conn.close()
                if alias:
                    add_material_alias(new_id, alias)
                fkey = file_key(new_id, code)
                if product:
                    save_uploads(product, fkey, "产品照片")
                if drawing:
                    save_uploads(drawing, fkey, "图纸")
                if init_qty > 0:
                    add_record("期初", new_id, init_qty, "系统", "建账期初")
                st.session_state.pop("loc_num", None)   # 让位置编号重新自动计算
                st.success(f"物料 {code}（{customer}）已添加，位置 {location}"
                           + (f"，别名 {alias}" if alias else ""))
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

    # ---- 释放库位 ----
    with st.expander("释放库位", expanded=False):
        st.caption("位置属于料号、不属于库存数量：出库出到0不会自动腾位置，"
                   "要彻底不再用这个料号、把位置让给别的物料时，才用这里手动释放。"
                   "释放后该物料标记为停用、位置清空，库存流水历史仍然保留可查。")
        conn = get_conn()
        occ = conn.execute(
            "SELECT id, code, customer, location FROM materials"
            " WHERE location != '' AND (status IS NULL OR status = '在用')"
            " ORDER BY location, code").fetchall()
        conn.close()
        if not occ:
            st.caption("没有在用的物料占用位置。")
        else:
            rel_map = {f"{loc} | {code} | {cust}": (mid, code, cust, loc)
                      for mid, code, cust, loc in occ}
            rel_sel = st.selectbox("选择要释放的物料（位置 | 料号 | 客户）",
                                   list(rel_map.keys()), key="release_sel")
            r_mid, r_code, r_cust, r_loc = rel_map[rel_sel]
            conn = get_conn()
            r_stock = get_stock(conn, r_mid)
            conn.close()
            st.caption(f"当前库存：{r_stock:g}")
            if st.button("释放库位", key="release_btn"):
                confirm_release_location(r_mid, r_code, r_cust, r_loc, r_stock)

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
        cols = ["位置", "客户", "料号", "颜色", "库存", "备注", "建档时间"]
        # 分表规则：选了多个排就一排一个工作表，否则一整张表。带不带照片都用这一份，保证两个文件长得一样
        if exp_pre and len(exp_pre) > 1:
            groups = [(f"{p}排", exp[exp["位置"].fillna("").str[:1] == p]) for p in exp_pre]
            groups = [(n, s) for n, s in groups if len(s)]
        else:
            groups = [("物料清单", exp)]
        day_str = str(exp_date) if only_day else "全部"
        fbase = f"物料清单_{day_str}_{'-'.join(exp_pre) or '全部'}排"

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            for name, sub in groups:
                sub[cols].to_excel(w, index=False, sheet_name=name)
        dc1, dc2 = st.columns(2)
        with dc1:
            st.download_button(
                f"下载 Excel（{len(exp)} 种）", buf.getvalue(),
                file_name=f"{fbase}.xlsx", mime=XLSX_MIME, type="primary")
            st.caption("纯文字，秒下，文件小")
        with dc2:
            # 带照片的要现取现压几百张图，慢（大概几十秒），所以做成点一下才生成，
            # 生成完的结果存 session_state，页面因为别的操作重跑也不用再等一遍
            if st.button(f"生成带照片的 Excel（{len(exp)} 种）", key="exp_photo_btn"):
                bar = st.progress(0.0, text="正在取照片…")
                done = {"n": 0}

                def tick():
                    done["n"] += 1
                    bar.progress(min(done["n"] / len(exp), 1.0),
                                 text=f"正在取照片… {done['n']}/{len(exp)}")

                st.session_state["photo_xlsx"] = (
                    f"{fbase}_带照片.xlsx", build_photo_excel(cols, groups, progress=tick))
                bar.empty()
            got = st.session_state.get("photo_xlsx")
            if got and got[0].startswith(fbase):      # 换了筛选条件，旧文件就不给下了，免得下错
                st.download_button(
                    f"下载带照片 Excel（{len(got[1]) / 1048576:.1f} MB）", got[1],
                    file_name=got[0], mime=XLSX_MIME, key="exp_photo_dl")
            st.caption("每种最多贴 6 张产品照片 + 6 张图纸；贴的是缩略图，"
                       "看清晰大图请到系统里点物料详情")

    # ---- 批量生成二维码标签（打印后裁开贴标签袋，替代图纸/纸质标签） ----
    with st.expander("生成二维码标签（批量，打印后裁开贴标签袋）", expanded=False):
        st.caption("扫码能看到「料号 | 客户」；筛选出要打印的物料，生成一张标签图，"
                   "下载后打印、裁开，一张贴一个标签袋，原来的图纸/纸质标签就可以扔了")
        qc1, qc2 = st.columns(2)
        with qc1:
            q_pre = st.multiselect("位置排（默认全部）", all_pre, default=all_pre,
                                   key="qr_pre")
        with qc2:
            q_cols = st.number_input("每行排几个", min_value=1, max_value=10, value=4,
                                     step=1, key="qr_cols")
        q_df = df_all if not q_pre else df_all[df_all["位置"].fillna("").str[:1].isin(q_pre)]
        st.caption(f"命中 **{len(q_df)}** 种物料")
        if len(q_df) and st.button(f"生成标签图（{len(q_df)} 个）", key="qr_gen"):
            items = [(int(r["id"]), r["料号"], r["客户"], r["位置"])
                     for _, r in q_df.iterrows()]
            sheet = build_label_sheet(items, cols=int(q_cols))
            st.image(sheet, caption=f"共 {len(items)} 个标签", use_container_width=True)
            st.download_button(
                "下载标签图（PNG）", image_to_png_bytes(sheet),
                file_name=f"二维码标签_{'-'.join(q_pre) or '全部'}排.png",
                mime="image/png", type="primary", key="qr_dl_sheet")

    # ---- 现有物料清单（直接编辑全部字段 + 盘点录入） ----
    st.subheader("现有物料（点击单元格直接改；盘点时填「实盘数量」）")
    st.caption("位置/客户/料号/颜色/安全库存都能直接改；「库存」由流水算出不能手改，"
               "盘点时在「实盘数量」填实际数，保存后自动生成盘点流水把库存调对。")
    df = inventory_df()
    edit_df = df[["位置", "客户", "料号", "颜色", "库存",
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
            old_code = str(df.iloc[i]["料号"] or "").strip()
            new_code = str(row["料号"] or "").strip()
            was_legacy_owner = False
            if old_code != new_code:
                min_id = conn.execute(
                    "SELECT MIN(id) FROM materials WHERE code=?", (old_code,)).fetchone()[0]
                was_legacy_owner = (min_id == mid)
            try:
                conn.execute(
                    "UPDATE materials SET location=?, customer=?, code=?, color=?,"
                    " safety_stock=? WHERE id=?",
                    (str(row["位置"] or "").strip(), cust,
                     new_code, str(row["颜色"] or "").strip(),
                     float(row["安全库存"] or 0), mid))
            except sqlite3.IntegrityError:
                errs.append(f"{row['料号']}（{cust}）与其他物料撞号，未保存")
                continue
            rename_material_files(mid, old_code, new_code, was_legacy_owner)
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

    # ---- 料号别名管理：同一个物料，客户/别的单子上有时候叫别的编号 ----
    st.subheader("料号别名管理（点一下表格里要管理的那一行）")
    st.caption("同一个物料有时候客户/别的单子上会叫别的编号，这里给它挂上「别名料号」备查，"
              "不影响出入库和主料号。")
    alias_src = inventory_df()
    alias_show = alias_src[["位置", "客户", "料号", "颜色"]].copy()
    alias_event = st.dataframe(
        alias_show, use_container_width=True, hide_index=True,
        on_select="rerun", selection_mode="single-row", key="mat_alias_table")
    alias_rows = alias_event.selection.rows if alias_event and alias_event.selection else []
    if alias_rows:
        amid = int(alias_src.iloc[alias_rows[0]]["id"])
        acode = alias_src.iloc[alias_rows[0]]["料号"]
        acust = alias_src.iloc[alias_rows[0]]["客户"]
        st.markdown(f"当前选中：**{acode}**（{acust}）")
        aliases = get_material_aliases(amid)
        if aliases:
            for a_id, a_code in aliases:
                ac1, ac2 = st.columns([5, 1])
                ac1.write(a_code)
                if ac2.button("删除", key=f"mat_del_alias_{a_id}"):
                    delete_material_alias(a_id)
                    st.rerun()
        else:
            st.caption("还没有别名料号")
        new_alias = st.text_input("新增别名料号", key=f"mat_new_alias_{amid}")
        if st.button("添加别名", key=f"mat_add_alias_btn_{amid}"):
            ok, msg = add_material_alias(amid, new_alias)
            if ok:
                st.success("已添加")
                st.rerun()
            else:
                st.error(msg)
    else:
        st.caption("先点上面表格里的一行，选中要管理别名的物料")


# ---------------- 页面：辅料/配件（探针、U叉、卡扣、线材等通用耗材，跟客户物料分开管理） ----------------

def page_supply_management():
    """辅料/配件管理：独立页面，跟「物料管理」平级，不是塞在物料管理里面"""
    st.header("辅料/配件管理")
    st.caption("探针、U叉、卡扣、线材等通用耗材，跟「物料管理」（认客户的连接器类）分开记账，"
              "不认客户、按盒子/库位存")
    page_supply_manage()


def page_supply_inventory():
    st.subheader("辅料/配件库存")
    st.caption("探针、U叉、卡扣、线材等通用耗材，跟上面的物料是分开记账的")
    df = supply_inventory_df()
    if df.empty:
        st.caption("还没有辅料/配件数据，请到「辅料/配件管理」添加。")
        return

    low = df[df["状态"] == "需补货"]
    if len(low) > 0:
        st.error(f"有 {len(low)} 条低于安全库存，需要补货：" +
                 "、".join(f"{r['型号']}（{r['位置'] or '未分配'}，剩{r['库存']:g}{r['单位']}）"
                          for _, r in low.iterrows()))

    keyword = st.text_input("搜索（类别 / 型号 / 位置）",
                            placeholder="例如：探针、LK165、探针1-1", key="sup_inv_kw")
    show = df
    if keyword:
        mask = (df["类别"].str.contains(keyword, case=False, na=False)
                | df["型号"].str.contains(keyword, case=False, na=False)
                | df["位置"].str.contains(keyword, case=False, na=False))
        show = df[mask]

    def highlight(row):
        return ["background-color: #FCEBEB" if row["状态"] == "需补货" else ""] * len(row)

    # 跟物料的库存查询一个交互：点表格里的一行，直接跳到下面的详情
    st.caption("提示：点击表格中的一行，可直接跳转到下方该行的详情。")
    event = st.dataframe(show.drop(columns=["id"]).style.apply(highlight, axis=1),
                use_container_width=True, hide_index=True,
                on_select="rerun", selection_mode="single-row", key="sup_inv_table")

    if len(show) > 0:
        csv = show.drop(columns=["id"]).to_csv(index=False).encode("utf-8-sig")
        st.download_button("导出当前结果为 CSV", csv,
                           file_name=f"辅料库存_{date.today()}.csv", mime="text/csv",
                           key="sup_inv_export")

    pick_map = {f"{r['类别'] or '-'} | {r['型号']} | {r['位置'] or '未分配'}": int(r["id"])
               for _, r in df.iterrows()}
    pick_key = "sup_inv_pick"

    selected_rows = event.selection.rows if event and event.selection else []
    if selected_rows and selected_rows[0] < len(show):
        clicked_id = int(show.iloc[selected_rows[0]]["id"])
        if clicked_id != st.session_state.get("_sup_inv_last_table_pick"):
            st.session_state["_sup_inv_last_table_pick"] = clicked_id
            label = next((k for k, v in pick_map.items() if v == clicked_id), None)
            if label:
                st.session_state[pick_key] = label
                st.session_state["_sup_inv_scroll_to_detail"] = True

    st.subheader("选择辅料/配件", anchor="supply-detail")
    pick = st.selectbox("选择辅料/配件", ["（请选择）"] + list(pick_map.keys()), key=pick_key)
    if st.session_state.pop("_sup_inv_scroll_to_detail", False):
        components.html("""
            <script>
                var el = window.parent.document.getElementById('supply-detail');
                if (el) { el.scrollIntoView({behavior: 'smooth', block: 'start'}); }
            </script>
        """, height=0)
    if pick != "（请选择）":
        sid = pick_map[pick]
        row = df[df["id"] == sid].iloc[0]
        st.markdown(
            f"**类别**：{row['类别'] or '-'}　**型号**：{row['型号']}　"
            f"**位置**：{row['位置'] or '-'}　**库存**：{row['库存']:g}{row['单位']}　"
            f"**备注**：{row['备注'] or '-'}")
        show_supply_files(sid, row["型号"])


def page_supply_record(rtype):
    conn = get_conn()
    sup = conn.execute(
        "SELECT id, category, code, location, unit FROM supplies"
        " WHERE status IS NULL OR status = '在用'"
        " ORDER BY " + SUPPLY_LOCATION_ORDER_SQL
    ).fetchall()
    conn.close()
    if not sup:
        st.warning("还没有辅料/配件档案，请先到「辅料/配件管理」添加。")
        return

    options = {}
    for sid, cat, code, loc, unit in sup:
        label = f"{cat or '-'} | {code} | {loc or '-'}"
        options[label] = (sid, code, unit)
    label = st.selectbox("选择辅料/配件（类别 | 型号 | 位置）", list(options.keys()),
                         key=f"sup_{rtype}_pick")
    sid, code, unit = options[label]

    conn = get_conn()
    current = get_supply_stock(conn, sid)
    conn.close()
    st.info(f"当前库存：{current:g} {unit}")

    with st.form(f"sup_form_{rtype}", clear_on_submit=True):
        qty = st.number_input(f"{rtype}数量（{unit}）", min_value=0.0, step=1.0, format="%g")
        operator = st.text_input("经手人" if rtype == "入库" else "领用人",
                                 placeholder="谁办的这事")
        note = st.text_input("备注", placeholder="例如：车间领用 / 供应商到货")
        submitted = st.form_submit_button(f"确认{rtype}", type="primary")

    if submitted:
        if qty <= 0:
            st.error("数量必须大于 0")
        elif not operator.strip():
            st.error("请填写经手人" if rtype == "入库" else "请填写领用人")
        else:
            ok, msg, _rid = add_supply_record(rtype, sid, qty, operator.strip(), note.strip())
            if ok:
                st.success(f"{rtype}成功：{code} {qty:g}{unit}")
                st.rerun()
            else:
                st.error(msg)


def page_supply_records():
    st.subheader("辅料/配件出入流水")
    col1, col2 = st.columns(2)
    with col1:
        days = st.selectbox("时间范围", ["今天", "最近7天", "最近30天", "全部"], index=1,
                            key="sup_rec_days")
    with col2:
        rtype = st.selectbox("类型", ["全部", "入库", "出库", "期初"], key="sup_rec_type")

    sql = """
        SELECT r.created_at AS 时间, r.type AS 类型,
               s.category AS 类别, s.code AS 型号, s.location AS 位置,
               r.quantity AS 数量, s.unit AS 单位,
               r.operator AS 经手人, r.note AS 备注
        FROM supply_records r JOIN supplies s ON s.id = r.supply_id
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

    keyword = st.text_input("搜索（型号 / 位置 / 经手人）", key="sup_rec_kw")
    if keyword:
        df = df[df["型号"].str.contains(keyword, case=False, na=False)
                | df["位置"].str.contains(keyword, case=False, na=False)
                | df["经手人"].str.contains(keyword, case=False, na=False)]
    st.dataframe(df, use_container_width=True, hide_index=True)

    if len(df) > 0:
        csv = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("导出当前结果为 CSV", csv,
                           file_name=f"辅料流水_{date.today()}.csv", mime="text/csv",
                           key="sup_rec_export")


def page_supply_manage():
    with st.expander("从 Excel 批量导入（推荐）", expanded=False):
        st.caption("列名需包含：类别（探针/U叉/卡扣/线材...）、型号/编号、位置（盒子/库位）、"
                  "数量（作为期初库存导入）；备注、安全库存、单位可选。"
                  "同一型号分装在多个位置，就分成多行，每行一个位置的数量。")
        up = st.file_uploader("选择 Excel 或 CSV 文件", type=["xlsx", "xls", "csv"],
                              key="sup_import_file")
        if up is not None:
            try:
                imp = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
                colmap = {}
                for c in imp.columns:
                    cs = str(c).strip()
                    if cs in ("类别", "分类", "类型"):
                        colmap[c] = "category"
                    elif cs in ("型号", "编号", "料号", "物料名称"):
                        colmap[c] = "code"
                    elif cs in ("位置", "库位", "盒子", "盒子编号"):
                        colmap[c] = "location"
                    elif cs in ("单位",):
                        colmap[c] = "unit"
                    elif cs in ("备注",):
                        colmap[c] = "note"
                    elif cs in ("安全库存",):
                        colmap[c] = "safety_stock"
                    elif cs in ("库存", "期初库存", "数量", "总数量"):
                        colmap[c] = "init_qty"
                imp = imp.rename(columns=colmap)
                if "code" not in imp.columns:
                    st.error("没找到「型号/编号」列，请检查列名。")
                else:
                    st.write(f"识别到 {len(imp)} 行，预览：")
                    st.dataframe(imp.head(10), use_container_width=True, hide_index=True)
                    if st.button("确认导入", type="primary", key="sup_import_confirm"):
                        conn = get_conn()
                        n_ok, n_skip = 0, 0
                        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        for _, r in imp.iterrows():
                            code = str(r.get("code", "")).strip()
                            if not code or code.lower() == "nan":
                                n_skip += 1
                                continue
                            category = str(r.get("category", "") or "").strip()
                            location = str(r.get("location", "") or "").strip()
                            exists = conn.execute(
                                "SELECT 1 FROM supplies WHERE category=? AND code=? AND location=?",
                                (category, code, location)).fetchone()
                            if exists:
                                n_skip += 1
                                continue
                            safety = r.get("safety_stock")
                            safety = float(safety) if pd.notna(safety) else 0
                            unit = str(r.get("unit", "") or "").strip() or "PCS"
                            cur = conn.execute(
                                "INSERT INTO supplies (category,code,location,unit,"
                                "safety_stock,note,created_at)"
                                " VALUES (?,?,?,?,?,?,?)",
                                (category, code, location, unit, safety,
                                 str(r.get("note", "") or ""), now))
                            init_qty = r.get("init_qty")
                            if pd.notna(init_qty):
                                try:
                                    qv = float(str(init_qty).strip().split()[0])
                                except (ValueError, IndexError):
                                    qv = 0
                                if qv > 0:
                                    conn.execute(
                                        "INSERT INTO supply_records (type,supply_id,quantity,"
                                        "operator,note,created_at) VALUES ('期初',?,?, '系统','建账期初',?)",
                                        (cur.lastrowid, qv, now))
                            n_ok += 1
                        conn.commit()
                        conn.close()
                        st.success(f"导入完成：新增 {n_ok} 条，跳过 {n_skip} 条（已存在或为空）")
                        st.rerun()
            except Exception as e:
                st.error(f"读取文件失败：{e}")

    with st.expander("手动新增单条", expanded=False):
        conn = get_conn()
        all_cats = [r[0] for r in conn.execute(
            "SELECT DISTINCT category FROM supplies WHERE category != '' ORDER BY category")]
        conn.close()
        cat_opts = all_cats + ["（新类别）"]
        cat_sel = st.selectbox("类别", cat_opts, key="sup_new_cat_sel") if all_cats else "（新类别）"
        if cat_sel == "（新类别）":
            category = st.text_input("新类别名称", placeholder="如 探针/U叉/卡扣/线材",
                                     key="sup_new_cat_new").strip()
        else:
            category = cat_sel

        # 位置两种写法：「类别-编号」（如 金属治具-1，一个大箱子里按号放，只选号）或者自己填
        # （探针1-1 这种分盒分排的）。这个类别已有的位置不是「类别-数字」写法的，默认自己填。
        used = {}             # 编号 → 放在这个号上的型号（只算在用的）
        has_other_loc = False
        if category:
            box_re = re.compile(rf"^{re.escape(category)}-(\d+)$")
            conn = get_conn()
            for loc, c, status in conn.execute(
                    "SELECT location, code, status FROM supplies"
                    " WHERE category=? AND location != ''", (category,)):
                m = box_re.match(loc.strip())
                if not m:
                    has_other_loc = True
                elif status in (None, "在用"):
                    used.setdefault(int(m.group(1)), []).append(c)
            conn.close()
        box_mode = st.radio(
            "位置写法", ["编号", "自己填"], horizontal=True,
            index=1 if has_other_loc and not used else 0,
            format_func=lambda x: f"{category or '类别'}-编号" if x == "编号" else "自己填（如 探针1-1）",
            key=f"sup_new_locmode_{category}") == "编号"

        num = None
        if box_mode and category:
            next_num = max(used) + 1 if used else 1
            # 不给 key：存完一条 next_num 变了，输入框自动换成下一个号
            num = int(st.number_input(f"位置：{category}- 几号", min_value=1, step=1,
                                      value=next_num))
            if num in used:
                st.warning(f"{category}-{num} 已经放了 {'、'.join(used[num])}，换个号"
                           f"（下一个空号是 {next_num}）")
            else:
                st.caption(f"会存成「{category}-{num}」"
                           + (f"；{category} 已用到 -{max(used)}，共 {len(used)} 个号"
                              if used else "；这个类别还没编过号，从 1 开始"))

        with st.form("sup_form_new", clear_on_submit=True):
            code = st.text_input("型号/编号 *", placeholder="如 LK165-F1.5")
            if box_mode:
                location = f"{category}-{num}" if num else ""
            else:
                location = st.text_input("位置（盒子/库位）", placeholder="如 探针1-1")
            c1, c2 = st.columns(2)
            with c1:
                unit = st.text_input("单位", value="PCS")
                init_qty = st.number_input("库存", min_value=0.0, step=1.0, format="%g")
            with c2:
                safety = st.number_input("安全库存", min_value=0.0, step=1.0, format="%g")
            note = st.text_input("备注")
            st.caption("手机上点上传可直接调相机拍摄，现场拍完直接建档")
            photos = st.file_uploader("照片（可多选，如有）",
                                      type=["jpg", "jpeg", "png", "webp"],
                                      accept_multiple_files=True, key="sup_new_pf")
            submitted = st.form_submit_button("保存", type="primary")

        if submitted:
            if not code.strip():
                st.error("型号/编号必填")
            elif not category.strip():
                st.error("请选择或填写类别")
            elif box_mode and num in used:
                st.error(f"{location} 已经放了 {'、'.join(used[num])}，换个号再存")
            else:
                conn = get_conn()
                exists = conn.execute(
                    "SELECT 1 FROM supplies WHERE category=? AND code=? AND location=?",
                    (category.strip(), code.strip(), location.strip())).fetchone()
                if exists:
                    conn.close()
                    st.error("这个类别+型号+位置的组合已经存在了")
                else:
                    cur = conn.execute(
                        "INSERT INTO supplies (category,code,location,unit,safety_stock,"
                        "note,created_at) VALUES (?,?,?,?,?,?,?)",
                        (category.strip(), code.strip(), location.strip(),
                         unit.strip() or "PCS", safety, note.strip(),
                         datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    new_id = cur.lastrowid
                    conn.commit()
                    conn.close()
                    # 照片要等 INSERT 拿到 id 才能命名（文件名前缀是 supply{id}_型号），
                    # 所以放在建档之后存
                    n_photo = save_uploads(photos, supply_file_key(new_id, code.strip()),
                                           "照片") if photos else 0
                    if init_qty > 0:
                        add_supply_record("期初", new_id, init_qty, "系统", "建账期初")
                    st.success(f"已添加：{category} {code}（{location or '未分配位置'}）"
                               + (f"，照片 {n_photo} 张" if n_photo else ""))
                    st.rerun()

    with st.expander("停用 / 重新启用", expanded=False):
        st.caption("停用后不会再出现在入库/出库的选择列表里，历史流水仍然保留可查；不影响其它同型号的行。")
        conn = get_conn()
        active = conn.execute(
            "SELECT id, category, code, location FROM supplies"
            " WHERE status IS NULL OR status = '在用' ORDER BY " + SUPPLY_LOCATION_ORDER_SQL
        ).fetchall()
        inactive = conn.execute(
            "SELECT id, category, code, location FROM supplies"
            " WHERE status = '停用' ORDER BY " + SUPPLY_LOCATION_ORDER_SQL
        ).fetchall()
        conn.close()

        if active:
            deact_map = {f"{cat or '-'} | {code} | {loc or '-'}": sid
                        for sid, cat, code, loc in active}
            deact_sel = st.selectbox("选择要停用的行", list(deact_map.keys()), key="sup_deact_sel")
            if st.button("停用", key="sup_deact_btn"):
                conn = get_conn()
                conn.execute("UPDATE supplies SET status='停用' WHERE id=?",
                            (deact_map[deact_sel],))
                conn.commit()
                conn.close()
                st.success("已停用")
                st.rerun()
        else:
            st.caption("没有在用的行。")

        if inactive:
            react_map = {f"{cat or '-'} | {code} | {loc or '-'}": sid
                        for sid, cat, code, loc in inactive}
            react_sel = st.selectbox("选择要重新启用的行", list(react_map.keys()), key="sup_react_sel")
            if st.button("重新启用", key="sup_react_btn"):
                conn = get_conn()
                conn.execute("UPDATE supplies SET status='在用' WHERE id=?",
                            (react_map[react_sel],))
                conn.commit()
                conn.close()
                st.success("已重新启用")
                st.rerun()

    with st.expander("给已有辅料/配件上传照片", expanded=False):
        conn = get_conn()
        rows = conn.execute(
            "SELECT id, category, code, location FROM supplies"
            " ORDER BY " + SUPPLY_LOCATION_ORDER_SQL
        ).fetchall()
        conn.close()
        if not rows:
            st.caption("还没有辅料/配件档案")
        else:
            up_map = {f"{cat or '-'} | {code} | {loc or '-'}": (sid, code)
                     for sid, cat, code, loc in rows}
            sel = st.selectbox("选择一行（类别 | 型号 | 位置）", list(up_map.keys()),
                               key="sup_upload_sel")
            st.caption("在手机上打开本页面，点「选择文件」会直接弹出相机，可现场连拍")
            pf = st.file_uploader("照片（可多选，会覆盖这一行原有照片）",
                                  type=["jpg", "jpeg", "png", "webp"],
                                  accept_multiple_files=True, key="sup_pf")
            if st.button("保存照片", type="primary", key="sup_upload_btn"):
                if not pf:
                    st.warning("请先选择文件")
                else:
                    sid, code = up_map[sel]
                    n = save_uploads(pf, supply_file_key(sid, code), "照片")
                    st.success(f"已保存 {n} 张")
                    st.rerun()



# ---------------- 页面：扫码详情（手机端） ----------------

def page_scan_detail(id_param):
    st.markdown("""
        <style>
        .stButton > button { font-size: 1.15rem; padding: 0.7rem; }
        h1, h2, h3 { font-size: 1.4rem !important; }
        </style>
    """, unsafe_allow_html=True)

    if st.button("← 返回系统首页"):
        st.query_params.clear()
        st.rerun()

    try:
        mid = int(id_param)
    except (TypeError, ValueError):
        st.error("二维码链接不对（id 不是数字）")
        return

    conn = get_conn()
    info = conn.execute(
        "SELECT code,customer,color,location,note"
        " FROM materials WHERE id=?", (mid,)).fetchone()
    conn.close()
    if not info:
        st.error("没找到这个物料，可能已经被删除了")
        return
    code, customer, color, location, note = info
    conn = get_conn()
    stock = get_stock(conn, mid)
    unit = conn.execute("SELECT unit FROM materials WHERE id=?", (mid,)).fetchone()[0]
    conn.close()

    st.title(code)

    # ---- 顶部：产品照片轮播 ----
    tiles = material_photo_tiles(mid, code, "产品照片")
    if tiles:
        idx_key = f"scan_photo_idx_{mid}"
        idx = st.session_state.get(idx_key, 0) % len(tiles)
        st.image(tiles[idx], use_container_width=True)
        if len(tiles) > 1:
            pc1, pc2, pc3 = st.columns([1, 1, 1])
            with pc1:
                if st.button("◀ 上一张", use_container_width=True, key=f"prev_{mid}"):
                    st.session_state[idx_key] = (idx - 1) % len(tiles)
                    st.rerun()
            with pc2:
                st.markdown(f"<div style='text-align:center'>{idx+1}/{len(tiles)}</div>",
                           unsafe_allow_html=True)
            with pc3:
                if st.button("下一张 ▶", use_container_width=True, key=f"next_{mid}"):
                    st.session_state[idx_key] = (idx + 1) % len(tiles)
                    st.rerun()
    else:
        st.caption("还没有产品照片")

    st.markdown(
        f"**客户**：{customer or '-'}　**颜色**：{color or '-'}  \n"
        f"**位置**：{location or '-'}　**当前库存**：{stock:g}{unit}")
    if note:
        st.caption(f"备注：{note}")
    report_issue_widget(mid, "scan")

    # ---- 中部：四个大按钮 ----
    st.markdown("#### 操作")
    action_key = f"scan_action_{mid}"
    cur_action = st.session_state.get(action_key)

    def toggle(name):
        st.session_state[action_key] = None if cur_action == name else name
        st.rerun()

    ac1, ac2 = st.columns(2)
    with ac1:
        if st.button("📤 出库", use_container_width=True, key=f"a_out_{mid}"):
            toggle("出库")
        if st.button("📐 查看图纸", use_container_width=True, key=f"a_draw_{mid}"):
            toggle("图纸")
    with ac2:
        if st.button("📥 入库", use_container_width=True, key=f"a_in_{mid}"):
            toggle("入库")
        if st.button("➕ 添加图纸", use_container_width=True, key=f"a_adddraw_{mid}"):
            toggle("添加图纸")

    if cur_action == "出库":
        st.markdown("---")
        mode = st.radio("出库方式", ["直接出库（知道密码）", "出库提报（不用密码，等负责人确认）"],
                        key=f"scan_out_mode_{mid}")
        if mode == "直接出库（知道密码）":
            if require_full_access():
                st.info(f"当前库存：{stock:g}{unit}")
                staff_list = get_staff_list()
                op_pick = st.selectbox("领用人 *", staff_list + ["（新增人员）"],
                                       key=f"scan_op_pick_{mid}")
                operator_input = (st.text_input("新人员姓名", key=f"scan_op_new_{mid}").strip()
                                 if op_pick == "（新增人员）" else op_pick)
                with st.form(f"scan_form_出库_{mid}"):
                    qty = st.number_input(f"数量（{unit}）", min_value=0.0, step=1.0, format="%g")
                    note_in = st.text_input("备注", key=f"scan_note_{mid}")
                    submitted = st.form_submit_button("确认出库", type="primary",
                                                      use_container_width=True)
                if submitted:
                    if qty <= 0:
                        st.error("数量必须大于 0")
                    elif not operator_input.strip():
                        st.error("请填写领用人")
                    else:
                        ok, msg, _rid = add_record(
                            "出库", mid, qty, operator_input.strip(), note_in.strip(),
                            sub_type="领用")
                        if ok:
                            add_staff(operator_input.strip())
                            st.session_state[action_key] = None
                            st.success(f"出库成功：{code} {qty:g}{unit}")
                            st.rerun()
                        else:
                            st.error(msg)
        else:
            st.info(f"当前库存：{stock:g}{unit}　提交后不会立刻扣库存，要等负责人确认")
            who = st.text_input("你的名字 *", value=st.session_state.get("_last_reporter", ""),
                                key=f"scan_submit_who_{mid}")
            with st.form(f"scan_submit_out_{mid}"):
                qty = st.number_input(f"数量（{unit}）", min_value=0.0, step=1.0, format="%g")
                note_in = st.text_input("备注", key=f"scan_submit_note_{mid}")
                submitted = st.form_submit_button("提交提报", type="primary",
                                                  use_container_width=True)
            if submitted:
                if qty <= 0:
                    st.error("数量必须大于 0")
                elif not who.strip():
                    st.error("请填写你的名字")
                else:
                    add_submission("出库提报", mid, who.strip(), quantity=qty,
                                   message=note_in.strip())
                    st.session_state["_last_reporter"] = who.strip()
                    st.session_state[action_key] = None
                    st.success(f"已提交：{code} {qty:g}{unit}，等负责人确认")
                    st.rerun()

    elif cur_action == "入库":
        st.markdown("---")
        if require_full_access():
            st.info(f"当前库存：{stock:g}{unit}")
            staff_list = get_staff_list()
            op_pick = st.selectbox("经手人 *", staff_list + ["（新增人员）"],
                                   key=f"scan_in_pick_{mid}")
            operator_input = (st.text_input("新人员姓名", key=f"scan_in_new_{mid}").strip()
                             if op_pick == "（新增人员）" else op_pick)
            with st.form(f"scan_form_入库_{mid}"):
                qty = st.number_input(f"数量（{unit}）", min_value=0.0, step=1.0, format="%g")
                note_in = st.text_input("备注", key=f"scan_in_note_{mid}")
                submitted = st.form_submit_button("确认入库", type="primary",
                                                  use_container_width=True)
            if submitted:
                if qty <= 0:
                    st.error("数量必须大于 0")
                elif not operator_input.strip():
                    st.error("请填写经手人")
                else:
                    ok, msg, _rid = add_record(
                        "入库", mid, qty, operator_input.strip(), note_in.strip())
                    if ok:
                        add_staff(operator_input.strip())
                        st.session_state[action_key] = None
                        st.success(f"入库成功：{code} {qty:g}{unit}")
                        st.rerun()
                    else:
                        st.error(msg)

    elif cur_action == "图纸":
        st.markdown("---")
        draw_tiles = material_photo_tiles(mid, code, "图纸")
        if not draw_tiles:
            st.caption("还没有上传图纸，可以点上面的「➕ 添加图纸」补一张")
        else:
            for t in draw_tiles:
                st.image(t, use_container_width=True)

    elif cur_action == "添加图纸":
        st.markdown("---")
        st.caption("手机上点「Browse files」会直接弹出相机，可以对着图纸拍。"
                   "新图纸是**追加**在原有图纸后面的，不会覆盖掉已有的。")
        new_draw = st.file_uploader("图纸照片（可多选）",
                                    type=["jpg", "jpeg", "png", "webp", "pdf"],
                                    accept_multiple_files=True,
                                    key=f"scan_draw_up_{mid}")
        if st.button("保存图纸", type="primary", use_container_width=True,
                     key=f"scan_draw_save_{mid}"):
            if not new_draw:
                st.warning("请先选择文件")
            else:
                n = append_uploads(new_draw, file_key(mid, code), "图纸")
                st.session_state[action_key] = "图纸"     # 存完直接切到查看，能马上看到效果
                st.success(f"已添加 {n} 张图纸")
                st.rerun()

    # ---- 底部：最近10条流水 ----
    st.markdown("#### 最近流水")
    conn = get_conn()
    rows = conn.execute(
        "SELECT created_at, type, sub_type, quantity, operator, note"
        " FROM records WHERE material_id=? ORDER BY id DESC LIMIT 10", (mid,)).fetchall()
    conn.close()
    if not rows:
        st.caption("还没有出入库记录")
    else:
        hist_df = pd.DataFrame(rows, columns=["时间", "类型", "子类型", "数量", "经手人", "备注"])
        st.dataframe(hist_df, use_container_width=True, hide_index=True)


# ---------------- 页面：批量打印二维码 ----------------

def a4_label_grid(size_mm, margin_mm=8):
    """A4页在给定标签边长(mm)下，边距margin_mm，能整齐排下几列几行（不做像素换算，纯毫米算）"""
    usable = 210 - margin_mm * 2, 297 - margin_mm * 2
    return max(1, int(usable[0] // size_mm)), max(1, int(usable[1] // size_mm))


def build_print_html(items, size_mm=35):
    """items: [(材料id, 料号, 客户, 位置), ...]；生成一份自带打印按钮的HTML，标签边长size_mm，
    按A4自动分页排布，浏览器打印时只会打印这段内容，不会带上外面 Streamlit 的侧边栏/按钮
    （那些本来就不在这个 iframe 里）"""
    cols, rows = a4_label_grid(size_mm)
    per_page = max(1, cols * rows)
    pages = []
    for i in range(0, len(items), per_page):
        chunk = items[i:i + per_page]
        cells = []
        for mid, code, customer, location in chunk:
            qr_img = make_qr_image(qr_payload(mid), box_size=5)
            buf = io.BytesIO()
            qr_img.save(buf, format="PNG")
            b64 = base64.b64encode(buf.getvalue()).decode()
            cells.append(f'''
                <div class="cell">
                    <img src="data:image/png;base64,{b64}" />
                    <div class="code">{code}</div>
                    <div class="meta">{customer}</div>
                    <div class="meta">{location}</div>
                </div>''')
        pages.append(
            f'<div class="sheet" style="grid-template-columns:repeat({cols},1fr);'
            f'grid-template-rows:repeat({rows},1fr);">' + "".join(cells) + "</div>")

    return f"""<!doctype html><html><head><meta charset="utf-8"><style>
        body {{ font-family: "Microsoft YaHei", sans-serif; margin: 0; background: #ccc; }}
        .toolbar {{ padding: 12px; background: #fff; }}
        .toolbar button {{ font-size: 16px; padding: 8px 24px; cursor: pointer; }}
        .sheet {{
            display: grid; width: {cols * size_mm}mm; height: {rows * size_mm}mm; margin: 8mm auto;
            background: #fff; page-break-after: always;
        }}
        .cell {{
            display: flex; flex-direction: column; align-items: center; justify-content: center;
            border: 1px dashed #999; padding: 1mm; box-sizing: border-box; overflow: hidden;
        }}
        .cell img {{ width: 60%; }}
        .code {{ font-size: 10px; font-weight: bold; margin-top: 1px; }}
        .meta {{ font-size: 8px; color: #333; }}
        @page {{ size: A4; margin: 0; }}
        @media print {{
            body {{ background: #fff; }}
            .toolbar {{ display: none; }}
            .sheet {{ margin: 0; }}
        }}
        </style></head>
        <body>
        <div class="toolbar">
            <button onclick="window.print()">🖨 打印</button>
            <span>共 {len(items)} 个标签，{len(pages)} 页，每页 {cols}×{rows}</span>
        </div>
        {''.join(pages)}
        </body></html>"""


def build_a4_sheets(items, size_mm=35, dpi=300):
    """items: [(材料id, 料号, 客户, 位置), ...]；每个标签固定 size_mm×size_mm，按真实毫米在
    A4(210x297mm)上尽量铺满、自动分页，直接复用 make_qr_label 保证跟单张标签样式一致。
    可以直接当照片下载、打印时选"实际大小"（不用"缩放至页面"）就是准确尺寸"""
    def mm(v):
        return int(v / 25.4 * dpi)

    cols, rows = a4_label_grid(size_mm)
    page_w, page_h = mm(210), mm(297)
    margin = mm(8)
    cell = mm(size_mm)
    per_page = cols * rows

    sheets = []
    for i in range(0, len(items), per_page):
        chunk = items[i:i + per_page]
        page = Image.new("RGB", (page_w, page_h), "white")
        draw = ImageDraw.Draw(page)
        for idx, (mid, code, customer, location) in enumerate(chunk):
            r, c = divmod(idx, cols)
            x0, y0 = margin + c * cell, margin + r * cell
            page.paste(make_qr_label(mid, code, customer, location, size_mm=size_mm, dpi=dpi), (x0, y0))
            draw.rectangle([x0, y0, x0 + cell, y0 + cell], outline=(200, 200, 200))
        sheets.append(page)
    return sheets


def list_printers():
    """列出这台电脑能看到的打印机（含别的电脑共享出来、已经"添加"过的），
    找不到 pywin32 或没权限就返回空列表，不报错"""
    try:
        import win32print
        return [p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    except Exception:
        return []


PRINT_DOC_NAME = "仓库标签"


def get_print_offset():
    """打印位置微调（毫米）：正数往右/往下挪，负数往左/往上挪。
    标签机送纸和驱动定位常年有一两毫米的固定偏差，图案整体偏一边，一边留白一边被裁——
    这个偏差在软件里没法自动测出来，所以做成设置项，打一张看偏多少就填多少反向补回来。
    存在设置表里，所有走「直接打印」的地方（物料标签/位置标签）都会自动用上"""
    def val(key):
        try:
            return float(get_setting(key, "0") or 0)
        except ValueError:
            return 0.0
    return val("print_offset_x_mm"), val("print_offset_y_mm")


def print_label_direct(img, printer_name, size_mm=35, height_mm=None,
                       offset_x_mm=None, offset_y_mm=None):
    """把标签图片按真实物理尺寸直接发给Windows打印机，不弹打印对话框、不用先下载。
    size_mm 当宽度用；height_mm 不传就跟 size_mm 一样（正方形标签，原来的用法不用改）。
    要求：打印机驱动里的纸张/标签尺寸必须跟这里的宽高完全一样，不然图案在实际标签纸上
    的位置会整体偏移（一边留白一边被裁掉）——这个要在Windows的打印机属性里设，系统内部设不了。
    打印前先拿驱动汇报的实际纸张物理尺寸（PHYSICALWIDTH/HEIGHT）跟这里要打的尺寸对一下，
    对不上就直接报错说清楚"驱动现在是多少、你要的是多少"，不然会是"打出来了但显示不全/
    被裁切"这种不知道哪里错的情况（驱动没汇报物理尺寸的极少数情况下跳过这个检查）。
    这个函数不抛异常不代表打印机真的打出来了，只代表任务成功排进了Windows打印队列，
    配合 print_job_problem 能多查一层队列里的实际状态"""
    import win32ui
    from PIL import ImageWin

    height_mm = size_mm if height_mm is None else height_mm
    if offset_x_mm is None or offset_y_mm is None:
        saved_x, saved_y = get_print_offset()
        offset_x_mm = saved_x if offset_x_mm is None else offset_x_mm
        offset_y_mm = saved_y if offset_y_mm is None else offset_y_mm
    pdc = win32ui.CreateDC()
    pdc.CreatePrinterDC(printer_name)
    dpi_x = pdc.GetDeviceCaps(88)   # LOGPIXELSX
    dpi_y = pdc.GetDeviceCaps(90)   # LOGPIXELSY
    w_px = int(size_mm / 25.4 * dpi_x)
    h_px = int(height_mm / 25.4 * dpi_y)

    phys_w = pdc.GetDeviceCaps(110)   # PHYSICALWIDTH：驱动当前纸张设置的实际物理宽度（像素）
    phys_h = pdc.GetDeviceCaps(111)   # PHYSICALHEIGHT
    if phys_w and phys_h:
        actual_w_mm = phys_w / dpi_x * 25.4
        actual_h_mm = phys_h / dpi_y * 25.4
        if abs(actual_w_mm - size_mm) > 2 or abs(actual_h_mm - height_mm) > 2:
            pdc.DeleteDC()
            raise ValueError(
                f"打印机「{printer_name}」当前纸张设置是 {actual_w_mm:.0f}×{actual_h_mm:.0f}mm，"
                f"跟要打印的 {size_mm:.0f}×{height_mm:.0f}mm 不一致，会被裁切/显示不全。"
                f"请先去 Windows「打印机属性→首选项」把纸张/标签尺寸改成 "
                f"{size_mm:.0f}×{height_mm:.0f}mm，再重新打印")

    pdc.StartDoc(PRINT_DOC_NAME)
    pdc.StartPage()
    # 按微调量整体平移之后再画：图片本身尺寸不变，只是在标签纸上挪个位置
    off_x = int(offset_x_mm / 25.4 * dpi_x)
    off_y = int(offset_y_mm / 25.4 * dpi_y)
    dib = ImageWin.Dib(img)
    dib.draw(pdc.GetHandleOutput(), (off_x, off_y, w_px + off_x, h_px + off_y))
    pdc.EndPage()
    pdc.EndDoc()
    pdc.DeleteDC()


def print_job_problem(printer_name):
    """打印任务发出去之后，去Windows打印队列里查一下最近这个任务的实际状态——
    "已发送"只代表任务被Windows接收排队了，不代表打印机真的吐出了标签，
    队列卡住（缺标签/打印机离线或暂停/驱动报错）在Python这边不会有任何异常，
    是"点了显示成功、但打印机没反应"这类问题的常见原因之一。win32ui的StartDoc
    不会可靠返回job id，所以这里改成按文档名（PRINT_DOC_NAME）在队列里找最近一条
    任务来查，查到了具体卡在哪一步就返回中文描述列表；任务已经打完从队列消失、
    或者查不到（权限/驱动不支持）就返回None，不瞎报"""
    import time
    import win32print
    time.sleep(0.5)   # 给Windows spooler一点时间把状态更新出来，太快查可能还是初始状态
    try:
        h = win32print.OpenPrinter(printer_name)
        try:
            jobs = win32print.EnumJobs(h, 0, -1, 1)
        finally:
            win32print.ClosePrinter(h)
    except Exception:
        return None
    matches = [j for j in jobs if j.get("pDocument") == PRINT_DOC_NAME]
    if not matches:
        return None   # 队列里已经没有了，大概率已经处理完了，看不出问题就不报
    status = matches[-1].get("Status", 0)
    flags = {
        win32print.JOB_STATUS_PAUSED: "打印任务被暂停了",
        win32print.JOB_STATUS_ERROR: "打印任务出错",
        win32print.JOB_STATUS_OFFLINE: "打印机显示离线",
        win32print.JOB_STATUS_PAPEROUT: "缺纸/缺标签",
        win32print.JOB_STATUS_BLOCKED_DEVQ: "被驱动队列阻塞",
        win32print.JOB_STATUS_USER_INTERVENTION: "打印机需要人工处理（比如卡纸、缺标签）",
    }
    problems = [text for flag, text in flags.items() if status & flag]
    return problems or None


def page_print_labels():
    st.header("批量打印二维码")
    st.caption("筛选物料、勾选要打印的，生成一份按 A4 不干胶排版的打印页面，"
              "在下面预览区里点「打印」用浏览器自带的打印功能——只会打印标签内容，"
              "不会带上这个系统左边的菜单栏。二维码内容是网址链接（带物料编号），"
              "手机相机/扫码枪直接扫开就能看详情；生产系统那边也能从网址里解析出编号。")

    with st.expander("访问地址设置（改了要重新打印标签）", expanded=False):
        cur_base = get_setting("base_url", "")
        new_base = st.text_input(
            "本机局域网访问地址（含端口，不要带结尾斜杠）",
            value=cur_base, placeholder="http://192.168.1.34:8501")
        if st.button("保存访问地址"):
            set_setting("base_url", new_base.strip())
            st.success("已保存")
            st.rerun()
        if not get_setting("base_url"):
            st.warning("还没设置访问地址，现在生成的二维码扫了会打不开，请先填上面这一项。")

    with st.expander("标签打印机设置（选好了才能用下面的「直接打印」，跳过下载/切电脑那一步）"):
        printers = list_printers()
        if not printers:
            st.caption("这台电脑没找到能用的打印机（本地USB接的，或者别的电脑共享、"
                      "已经在Windows里「添加」过的都算）。没有的话「直接打印」用不了，"
                      "但「生成打印预览」「下载A4图片」不受影响。")
        else:
            cur_printer = get_setting("printer_name", "")
            options = printers if cur_printer in printers else [cur_printer] + printers if cur_printer else printers
            idx = options.index(cur_printer) if cur_printer in options else 0
            pick_printer = st.selectbox("选择标签打印机", options, index=idx)
            if st.button("保存打印机选择"):
                set_setting("printer_name", pick_printer)
                st.success("已保存")
                st.rerun()
        st.caption("提示：打印机驱动属性里的纸张/标签尺寸最好也设成跟下面「标签尺寸」一样大，"
                  "不然打出来可能被裁掉一块或者留一大片空白——这个要在Windows打印机设置里调，"
                  "系统内部调不了。建议先「测试打印1张」看看对不对，别直接打一整批。")

        st.markdown("**打印位置微调**（打出来整体偏一边时用，正数往右/往下，负数往左/往上）")
        cur_off_x, cur_off_y = get_print_offset()
        oc1, oc2 = st.columns(2)
        with oc1:
            off_x = st.number_input("左右微调（mm）", min_value=-5.0, max_value=5.0,
                                    value=float(cur_off_x), step=0.5,
                                    help="内容偏右就填负数，比如偏右1.5mm就填 -1.5")
        with oc2:
            off_y = st.number_input("上下微调（mm）", min_value=-5.0, max_value=5.0,
                                    value=float(cur_off_y), step=0.5,
                                    help="内容偏下就填负数")
        if st.button("保存位置微调"):
            set_setting("print_offset_x_mm", str(off_x))
            set_setting("print_offset_y_mm", str(off_y))
            st.success("已保存，物料标签和位置标签的「直接打印」都会按这个挪")
            st.rerun()

    df_all = inventory_df()
    if df_all.empty:
        st.warning("还没有物料数据。")
        return

    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        cust_labels = ["（无客户）" if not c else c for c in sorted(df_all["客户"].unique())]
        cust_label_map = {("（无客户）" if not c else c): c for c in df_all["客户"].unique()}
        pick_cust_labels = st.multiselect("按客户筛选（默认全部）", cust_labels, default=cust_labels)
        pick_custs = [cust_label_map[l] for l in pick_cust_labels]
    with fc2:
        cat_labels = ["（无类别）" if not c else c for c in sorted(df_all["类别"].unique())]
        cat_label_map = {("（无类别）" if not c else c): c for c in df_all["类别"].unique()}
        pick_cat_labels = st.multiselect("按类别筛选（默认全部）", cat_labels, default=cat_labels)
        pick_cats = [cat_label_map[l] for l in pick_cat_labels]
    with fc3:
        all_pre = sorted({str(l)[0] for l in df_all["位置"] if l and str(l)[0].isalpha()})
        pick_pre = st.multiselect("按位置排筛选（默认全部）", all_pre, default=all_pre)

    rc1, rc2 = st.columns(2)
    with rc1:
        num_from = st.number_input("位置编号从（0=不限制）", min_value=0, value=0, step=1)
    with rc2:
        num_to = st.number_input("位置编号到（0=不限制）", min_value=0, value=0, step=1)

    filtered = df_all.copy()
    filtered = filtered[filtered["客户"].isin(pick_custs)]
    filtered = filtered[filtered["类别"].isin(pick_cats)]
    if pick_pre:
        filtered = filtered[filtered["位置"].fillna("").str[:1].isin(pick_pre)]
    if num_from or num_to:
        def in_range(loc):
            m = re.match(r"^[A-Za-z]+-(\d+)$", str(loc or ""))
            if not m:
                return False
            n = int(m.group(1))
            return (n >= num_from) and (num_to == 0 or n <= num_to)
        filtered = filtered[filtered["位置"].apply(in_range)]

    st.caption(f"筛选出 **{len(filtered)}** 种物料，下面勾选要打印的（默认全选）")
    pick_df = filtered[["id", "位置", "客户", "类别", "料号", "颜色"]].copy()
    pick_df.insert(0, "打印", True)
    edited = st.data_editor(
        pick_df, use_container_width=True, hide_index=True,
        disabled=["id", "位置", "客户", "类别", "料号", "颜色"],
        column_config={"打印": st.column_config.CheckboxColumn("打印")})
    selected = edited[edited["打印"]]

    size_mm = st.number_input("标签尺寸（mm，正方形边长，打印机最小30mm）", min_value=30, max_value=60,
                              value=35, step=1)
    grid_cols, grid_rows = a4_label_grid(int(size_mm))
    st.caption(f"A4纸每页能整齐排下 {grid_cols}×{grid_rows} = {grid_cols * grid_rows} 个标签"
              f"（标签内容：二维码 + 料号 + 位置编号）")

    pc1, pc2 = st.columns(2)
    with pc1:
        gen_html = st.button(f"生成打印预览（{len(selected)} 个标签）", type="primary",
                             use_container_width=True)
    with pc2:
        gen_a4 = st.button(f"下载A4图片（{len(selected)} 个标签）", use_container_width=True)

    cur_printer = get_setting("printer_name", "")
    pc3, pc4 = st.columns(2)
    with pc3:
        test_print = st.button("测试打印1张", use_container_width=True,
                               disabled=not cur_printer or len(selected) == 0)
    with pc4:
        direct_print = st.button(f"直接打印全部（{len(selected)} 个标签）",
                                 use_container_width=True,
                                 disabled=not cur_printer or len(selected) == 0)
    if not cur_printer:
        st.caption("没选打印机，「测试打印」「直接打印」用不了——去上面「标签打印机设置」选一个")

    if (gen_html or gen_a4) and len(selected) == 0:
        st.warning("先勾选至少一个物料")
    elif gen_html:
        items = [(int(r["id"]), r["料号"], r["客户"], r["位置"]) for _, r in selected.iterrows()]
        html = build_print_html(items, int(size_mm))
        components.html(html, height=800, scrolling=True)
    elif gen_a4:
        items = [(int(r["id"]), r["料号"], r["客户"], r["位置"]) for _, r in selected.iterrows()]
        sheets = build_a4_sheets(items, int(size_mm))
        st.image(sheets[0], caption=f"预览第1页（共{len(sheets)}页）", use_container_width=True)
        if len(sheets) == 1:
            st.download_button(
                "下载A4图片（PNG）", image_to_png_bytes(sheets[0]),
                file_name="二维码标签_A4.png", mime="image/png", type="primary")
        else:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w") as zf:
                for i, sheet in enumerate(sheets, 1):
                    zf.writestr(f"二维码标签_第{i}页.png", image_to_png_bytes(sheet))
            st.download_button(
                f"下载A4图片（共{len(sheets)}页，打包zip）", buf.getvalue(),
                file_name="二维码标签_A4.zip", mime="application/zip", type="primary")
    elif test_print:
        r = selected.iloc[0]
        img = make_qr_label(int(r["id"]), r["料号"], r["客户"], r["位置"], size_mm=int(size_mm))
        try:
            print_label_direct(img, cur_printer, size_mm=int(size_mm))
            problems = print_job_problem(cur_printer)
            if problems:
                st.warning(f"发到「{cur_printer}」了，但Windows打印队列显示：{'、'.join(problems)}"
                          f"——去电脑「设备和打印机」看这台打印机的队列，处理掉再试")
            else:
                st.success(f"已发送1张到「{cur_printer}」，看看打出来尺寸、位置对不对")
        except Exception as e:
            st.error(f"打印失败：{e}")
    elif direct_print:
        items = [(int(r["id"]), r["料号"], r["客户"], r["位置"]) for _, r in selected.iterrows()]
        ok_count, fail_count = 0, 0
        last_problems = None
        progress = st.progress(0, text="打印中…")
        for i, (mid, code, customer, location) in enumerate(items, 1):
            img = make_qr_label(mid, code, customer, location, size_mm=int(size_mm))
            try:
                print_label_direct(img, cur_printer, size_mm=int(size_mm))
                ok_count += 1
                if i == 1:
                    last_problems = print_job_problem(cur_printer)
                    if last_problems:
                        st.warning(f"Windows打印队列显示：{'、'.join(last_problems)}"
                                  f"——先去电脑「设备和打印机」处理掉，不然后面几张也打不出来")
                        break
            except Exception as e:
                fail_count += 1
                st.error(f"{code} 打印失败：{e}")
                break
            progress.progress(i / len(items), text=f"打印中…{i}/{len(items)}")
        progress.empty()
        if last_problems:
            pass   # 上面已经提示过队列卡住了，这里不用再重复一条容易误导的"已发送"
        elif fail_count == 0:
            st.success(f"已发送 {ok_count} 张到「{cur_printer}」")
        else:
            st.warning(f"发送了 {ok_count} 张，中途失败停止")

    st.markdown("---")
    with st.expander("生成位置二维码（贴在盒子/货架上，跟物料无关，贴一次长期有效）", expanded=False):
        st.caption("这个码扫出来不是网址，是固定的位置编号，配合「扫码上架」页用："
                  "拍这个码就能把选好的物料放进这个位置。位置不变（盒子没挪窝）就不用重贴。")

        st.caption("批量生成：填前缀+起止编号，点「生成」自动加到下面的清单里（可以点好几次拼不同排）")
        rgc1, rgc2, rgc3, rgc4 = st.columns([2, 1, 1, 1.2])
        with rgc1:
            range_pre = st.text_input("前缀", placeholder="如 B", key="loc_range_pre")
        with rgc2:
            range_from = st.number_input("从", min_value=1, value=1, step=1, key="loc_range_from")
        with rgc3:
            range_to = st.number_input("到", min_value=1, value=10, step=1, key="loc_range_to")
        with rgc4:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            range_gen = st.button("生成到清单", use_container_width=True, key="loc_range_gen")
        if range_gen:
            if not range_pre.strip():
                st.warning("先填前缀，比如 B")
            elif range_to < range_from:
                st.warning("「到」不能比「从」小")
            else:
                new_codes = [f"{range_pre.strip().upper()}-{n}"
                            for n in range(int(range_from), int(range_to) + 1)]
                existing = st.session_state.get("loc_label_text", "")
                st.session_state["loc_label_text"] = (
                    existing.rstrip() + "\n" + "\n".join(new_codes)).strip()
                st.rerun()

        loc_text = st.text_area(
            "位置编号清单（每行一个，也可以直接在这手动加/删/改）", height=100, key="loc_label_text",
            placeholder="B-14\nB-15\nB-16")
        locs = [l.strip() for l in loc_text.splitlines() if l.strip()]

        lsz1, lsz2, lsz3 = st.columns([1, 1, 1])
        with lsz1:
            loc_width_mm = st.number_input(
                "标签宽度（mm）", min_value=15, max_value=200, value=35, step=1, key="loc_width_mm")
        with lsz2:
            loc_height_mm = st.number_input(
                "标签高度（mm）", min_value=15, max_value=200, value=35, step=1, key="loc_height_mm")
        with lsz3:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            loc_show_qr = st.checkbox("包含二维码", value=True, key="loc_show_qr")
        if not loc_show_qr:
            st.caption("⚠ 不带二维码的标签只能靠人眼看、手动输入位置编号——"
                      "「扫码上架」页的扫码枪没法识别，适合纯粹给人看的大牌子。")

        lgc1, lgc2 = st.columns(2)
        with lgc1:
            loc_gen = st.button(f"生成位置标签图（{len(locs)} 个）",
                                use_container_width=True, disabled=not locs, key="loc_gen")
        with lgc2:
            loc_printer = get_setting("printer_name", "")
            loc_direct_print = st.button(
                f"直接打印这些位置标签（{len(locs)} 个）", use_container_width=True,
                disabled=not locs or not loc_printer, key="loc_direct_print")

        if loc_gen:
            sheet = build_location_label_sheet(
                locs, width_mm=int(loc_width_mm), height_mm=int(loc_height_mm), show_qr=loc_show_qr)
            st.image(sheet, use_container_width=True)
            st.download_button(
                "下载位置标签图（PNG）", image_to_png_bytes(sheet),
                file_name="位置标签.png", mime="image/png",
                type="primary", key="loc_dl")
        elif loc_direct_print:
            ok_count, fail_count = 0, 0
            last_problems = None
            for i, loc in enumerate(locs, 1):
                img = make_location_label(
                    loc, width_mm=int(loc_width_mm), height_mm=int(loc_height_mm), show_qr=loc_show_qr)
                try:
                    print_label_direct(img, loc_printer,
                                       size_mm=int(loc_width_mm), height_mm=int(loc_height_mm))
                    ok_count += 1
                    if i == 1:
                        last_problems = print_job_problem(loc_printer)
                        if last_problems:
                            st.warning(f"Windows打印队列显示：{'、'.join(last_problems)}"
                                      f"——先去电脑「设备和打印机」处理掉，不然后面也打不出来")
                            break
                except Exception as e:
                    fail_count += 1
                    st.error(f"{loc} 打印失败：{e}")
                    break
            if last_problems:
                pass
            elif fail_count == 0:
                st.success(f"已发送 {ok_count} 张到「{loc_printer}」")
            else:
                st.warning(f"发送了 {ok_count} 张，中途失败停止")


# ---------------- 页面：扫码上架 ----------------

MAT_QR_RE = re.compile(r"[?&]id=(\d+)")        # 物料标签二维码里的网址：.../?id=123
SUPPLY_LOC_PREFIXES = {"探针", "治具"}          # 配件库的位置前缀，supplies 表里现有的位置还会自动补充
MAT_BRIEF_COLS = ["id", "code", "customer", "location", "status", "color"]
CUST_NONE, CUST_NEW = "（不填客户）", "（新客户）"
CAT_NONE, CAT_NEW = "（不填类别）", "（新类别）"


def _sk(name):
    """扫码上架输入框的 key 带一个编号，每存好一个料编号 +1：换了 key 就是全新的空输入框。
    （只删 session_state 里的值不行——浏览器那边还记着旧内容，下次一点别的又会被当成新输入提交回来）"""
    return f"{name}_{st.session_state.get('shelve_nonce', 0)}"


def normalize_location(raw):
    """位置统一写法，只改写法、不改编号：去掉位置码的 LOC: 前缀，全角转半角（Ａ－９９→A-99），
    各种横杠统一成 -、去掉横杠两边的空格，「字母-数字」格式的字母转大写（a-97→A-97）"""
    s = unicodedata.normalize("NFKC", raw or "").strip()
    if s.upper().startswith("LOC:"):
        s = s[4:].strip()
    s = re.sub(r"\s*[-‐‑‒–—―−]\s*", "-", s)
    m = LOC_RE.match(s)
    return f"{m.group(1).upper()}-{m.group(2)}" if m else s


def is_supply_location(loc):
    """探针1-1、治具1-1 这类位置属于配件库（辅料/配件），不是物料的格子"""
    conn = get_conn()
    locs = [r[0] for r in conn.execute("SELECT DISTINCT location FROM supplies WHERE location != ''")]
    conn.close()
    prefixes = set(SUPPLY_LOC_PREFIXES)
    for l in locs:
        m = re.match(r"^[^\x00-\x7f]+", l.strip())  # 开头的中文部分，如「探针」
        if m:
            prefixes.add(m.group(0))
    return any(loc.startswith(p) for p in prefixes)


def material_brief(mid):
    conn = get_conn()
    r = conn.execute("SELECT id, code, customer, location, status, color FROM materials WHERE id=?",
                     (mid,)).fetchone()
    conn.close()
    return dict(zip(MAT_BRIEF_COLS, r)) if r else None


def find_shelve_material(code, customer):
    """按 料号+客户 找库里已有的物料：不分大小写，别名料号也算。
    返回 (匹配到的物料 或 None, 同料号但客户不一样的物料列表)"""
    conn = get_conn()
    rows = conn.execute(
        "SELECT DISTINCT m.id, m.code, m.customer, m.location, m.status, m.color"
        " FROM materials m LEFT JOIN material_aliases a ON a.material_id = m.id"
        " WHERE m.code = ? COLLATE NOCASE OR a.code = ? COLLATE NOCASE",
        (code, code)).fetchall()
    conn.close()
    mats = [dict(zip(MAT_BRIEF_COLS, r)) for r in rows]
    # 同客户里：一模一样的写法优先，其次只是大小写不同，最后是别名
    same = sorted((m for m in mats if m["customer"] == customer),
                  key=lambda m: (m["code"] != code, m["code"].lower() != code.lower()))
    return (same[0] if same else None), [m for m in mats if m["customer"] != customer]


def location_occupants(loc, exclude_mid=None):
    """这个位置上现在登记着的在用物料（不含正在上架的这个料），带库存"""
    df = location_view_df()
    hit = df[df["位置"].map(normalize_location) == loc]
    if exclude_mid is not None:
        hit = hit[hit["id"] != exclude_mid]
    return hit


def focus_input(label_prefix, only_from=None):
    """把光标放进某个输入框（按标签开头找）。扫码枪其实就是个键盘，光标在哪就往哪打字，
    所以扫完物料码自动跳到「位置」栏、保存完自动回到「料号」栏，就能一路扫下去不用点屏幕。
    脚本里带个递增的编号，内容每次不一样，Streamlit 才会重新执行它。
    only_from：光标还停在这个框里（或哪个框都不在）才挪；人已经自己点进别的框了就别拽走"""
    n = st.session_state["_focus_n"] = st.session_state.get("_focus_n", 0) + 1
    components.html(f"""<script>
        // {n}
        const want = {json.dumps(label_prefix)};
        const from = {json.dumps(only_from)};
        const a = window.parent.document.activeElement;
        if (from !== null && a && a.tagName === 'INPUT'
            && !(a.getAttribute('aria-label') || '').startsWith(from)) {{ throw 'skip'; }}
        let tries = 0;
        const t = setInterval(() => {{
            const el = [...window.parent.document.querySelectorAll('input')]
                .find(i => (i.getAttribute('aria-label') || '').startsWith(want));
            if (el || ++tries > 20) {{ clearInterval(t); if (el) {{ el.focus(); el.select(); }} }}
        }}, 100);
    </script>""", height=0)


def _shelve_changed():
    """输入变了，之前对提示问题的回答就不算数了"""
    st.session_state.pop("shelve_decide", None)


def _shelve_code_changed():
    _shelve_changed()
    ss = st.session_state
    ss["shelve_code_changed"] = True
    # 先扫了位置、后扫物料码：两样都齐了，直接保存（顺序反了也行）
    if MAT_QR_RE.search(ss.get(_sk("shelve_code"), "")) and ss.get(_sk("shelve_location"), "").strip():
        ss["shelve_go"] = True


def _shelve_go():
    st.session_state["shelve_go"] = True


def _shelve_loc_entered():
    # 扫码枪扫完位置码会自带一个回车 → 直接保存
    _shelve_changed()
    _shelve_go()


def _shelve_decide(key, value):
    d = dict(st.session_state.get("shelve_decide") or {})
    d[key] = value
    st.session_state["shelve_decide"] = d
    st.session_state["shelve_go"] = True


def show_shelve_result(last):
    """上一个保存结果 + 标签。存在 session_state 里，点「直接打印」刷新页面后还在，打印才点得动"""
    mid, code, customer, loc = last["mid"], last["code"], last["customer"], last["loc"]
    with st.container(border=True):
        st.success(last["msg"])
        for n in last["notes"]:
            st.caption(n)
        label_img = make_qr_label(mid, code, customer, loc)
        lc1, lc2 = st.columns([1, 2])
        with lc1:
            st.image(label_img, width=180)
        with lc2:
            st.download_button(
                "下载标签", image_to_png_bytes(label_img),
                file_name=f"{code}_{customer}_二维码.png", mime="image/png",
                key=f"shelve_label_dl_{mid}_{loc}")
            printer_name = get_setting("printer_name", "")
            if printer_name:
                if st.button(f"直接打印到「{printer_name}」", type="primary",
                             key=f"shelve_print_{mid}_{loc}"):
                    try:
                        print_label_direct(label_img, printer_name)
                        problems = print_job_problem(printer_name)
                        if problems:
                            st.warning(f"发到「{printer_name}」了，但Windows打印队列显示："
                                       f"{'、'.join(problems)}——去电脑「设备和打印机」看这台"
                                       f"打印机的队列，处理掉再试")
                        else:
                            st.success("已发送到打印机")
                    except Exception as e:
                        st.error(f"打印失败：{e}")
            else:
                st.caption("没设置打印机——去「批量打印二维码」页的「标签打印机设置」"
                           "选一个，就能一键直接打印")


def shelve_submit(f):
    """保存前逐项检查；需要你拍板的（客户对不上 / 格子被占）先问，点了回答再存"""
    ss = st.session_state
    code_in, loc = f["code_in"], f["loc"]
    if not code_in and not f["loc_in"].strip():
        return  # 什么都没填（比如刚保存完页面清空时的一次多余点击），不提示
    if not code_in:
        st.error("料号必填：扫物料标签的二维码，或者直接输入料号")
        return
    if ("://" in code_in or MAT_QR_RE.search(code_in)) and not f["scanned"]:
        st.error("「料号」栏里是个网址，但系统里找不到对应的物料，请直接输入料号")
        return
    if f["cust_new_missing"]:
        st.error("选了「新客户」，请填上客户名称")
        return
    if f["cat_new_missing"]:
        st.error("选了「新类别」，请填上类别名称")
        return
    if not loc:
        st.error("位置必填：点进「位置」栏扫盒子上的位置码，或手动输入")
        return
    if "://" in loc or MAT_QR_RE.search(loc):
        st.error("「位置」栏扫到的是物料标签的二维码，不是位置码。请清空这一栏，扫盒子上的位置码。")
        return
    if is_supply_location(loc):
        st.error(f"「{loc}」是探针/治具这些配件的位置。扫码上架只管物料，配件请到「辅料/配件管理」录入。")
        return

    decide = ss.get("shelve_decide") or {}

    # 1. 定下是哪个物料：扫到的 / 同客户已有的 / 同料号但客户不一样的要问一句
    mat = f["scanned"] or f["match"]
    if not mat and f["others"]:
        ans = decide.get("customer")
        if ans is None:
            who = f["customer"] or "不填客户"
            st.warning(f"库里已经有同料号的物料，但客户不是你选的「{who}」。是同一个料吗？")
            for m in f["others"]:
                where = f"，现在在 {m['location']}" if m["location"] else ""
                st.button(f"是同一个：用 {m['code']}（{m['customer'] or '无客户'}）这条{where}",
                          key=f"shelve_use_{m['id']}", on_click=_shelve_decide,
                          args=("customer", m["id"]))
            st.button(f"不是，新建一条（客户：{who}）", key="shelve_use_new",
                      on_click=_shelve_decide, args=("customer", "new"))
            return
        if ans != "new":
            mat = next((m for m in f["others"] if m["id"] == ans), None)

    # 2. 格子被别的料占着：问是释放旧的，还是一起放（有的格子本来就放好几种料）
    occ = location_occupants(loc, exclude_mid=mat["id"] if mat else None)
    release = False
    if not occ.empty:
        ans = decide.get("occupied")
        if ans is None:
            st.warning(f"格子 **{loc}** 里已经登记了下面这些料：")
            st.dataframe(occ[["料号", "客户", "颜色", "库存", "单位"]], hide_index=True)
            st.button("把它们释放掉，放这个料进来", key="shelve_occ_release",
                      on_click=_shelve_decide, args=("occupied", "release"))
            st.caption("释放 = 旧料标成停用、位置清空，库存数不动；东西已经没了的话，再去「库位视图」把它清零。")
            st.button("一起放（这个格子本来就放好几种料）", key="shelve_occ_together",
                      on_click=_shelve_decide, args=("occupied", "together"))
            return
        release = ans == "release"

    # 3. 保存：释放旧料 + 新建/更新，放在一个事务里，要么全成功要么全不写
    conn = get_conn()
    released = []
    try:
        conn.execute("BEGIN IMMEDIATE")
        if release:
            for _, r in occ.iterrows():
                conn.execute("UPDATE materials SET status='停用', location='' WHERE id=?",
                             (int(r["id"]),))
                released.append(f"{r['料号']}（{r['客户'] or '无客户'}）")
        if mat:
            conn.execute("UPDATE materials SET location=?, status='在用' WHERE id=?",
                         (loc, mat["id"]))
            mid, code, customer, is_new = mat["id"], mat["code"], mat["customer"], False
        else:
            cur = conn.execute(
                "INSERT INTO materials (code,name,customer,category,color,spec,unit,location,"
                "safety_stock,note,created_at) VALUES (?,?,?,?,?, '', ?,?,?,?,?)",
                (code_in, code_in, f["customer"], f["category"], f["color"], f["unit"] or "个",
                 loc, f["safety"], f["note"], datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            mid, code, customer, is_new = cur.lastrowid, code_in, f["customer"], True
        conn.commit()
    except Exception as e:
        conn.rollback()
        st.error(f"保存失败：{e}")
        return
    finally:
        conn.close()

    if is_new and f["init_qty"] > 0:
        add_record("期初", mid, f["init_qty"], "系统", "建账期初")
    # 照片一律追加，不覆盖——已有物料原来的照片不能因为补拍一张就没了
    if f["photo"]:
        append_uploads(f["photo"], file_key(mid, code), "产品照片")
    if f["drawing"]:
        append_uploads(f["drawing"], file_key(mid, code), "图纸")

    cust_s = f"（{customer}）" if customer else ""
    if is_new:
        msg = f"已新增 **{code}**{cust_s}，位置 **{loc}**"
    else:
        old = mat["location"]
        moved = f"{old} → **{loc}**" if old and normalize_location(old) != loc else f"位置 **{loc}**"
        msg = f"已更新 **{code}**{cust_s}：{moved}"
        if mat["status"] == "停用":
            msg += "（已恢复在用）"
    notes = []
    if released:
        notes.append("已释放：" + "、".join(released))
    if f["photo"] or f["drawing"]:
        notes.append(f"照片追加 {len(f['photo'] or [])} 张、图纸追加 {len(f['drawing'] or [])} 张")
    ss["shelve_last"] = dict(mid=mid, code=code, customer=customer, loc=loc, msg=msg, notes=notes)
    ss["shelve_reset"] = True
    st.rerun()


def page_scan_shelve():
    st.header("扫码上架")
    st.caption("① 在「料号」栏扫物料标签上的二维码（库里已有的料），或者直接输入料号（新料）；"
               "② 点进「位置」栏（扫物料码后光标会自动跳过去），用扫码枪扫盒子上的位置码，扫完自动保存。"
               "库里已有的料只更新位置，照片是追加、不会覆盖原来的。"
               "探针、治具这类配件不在这里上架，请到「辅料/配件管理」录入。"
               "位置码要先在「批量打印二维码」页生成、打印贴到盒子上。")

    ss = st.session_state
    if ss.pop("shelve_reset", False):
        # 上一个存好了：换一套新的空输入框准备扫下一个（见 _sk），光标放回「料号」栏
        ss.pop("shelve_decide", None)
        ss["shelve_nonce"] = ss.get("shelve_nonce", 0) + 1
        focus_input("料号")

    if ss.get("shelve_last"):
        show_shelve_result(ss["shelve_last"])
        st.markdown("**下一个**")

    conn = get_conn()
    all_custs = [r[0] for r in conn.execute(
        "SELECT DISTINCT customer FROM materials WHERE customer != '' ORDER BY customer")]
    all_cats = [r[0] for r in conn.execute(
        "SELECT DISTINCT category FROM materials WHERE category != '' ORDER BY category")]
    conn.close()

    code_in = st.text_input("料号 *（扫物料标签的二维码，或直接输入）", key=_sk("shelve_code"),
                            placeholder="如 KLP201062", on_change=_shelve_code_changed).strip()

    # ---- 这是哪个料 ----
    scanned, match, others = None, None, []
    customer, cust_new_missing = "", False
    qr = MAT_QR_RE.search(code_in)
    if qr or "://" in code_in:
        scanned = material_brief(int(qr.group(1))) if qr else None
        if scanned:
            where = f"现在在 **{scanned['location']}**" if scanned["location"] else "现在没有位置"
            stop = "（已停用，上架后恢复在用）" if scanned["status"] == "停用" else ""
            st.success(f"已识别：**{scanned['code']}**（{scanned['customer'] or '无客户'}），{where}{stop}。"
                       "这次只更新位置。")
            if ss.pop("shelve_code_changed", False):
                focus_input("位置")
        else:
            st.error("扫到的是网址，但系统里找不到对应的物料（可能是别的系统打的标签）。请直接输入料号。")
    else:
        ss.pop("shelve_code_changed", None)
        cust_sel = st.selectbox("客户", [CUST_NONE] + all_custs + [CUST_NEW],
                                key=_sk("shelve_cust_sel"), on_change=_shelve_changed)
        if cust_sel == CUST_NEW:
            customer = st.text_input("新客户名称 *", key=_sk("shelve_cust_new"),
                                     on_change=_shelve_changed).strip()
            cust_new_missing = not customer
            if customer and customer not in all_custs:
                close = difflib.get_close_matches(customer, all_custs, n=1, cutoff=0.5)
                if close:
                    st.warning(f"库里已有客户「{close[0]}」，是同一个的话请在上面直接选它——"
                               "写法不一样会被当成两个客户。")
        elif cust_sel != CUST_NONE:
            customer = cust_sel

        if code_in and not cust_new_missing:
            match, others = find_shelve_material(code_in, customer)
            if match:
                how = ""
                if match["code"] != code_in:
                    how = ("（按库里的写法）" if match["code"].lower() == code_in.lower()
                           else f"（「{code_in}」是它的别名料号）")
                where = f"现在在 **{match['location']}**" if match["location"] else "现在没有位置"
                stop = "，已停用，上架后恢复在用" if match["status"] == "停用" else ""
                st.info(f"库里已有这个料{how}：**{match['code']}**（{match['customer'] or '无客户'}），"
                        f"{where}{stop}。这次只更新位置；颜色、数量这些不会改，要改请到「物料管理」。")
            elif others:
                lst = "、".join(f"**{m['code']}**（{m['customer'] or '无客户'}）" for m in others)
                st.warning(f"库里有 {lst}，但客户跟你选的不一样。是同一个料的话，把上面的客户改过来；"
                           "确实是另一个客户的新料就继续，保存时会再问一次。")
            else:
                st.caption("库里没有这个料，保存时会新建。")

    # ---- 新料才需要填的信息 ----
    new_mode = bool(code_in) and not scanned and not match and not qr and "://" not in code_in
    category, cat_new_missing = "", False
    color, unit, init_qty, safety, note = "", "个", 0.0, 0.0, ""
    if new_mode:
        cat_sel = st.selectbox("类别（可选）", [CAT_NONE] + all_cats + [CAT_NEW], key=_sk("shelve_cat_sel"))
        if cat_sel == CAT_NEW:
            category = st.text_input("新类别名称 *", key=_sk("shelve_cat_new")).strip()
            cat_new_missing = not category
        elif cat_sel != CAT_NONE:
            category = cat_sel
        with st.expander("更多信息（可选）", expanded=False):
            color = st.text_input("颜色", key=_sk("shelve_color")).strip()
            unit = st.text_input("单位", value="个", key=_sk("shelve_unit")).strip()
            init_qty = st.number_input("库存数量", min_value=0.0, step=1.0, format="%g",
                                       key=_sk("shelve_qty"))
            safety = st.number_input("安全库存", min_value=0.0, step=1.0, format="%g",
                                     key=_sk("shelve_safety"))
            note = st.text_input("备注", key=_sk("shelve_note")).strip()

    has_old = bool(scanned or match)
    st.caption("手机上点「选择文件」可直接调相机拍照（可选，不传也能保存）"
               + ("；会追加到这个料原有的照片后面，不会覆盖" if has_old else ""))
    photo = st.file_uploader("产品照片", type=["jpg", "jpeg", "png", "webp", "pdf"],
                             accept_multiple_files=True, key=_sk("shelve_photo"))
    drawing = st.file_uploader("图纸照片", type=["jpg", "jpeg", "png", "webp", "pdf"],
                               accept_multiple_files=True, key=_sk("shelve_drawing"))

    loc_in = st.text_input("位置 *（点进这栏用扫码枪扫盒子上的位置码，扫完自动保存；"
                           "手动输入如 B-14 后按回车）",
                           key=_sk("shelve_location"), on_change=_shelve_loc_entered)
    loc = normalize_location(loc_in)
    raw = loc_in.strip()
    if raw and loc and loc != raw and not raw.upper().startswith("LOC:"):
        st.caption(f"会按统一写法存成 **{loc}**")
    st.button("保存并打印标签", type="primary", width="stretch", on_click=_shelve_go)

    if ss.pop("shelve_go", False):
        shelve_submit(dict(
            code_in=code_in, scanned=scanned, match=match, others=others,
            customer=customer, cust_new_missing=cust_new_missing,
            category=category, cat_new_missing=cat_new_missing,
            color=color, unit=unit, init_qty=init_qty, safety=safety, note=note,
            photo=photo, drawing=drawing, loc_in=loc_in, loc=loc))


# ---------------- 页面：批量扫码出入库 ----------------
# 一次处理好几个料：每扫一个加一行、填数量，最后整单一起确认，放在一个事务里写流水，
# 要么全成功要么一条都不写。扫完光标自动跳到这一行的数量框，填完按回车跳回扫码框。

STAFF_NEW = "（新增人员）"


def _bs_sk(name):
    """扫码框的 key 带编号，每扫一个 +1，换成全新的空框（道理同扫码上架的 _sk）"""
    return f"{name}_{st.session_state.get('bs_nonce', 0)}"


def _bs_qty(raw):
    """数量框里的字 → 数字；空的、不是数、不大于 0 的都返回 None。全角数字也认"""
    try:
        q = float(unicodedata.normalize("NFKC", raw or "").strip())
    except ValueError:
        return None
    return q if 0 < q < 1e9 else None


def batch_scan_lookup(raw):
    """扫到/输入的内容 → (候选物料列表, 出错说明)。认三种：物料标签二维码（网址带 ?id=）、
    位置码（LOC:A-58，或直接输 A-58）、料号/别名（不分大小写）"""
    qr = MAT_QR_RE.search(raw)
    if qr:
        m = material_brief(int(qr.group(1)))
        return ([m], "") if m else ([], "扫到的是网址，但系统里找不到对应的物料（可能是别的系统打的标签）")
    if "://" in raw:
        return [], "扫到的网址不是本系统的物料标签"
    is_loc_code = raw.upper().startswith("LOC:")
    if not is_loc_code:
        conn = get_conn()
        rows = conn.execute(
            "SELECT DISTINCT m.id, m.code, m.customer, m.location, m.status, m.color"
            " FROM materials m LEFT JOIN material_aliases a ON a.material_id = m.id"
            " WHERE m.code = ? COLLATE NOCASE OR a.code = ? COLLATE NOCASE ORDER BY m.id",
            (raw, raw)).fetchall()
        conn.close()
        if rows:
            return [dict(zip(MAT_BRIEF_COLS, r)) for r in rows], ""
    loc = normalize_location(raw)
    if is_supply_location(loc):
        return [], f"「{loc}」是探针/治具这些配件的位置，批量扫码目前只管物料"
    if is_loc_code or LOC_RE.match(loc):
        occ = location_occupants(loc)
        if occ.empty:
            return [], f"位置 {loc} 上没有登记在用的物料"
        return [material_brief(int(i)) for i in occ["id"]], ""
    return [], f"找不到「{raw}」：不是物料标签也不是位置码，库里也没有这个料号/别名"


def _bs_add(mat, focus=True):
    """把一个料加进单子；已经在单子里的不重复加，光标跳到它的数量框。返回这一行的编号"""
    ss = st.session_state
    rows = ss.setdefault("bs_rows", [])
    ss.pop("bs_pick", None)
    old = next((r for r in rows if r["id"] == mat["id"]), None)
    if old:
        rid = old["rid"]
        ss["bs_msg"] = ("info", f"**{mat['code']}** 已经在单子里了（第 {rows.index(old) + 1} 行），"
                                "要改数量直接改")
    else:
        rid = ss["bs_serial"] = ss.get("bs_serial", 0) + 1
        rows.append(dict(mat, rid=rid))
        where = f"{mat['location']} " if mat["location"] else ""
        ss["bs_msg"] = ("success", f"已加入：{where}**{mat['code']}**（{mat['customer'] or '无客户'}），"
                                   "填上数量按回车")
    if focus:
        ss["bs_focus"] = (f"数量#{rid}#", None)
    return rid


def _bs_add_all(mats):
    rids = [_bs_add(m, focus=False) for m in mats]
    st.session_state["bs_msg"] = ("success", f"已加入 {len(mats)} 种，挨个填上数量")
    st.session_state["bs_focus"] = (f"数量#{rids[0]}#", None)


def _bs_pick_cancel():
    st.session_state.pop("bs_pick", None)
    st.session_state["bs_focus"] = ("扫码", None)


def _bs_handle(raw):
    """处理一次扫码/输入：唯一对上就加进单子，对上好几个（一格放了几种料、同料号几个客户）让人挑"""
    ss = st.session_state
    ss.pop("bs_pick", None)
    mats, err = batch_scan_lookup(raw)
    if err:
        ss["bs_msg"] = ("error", err)
        ss["bs_focus"] = ("扫码", None)
    elif len(mats) == 1:
        _bs_add(mats[0])
    else:
        ss["bs_pick"] = dict(raw=raw, mats=mats)


def _bs_scan_entered():
    ss = st.session_state
    raw = (ss.get(_bs_sk("bs_scan")) or "").strip()
    ss["bs_nonce"] = ss.get("bs_nonce", 0) + 1       # 下一轮换个新的空扫码框
    if raw:
        _bs_handle(raw)


def _bs_qty_changed(rid):
    """数量填完：光标跳回扫码框。要是数量框里冒出网址/位置码——忘了按回车就扫了下一个——
    把扫进来的那截拿去当扫码处理，数量框只留前面的数字"""
    ss = st.session_state
    key, ok_key = f"bs_qty_{rid}", f"bs_qty_ok_{rid}"
    val = (ss.get(key) or "").strip()
    hit = re.search(r"https?://|LOC:", val, re.I)
    if hit or MAT_QR_RE.search(val):
        start = hit.start() if hit else 0
        head = val[:start].strip()
        keep = head if _bs_qty(head) else ss.get(ok_key, "")
        ss[key] = ss[ok_key] = keep
        _bs_handle(val[start:])
        return
    ss[ok_key] = val
    ss["bs_focus"] = ("扫码", f"数量#{rid}#")


def _bs_remove(rid):
    ss = st.session_state
    ss["bs_rows"] = [r for r in ss.get("bs_rows", []) if r["rid"] != rid]
    ss["bs_focus"] = ("扫码", None)


def _bs_clear():
    ss = st.session_state
    ss["bs_rows"] = []
    ss.pop("bs_pick", None)
    ss["bs_focus"] = ("扫码", None)


def _bs_check(rtype):
    """逐行检查数量/库存，返回 (能写的 [(物料id, 数量, 料号, 单位)], 问题列表)"""
    ss = st.session_state
    items, problems = [], []
    conn = get_conn()
    for i, r in enumerate(ss.get("bs_rows", []), 1):
        stock = get_stock(conn, r["id"])
        unit = conn.execute("SELECT unit FROM materials WHERE id=?", (r["id"],)).fetchone()[0] or "个"
        raw = (ss.get(f"bs_qty_{r['rid']}") or "").strip()
        qty = _bs_qty(raw)
        name = f"第 {i} 行 {r['code']}"
        if not raw:
            problems.append(f"{name} 还没填数量")
        elif qty is None:
            problems.append(f"{name} 的数量「{raw}」不对，要填大于 0 的数字")
        elif rtype == "出库" and qty > stock:
            problems.append(f"{name} 库存只有 {stock:g}{unit}，不够出 {qty:g}")
        else:
            items.append((r["id"], qty, r["code"], unit))
    conn.close()
    return items, problems


def add_records_batch(rtype, items, operator, note, sub_type="", created_at=None):
    """一次写多条流水，放在一个事务里：出库时逐条再核一遍库存，有一条不够就全部不写。
    items = [(物料id, 数量, 料号), ...]，返回 (是否成功, 消息)"""
    ts = created_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_conn()
    try:
        conn.execute("BEGIN IMMEDIATE")
        for mid, qty, code in items:
            if rtype == "出库":
                current = get_stock(conn, mid)
                if qty > current:
                    conn.rollback()
                    return False, (f"{code} 库存不足：当前只剩 {current:g}，不能出库 {qty:g}。"
                                   "整单都没写，改好数量再确认")
            conn.execute(
                "INSERT INTO records (type, material_id, quantity, operator, note,"
                " sub_type, link_id, expect_return, created_at)"
                " VALUES (?,?,?,?,?,?,?,?,?)",
                (rtype, mid, qty, operator, note, sub_type, None, "", ts))
        conn.commit()
        return True, "ok"
    except Exception as e:
        conn.rollback()
        return False, f"写入失败：{e}（整单都没写）"
    finally:
        conn.close()


def _bs_confirm():
    ss = st.session_state
    rtype = ss.get("bs_rtype", "出库")
    if not ss.get("bs_rows"):
        return
    items, problems = _bs_check(rtype)
    if rtype == "出库":
        pick = ss.get("bs_op_pick") or ""
        operator = (ss.get("bs_op_new") or "").strip() if pick == STAFF_NEW else pick
    else:
        operator = (ss.get("bs_op_in") or "").strip()
    if not operator:
        problems.append("请填写领用人" if rtype == "出库" else "请填写经手人")
    if problems:
        ss["bs_errors"] = problems
        return
    created_at = None
    if rtype == "出库":
        d = ss.get("bs_date") or date.today()
        created_at = datetime.combine(d, datetime.now().time()).strftime("%Y-%m-%d %H:%M:%S")
    ok, msg = add_records_batch(
        rtype, [(mid, qty, code) for mid, qty, code, _unit in items], operator,
        (ss.get("bs_note") or "").strip(),
        sub_type=ss.get("bs_sub", "领用") if rtype == "出库" else "", created_at=created_at)
    if not ok:
        ss["bs_errors"] = [msg]
        return
    if rtype == "出库":
        add_staff(operator)
        ss["bs_op_pick"] = operator
    ss["bs_done"] = (f"已{rtype} {len(items)} 种（{operator}）：" +
                     "、".join(f"{code} ×{qty:g}{unit}" for _mid, qty, code, unit in items))
    ss["bs_rows"] = []
    ss["bs_note"] = ""
    ss["bs_focus"] = ("扫码", None)


def page_batch_scan():
    st.header("批量扫码出入库")
    st.caption("一次处理好几个料：扫一个加一行，填上数量，最后整单一起确认。"
               "认物料标签的二维码、盒子上的位置码，也可以直接输入料号/别名后按回车。"
               "扫完光标会自动跳到这一行的数量框，填完按回车又跳回扫码框，接着扫下一个。"
               "整单要么全部写进去，要么一条都不写（比如有一个料库存不够）。")
    ss = st.session_state
    rows = ss.setdefault("bs_rows", [])
    done = ss.pop("bs_done", None)
    if done:
        st.success(done)

    rtype = st.radio("这一单是", ["出库", "入库"], horizontal=True, key="bs_rtype")
    st.text_input("扫码（扫物料标签或位置码，也可以输入料号/别名后按回车）", key=_bs_sk("bs_scan"),
                  on_change=_bs_scan_entered, placeholder="光标放在这里，用扫码枪扫")
    msg = ss.pop("bs_msg", None)
    if msg:
        getattr(st, msg[0])(msg[1])

    pick = ss.get("bs_pick")
    if pick:
        st.warning(f"「{pick['raw']}」对上了好几个料，加哪个？")
        for m in pick["mats"]:
            extra = (f" · {m['color']}" if m["color"] else "") + ("（已停用）" if m["status"] == "停用" else "")
            st.button(f"{m['location'] or '无位置'}　{m['code']}（{m['customer'] or '无客户'}）{extra}",
                      key=f"bs_pick_{m['id']}", on_click=_bs_add, args=(m,))
        pc1, pc2 = st.columns(2)
        pc1.button("都加上", key="bs_pick_all", on_click=_bs_add_all, args=(pick["mats"],),
                   width="stretch")
        pc2.button("都不要", key="bs_pick_none", on_click=_bs_pick_cancel, width="stretch")

    if not rows:
        st.caption("单子还是空的，扫第一个吧。")
    else:
        st.markdown(f"**这一单：{len(rows)} 种**")
        widths = [1.2, 3.2, 1.6, 1.6, 0.8]
        for c, t in zip(st.columns(widths), ["位置", "料号", "当前库存", f"{rtype}数量", ""]):
            c.caption(t)
        conn = get_conn()
        for r in rows:
            stock = get_stock(conn, r["id"])
            unit = conn.execute("SELECT unit FROM materials WHERE id=?", (r["id"],)).fetchone()[0] or "个"
            c = st.columns(widths, vertical_alignment="center")
            c[0].write(r["location"] or "-")
            extra = (f" · {r['color']}" if r["color"] else "") + ("（已停用）" if r["status"] == "停用" else "")
            c[1].write(f"**{r['code']}**（{r['customer'] or '无客户'}）{extra}")
            val = c[3].text_input(f"数量#{r['rid']}#", key=f"bs_qty_{r['rid']}",
                                  placeholder=f"填数量（{unit}）", label_visibility="collapsed",
                                  on_change=_bs_qty_changed, args=(r["rid"],))
            qty = _bs_qty(val)
            if rtype == "出库" and qty is not None and qty > stock:
                c[2].markdown(f":red[{stock:g} {unit}，不够]")
            else:
                c[2].write(f"{stock:g} {unit}")
            c[4].button("删除", key=f"bs_del_{r['rid']}", on_click=_bs_remove, args=(r["rid"],))
        conn.close()

        st.markdown("---")
        if rtype == "出库":
            st.radio("出库类型", ["领用", "报废"], horizontal=True, key="bs_sub")
            op_pick = st.selectbox("领用人 *", get_staff_list() + [STAFF_NEW], key="bs_op_pick")
            if op_pick == STAFF_NEW:
                st.text_input("新人员姓名 *", key="bs_op_new")
            st.date_input("日期（默认今天，事后补录可以改成实际日期）", value=date.today(), key="bs_date")
        else:
            st.text_input("经手人 *", placeholder="谁办的这事", key="bs_op_in")
        st.text_input("备注（整单共用，可选）", placeholder="例如：供应商到货 / 线束车间领用", key="bs_note")
        bc1, bc2 = st.columns([3, 1])
        bc1.button(f"确认{rtype}（{len(rows)} 种）", type="primary", width="stretch", on_click=_bs_confirm)
        bc2.button("清空单子", width="stretch", on_click=_bs_clear)
        for e in ss.pop("bs_errors", None) or []:
            st.error(e)

    f = ss.pop("bs_focus", None)
    if f:
        focus_input(*f)


# ---------------- 主程序 ----------------

def main():
    scan_id = st.query_params.get("id")
    st.set_page_config(page_title="仓库管理系统", page_icon="W",
                       layout="centered" if scan_id else "wide",
                       initial_sidebar_state="collapsed" if scan_id else "auto")
    init_db()
    auto_backup()
    seed_demo_data()
    log_visit()
    check_magic_link()

    if scan_id:
        # 扫码进来的：直接看这一个物料的详情页，不进正常的多页导航，避免手机上误点别的功能
        page_scan_detail(scan_id)
        return

    st.sidebar.title("仓库管理系统")
    # 「在外未还」已从菜单撤下（2026-09-12）：借出没人登记，用不上。
    # 函数 page_outstanding 还留着，以后要用把名字加回这个列表、下面补上 elif 即可。
    page = st.sidebar.radio("功能", ["库存查询", "库位视图", "扫码上架", "入库登记", "出库登记",
                                     "批量扫码出入库", "出库提报", "出入流水", "AI 问答", "物料管理",
                                     "辅料/配件管理", "批量打印二维码", "提报处理"])
    st.sidebar.markdown("---")
    st.sidebar.caption("数据文件：warehouse.db\n备份目录：backups/（每天自动备份）\n照片目录：uploads/\n物料码：网址链接（带编号），改了访问地址要重新打印标签")
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
    elif page == "扫码上架":
        if require_full_access():
            page_scan_shelve()
    elif page == "入库登记":
        if require_full_access():
            page_record("入库")
    elif page == "出库登记":
        if require_full_access():
            page_record("出库")
    elif page == "批量扫码出入库":
        if require_full_access():
            page_batch_scan()
    elif page == "出库提报":
        page_submit_out()
    elif page == "出入流水":
        page_records()
    elif page == "AI 问答":
        page_ai()
    elif page == "物料管理":
        if require_full_access():
            page_materials()
    elif page == "辅料/配件管理":
        if require_full_access():
            page_supply_management()
    elif page == "批量打印二维码":
        if require_full_access():
            page_print_labels()
    elif page == "提报处理":
        if require_full_access():
            page_report_review()


if __name__ == "__main__":
    main()
