# -*- coding: utf-8 -*-
"""
标签打印小工具 —— 跟仓库系统完全无关，独立运行，随便什么场合要打小标签都能用
（发货治具包装袋、文件夹、临时物品……只是打印文字+可选二维码，不碰仓库数据库）。

启动：双击「start_label_tool.bat」，或命令行 streamlit run label_tool.py --server.port 8503
"""
import io
import json
import os

import pandas as pd
import qrcode
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_PATH = os.path.join(BASE_DIR, "label_tool_settings.json")


def load_settings():
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_settings(d):
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)


# ---------------- 标签绘制（跟仓库系统的位置标签同一套逻辑，独立一份不互相依赖） ----------------

def _label_font(size):
    """找一个能显示中文的字体；实在找不到就退回默认字体"""
    for path in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\simhei.ttf",
                 r"C:\Windows\Fonts\simsun.ttc",
                 "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"):
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()


def make_qr_image(text, box_size=10):
    qr = qrcode.QRCode(border=2, box_size=box_size)
    qr.add_data(text)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("RGB")


def fit_text_multiline(draw, text, max_w, max_h_px, max_lines=3, min_px=None):
    """从大字号往下试，找一行塞不下就自动换行（最多 max_lines 行）的最大字号；
    只有换到 max_lines 行、字号缩到下限还是装不下，才会截断最后一行加省略号——
    宁可字小、换行，也不该悄悄把内容截没了。返回 (font, [每行要显示的文字])"""
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


def make_label(text, width_mm=50, height_mm=30, show_qr=True, qr_content=None, dpi=300):
    """通用标签：文字（可选二维码），宽高随便填。qr_content 不传就直接拿 text 当二维码内容。
    width_mm/height_mm 必须跟打印机驱动里设置的纸张/标签尺寸一致，不然打出来会错位/被裁切。"""
    def mm(v):
        return int(v / 25.4 * dpi)

    w, h = mm(width_mm), mm(height_mm)
    margin_top = mm(1.5)
    margin_bottom = mm(1.5)
    pad = mm(0.6)
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)

    if show_qr:
        usable_h = h - margin_top - margin_bottom
        qr_size = int(min(usable_h * 0.62, w - pad * 2))
        qr_img = make_qr_image(qr_content if qr_content is not None else text).resize(
            (qr_size, qr_size), Image.NEAREST)
        img.paste(qr_img, ((w - qr_size) // 2, margin_top))
        text_top = margin_top + qr_size + mm(0.5)
        max_w = w - pad * 2
        max_h = h - margin_bottom - text_top
        text_h_ratio, max_lines = 0.72, 2
    else:
        text_top = margin_top
        max_w = w - pad * 2
        max_h = h - margin_top - margin_bottom
        text_h_ratio, max_lines = 0.85, 3   # 没有二维码，整张纸都给文字，字能放更大、行也能多留一行

    font, lines = fit_text_multiline(draw, text, max_w, max_h * text_h_ratio, max_lines=max_lines)
    draw.multiline_text((w // 2, text_top + max_h // 2), "\n".join(lines),
                        fill="black", font=font, anchor="mm", align="center",
                        spacing=int(font.size * 0.2))
    return img


def build_image_sheet(images, cols=4):
    """把多张标签图排成一张网格图，方便一次性打印后裁开分贴/整体下载预览"""
    cell_w = max(im.width for im in images) + 20
    cell_h = max(im.height for im in images) + 20
    rows = (len(images) + cols - 1) // cols
    sheet = Image.new("RGB", (cell_w * cols, cell_h * rows), "white")
    for i, im in enumerate(images):
        r, c = divmod(i, cols)
        x = c * cell_w + (cell_w - im.width) // 2
        y = r * cell_h + (cell_h - im.height) // 2
        sheet.paste(im, (x, y))
    return sheet


def image_to_png_bytes(img):
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ---------------- 打印机 ----------------

def list_printers():
    """列出这台电脑能看到的打印机，找不到 pywin32 或没权限就返回空列表，不报错"""
    try:
        import win32print
        return [p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    except Exception:
        return []


PRINT_DOC_NAME = "标签打印小工具"


def print_label_direct(img, printer_name, size_mm, height_mm=None):
    """把标签图片按真实物理尺寸直接发给Windows打印机。打印前先跟驱动汇报的实际纸张尺寸
    （PHYSICALWIDTH/HEIGHT）对一下，对不上就直接报错说清楚"驱动现在是多少、你要的是多少"，
    不然会是"打出来了但显示不全/被裁切"这种不知道哪里错的情况（驱动没汇报物理尺寸的
    极少数情况下跳过这个检查）。"""
    import win32ui
    from PIL import ImageWin

    height_mm = size_mm if height_mm is None else height_mm
    pdc = win32ui.CreateDC()
    pdc.CreatePrinterDC(printer_name)
    dpi_x = pdc.GetDeviceCaps(88)   # LOGPIXELSX
    dpi_y = pdc.GetDeviceCaps(90)   # LOGPIXELSY
    w_px = int(size_mm / 25.4 * dpi_x)
    h_px = int(height_mm / 25.4 * dpi_y)

    phys_w = pdc.GetDeviceCaps(110)   # PHYSICALWIDTH
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
    dib = ImageWin.Dib(img)
    dib.draw(pdc.GetHandleOutput(), (0, 0, w_px, h_px))
    pdc.EndPage()
    pdc.EndDoc()
    pdc.DeleteDC()


def print_job_problem(printer_name):
    """打印任务发出去之后，去Windows打印队列里查一下最近这个任务的实际状态"""
    try:
        import win32print
        h = win32print.OpenPrinter(printer_name)
        try:
            jobs = win32print.EnumJobs(h, 0, 5, 1)
        finally:
            win32print.ClosePrinter(h)
        problems = []
        for j in jobs:
            status = j.get("Status", 0)
            if status & 0x00000001:   # JOB_STATUS_PAUSED
                problems.append("任务被暂停")
            if status & 0x00000002:   # JOB_STATUS_ERROR
                problems.append("任务出错")
            if status & 0x00000200:   # JOB_STATUS_OFFLINE
                problems.append("打印机离线")
            if status & 0x00000400:   # JOB_STATUS_PAPEROUT
                problems.append("缺纸/缺标签")
        return list(dict.fromkeys(problems))
    except Exception:
        return []


# ---------------- 页面 ----------------

def main():
    st.set_page_config(page_title="标签打印小工具", page_icon="🏷️", layout="wide")
    st.title("🏷️ 标签打印小工具")
    st.caption("跟仓库系统无关，独立运行——治具包装袋、文件夹、临时物品，随便什么标签都能打。"
              "每行填一个标签内容，宽高、要不要二维码自己设，字太长会自动换行，不会被截断。")

    settings = load_settings()

    with st.expander("从 Excel/CSV 批量导入（一行一条记录的表格，比如治具明细、物料清单）", expanded=True):
        st.caption("支持 .xlsx / .xls / .csv。导入后勾选要印到标签上的列（可以选多列，"
                  "会按你勾选的顺序拼在一起），下面会先给预览，确认没问题再加进标签清单。")
        up = st.file_uploader("选择文件", type=["xlsx", "xls", "csv"], key="imp_file")
        if up is not None:
            try:
                if up.name.lower().endswith(".csv"):
                    imp = pd.read_csv(up)
                elif up.name.lower().endswith(".xls"):
                    imp = pd.read_excel(up, engine="xlrd")
                else:
                    imp = pd.read_excel(up)
                imp = imp.dropna(how="all")
                imp.columns = [str(c).strip() for c in imp.columns]
                st.write(f"识别到 **{len(imp)}** 行，**{len(imp.columns)}** 列")
                st.dataframe(imp.head(10), use_container_width=True)

                cols = list(imp.columns)
                priority = ["名称", "型号", "序号", "规格", "编号", "料号"]
                default_cols = [c for c in cols if any(p in c for p in priority)][:3] or cols[:2]
                pick_cols = st.multiselect("勾选要印到标签上的列（按勾选顺序拼接）",
                                           cols, default=default_cols, key="imp_pick_cols")
                sep = st.text_input("列之间的分隔符", value=" ", key="imp_sep")

                if pick_cols:
                    preview_lines = imp[pick_cols].astype(str).apply(
                        lambda r: sep.join(v for v in r if v and v != "nan"), axis=1).tolist()
                    st.text_area(f"预览将生成的 {len(preview_lines)} 条标签内容",
                                value="\n".join(preview_lines), height=120,
                                disabled=True, key="imp_preview_text")
                    if st.button(f"加到下面的标签清单（{len(preview_lines)} 条）", type="primary"):
                        existing = st.session_state.get("label_text", "")
                        st.session_state["label_text"] = (
                            existing.rstrip() + "\n" + "\n".join(preview_lines)).strip()
                        st.rerun()
                else:
                    st.warning("至少勾选一列")
            except Exception as e:
                st.error(f"读取文件失败：{e}")

    st.caption("批量生成：填前缀+起止编号，点「生成」自动加到下面清单里（可以点好几次拼不同批次）")
    rgc1, rgc2, rgc3, rgc4 = st.columns([2, 1, 1, 1.2])
    with rgc1:
        range_pre = st.text_input("前缀", placeholder="如 发货治具-", key="range_pre")
    with rgc2:
        range_from = st.number_input("从", min_value=1, value=1, step=1, key="range_from")
    with rgc3:
        range_to = st.number_input("到", min_value=1, value=10, step=1, key="range_to")
    with rgc4:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        range_gen = st.button("生成到清单", use_container_width=True, key="range_gen")
    if range_gen:
        if not range_pre.strip():
            st.warning("先填前缀，比如「发货治具-」")
        elif range_to < range_from:
            st.warning("「到」不能比「从」小")
        else:
            new_lines = [f"{range_pre.strip()}{n}" for n in range(int(range_from), int(range_to) + 1)]
            existing = st.session_state.get("label_text", "")
            st.session_state["label_text"] = (existing.rstrip() + "\n" + "\n".join(new_lines)).strip()
            st.rerun()

    text_area = st.text_area(
        "标签内容清单（每行一个，也可以直接在这手动加/删/改，或从 Excel/记事本粘贴）",
        height=140, key="label_text",
        placeholder="发货治具-1\n发货治具-2\n文件夹A\n...")
    items_text = [l.strip() for l in text_area.splitlines() if l.strip()]

    c1, c2, c3 = st.columns(3)
    with c1:
        width_mm = st.number_input("标签宽度（mm）", min_value=10, max_value=300,
                                   value=int(settings.get("width_mm", 50)), step=1, key="width_mm")
    with c2:
        height_mm = st.number_input("标签高度（mm）", min_value=10, max_value=300,
                                    value=int(settings.get("height_mm", 30)), step=1, key="height_mm")
    with c3:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        show_qr = st.checkbox("包含二维码", value=bool(settings.get("show_qr", True)), key="show_qr")

    printers = list_printers()
    if printers:
        saved_printer = settings.get("printer_name", "")
        options = [saved_printer] + [p for p in printers if p != saved_printer] \
            if saved_printer in printers else printers
        printer_name = st.selectbox("打印机", options, key="printer_name")
    else:
        printer_name = ""
        st.caption("没找到能用的打印机（本地USB接的，或者已经在Windows里「添加」过的共享打印机）——"
                  "生成预览、下载图片不受影响，只是不能直接打印。")

    if st.button("把当前设置存为默认值（下次打开自动带出）"):
        save_settings({"width_mm": width_mm, "height_mm": height_mm,
                       "show_qr": show_qr, "printer_name": printer_name})
        st.success("已保存")

    st.markdown("---")
    b1, b2, b3 = st.columns(3)
    with b1:
        gen = st.button(f"生成预览（{len(items_text)} 个）", type="primary",
                        use_container_width=True, disabled=not items_text)
    with b2:
        test_print = st.button("测试打印1张", use_container_width=True,
                               disabled=not items_text or not printer_name)
    with b3:
        direct_print = st.button(f"直接打印全部（{len(items_text)} 个）", use_container_width=True,
                                 disabled=not items_text or not printer_name)

    if gen:
        imgs = [make_label(t, width_mm, height_mm, show_qr) for t in items_text]
        sheet = build_image_sheet(imgs)
        st.image(sheet, use_container_width=True)
        st.download_button("下载标签图（PNG）", image_to_png_bytes(sheet),
                           file_name="标签.png", mime="image/png", type="primary")
    elif test_print:
        img = make_label(items_text[0], width_mm, height_mm, show_qr)
        try:
            print_label_direct(img, printer_name, size_mm=width_mm, height_mm=height_mm)
            problems = print_job_problem(printer_name)
            if problems:
                st.warning(f"发到「{printer_name}」了，但Windows打印队列显示：{'、'.join(problems)}"
                          f"——去电脑「设备和打印机」看这台打印机的队列，处理掉再试")
            else:
                st.success(f"已发送1张到「{printer_name}」，看看打出来尺寸、位置对不对")
        except Exception as e:
            st.error(f"打印失败：{e}")
    elif direct_print:
        ok_count, fail_count = 0, 0
        last_problems = None
        progress = st.progress(0, text="打印中…")
        for i, t in enumerate(items_text, 1):
            img = make_label(t, width_mm, height_mm, show_qr)
            try:
                print_label_direct(img, printer_name, size_mm=width_mm, height_mm=height_mm)
                ok_count += 1
                if i == 1:
                    last_problems = print_job_problem(printer_name)
                    if last_problems:
                        st.warning(f"Windows打印队列显示：{'、'.join(last_problems)}"
                                  f"——先去电脑「设备和打印机」处理掉，不然后面也打不出来")
                        break
            except Exception as e:
                fail_count += 1
                st.error(f"「{t}」打印失败：{e}")
                break
            progress.progress(i / len(items_text), text=f"打印中…{i}/{len(items_text)}")
        progress.empty()
        if last_problems:
            pass
        elif fail_count == 0:
            st.success(f"已发送 {ok_count} 张到「{printer_name}」")
        else:
            st.warning(f"发送了 {ok_count} 张，中途失败停止")


if __name__ == "__main__":
    main()
