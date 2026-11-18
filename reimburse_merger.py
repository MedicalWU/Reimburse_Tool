# -*- coding: utf-8 -*-
"""
报销材料合并工具（CUHK Surgery / MRC 报销材料归档）

- 按报销材料类型分槽位（① 订单截图 ② 支付记录 ③ 收据/发票 ④ 补充材料），每个槽位可放多个文件
- 文件可直接拖入（支持整个文件夹，自动展开其中的 PDF / 图片）
- PDF 与图片（png/jpg/bmp/tif/webp/heic…）可混合合并，图片自动按 A4 排版并摆正 EXIF 方向
- 合并前即可预览：缩略图看全貌 + 点击放大逐页翻阅
- 每一页都能拖拽换位、单独删除、单页旋转 90°，导出的 PDF 与预览完全一致
- 保存位置 = 文件夹 + 文件名（同一批材料通常只在文件名上区分）
"""

import io
import json
import os
import sys
import tempfile
import threading
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz  # PyMuPDF
from PIL import Image, ImageOps, ImageTk

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except Exception:
    pass

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except Exception:          # 没有 tkdnd 时退化为只支持按钮导入
    TkinterDnD = None
    DND_FILES = None
    HAS_DND = False

APP_TITLE = "报销材料合并工具"

# ------------------------------------------------ 视觉：CUHK 官方色彩（Pantone 255C / 117C）
PU = "#72246C"          # CUHK Purple
PU_DARK = "#4C1649"
PU_PRESS = "#3A1038"
PU_TINT = "#F5EDF4"     # 淡紫底
PU_LINE = "#E3D5E2"     # 描边
GOLD = "#C99700"        # CUHK Gold
GOLD_TINT = "#FBF3DF"
INK = "#2B262C"
INK_2 = "#6C6572"
INK_3 = "#9A939D"
BG = "#FAF8FB"
CARD = "#FFFFFF"
FONT = "Microsoft YaHei UI"


def strip_number(title):
    """去掉标题前面的 ①②③④ 圈号（改由紫色序号块呈现）"""
    if title and title[0] in "①②③④⑤⑥⑦⑧⑨":
        return title[1:].strip()
    return title


def setup_style(root):
    """统一 ttk 样式（clam 主题才能自定义背景色）"""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD)
    style.configure("TLabel", background=BG, foreground=INK, font=(FONT, 9))
    style.configure("Card.TLabel", background=CARD, foreground=INK, font=(FONT, 9))
    style.configure("Muted.TLabel", background=CARD, foreground=INK_2, font=(FONT, 8))
    style.configure("H1.TLabel", background=CARD, foreground=PU_DARK, font=(FONT, 17, "bold"))
    style.configure("H2.TLabel", background=CARD, foreground=PU, font=(FONT, 11, "bold"))
    style.configure("Body.TLabel", background=BG, foreground=INK, font=(FONT, 9))

    style.configure("TButton", font=(FONT, 9), padding=(10, 5), borderwidth=0,
                    background="#F2ECF2", foreground=PU_DARK)
    style.map("TButton",
              background=[("active", PU_LINE), ("pressed", "#E0D2DF")],
              foreground=[("disabled", "#B5AEBA")])
    style.configure("Primary.TButton", font=(FONT, 10, "bold"), padding=(20, 9), borderwidth=0,
                    background=PU, foreground="#FFFFFF")
    style.map("Primary.TButton",
              background=[("active", PU_DARK), ("pressed", PU_PRESS), ("disabled", "#CFC6D0")],
              foreground=[("disabled", "#FFFFFF")])
    style.configure("Ghost.TButton", font=(FONT, 9), padding=(7, 2), borderwidth=0,
                    background="#F2ECF2", foreground=PU)
    style.map("Ghost.TButton", background=[("active", PU_LINE)], foreground=[("active", PU_DARK)])

    style.configure("TEntry", fieldbackground="#FFFFFF", foreground=INK, padding=5,
                    bordercolor=PU_LINE, lightcolor=PU_LINE, darkcolor=PU_LINE, insertcolor=PU)
    style.configure("TCombobox", fieldbackground="#FFFFFF", background="#F2ECF2",
                    foreground=INK, padding=3, arrowcolor=PU)
    style.map("TCombobox", fieldbackground=[("readonly", "#FFFFFF")])
    style.configure("Vertical.TScrollbar", background=PU_LINE, troughcolor=BG,
                    bordercolor=BG, arrowcolor=PU, borderwidth=0, width=10, arrowsize=12)
    style.configure("Horizontal.TProgressbar", background=PU, troughcolor=PU_TINT,
                    bordercolor=PU_TINT, lightcolor=PU, darkcolor=PU, thickness=8)

PDF_EXTS = {".pdf"}
IMG_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".heic", ".heif", ".gif"}
ALL_EXTS = PDF_EXTS | IMG_EXTS

SLOT_DEFS = [
    ("order", "① 采购 / 订单截图明细",
     "需显示：总价 + 支付时间 + 明细 + 店铺或商品名称"),
    ("pay", "② 支付记录",
     "微信/支付宝截图＋支付证明；或含持卡人姓名的月结单"),
    ("receipt", "③ 收据 / 发票",
     "需显示：抬头 + 明细 + 金额 + 日期 + 章"),
    ("extra", "④ 补充材料（可选）",
     "行程单、登机牌、邀请邮件、badge、比价截图等"),
]

ENTITY_DEFS = [
    ("surgery", "Surgery（外科）", "香港中文大学医学院"),
    ("mrc", "MRC（医疗机器人创新中心）", "Multi-Scale Medical Robotics Center 医疗机器人创新技术中心"),
]

THUMB_W = 130          # 缩略图目标宽度（像素）
A4_W, A4_H = 595.0, 842.0
THUMB_MAX_H = 1600.0


# ---------------------------------------------------------------- 核心逻辑
def is_pdf(path):
    return os.path.splitext(path)[1].lower() in PDF_EXTS


def page_count(path):
    if is_pdf(path):
        doc = fitz.open(path)
        try:
            if doc.needs_pass:
                raise ValueError("PDF 已加密，需密码才能读取")
            return doc.page_count
        finally:
            doc.close()
    return 1


def _add_image_page(doc, path):
    """把一张图片铺成一页（保持比例、按 A4 宽度排版）"""
    img = Image.open(path)
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    if img.mode in ("RGBA", "LA", "P"):
        bg = Image.new("RGB", img.size, "white")
        conv = img.convert("RGBA")
        bg.paste(conv, mask=conv.split()[-1])
        img = bg
    else:
        img = img.convert("RGB")

    w, h = img.size
    if w == 0 or h == 0:
        raise ValueError("图片尺寸异常：%s" % os.path.basename(path))

    if h >= w:                      # 竖图（手机截图、票据照片）
        pw = A4_W
        ph = min(A4_W * h / w, THUMB_MAX_H)
    else:                           # 横图（截图、表格）
        pw = A4_H
        ph = A4_H * h / w
    if ph < 120:                    # 极扁的图给个下限，避免页面过窄
        ph = 120.0

    page = doc.new_page(width=pw, height=ph)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=92)
    page.insert_image(fitz.Rect(0, 0, pw, ph), stream=buf.getvalue())


def build_document(slots, rots=None, order=None):
    """生成合并后的 fitz.Document（调用方负责 close）

    rots : {(文件的绝对路径, 文件内页码): 追加旋转角 0/90/180/270}
    order: [(文件绝对路径, 文件内页码), ...] 显式页序；None 表示按槽位顺序整份并入
    """
    out = fitz.open()
    cache = {}
    try:
        if order is not None:
            for path, page_no in order:
                if not os.path.exists(path):
                    raise ValueError("文件不存在：%s" % path)
                start = out.page_count
                if is_pdf(path):
                    src = cache.get(path)
                    if src is None:
                        src = fitz.open(path)
                        cache[path] = src
                        if src.needs_pass:
                            raise ValueError("PDF 已加密，无法合并：%s" % os.path.basename(path))
                    if page_no > src.page_count:
                        continue
                    out.insert_pdf(src, from_page=page_no - 1, to_page=page_no - 1)
                else:
                    _add_image_page(out, path)
                if rots:
                    add = rots.get((path, page_no), 0)
                    if add and out.page_count > start:
                        page = out[start]
                        page.set_rotation((page.rotation + add) % 360)
            return out

        for key, _title, _hint in SLOT_DEFS:
            for entry in slots.get(key, []):
                path = entry["path"]
                if not os.path.exists(path):
                    raise ValueError("文件不存在：%s" % path)
                start = out.page_count
                if is_pdf(path):
                    src = fitz.open(path)
                    try:
                        if src.needs_pass:
                            raise ValueError("PDF 已加密，无法合并：%s" % os.path.basename(path))
                        out.insert_pdf(src)
                    finally:
                        src.close()
                else:
                    try:
                        _add_image_page(out, path)
                    except Exception as e:
                        raise ValueError("图片读取失败（%s）：%s" % (os.path.basename(path), e))
                if rots:
                    ap = os.path.abspath(path)
                    for p in range(out.page_count - start):
                        add = rots.get((ap, p + 1), 0)
                        if add:
                            page = out[start + p]
                            page.set_rotation((page.rotation + add) % 360)
    except Exception:
        out.close()
        raise
    finally:
        for d in cache.values():
            try:
                d.close()
            except Exception:
                pass
    return out


def merge_to_file(slots, output_path, rots=None, order=None):
    """生成并保存合并后的 PDF，返回 (总页数, 文件大小字节)"""
    doc = build_document(slots, rots, order)
    try:
        if doc.page_count == 0:
            raise ValueError("没有可合并的内容，请先添加文件")
        tmp = output_path + ".part"
        doc.save(tmp, garbage=3, deflate=True)
        total = doc.page_count
    finally:
        doc.close()
    if os.path.exists(output_path):
        os.remove(output_path)
    os.replace(tmp, output_path)
    return total, os.path.getsize(output_path)


def merge_units(slots):
    """给预览用：按槽位顺序返回每一页的信息（不含用户自定义排序）"""
    units = []
    n = 0
    for key, title, _hint in SLOT_DEFS:
        for entry in slots.get(key, []):
            ap = os.path.abspath(entry["path"])
            try:
                cnt = entry.get("pages") or page_count(entry["path"])
            except Exception:
                cnt = 1
            for p in range(cnt):
                n += 1
                units.append({
                    "key": (ap, p + 1),
                    "title": title,
                    "name": os.path.basename(entry["path"]),
                    "page_no": p + 1,
                    "g": n,
                })
    return units


def human_size(n):
    if n < 1024 * 1024:
        return "%.0f KB" % (n / 1024.0)
    return "%.1f MB" % (n / 1024.0 / 1024.0)


def truncate(text, limit=26):
    return text if len(text) <= limit else text[:limit - 1] + "…"


def natural_key(text):
    """1.pdf < 2.pdf < 10.pdf"""
    import re
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r"(\d+)", text)]


def collect_drop_paths(paths):
    """把拖入的路径整理成文件列表：目录会被展开为其中的 PDF/图片"""
    files = []
    for raw in paths:
        p = str(raw).strip().strip('"')
        if not p:
            continue
        if os.path.isdir(p):
            found = []
            for dirpath, dirnames, filenames in os.walk(p):
                dirnames[:] = [d for d in dirnames if not d.startswith(".")]
                for fn in filenames:
                    if os.path.splitext(fn)[1].lower() in ALL_EXTS:
                        found.append(os.path.join(dirpath, fn))
            found.sort(key=lambda x: natural_key(os.path.relpath(x, p)))
            files.extend(found[:300])
        elif os.path.splitext(p)[1].lower() in ALL_EXTS:
            files.append(p)
    return files


# ---------------------------------------------------------------- 配置
def config_path():
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    d = os.path.join(base, "ReimburseMerger")
    return os.path.join(d, "config.json")


def load_config():
    try:
        with open(config_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg):
    try:
        os.makedirs(os.path.dirname(config_path()), exist_ok=True)
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ---------------------------------------------------------------- 放大预览窗
class PageViewer(tk.Toplevel):
    def __init__(self, master, app, index):
        super().__init__(master)
        self.app = app
        self.index = index
        self.photo = None
        self.fit_mode = True
        self.title("预览")
        self.geometry("900x760")
        self.minsize(600, 480)
        self.transient(master)
        self._center_on(master)

        bar = ttk.Frame(self, padding=(8, 6))
        bar.pack(fill="x")
        ttk.Button(bar, text="◀ 上一页", width=9, command=lambda: self.step(-1)).pack(side="left")
        ttk.Button(bar, text="下一页 ▶", width=9, command=lambda: self.step(1)).pack(side="left", padx=(6, 0))
        self.pos_var = tk.StringVar()
        ttk.Label(bar, textvariable=self.pos_var, font=(FONT, 10, "bold")).pack(side="left", padx=12)
        ttk.Button(bar, text="适应窗口", width=9, command=self.set_fit).pack(side="left", padx=(10, 0))
        ttk.Button(bar, text="实际大小", width=9, command=self.set_actual).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="↻ 旋转 90°", style="Ghost.TButton",
                   command=lambda: self.rotate(90)).pack(side="left", padx=(10, 0))
        ttk.Button(bar, text="打开原文件", width=11, command=self.open_source).pack(side="right")

        bar2 = ttk.Frame(self, padding=(8, 0))
        bar2.pack(fill="x", pady=(0, 6))
        ttk.Button(bar2, text="◀ 前移一位", style="Ghost.TButton",
                   command=lambda: self.move(-1)).pack(side="left")
        ttk.Button(bar2, text="后移一位 ▶", style="Ghost.TButton",
                   command=lambda: self.move(1)).pack(side="left", padx=(6, 0))
        ttk.Button(bar2, text="移到最前", style="Ghost.TButton",
                   command=lambda: self.move_to(0)).pack(side="left", padx=(6, 0))
        ttk.Button(bar2, text="移到最后", style="Ghost.TButton",
                   command=lambda: self.move_to(-1)).pack(side="left", padx=(6, 0))
        ttk.Button(bar2, text="✕ 删掉这页", style="Ghost.TButton",
                   command=self.delete_current).pack(side="left", padx=(12, 0))
        ttk.Label(bar2, text="调整页序 / 删页都会实时同步到主界面缩略图",
                  foreground="#777", font=(FONT, 8)).pack(side="left", padx=(10, 0))

        self.canvas = tk.Canvas(self, background="#3a3a3a", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self.render())
        self.bind("<Left>", lambda e: self.step(-1))
        self.bind("<Right>", lambda e: self.step(1))
        self.bind("<Escape>", lambda e: self.destroy())

        # 模态：避免主窗在查看时重建文档（窗口尚未显示时 grab 会失败，稍后重试）
        self.after(120, self._try_grab)
        self.render()

    def _center_on(self, master):
        """相对主窗居中，避免预览窗弹到屏幕角落"""
        try:
            master.update_idletasks()
            w, h = 900, 760
            x = master.winfo_rootx() + max(0, (master.winfo_width() - w) // 2)
            y = master.winfo_rooty() + max(0, (master.winfo_height() - h) // 2)
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            x = max(0, min(x, sw - w - 10))
            y = max(0, min(y, sh - h - 60))
            self.geometry("900x760+%d+%d" % (x, y))
        except Exception:
            pass

    def _try_grab(self):
        try:
            if self.winfo_exists():
                self.grab_set()
        except Exception:
            pass

    def step(self, delta):
        total = self.app.doc_page_count()
        if total == 0:
            return
        self.index = max(0, min(total - 1, self.index + delta))
        self.render()

    def set_fit(self):
        self.fit_mode = True
        self.render()

    def set_actual(self):
        self.fit_mode = False
        self.render()

    def open_source(self):
        try:
            self.app.open_current_source(self.index)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)

    def rotate(self, delta):
        """旋转当前页 90°，并立即刷新放大视图与主界面缩略图"""
        try:
            self.app.rotate_page(self.index, delta, keep_viewer=True)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "旋转失败：%s" % e, parent=self)
            return
        self.render()

    def move(self, delta):
        """当前页前移/后移一位"""
        try:
            new_index = self.app.move_page(self.index, delta, keep_viewer=True)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "调整顺序失败：%s" % e, parent=self)
            return
        if new_index is not None:
            self.index = new_index
        self.render()

    def move_to(self, target):
        try:
            new_index = self.app.move_page_to(self.index, target, keep_viewer=True)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "调整顺序失败：%s" % e, parent=self)
            return
        if new_index is not None:
            self.index = new_index
        self.render()

    def delete_current(self):
        """删掉当前页；删完还有页就停在新位置，全删完就关窗"""
        if not self.app.units:
            return
        try:
            self.app.delete_page(self.index, keep_viewer=True)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "删除失败：%s" % e, parent=self)
            return
        if not self.app.units:
            self.destroy()
            return
        self.index = min(self.index, len(self.app.units) - 1)
        self.render()

    def render(self):
        total = self.app.doc_page_count()
        if total == 0:
            self.pos_var.set("无内容")
            return
        self.index = max(0, min(total - 1, self.index))

        cw = max(self.canvas.winfo_width(), 200)
        ch = max(self.canvas.winfo_height(), 200)
        page = self.app.doc[self.index]
        pw, ph = page.rect.width, page.rect.height

        if self.fit_mode:
            zoom = min((cw - 24) / pw, (ch - 24) / ph)
        else:
            zoom = 96.0 / 72.0
        zoom = max(zoom, 0.1)

        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        self.photo = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        x = max(cw // 2, pix.width // 2 + 8)
        y = max(ch // 2, pix.height // 2 + 8)
        self.canvas.create_image(x, y, image=self.photo, anchor="center")
        title, name, page_no, _g, _ap = (self.app.units[self.index]["title"],
                                         self.app.units[self.index]["name"],
                                         self.app.units[self.index]["page_no"],
                                         self.app.units[self.index]["g"],
                                         self.app.units[self.index]["key"][0])
        self.title("预览 · 第 %d/%d 页 · %s · %s" % (
            self.index + 1, total, truncate(title, 10), truncate(name, 18)))
        self.pos_var.set("第 %d / %d 页" % (self.index + 1, total))


# ---------------------------------------------------------------- 主界面
class App:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()
        self.slots = {key: [] for key, _t, _h in SLOT_DEFS}
        self.doc = fitz.open()          # 预览用文档（= 合并结果）
        self.units = []                 # 页 -> 来源映射
        self.thumbs = []                # 缩略图 PhotoImage 引用
        self._render_token = 0
        self._viewer = None
        self._dir_auto = True          # 目录仍由程序自动推断（用户手动选/改后置 False）
        self._name_auto = True         # 文件名仍是自动生成
        self._auto_name_last = ""
        self.dnd_ready = False         # 拖放是否注册成功
        self.page_order = []           # 用户可自定义的页序 [(文件绝对路径, 文件内页码)]
        self._order_touched = False    # 用户是否手动调整过页序
        self.hidden = set()           # 用户单独删掉的页 (绝对路径, 文件内页码)
        self._cells = []              # 缩略图卡片（拖拽时用来判断落点）
        self._drag = None
        self.rots = {}                 # {(文件绝对路径, 文件内页码): 追加旋转角}
        for k, v in (self.cfg.get("rots") or {}).items():
            try:
                path, page = k.rsplit("|", 1)
                self.rots[(path, int(page))] = int(v) % 360
            except Exception:
                continue

        root.title(APP_TITLE + " · CUHK")
        root.geometry("1200x820")
        root.minsize(1020, 660)
        root.configure(background=BG)

        self._build_ui()
        self.refresh_all()

    # ---------------- UI 搭建
    def _build_ui(self):
        setup_style(self.root)
        self.root.configure(background=BG)

        outer = tk.Frame(self.root, background=BG)
        outer.pack(fill="both", expand=True)
        outer.grid_rowconfigure(1, weight=1)
        outer.grid_columnconfigure(0, weight=1)

        # ---- 顶部品牌条
        headwrap = tk.Frame(outer, background=CARD)
        headwrap.grid(row=0, column=0, sticky="ew")
        head = tk.Frame(headwrap, background=CARD, padx=18, pady=7)
        head.pack(fill="x")

        badge = tk.Label(head, text="CUHK", background=PU, foreground="#FFFFFF",
                         font=(FONT, 13, "bold"), padx=14, pady=10)
        badge.pack(side="left")

        titles = tk.Frame(head, background=CARD, padx=14)
        titles.pack(side="left")
        tk.Label(titles, text=APP_TITLE, background=CARD, foreground=PU_DARK,
                 font=(FONT, 17, "bold")).pack(anchor="w")
        tk.Label(titles, text="The Chinese University of Hong Kong · Surgery / MRC 报销材料归档",
                 background=CARD, foreground=INK_2, font=(FONT, 9)).pack(anchor="w", pady=(2, 0))

        tk.Button(head, text="一键全部移除", command=self.clear_all_slots,
                  background=CARD, foreground=PU, activebackground=CARD, activeforeground=PU_DARK,
                  font=(FONT, 9), relief="flat", bd=0, cursor="hand2").pack(side="right")
        tk.Button(head, text="使用说明", command=self.show_help,
                  background=CARD, foreground=INK_2, activebackground=CARD,
                  font=(FONT, 9), relief="flat", bd=0, cursor="hand2").pack(side="right", padx=(0, 14))

        # 金色细线：CUHK 校色点缀
        tk.Frame(headwrap, background=GOLD, height=3).pack(fill="x")
        tk.Frame(headwrap, background=PU_LINE, height=1).pack(fill="x")

        # ---- 主体
        body = tk.Frame(outer, background=BG)
        body.grid(row=1, column=0, sticky="nsew", padx=14, pady=(12, 0))
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, minsize=470, weight=0)
        body.grid_columnconfigure(1, weight=1)

        self._build_slots(body)
        self._build_preview(body)

        # ---- 底部操作区
        bottom = tk.Frame(outer, background=BG, padx=14, pady=12)
        bottom.grid(row=2, column=0, sticky="ew")
        self._build_bottom(bottom)

        self._update_entity_hint()
        self.root.bind_all("<MouseWheel>", self._on_wheel)

    def _card(self, master, **kw):
        """带描边的白色卡片"""
        return tk.Frame(master, background=CARD, padx=10, pady=8,
                        highlightbackground=PU_LINE, highlightthickness=1, **kw)

    @staticmethod
    def _short(text, n=21):
        return text if len(text) <= n else text[:n - 1] + "…"

    def _tip(self, widget, text):
        """鼠标悬停显示完整说明"""
        tip = {"win": None}

        def show(_e=None):
            if tip["win"] or not text:
                return
            w = tk.Toplevel(self.root)
            w.wm_overrideredirect(True)
            w.configure(background=PU_DARK)
            tk.Label(w, text=text, background=PU_DARK, foreground="#FFFFFF",
                     font=(FONT, 8), padx=8, pady=4).pack()
            w.wm_geometry("+%d+%d" % (widget.winfo_rootx(),
                                      widget.winfo_rooty() + widget.winfo_height() + 4))
            tip["win"] = w

        def hide(_e=None):
            if tip["win"]:
                try:
                    tip["win"].destroy()
                except Exception:
                    pass
                tip["win"] = None

        widget.bind("<Enter>", show, add="+")
        widget.bind("<Leave>", hide, add="+")

    def show_help(self):
        messagebox.showinfo(
            "使用说明",
            "1) 把材料直接拖进左栏对应栏目（也可点「＋ 添加」）\n"
            "2) 同一栏目内用 ↑ ↓ 调整顺序，拖动整个文件夹会自动展开\n"
            "3) 右栏缩略图即最终顺序，点开可放大核对抬头/金额/日期\n"
            "4) 选好文件夹 + 填文件名 → 开始合并\n\n"
            "抬头：Surgery 写「香港中文大学医学院」；\n"
            "      MRC 写「Multi-Scale Medical Robotics Center 医疗机器人创新技术中心」。",
        )

    def _build_bottom(self, parent):
        card = self._card(parent)
        card.pack(fill="x")

        # 抬头提醒
        r0 = tk.Frame(card, background=CARD)
        r0.pack(fill="x")
        tk.Label(r0, text="抬头提醒", background=CARD, foreground=INK_2,
                 font=(FONT, 9, "bold")).pack(side="left")
        self.entity_var = tk.StringVar()
        self.entity_combo = ttk.Combobox(r0, textvariable=self.entity_var, state="readonly",
                                         values=[e[1] for e in ENTITY_DEFS], width=22)
        self.entity_combo.current(0)
        self.entity_combo.pack(side="left", padx=(10, 6))
        self.entity_combo.bind("<<ComboboxSelected>>", lambda e: self._update_entity_hint())
        ttk.Button(r0, text="复制抬头", style="Ghost.TButton", command=self.copy_entity).pack(side="left")
        self.entity_hint = tk.Label(r0, text="", background=CARD, foreground=PU,
                                    font=(FONT, 9))
        self.entity_hint.pack(side="left", padx=(12, 0))

        # 保存位置
        r1 = tk.Frame(card, background=CARD)
        r1.pack(fill="x", pady=(6, 0))
        tk.Label(r1, text="文件夹", width=5, anchor="w", background=CARD,
                 foreground=INK_2).pack(side="left")
        self.dir_var = tk.StringVar()
        self.dir_entry = ttk.Entry(r1, textvariable=self.dir_var)
        self.dir_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(r1, text="选择文件夹…", style="Ghost.TButton",
                   command=self.choose_dir).pack(side="left", padx=(8, 0))

        r1b = tk.Frame(card, background=CARD)
        r1b.pack(fill="x", pady=(5, 0))
        tk.Label(r1b, text="文件名", width=5, anchor="w", background=CARD,
                 foreground=INK_2).pack(side="left")
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(r1b, textvariable=self.name_var)
        self.name_entry.pack(side="left", fill="x", expand=True)
        tk.Label(r1b, text=".pdf", background=CARD, foreground=INK_3).pack(side="left", padx=(4, 6))
        self.merge_btn = ttk.Button(r1b, text="开始合并", style="Primary.TButton",
                                   command=self.start_merge)
        self.merge_btn.pack(side="left")

        self.dir_var.trace_add("write", lambda *_: self._update_target_preview())
        self.name_var.trace_add("write", lambda *_: self._update_target_preview())
        self.dir_entry.bind("<KeyRelease>", lambda e: setattr(self, "_dir_auto", False))
        self.name_entry.bind("<KeyRelease>", lambda e: setattr(self, "_name_auto", False))

        # 目标路径 + 状态 + 进度
        r2 = tk.Frame(card, background=CARD)
        r2.pack(fill="x", pady=(5, 0))
        self.status_var = tk.StringVar(value="就绪")
        tk.Label(r2, textvariable=self.status_var, background=CARD, foreground=PU,
                 font=(FONT, 9, "bold")).pack(side="right")
        self.target_var = tk.StringVar(value="")
        tk.Label(r2, textvariable=self.target_var, background=CARD, foreground=INK_3,
                 font=(FONT, 8), anchor="w").pack(side="left", fill="x", expand=True)

        self.progress = ttk.Progressbar(card, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(6, 0))

    def _on_wheel(self, event):
        """按鼠标所在区域决定滚动哪个面板"""
        target = None
        w = event.widget
        depth = 0
        while w is not None and depth < 12:
            if w is getattr(self, "pv_canvas", None):
                target = self.pv_canvas
                break
            if w is getattr(self, "slot_canvas", None):
                target = self.slot_canvas
                break
            w = getattr(w, "master", None)
            depth += 1
        if target is None:
            target = self.slot_canvas
        try:
            target.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    def _build_slots(self, parent):
        col = tk.Frame(parent, background=BG)
        col.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        head = tk.Frame(col, background=BG)
        head.pack(fill="x", pady=(0, 6))
        tk.Label(head, text="材料归位", background=BG, foreground=PU_DARK,
                 font=(FONT, 11, "bold")).pack(side="left")
        tk.Label(head, text="文件可直接拖进对应栏目 · 同栏内用 ↑ ↓ 排序", background=BG,
                 foreground=INK_3, font=(FONT, 8)).pack(side="left", padx=(8, 0))

        wrap = tk.Frame(col, background=BG)
        wrap.pack(fill="both", expand=True)

        canvas = tk.Canvas(wrap, background=BG, highlightthickness=0, width=452)
        self.slot_canvas = canvas
        sb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, background=BG)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.slot_widgets = {}
        for idx, (key, title, hint) in enumerate(SLOT_DEFS):
            box = self._card(inner)
            box.pack(fill="x", padx=2, pady=(0 if idx == 0 else 8, 0))

            top = tk.Frame(box, background=CARD)
            top.pack(fill="x")
            tk.Label(top, text=str(idx + 1), background=PU, foreground="#FFFFFF",
                     font=(FONT, 9, "bold"), width=2, padx=2, pady=0).pack(side="left")
            tk.Label(top, text=strip_number(title), background=CARD, foreground=PU_DARK,
                     font=(FONT, 10, "bold")).pack(side="left", padx=(7, 8))
            count = tk.Label(top, text="", background=CARD, foreground=INK_3, font=(FONT, 8))
            count.pack(side="right")
            hint_lbl = tk.Label(top, text=self._short(hint), background=CARD,
                                foreground=INK_3, font=(FONT, 8))
            hint_lbl.pack(side="left")
            self._tip(hint_lbl, hint)

            lb = tk.Listbox(box, selectmode="extended", activestyle="none", height=2,
                            font=(FONT, 9), background="#FFFFFF", foreground=INK,
                            selectbackground=PU_TINT, selectforeground=PU_DARK,
                            borderwidth=0, highlightthickness=1,
                            highlightbackground=PU_LINE, highlightcolor=PU,
                            relief="flat")
            lb.pack(fill="x", pady=(4, 0))
            lb.bind("<Double-Button-1>", lambda e, k=key: self.open_slot_file(k))
            lb.bind("<Delete>", lambda e, k=key: self.remove_selected(k))

            if HAS_DND:
                for target in (lb, box):
                    try:
                        target.drop_target_register(DND_FILES)
                        target.dnd_bind("<<Drop>>", lambda e, k=key: self._on_drop(k, e))
                        target.dnd_bind("<<DropEnter>>", lambda e, w=lb: w.configure(background=PU_TINT))
                        target.dnd_bind("<<DropLeave>>", lambda e, w=lb: w.configure(background="#FFFFFF"))
                        self.dnd_ready = True
                    except Exception:
                        pass

            btns = tk.Frame(box, background=CARD)
            btns.pack(fill="x", pady=(4, 0))
            ttk.Button(btns, text="＋ 添加", style="Ghost.TButton",
                       command=lambda k=key: self.add_files(k)).pack(side="left")
            ttk.Button(btns, text="移除选中", style="Ghost.TButton",
                       command=lambda k=key: self.remove_selected(k)).pack(side="left", padx=(4, 0))
            ttk.Button(btns, text="清空", style="Ghost.TButton",
                       command=lambda k=key: self.clear_slot(k)).pack(side="left", padx=(4, 0))
            ttk.Button(btns, text="↑", style="Ghost.TButton", width=2,
                       command=lambda k=key: self.move(k, -1)).pack(side="right")
            ttk.Button(btns, text="↓", style="Ghost.TButton", width=2,
                       command=lambda k=key: self.move(k, 1)).pack(side="right", padx=(0, 4))

            self.slot_widgets[key] = {"listbox": lb, "count": count}

    def _build_preview(self, parent):
        card = self._card(parent)
        card.grid(row=0, column=1, sticky="nsew")

        bar = tk.Frame(card, background=CARD)
        bar.pack(fill="x")
        tk.Label(bar, text="合并预览", background=CARD, foreground=PU_DARK,
                 font=(FONT, 11, "bold")).pack(side="left")
        tk.Label(bar, text="直接拖动缩略图即可换位 · 点缩略图放大 · ↻ 转 90° · ✕ 删掉该页", background=CARD,
                 foreground=INK_3, font=(FONT, 8)).pack(side="left", padx=(8, 0))
        self.page_var = tk.StringVar(value="")
        self.page_chip = tk.Label(bar, textvariable=self.page_var, background=PU_TINT,
                                  foreground=PU, font=(FONT, 8, "bold"), padx=8, pady=2)
        self.page_chip.pack(side="right")
        self.restore_lbl = tk.Label(bar, text="", background=CARD, foreground=GOLD,
                                    font=(FONT, 8, "bold"), cursor="hand2", padx=6)
        self.restore_lbl.pack(side="right")
        self.restore_lbl.bind("<Button-1>", lambda e: self.restore_hidden_pages())

        wrap = tk.Frame(card, background=CARD)
        wrap.pack(fill="both", expand=True, pady=(8, 0))
        self.pv_canvas = tk.Canvas(wrap, background="#FFFFFF", highlightthickness=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.pv_canvas.yview)
        self.pv_inner = tk.Frame(self.pv_canvas, background="#FFFFFF")
        self.pv_inner.bind("<Configure>",
                           lambda e: self.pv_canvas.configure(scrollregion=self.pv_canvas.bbox("all")))
        win = self.pv_canvas.create_window((0, 0), window=self.pv_inner, anchor="nw")
        self._pv_win = win
        self._thumb_cols_cache = None
        self.pv_canvas.bind("<Configure>", self._on_pv_resize)
        self._pv_window = win
        self.pv_canvas.configure(yscrollcommand=sb.set)
        self.pv_canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.empty_hint = tk.Label(self.pv_inner,
                                   text="把材料拖进左栏对应栏目，这里会按最终合并顺序显示每一页",
                                   background="#FFFFFF", foreground=INK_3, font=(FONT, 9))
        self.empty_hint.pack(pady=60)

        if HAS_DND:
            try:
                self.pv_canvas.drop_target_register(DND_FILES)
                self.pv_canvas.dnd_bind("<<Drop>>", self._on_drop_preview)
            except Exception:
                pass

    def _on_pv_resize(self, event):
        """预览区宽度变化时：跟随宽度铺满，并在列数变化时重排缩略图"""
        try:
            self.pv_canvas.itemconfigure(self._pv_win, width=event.width)
        except Exception:
            pass
        cols = max(2, min(6, int((event.width - 14) // (THUMB_W + 26))))
        if self.units and cols != getattr(self, "_thumb_cols_cache", None):
            self._thumb_cols_cache = cols
            self._render_token += 1
            for w in self.pv_inner.winfo_children():
                w.destroy()
            self.thumbs = []
            self._layout_thumbs(self._render_token)

    def _on_drop_preview(self, event):
        """拖到右侧预览区时给出提示，避免用户以为没反应"""
        self.status_var.set("请把文件拖进左栏对应栏目（① 订单 / ② 支付 / ③ 发票 / ④ 补充）")
        return "break"

    # ---------------- 槽位操作
    def add_files(self, key):
        files = filedialog.askopenfilenames(
            title="添加材料（可多选，支持 PDF 与图片）",
            filetypes=[("PDF 与图片", "*.pdf *.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.heic *.heif *.gif"),
                       ("PDF 文件", "*.pdf"),
                       ("图片", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.heic *.heif *.gif"),
                       ("所有文件", "*.*")],
        )
        if not files:
            return
        self._add_paths(key, list(files))

    def _add_paths(self, key, paths):
        """把一批路径加入槽位（按钮导入与拖入共用）"""
        added, duplicated, bad = 0, 0, []
        existing = [os.path.abspath(e["path"]) for e in self.slots[key]]
        for f in paths:
            ext = os.path.splitext(f)[1].lower()
            if ext not in ALL_EXTS:
                bad.append(os.path.basename(f))
                continue
            if os.path.abspath(f) in existing:
                duplicated += 1
                continue
            try:
                cnt = page_count(f)
            except Exception as e:
                bad.append("%s（%s）" % (os.path.basename(f), e))
                continue
            self.slots[key].append({"path": f, "pages": cnt})
            existing.append(os.path.abspath(f))
            added += 1

        self.refresh_all()
        msg = []
        if added:
            msg.append("已添加 %d 个文件" % added)
        if duplicated:
            msg.append("跳过重复 %d 个" % duplicated)
        if bad:
            msg.append("无法添加 %d 个" % len(bad))
        if msg:
            self.status_var.set("，".join(msg))
        if bad:
            messagebox.showwarning(APP_TITLE, "以下文件无法添加：\n" + "\n".join(bad))
        return added

    def _on_drop(self, key, event):
        """把拖到某个槽位上的文件/文件夹加入该槽位"""
        try:
            raw = self.root.tk.splitlist(event.data)
        except Exception:
            raw = [event.data]
        files = collect_drop_paths(raw)
        if not files:
            self.status_var.set("拖入的内容里没有 PDF 或图片")
            return "break"
        title = dict((k, t) for k, t, _h in SLOT_DEFS)[key]
        n = self._add_paths(key, files)
        self.status_var.set("已拖入「%s」：%d 个文件" % (title, n))
        return "copy"

    def _slot(self, key):
        return self.slots[key]

    def remove_selected(self, key):
        lb = self.slot_widgets[key]["listbox"]
        idxs = list(lb.curselection())
        if not idxs:
            self.status_var.set("请先选中要移除的文件")
            return
        for i in reversed(idxs):
            del self.slots[key][i]
        self.refresh_all()
        self.status_var.set("已移除 %d 个文件" % len(idxs))

    def move(self, key, delta):
        lb = self.slot_widgets[key]["listbox"]
        idxs = list(lb.curselection())
        if not idxs:
            self.status_var.set("请先选中要调整顺序的文件")
            return
        lst = self.slots[key]
        order = sorted(idxs) if delta < 0 else sorted(idxs, reverse=True)
        for i in order:
            t = i + delta
            if 0 <= t < len(lst):
                lst[i], lst[t] = lst[t], lst[i]
        self.refresh_all()
        sel = max(0, min(len(lst) - 1, (min(idxs) + delta) if delta < 0 else (max(idxs) + delta)))
        if lst:
            lb.selection_set(sel)

    def open_slot_file(self, key):
        lb = self.slot_widgets[key]["listbox"]
        idxs = list(lb.curselection())
        if not idxs:
            return
        path = self.slots[key][idxs[0]]["path"]
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "无法打开：%s" % e)

    def clear_slot(self, key):
        """一键清空某个槽位里的全部文件"""
        n = len(self.slots[key])
        if not n:
            self.status_var.set("该槽位已经是空的")
            return
        if not messagebox.askyesno(APP_TITLE, "确定移除「%s」里的 %d 个文件吗？" % (
                dict((k, t) for k, t, _h in SLOT_DEFS)[key], n)):
            return
        self.slots[key] = []
        self.refresh_all()
        self.status_var.set("已移除 %d 个文件" % n)

    def clear_all_slots(self):
        """一键移除全部槽位"""
        total = sum(len(v) for v in self.slots.values())
        if not total:
            self.status_var.set("列表已经是空的")
            return
        if not messagebox.askyesno(APP_TITLE, "确定移除全部 %d 个文件吗？" % total):
            return
        for key, _t, _h in SLOT_DEFS:
            self.slots[key] = []
        self.refresh_all()
        self.status_var.set("已全部移除（%d 个文件）" % total)

    # ---------------- 预览
    def doc_page_count(self):
        return self.doc.page_count

    def open_current_source(self, index):
        if index >= len(self.units):
            return
        name = self.units[index]["name"]
        ap = self.units[index]["key"][0]
        if os.path.exists(ap):
            os.startfile(ap)
            return
        for lst in self.slots.values():
            for e in lst:
                if os.path.basename(e["path"]) == name:
                    os.startfile(e["path"])
                    return
        raise ValueError("找不到对应的源文件")

    def rotate_page(self, index, delta=90, keep_viewer=False):
        """把第 index 页（0 基）顺时针旋转 delta 度，并刷新预览"""
        if index < 0 or index >= len(self.units):
            return 0
        page_no = self.units[index]["page_no"]
        k = (self.units[index]["key"][0], page_no)
        newv = (self.rots.get(k, 0) + delta) % 360
        if newv:
            self.rots[k] = newv
        else:
            self.rots.pop(k, None)
        self._save_rots()
        self._rebuild_preview(keep_viewer=keep_viewer)
        return newv

    def move_page(self, index, delta, keep_viewer=False):
        """把第 index 页前移/后移若干位，返回该页新的下标"""
        if not self.page_order or index < 0 or index >= len(self.page_order):
            return None
        target = index + delta
        if target < 0 or target >= len(self.page_order):
            return index
        item = self.page_order.pop(index)
        self.page_order.insert(target, item)
        self._order_touched = True
        self._rebuild_preview(keep_viewer=keep_viewer)
        return target

    def move_page_to(self, index, target, keep_viewer=False):
        """把第 index 页移到指定位置（target 为 -1 表示移到最后）"""
        if not self.page_order or index < 0 or index >= len(self.page_order):
            return None
        if target < 0 or target >= len(self.page_order):
            target = len(self.page_order) - 1
        if target == index:
            return index
        item = self.page_order.pop(index)
        self.page_order.insert(target, item)
        self._order_touched = True
        self._rebuild_preview(keep_viewer=keep_viewer)
        return target

    def _save_rots(self):
        self.cfg["rots"] = {"%s|%d" % (p, n): v for (p, n), v in self.rots.items()}
        save_config(self.cfg)

    def refresh_all(self):
        self._refresh_slot_lists()
        self._rebuild_preview()

    def _refresh_slot_lists(self):
        for key, title, _hint in SLOT_DEFS:
            lb = self.slot_widgets[key]["listbox"]
            lb.delete(0, "end")
            entries = self.slots[key]
            for i, e in enumerate(entries, start=1):
                cnt = e.get("pages") or 1
                kind = "PDF" if is_pdf(e["path"]) else "图片"
                lb.insert("end", "%d. %s    [%s %d 页]" % (i, truncate(os.path.basename(e["path"]), 30), kind, cnt))
            n = len(entries)
            self.slot_widgets[key]["count"].configure(
                text=("已添加 %d 个" % n) if n else "尚未添加",
                foreground=(GOLD if n else INK_3))

    def _rebuild_preview(self, keep_viewer=False):
        self._render_token += 1
        token = self._render_token

        if self._viewer is not None and not keep_viewer:
            try:
                self._viewer.destroy()
            except Exception:
                pass
            self._viewer = None

        for w in self.pv_inner.winfo_children():
            w.destroy()
        self.thumbs = []

        try:
            if self.doc is not None:
                self.doc.close()
        except Exception:
            pass
        self.doc = fitz.open()

        total = sum(len(v) for v in self.slots.values())
        if total == 0:
            self.empty_hint = tk.Label(self.pv_inner,
                                       text="把材料拖进左栏对应栏目，这里会按最终合并顺序显示每一页",
                                       background="#FFFFFF", foreground=INK_3, font=(FONT, 9))
            self.empty_hint.pack(pady=60)
            self.units = []
            self.page_order = []
            self._order_touched = False
            self.hidden = set()
            self.page_var.set("")
            self._update_restore_label()
            self._auto_output_name()
            return

        try:
            self.doc.close()
            natural = merge_units(self.slots)
            self._sync_order(natural)
            by_key = {u["key"]: u for u in natural}
            self.units = []
            for i, k in enumerate(self.page_order, start=1):
                u = by_key[k]
                self.units.append({"key": u["key"], "title": u["title"], "name": u["name"],
                                   "page_no": u["page_no"], "g": i})
            self.doc = build_document(self.slots, self.rots, self.page_order)
        except Exception as e:
            self.units = []
            tk.Label(self.pv_inner, text="预览失败：%s" % e, background="#FFFFFF",
                     foreground="#C0392B", wraplength=600).pack(pady=30)
            self.page_var.set("")
            return

        self.page_var.set("共 %d 页" % self.doc.page_count)
        self._update_restore_label()
        self._layout_thumbs(token)
        self._auto_output_name()

    def _update_restore_label(self):
        try:
            n = len(self.hidden)
            if n:
                self.restore_lbl.configure(text="已删除 %d 页 · 点此恢复" % n)
                self.restore_lbl.pack(side="right", before=self.page_chip)
            else:
                self.restore_lbl.pack_forget()
        except Exception:
            pass

    def _sync_order(self, natural):
        """把用户自定义页序与当前材料对齐

        没手动排过 → 完全按槽位自然顺序；
        排过 → 保留用户已排好的相对次序，新加入的页追加到最后（用户再用 ◀ ▶ 挪到位）。
        被单独删掉的页（self.hidden）不进页序。
        """
        nat_keys = [u["key"] for u in natural]
        nat_set = set(nat_keys)
        self.hidden = {k for k in self.hidden if k in nat_set}   # 文件被删掉后，删除记录一并失效
        if not self._order_touched:
            self.page_order = [k for k in nat_keys if k not in self.hidden]
            return
        cur = [k for k in self.page_order if k in nat_set and k not in self.hidden]
        have = set(cur)
        cur.extend([k for k in nat_keys if k not in have and k not in self.hidden])
        self.page_order = cur

    def _thumb_cols(self):
        """按预览区实际宽度决定每行几张缩略图"""
        try:
            w = self.pv_canvas.winfo_width()
        except Exception:
            w = 0
        if w <= 1:
            w = 660
        unit = THUMB_W + 26
        return max(2, min(6, int((w - 14) // unit)))

    def _layout_thumbs(self, token):
        cols = self._thumb_cols()
        grid = tk.Frame(self.pv_inner, background="#FFFFFF")
        grid.pack(anchor="n", padx=4, pady=4)
        self._cells = []

        holders = []
        for i, u in enumerate(self.units):
            title, name, page_no, g, ap = u["title"], u["name"], u["page_no"], u["g"], u["key"][0]
            r, c = divmod(i, cols)
            cell = tk.Frame(grid, background=CARD, padx=6, pady=6,
                            highlightbackground=PU_LINE, highlightthickness=1)
            cell.grid(row=r, column=c, padx=5, pady=5, sticky="n")
            ph = tk.Frame(cell, width=THUMB_W, height=int(THUMB_W * 1.35), background="#FFFFFF")
            ph.pack()
            ph.pack_propagate(False)
            tk.Label(ph, text="渲染中…", background="#FFFFFF", foreground=INK_3,
                     font=(FONT, 8)).pack(expand=True, fill="both")
            self._cells.append(cell)

            foot = tk.Frame(cell, background=CARD)
            foot.pack(fill="x", pady=(5, 0))
            cap = tk.Label(foot, text="%s · 第 %d 页" % (truncate(strip_number(title), 8), g),
                           background=CARD, foreground=PU_DARK, font=(FONT, 8, "bold"))
            cap.pack(side="left")

            rot = self.rots.get((ap, page_no), 0)
            rot_btn = tk.Label(foot, text=("↻ %d°" % rot) if rot else "↻",
                               background=CARD, foreground=PU if not rot else GOLD,
                               font=(FONT, 8, "bold"), cursor="hand2", padx=2)
            rot_btn.pack(side="right")
            self._tip(rot_btn, "把这一页顺时针转 90°（已转 %d°）" % rot)

            move_bar = tk.Frame(cell, background=CARD)
            move_bar.pack(fill="x", pady=(3, 0))
            left = tk.Button(move_bar, text="◀ 前移", command=lambda idx=i: self._move_thumb(idx, -1),
                             background=CARD, foreground=PU, activebackground=PU_TINT,
                             activeforeground=PU_DARK, relief="flat", borderwidth=0,
                             font=(FONT, 8), padx=2, pady=0, cursor="hand2")
            left.pack(side="left")
            self._tip(left, "把这一页往前挪一位（和前一页交换位置）")
            right = tk.Button(move_bar, text="后移 ▶", command=lambda idx=i: self._move_thumb(idx, 1),
                              background=CARD, foreground=PU, activebackground=PU_TINT,
                              activeforeground=PU_DARK, relief="flat", borderwidth=0,
                              font=(FONT, 8), padx=2, pady=0, cursor="hand2")
            right.pack(side="left", padx=(6, 0))
            self._tip(right, "把这一页往后挪一位（和下一页交换位置）")
            if i == 0:
                left.configure(state="disabled", foreground=INK_3)
            if i == len(self.units) - 1:
                right.configure(state="disabled", foreground=INK_3)

            sub = tk.Label(cell, text="%s p%d" % (truncate(name, 14), page_no),
                           background=CARD, foreground=INK_3, font=(FONT, 7))
            sub.pack()
            for widget in (ph, cap, sub):
                self._bind_drag(widget, i)
            rot_btn.bind("<Button-1>", lambda e, idx=i: self._rotate_thumb(idx))
            holders.append(ph)

        self._render_thumbs(token, holders)

    def _make_del_badge(self, holder, index):
        """缩略图右上角的 ✕：单独删掉这一页"""
        badge = tk.Label(holder, text=" ✕ ", background=PU_TINT, foreground=PU,
                         font=(FONT, 8, "bold"), cursor="hand2")
        badge.place(relx=1.0, x=-4, y=4, anchor="ne")
        badge.bind("<Button-1>", lambda e, idx=index: self._delete_thumb(idx))
        self._tip(badge, "从合并结果里删掉这一页（文件仍在左栏，可整体恢复）")
        badge.bind("<Enter>", lambda e: badge.configure(background=PU, foreground="#FFFFFF"))
        badge.bind("<Leave>", lambda e: badge.configure(background=PU_TINT, foreground=PU))
        return badge

    def _delete_thumb(self, index):
        if index < 0 or index >= len(self.units):
            return
        key = self.units[index]
        self.delete_page(index)
        try:
            self.status_var.set("已删掉 %s 第 %d 页" % (truncate(key["name"], 20), key["page_no"]))
        except Exception:
            pass

    def _rotate_thumb(self, index):
        self.rotate_page(index, 90)
        try:
            self.status_var.set("已把第 %d 页旋转 90°（可继续点 ↻ 转回）" % (index + 1))
        except Exception:
            pass

    def delete_page(self, index, keep_viewer=False):
        """把第 index 页从合并结果里去掉（文件仍留在左栏，可随时恢复）"""
        if index < 0 or index >= len(self.units):
            return None
        key = self.units[index]["key"]
        self.hidden.add(key)
        if key in self.page_order:
            self.page_order.remove(key)
        self._order_touched = True
        self._rebuild_preview(keep_viewer=keep_viewer)
        try:
            self.status_var.set("已从合并结果中删掉第 %d 页（文件仍在左栏，可用右上角「已删除 N 页」恢复）"
                                % (index + 1))
        except Exception:
            pass
        return key

    def restore_hidden_pages(self):
        """把单独删掉的页全部放回去"""
        if not self.hidden:
            return
        self.hidden = set()
        self._rebuild_preview()
        try:
            self.status_var.set("已恢复被删掉的页面")
        except Exception:
            pass

    # ------------------------------------------------ 缩略图拖拽换位
    def _bind_drag(self, widget, index):
        widget.bind("<ButtonPress-1>", lambda e, i=index: self._drag_begin(e, i))
        widget.bind("<B1-Motion>", self._drag_motion)
        widget.bind("<ButtonRelease-1>", lambda e, i=index: self._drag_end(e, i))

    def _drag_begin(self, event, index):
        self._drag = {"index": index, "x": event.x_root, "y": event.y_root,
                      "moved": False, "target": index}

    def _drag_motion(self, event):
        d = self._drag
        if not d:
            return
        if abs(event.x_root - d["x"]) + abs(event.y_root - d["y"]) > 5:
            d["moved"] = True
        t = self._cell_at(event.x_root, event.y_root)
        if t is not None and t != d["target"]:
            d["target"] = t
            self._paint_drag(d["index"], t)
            try:
                self.status_var.set("拖动中：第 %d 页 → 第 %d 位" % (d["index"] + 1, t + 1))
            except Exception:
                pass

    def _drag_end(self, event, index):
        d = self._drag
        self._drag = None
        self._paint_drag(None, None)
        if not d:
            return
        if not d["moved"]:
            self.open_viewer(index)
            return
        target = self._cell_at(event.x_root, event.y_root)
        if target is None or target == index:
            return
        newi = self.move_page_to(index, target)
        if newi is not None:
            try:
                self.status_var.set("第 %d 页已拖到第 %d 位（合并时按新顺序输出）"
                                    % (index + 1, newi + 1))
            except Exception:
                pass

    def _cell_at(self, x_root, y_root):
        for i, cell in enumerate(self._cells):
            try:
                x0 = cell.winfo_rootx()
                y0 = cell.winfo_rooty()
                w = cell.winfo_width()
                h = cell.winfo_height()
            except Exception:
                continue
            if x0 <= x_root <= x0 + w and y0 <= y_root <= y0 + h:
                return i
        return None

    def _paint_drag(self, drag_index, target_index):
        for i, cell in enumerate(self._cells):
            try:
                if i == target_index and target_index is not None:
                    cell.configure(highlightbackground=GOLD, highlightthickness=2)
                elif i == drag_index and drag_index is not None:
                    cell.configure(highlightbackground=PU, highlightthickness=2)
                else:
                    cell.configure(highlightbackground=PU_LINE, highlightthickness=1)
            except Exception:
                pass

    def _move_thumb(self, index, delta):
        """缩略图上的 ◀ / ▶：把这一页与相邻页交换位置"""
        if not self.units:
            return
        if delta < 0 and index == 0:
            return
        if delta > 0 and index >= len(self.units) - 1:
            return
        new_index = self.move_page(index, delta)
        if new_index is None:
            return
        try:
            self.status_var.set("第 %d 页已移到第 %d 位（合并时按新顺序输出）"
                                % (index + 1, new_index + 1))
        except Exception:
            pass

    def _render_thumbs(self, token, holders, i=0):
        if token != self._render_token:
            return
        if i >= len(holders):
            return
        try:
            page = self.doc[i]
            pw, ph = max(page.rect.width, 1), max(page.rect.height, 1)
            zoom = min(THUMB_W / pw, (THUMB_W * 1.35) / ph)      # 等比缩进缩略框
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            photo = ImageTk.PhotoImage(img)
            self.thumbs.append(photo)
            holder = holders[i]
            for w in holder.winfo_children():
                w.destroy()
            lbl = tk.Label(holder, image=photo, borderwidth=1, relief="solid", cursor="hand2")
            lbl.pack(expand=True)          # 在固定尺寸的框里居中，横竖页脚注仍对齐
            self._bind_drag(lbl, i)
            self._make_del_badge(holder, i)
        except Exception:
            pass
        self.root.after(1, lambda: self._render_thumbs(token, holders, i + 1))

    def open_viewer(self, index):
        if self.doc.page_count == 0:
            return
        try:
            self._viewer = PageViewer(self.root, self, index)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "预览失败：%s" % e)

    # ---------------- 抬头 / 输出
    def _current_entity(self):
        name = self.entity_var.get()
        for key, label, holder in ENTITY_DEFS:
            if name == label:
                return key, label, holder
        return ENTITY_DEFS[0]

    def _update_entity_hint(self):
        key, label, holder = self._current_entity()
        self.entity_hint.configure(text="收据/发票抬头请写：%s" % holder)
        self._auto_output_name()

    def copy_entity(self):
        key, label, holder = self._current_entity()
        self.root.clipboard_clear()
        self.root.clipboard_append(holder)
        self.status_var.set("已复制抬头：%s" % holder)

    def _auto_output_name(self):
        if self._dir_auto:
            self.dir_var.set(self._default_dir())
        current = self.name_var.get().strip()
        if self._name_auto and (not current or current == self._auto_name_last):
            key, _label, _holder = self._current_entity()
            current = "报销材料_%s_%s" % (key.upper(), datetime.now().strftime("%Y%m%d_%H%M"))
            self.name_var.set(current)
            self._auto_name_last = current
        self._update_target_preview()

    def _default_dir(self):
        for key, _t, _h in SLOT_DEFS:
            if self.slots[key]:
                return os.path.dirname(self.slots[key][0]["path"])
        last = self.cfg.get("last_dir")
        if last and os.path.isdir(last):
            return last
        return os.path.expanduser("~") + os.sep + "Desktop"

    def _target_path(self):
        folder = self.dir_var.get().strip().strip('"')
        name = self.name_var.get().strip().strip('"')
        if not folder or not name:
            return ""
        if not name.lower().endswith(".pdf"):
            name += ".pdf"
        return os.path.join(folder, name)

    def _update_target_preview(self):
        path = self._target_path()
        if path:
            exists = os.path.exists(path)
            self.target_var.set("将保存为：%s%s" % (path, "　（同名文件已存在，会先询问是否覆盖）" if exists else ""))
        else:
            self.target_var.set("将保存为：（请先选择文件夹并填写文件名）")

    def choose_dir(self):
        initial = self.dir_var.get().strip() or self._default_dir()
        folder = filedialog.askdirectory(
            title="选择保存文件夹",
            initialdir=initial if os.path.isdir(initial) else self._default_dir(),
            mustexist=True,
        )
        if folder:
            self._dir_auto = False
            self.dir_var.set(os.path.normpath(folder))
            self.status_var.set("保存文件夹已设置")
            self._update_target_preview()

    # ---------------- 合并
    def start_merge(self):
        if not any(self.slots[k] for k, _t, _h in SLOT_DEFS):
            messagebox.showwarning(APP_TITLE, "请先在左栏添加材料")
            return
        out = self._target_path()
        if not out:
            if not self.dir_var.get().strip():
                self.choose_dir()
            if not self.name_var.get().strip():
                self._auto_output_name()
            out = self._target_path()
            if not out:
                messagebox.showwarning(APP_TITLE, "请先选择保存文件夹并填写文件名")
                return

        missing = [k for k, _t, _h in SLOT_DEFS if not self.slots[k]]
        if missing and missing != ["extra"]:
            names = [t for k, t, _h in SLOT_DEFS if k in missing]
            if not messagebox.askyesno(
                    APP_TITLE, "以下材料还没有添加：\n%s\n\n仍要继续合并吗？" % "\n".join(names)):
                return
        if os.path.exists(out):
            if not messagebox.askyesno(APP_TITLE, "目标文件已存在，是否覆盖？\n%s" % out):
                return
        parent = os.path.dirname(out)
        if parent and not os.path.isdir(parent):
            try:
                os.makedirs(parent, exist_ok=True)
            except Exception as e:
                messagebox.showerror(APP_TITLE, "无法创建保存目录：\n%s" % e)
                return

        self.merge_btn.configure(state="disabled")
        self.progress.configure(value=0)
        self.status_var.set("正在合并…")
        snapshot = {k: [dict(e) for e in v] for k, v in self.slots.items()}
        rots = dict(self.rots)
        order = list(self.page_order)
        threading.Thread(target=self._merge_worker, args=(snapshot, out, rots, order), daemon=True).start()

    def _merge_worker(self, slots, out, rots, order=None):
        try:
            total, size = merge_to_file(slots, out, rots, order)
            self.root.after(0, self._merge_done, out, total, size)
        except Exception as e:
            detail = traceback.format_exc()
            self.root.after(0, self._merge_failed, str(e), detail)

    def _merge_done(self, out, total, size):
        self.progress.configure(value=100)
        self.merge_btn.configure(state="normal")
        self.status_var.set("合并完成：%d 页，%s" % (total, human_size(size)))
        self.cfg["last_dir"] = os.path.dirname(out)
        save_config(self.cfg)
        if messagebox.askyesno(APP_TITLE,
                               "合并完成！\n\n共 %d 页，文件大小 %s\n保存位置：\n%s\n\n要打开所在文件夹吗？"
                               % (total, human_size(size), out)):
            try:
                os.startfile(os.path.dirname(out) or out)
            except Exception:
                pass

    def _merge_failed(self, msg, detail):
        self.progress.configure(value=0)
        self.merge_btn.configure(state="normal")
        self.status_var.set("合并失败")
        messagebox.showerror(APP_TITLE, "合并失败：\n%s" % msg)
        try:
            log = os.path.join(tempfile.gettempdir(), "reimburse_merge_error.log")
            with open(log, "a", encoding="utf-8") as f:
                f.write(detail + "\n")
        except Exception:
            pass


# ---------------------------------------------------------------- 自检
def _selftest():
    tmp = tempfile.mkdtemp()
    pdf1 = os.path.join(tmp, "订单截图.pdf")
    d = fitz.open()
    for _ in range(2):
        d.new_page(width=A4_W, height=A4_H)
    d.save(pdf1)
    d.close()

    img_path = os.path.join(tmp, "支付记录.png")
    Image.new("RGB", (900, 1600), (235, 240, 250)).save(img_path)

    slots = {"order": [{"path": pdf1}], "pay": [{"path": img_path}],
             "receipt": [], "extra": []}

    root = create_root()
    root.withdraw()
    app = App(root)
    app.slots = slots
    app.refresh_all()
    root.update()

    # 模拟一次拖放：把一个文件夹拖进 ④ 补充材料
    drop_dir = os.path.join(tmp, "拖入文件夹")
    os.makedirs(drop_dir, exist_ok=True)
    dropped = os.path.join(drop_dir, "邀请邮件.pdf")
    d = fitz.open()
    d.new_page(width=A4_W, height=A4_H)
    d.save(dropped)
    d.close()

    out = os.path.join(tmp, "合并结果.pdf")
    total, size = merge_to_file(slots, out)
    back = fitz.open(out).page_count
    thumb_ok = len(app.thumbs) >= 1
    units = merge_units(slots)

    # 旋转：把第一页顺时针转 90°，预览与导出文件都应带上旋转
    app._save_rots = lambda: None          # 自检不写用户配置
    app.rotate_page(0, 90)
    root.update()
    rot_preview = app.doc[0].rotation
    out2 = os.path.join(tmp, "合并结果_旋转.pdf")
    total2, _size2 = merge_to_file(slots, out2, app.rots)
    d2 = fitz.open(out2)
    rot_merged = d2[0].rotation
    d2.close()

    # 页序：把最后一页（支付记录图片）移到最前，预览与导出都应跟着变
    last_key = app.page_order[-1]
    app.move_page_to(len(app.page_order) - 1, 0)
    root.update()
    order_preview_ok = app.units[0]["key"] == last_key and app.doc.page_count == 3
    out3 = os.path.join(tmp, "合并结果_改序.pdf")
    _t3, _s3 = merge_to_file(slots, out3, app.rots, app.page_order)
    d3 = fitz.open(out3)
    first_is_image = d3[0].rect.height > 900          # 图片页（595×1058）挪到了第 1 页
    first_page_rot = d3[0].rotation
    d3.close()
    order_merged_ok = first_is_image and first_page_rot == 0

    # 再加一个文件到 ④：手动排过序后，新页应追加到最后，不打乱已排好的顺序
    import types
    app._on_drop("extra", types.SimpleNamespace(data="{%s}" % drop_dir))
    order = list(app.page_order)
    new_ok = len(order) == 4 and order[-1][0] == dropped and order[0] == last_key

    drop_ok = len(app.slots["extra"]) == 1 and app.slots["extra"][0]["path"] == dropped
    dnd_ok = bool(app.dnd_ready) or not HAS_DND

    # 单页删除：删掉第 2 页后预览与导出都少一页；恢复后回到 4 页
    victim = app.units[1]["key"]
    app.delete_page(1)
    root.update()
    del_preview_ok = (len(app.units) == 3 and victim in app.hidden
                      and app.doc.page_count == 3 and victim not in app.page_order)
    out4 = os.path.join(tmp, "合并结果_删页.pdf")
    _t4, _s4 = merge_to_file(slots, out4, app.rots, app.page_order)
    d4 = fitz.open(out4)
    del_merged_ok = d4.page_count == 3
    d4.close()
    app.restore_hidden_pages()
    root.update()
    restore_ok = len(app.units) == 4 and not app.hidden

    root.destroy()

    ok = (total == 3 and back == 3 and thumb_ok and size > 0 and drop_ok and dnd_ok
          and rot_preview == 90 and rot_merged == 90
          and order_preview_ok and order_merged_ok and new_ok
          and del_preview_ok and del_merged_ok and restore_ok
          and app._target_path().lower().endswith(".pdf"))
    with open(os.path.join(tempfile.gettempdir(), "merger_selftest.txt"), "w", encoding="utf-8") as f:
        f.write("ok=%s total=%s read_back=%s thumbs=%s drop=%s dnd=%s(has=%s) rot=%s/%s "
                "order=%s/%s insert=%s delete=%s/%s restore=%s page=%d size=%s target=%s\n"
                % (ok, total, back, len(app.thumbs), drop_ok, dnd_ok, HAS_DND,
                   rot_preview, rot_merged, order_preview_ok, order_merged_ok, new_ok,
                   del_preview_ok, del_merged_ok, restore_ok,
                   len(app.units), size, app._target_path()))


def create_root():
    """有 tkdnd 时用支持拖放的 Tk，否则退回普通 Tk"""
    if HAS_DND:
        try:
            return TkinterDnD.Tk()
        except Exception:
            pass
    return tk.Tk()


def main():
    if "--selftest" in sys.argv:
        _selftest()
        return
    try:
        fitz.TOOLS.mupdf_display_errors(False)
    except Exception:
        pass
    root = create_root()
    try:
        root.call("tk", "scaling", 1.2)
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()



