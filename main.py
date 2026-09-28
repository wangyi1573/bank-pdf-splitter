# -*- coding: utf-8 -*-
"""
银行 PDF 回单拆分工具（GUI）

用法：python main.py，或直接运行打包后的 exe。

布局：左侧功能区（选文件 / 解析 / 客户名匹配 / 导出）+ 右侧 PDF 预览主区。
预览区支持翻页、缩放，以及「划线拆分」——一页 A4 打印了 2/3 张回单时，
在预览图上画出分割线，把每张回单裁切为独立的 PDF 页。
详细设计见 SPEC.md 与代码内注释。
"""

from __future__ import annotations

import csv
import json
import logging
import multiprocessing
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import ttk, filedialog, messagebox

import receipt_parser as rp

try:
    from tkinterdnd2 import TkinterDnD
    _HAS_DND = True
except ImportError:
    _HAS_DND = False

APP_NAME = '银行PDF回单拆分工具'
APP_DIR = Path.home() / '.bank_pdf_splitter'
CONFIG_PATH = APP_DIR / 'config.json'
LOG_DIR = APP_DIR / 'logs'

QUERY_SPLIT_RE = re.compile(r'[,，;；、\n\r\t]+')

UI_FONT = '微软雅黑'
SIDE_W = 380
LINE_COLOR = '#e03131'


def setup_logging() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger('bank_pdf_splitter')
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
    fh = logging.FileHandler(LOG_DIR / f'{datetime.now():%Y-%m-%d}.log', encoding='utf-8')
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    return logger


LOG = setup_logging()


def load_config() -> dict:
    try:
        if CONFIG_PATH.exists():
            return json.loads(CONFIG_PATH.read_text(encoding='utf-8'))
    except Exception:
        LOG.warning('读取配置失败\n%s', traceback.format_exc())
    return {}


def save_config(cfg: dict):
    try:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception:
        LOG.warning('保存配置失败\n%s', traceback.format_exc())


def open_folder(path: str):
    try:
        if os.name == 'nt':
            os.startfile(path)  # noqa: S606
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', path])
        else:
            subprocess.Popen(['xdg-open', path])
    except Exception:
        LOG.warning('打开目录失败: %s', path)


def parse_queries(raw: str) -> list:
    """把输入框内容拆成多个查询词（去重保序）。"""
    parts = [p.strip() for p in QUERY_SPLIT_RE.split(raw or '')]
    out, seen = [], set()
    for p in parts:
        if p and p not in seen:
            seen.add(p)
            out.append(p)
    return out


_AppBase = TkinterDnD.Tk if _HAS_DND else tk.Tk


class App(_AppBase):

    def __init__(self):
        super().__init__()
        self.cfg = load_config()

        self.title(APP_NAME)
        self.geometry('1280x860')
        self.minsize(1024, 680)

        self.src_pdf_path = tk.StringVar()
        self.save_dir = tk.StringVar(value=self.cfg.get('last_dir', ''))
        self.query_text = tk.StringVar()
        self.status_text = tk.StringVar(value='就绪')
        self.merge_one = tk.BooleanVar(value=False)
        self.loose_match = tk.BooleanVar(value=True)
        self.avoid_overwrite = tk.BooleanVar(value=True)
        self.ocr_enabled = tk.BooleanVar(value=bool(self.cfg.get('ocr_enabled', False)))
        self.ocr_info = tk.StringVar(value='')
        self.filter_text = tk.StringVar()

        # 解析 / 匹配状态
        self.pages: list = []
        self.texts: list = []
        self._parsed_sig = None         # 已解析文件的签名，用于缓存失效判断
        self._candidates: list = []
        self._filtered: list = []
        self.matches: list = []
        self.busy = False
        self.parse_info = {'total': 0, 'empty_pages': 0, 'ocr_pages': 0, 'ocr_used': False}

        # 预览 / 划线状态
        self._pdoc = None               # 预览用的 PyMuPDF 文档（与解析独立）
        self._preview_sig = None
        self._page_no = 1               # 当前预览页（1 基）
        self._photo = None              # 防 PhotoImage 被 GC
        self._lines: dict = {}          # {1基页码: [0~1 分割线比例]}
        self._line_mode = False
        self._zoom_mode = tk.StringVar(value='适应宽度')
        self._render_after = None
        self.split_all_pages = tk.BooleanVar(value=False)

        self._q = queue.Queue()
        self._build_ui()
        self._poll_queue()

        if _HAS_DND:
            self._setup_dnd()
        self._probe_ocr_async()
        self.protocol('WM_DELETE_WINDOW', self._on_close)

        LOG.info('启动 %s', APP_NAME)

    # ── 界面 ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use('vista')
        except Exception:
            pass
        style.configure('TButton', font=(UI_FONT, 10))
        style.configure('TLabel', font=(UI_FONT, 10))
        style.configure('TCheckbutton', font=(UI_FONT, 9))
        style.configure('Treeview', font=(UI_FONT, 9), rowheight=24)
        style.configure('Treeview.Heading', font=(UI_FONT, 9, 'bold'))
        style.configure('TEntry', font=(UI_FONT, 10), padding=3)
        style.configure('TLabelframe.Label', font=(UI_FONT, 9, 'bold'))

        # 顶部标题条
        head = tk.Frame(self, bg='#0f4c9e', height=46)
        head.pack(fill='x')
        head.pack_propagate(False)
        tk.Label(head, text=APP_NAME, font=(UI_FONT, 14, 'bold'),
                 fg='white', bg='#0f4c9e').pack(side='left', padx=20)
        tk.Label(head, text='按客户拆分 · 划线拆分一页多单', font=(UI_FONT, 9),
                 fg='#bcd6f5', bg='#0f4c9e').pack(side='left')

        # 主区：左功能区 + 右预览
        body = tk.Frame(self)
        body.pack(fill='both', expand=True)
        self._build_sidebar(body)
        self._build_preview(body)

        # 底部状态栏
        foot = ttk.Frame(self, padding=(10, 4, 10, 6))
        foot.pack(fill='x')
        self.progress = ttk.Progressbar(foot, mode='determinate', maximum=100, length=220)
        self.progress.pack(side='left')
        ttk.Label(foot, textvariable=self.status_text, foreground='#333')\
            .pack(side='left', padx=(10, 0))

    # ── 左侧功能区 ──────────────────────────────────────────────────────────

    def _build_sidebar(self, parent):
        side = ttk.Frame(parent, width=SIDE_W)
        side.pack(side='left', fill='y', padx=(8, 4), pady=4)
        side.pack_propagate(False)
        side.columnconfigure(0, weight=1)

        # 1. 源文件
        box1 = ttk.LabelFrame(side, text=' 1. 选择源 PDF ', padding=8)
        box1.grid(row=0, column=0, sticky='ew')
        box1.columnconfigure(0, weight=1)
        ttk.Entry(box1, textvariable=self.src_pdf_path).grid(row=0, column=0, sticky='ew')
        ttk.Button(box1, text='选择文件', width=9, command=self.select_pdf)\
            .grid(row=0, column=1, padx=(6, 0))
        self.drop_label = tk.Label(
            box1, text='⬇ 把 PDF 拖到这里', font=(UI_FONT, 9),
            fg='#5a6b7d', bg='#eef4fc', relief='solid', borderwidth=1, pady=5)
        self.drop_label.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(6, 0))

        # 2. 解析与选项
        box2 = ttk.LabelFrame(side, text=' 2. 解析回单 ', padding=8)
        box2.grid(row=1, column=0, sticky='ew', pady=(8, 0))
        box2.columnconfigure(0, weight=1)
        row = ttk.Frame(box2)
        row.grid(row=0, column=0, sticky='ew')
        self.btn_parse = ttk.Button(row, text='① 解析回单', width=11, command=self.do_parse)
        self.btn_parse.pack(side='left')
        ttk.Checkbutton(row, text='扫描件启用 OCR', variable=self.ocr_enabled).pack(side='left', padx=(10, 0))
        self.ocr_label = ttk.Label(box2, textvariable=self.ocr_info,
                                   foreground='#999', font=(UI_FONT, 8))
        self.ocr_label.grid(row=1, column=0, sticky='w', pady=(2, 0))
        opts = ttk.Frame(box2)
        opts.grid(row=2, column=0, sticky='w', pady=(4, 0))
        ttk.Checkbutton(opts, text='宽松匹配', variable=self.loose_match,
                        command=self._on_match_option_changed).pack(side='left')
        ttk.Checkbutton(opts, text='合并导出', variable=self.merge_one).pack(side='left', padx=(10, 0))

        # 3. 客户名称
        box3 = ttk.LabelFrame(side, text=' 3. 客户名称或页码（企业/个人，第3页/p5，可多个） ', padding=8)
        box3.grid(row=2, column=0, sticky='ew', pady=(8, 0))
        box3.columnconfigure(0, weight=1)
        qrow = ttk.Frame(box3)
        qrow.grid(row=0, column=0, sticky='ew')
        qrow.columnconfigure(0, weight=1)
        ttk.Entry(qrow, textvariable=self.query_text).grid(row=0, column=0, sticky='ew')
        self.btn_preview = ttk.Button(qrow, text='② 预览', width=7, command=self.do_preview)
        self.btn_preview.grid(row=0, column=1, padx=(6, 0))

        cand = ttk.Frame(box3)
        cand.grid(row=1, column=0, sticky='ew', pady=(6, 0))
        cand.columnconfigure(0, weight=1)
        ch = ttk.Frame(cand)
        ch.grid(row=0, column=0, sticky='ew')
        ttk.Label(ch, text='候选客户：', foreground='#444').pack(side='left')
        self.cand_hint = ttk.Label(ch, text='尚未解析', foreground='#999')
        self.cand_hint.pack(side='left', padx=(4, 0))
        ttk.Entry(ch, textvariable=self.filter_text, width=10).pack(side='right')
        ttk.Label(ch, text='筛选:', foreground='#444').pack(side='right')
        self.filter_text.trace_add('write', lambda *_: self._refresh_candidates())
        self.cand_list = tk.Listbox(cand, selectmode='extended', height=6,
                                    font=(UI_FONT, 9), activestyle='none',
                                    exportselection=False,
                                    selectbackground='#cfe3fb', selectforeground='#123')
        self.cand_list.grid(row=1, column=0, sticky='ew', pady=(2, 0))
        # 候选可能上千条（一页含收付双方 + 长周期导出），配滚动条与滚轮
        cand_scroll = ttk.Scrollbar(cand, orient='vertical', command=self.cand_list.yview)
        cand_scroll.grid(row=1, column=1, sticky='ns', pady=(2, 0))
        self.cand_list.config(yscrollcommand=cand_scroll.set)
        self.cand_list.bind('<MouseWheel>',
                            lambda e: self.cand_list.yview_scroll(-1 * (e.delta // 120), 'units'))
        self.cand_list.bind('<Double-Button-1>', self._on_cand_double_click)
        ttk.Button(cand, text='加入名称框', width=9, command=self._use_selected)\
            .grid(row=2, column=0, sticky='w', pady=(3, 0))

        # 4. 匹配结果
        box4 = ttk.LabelFrame(side, text=' 4. 匹配预览 ', padding=8)
        box4.grid(row=3, column=0, sticky='nsew', pady=(8, 0))
        box4.columnconfigure(0, weight=1)
        box4.rowconfigure(0, weight=1)
        side.rowconfigure(3, weight=1)

        cols = ('page', 'query', 'party', 'date', 'amount')
        self.tree = ttk.Treeview(box4, columns=cols, show='headings', height=7)
        for cid, text, w, anchor in (
            ('page', '页码', 58, 'center'),
            ('query', '归属', 96, 'w'),
            ('party', '命中主体', 120, 'w'),
            ('date', '日期', 82, 'center'),
            ('amount', '金额', 76, 'e'),
        ):
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=w, anchor=anchor, stretch=(cid == 'party'))
        self.tree.grid(row=0, column=0, sticky='nsew')
        tsb = ttk.Scrollbar(box4, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=tsb.set)
        tsb.grid(row=0, column=1, sticky='ns')
        self.tree.bind('<Double-Button-1>', self._open_selected_page)

        # 操作区
        act = ttk.Frame(side)
        act.grid(row=4, column=0, sticky='ew', pady=(8, 0))
        act.columnconfigure(0, weight=1)
        self.btn_export = ttk.Button(act, text='③ 导出客户 PDF', command=self.do_export)
        self.btn_export.grid(row=0, column=0, sticky='ew')
        r2 = ttk.Frame(act)
        r2.grid(row=1, column=0, sticky='ew', pady=(4, 0))
        ttk.Button(r2, text='导出清单', command=self.do_export_csv).pack(side='left', expand=True, fill='x')
        ttk.Button(r2, text='打开目录', command=self._open_out_dir).pack(side='left', expand=True, fill='x', padx=(4, 0))
        ttk.Button(r2, text='清空', width=6, command=self.do_clear).pack(side='left', padx=(4, 0))

    # ── 右侧预览区 ──────────────────────────────────────────────────────────

    def _build_preview(self, parent):
        pane = ttk.Frame(parent)
        pane.pack(side='left', fill='both', expand=True, padx=(4, 8), pady=4)
        pane.rowconfigure(2, weight=1)
        pane.columnconfigure(0, weight=1)

        # 翻页 / 缩放
        nav = ttk.Frame(pane)
        nav.grid(row=0, column=0, sticky='ew')
        ttk.Button(nav, text='◀ 上一页', width=9, command=lambda: self._step_page(-1)).pack(side='left')
        self.page_label = ttk.Label(nav, text='－ / －', font=(UI_FONT, 10, 'bold'))
        self.page_label.pack(side='left', expand=True)
        ttk.Button(nav, text='下一页 ▶', width=9, command=lambda: self._step_page(1)).pack(side='left')
        ttk.Combobox(nav, textvariable=self._zoom_mode, width=8, state='readonly',
                     values=('适应宽度', '50%', '75%', '100%', '150%', '200%'))\
            .pack(side='right')

        # 划线工具条
        tools = ttk.LabelFrame(pane, text=' 划线拆分（一页多张回单时，画出分割线） ', padding=6)
        tools.grid(row=1, column=0, sticky='ew', pady=(6, 0))
        self.btn_line_mode = ttk.Button(tools, text='✂ 划线模式', width=10,
                                        command=self._toggle_line_mode)
        self.btn_line_mode.pack(side='left')
        ttk.Button(tools, text='撤销', width=5, command=self._undo_line).pack(side='left', padx=(4, 0))
        ttk.Button(tools, text='清空本页', width=8, command=self._clear_lines).pack(side='left', padx=(4, 0))
        ttk.Separator(tools, orient='vertical').pack(side='left', fill='y', padx=(8, 8))
        ttk.Label(tools, text='均分:').pack(side='left')
        for n in (2, 3, 4):
            ttk.Button(tools, text=f'{n}份', width=4,
                       command=lambda n=n: self._even_split(n)).pack(side='left', padx=(3, 0))
        ttk.Button(tools, text='复制到所有页', command=self._copy_lines_to_all).pack(side='left', padx=(8, 0))
        ttk.Checkbutton(tools, text='拆分时应用到所有页', variable=self.split_all_pages)\
            .pack(side='left', padx=(8, 0))
        self.btn_split = ttk.Button(tools, text='⚡ 拆分导出', command=self.do_split_export)
        self.btn_split.pack(side='right')
        self.line_hint = ttk.Label(tools, text='', foreground='#666', font=(UI_FONT, 8))
        self.line_hint.pack(side='right', padx=(0, 8))

        # 画布
        cf = ttk.Frame(pane)
        cf.grid(row=2, column=0, sticky='nsew', pady=(6, 0))
        cf.rowconfigure(0, weight=1)
        cf.columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(cf, bg='#3a3f44', highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky='nsew')
        vsb = ttk.Scrollbar(cf, orient='vertical', command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky='ns')
        hsb = ttk.Scrollbar(cf, orient='horizontal', command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=hsb.set)
        hsb.grid(row=1, column=0, sticky='ew')

        self.canvas.bind('<Button-1>', self._on_canvas_click)
        self.canvas.bind('<Button-3>', self._on_canvas_rightclick)
        self.canvas.bind('<Configure>', self._on_canvas_resize)
        self.canvas.bind('<Left>', lambda e: self._step_page(-1))
        self.canvas.bind('<Right>', lambda e: self._step_page(1))

    # ── PDF 预览 ────────────────────────────────────────────────────────────

    def _page_count(self) -> int:
        return self._pdoc.page_count if self._pdoc else 0

    def _step_page(self, delta):
        if not self._pdoc:
            return
        self._show_page(self._page_no + delta)

    def _show_page(self, pno: int):
        # 文档在 _render_preview 中懒加载，这里不能依赖 _pdoc 是否存在
        total = self._page_count()
        self._page_no = max(1, min(total, pno)) if total > 0 else max(1, pno)
        self._render_preview()

    def _zoom_factor(self, page_w: float) -> float:
        mode = self._zoom_mode.get()
        if mode == '适应宽度':
            cw = max(self.canvas.winfo_width(), 200) - 24
            return max(0.05, cw / page_w)
        try:
            return max(0.05, int(mode.rstrip('%')) / 100.0)
        except ValueError:
            return 1.0

    def _render_preview(self):
        """渲染当前页到位图，并叠加当前页的分割线。"""
        self._render_after = None
        self.canvas.delete('all')
        self._photo = None
        path = self.src_pdf_path.get().strip()
        if not path or not os.path.isfile(path):
            self.page_label.config(text='－ / －')
            self.canvas.create_text(400, 260, text='在左侧选择 / 拖入 PDF 后在此预览',
                                    fill='#c8cdd2', font=(UI_FONT, 12))
            self._update_line_hint()
            return
        try:
            import pymupdf
            sig = self._file_sig(path)
            if self._pdoc is None or sig != self._preview_sig:
                if self._pdoc:
                    self._pdoc.close()
                self._pdoc = pymupdf.open(path)
                self._preview_sig = sig
                self._page_no = 1
            total = self._page_count()
            if total == 0:
                raise ValueError('PDF 不含页面')
            self._page_no = max(1, min(total, self._page_no))
            page = self._pdoc[self._page_no - 1]

            z = self._zoom_factor(page.rect.width)
            pix = page.get_pixmap(matrix=pymupdf.Matrix(z, z), alpha=False)
            self._photo = tk.PhotoImage(data=pix.tobytes('ppm'))
            w, h = pix.width, pix.height
            self.canvas.create_image(0, 0, anchor='nw', image=self._photo)
            self.canvas.configure(scrollregion=(0, 0, w, h + 8))

            # 叠加分割线
            for f in sorted(self._lines.get(self._page_no, [])):
                y = f * h
                self.canvas.create_line(0, y, w, y, fill=LINE_COLOR, width=2,
                                        dash=(6, 3), tags='splitline')
            self.page_label.config(text=f'第 {self._page_no} / {total} 页')
        except Exception as ex:
            self.page_label.config(text='－ / －')
            self.canvas.create_text(400, 260, text=f'预览失败：{ex}',
                                    fill='#ffb4b4', font=(UI_FONT, 11))
            LOG.warning('预览失败: %s', ex)
        self._update_line_hint()

    def _on_canvas_resize(self, _e=None):
        if self._zoom_mode.get() != '适应宽度' or not self._pdoc:
            return
        # 防抖：拖动窗口时避免连续重渲染
        if self._render_after:
            self.after_cancel(self._render_after)
        self._render_after = self.after(150, self._render_preview)

    def _update_line_hint(self):
        n = len(self._lines.get(self._page_no, []))
        self.line_hint.config(text=f'本页 {n} 条线' if n else '')

    # ── 划线交互 ────────────────────────────────────────────────────────────

    def _toggle_line_mode(self):
        self._line_mode = not self._line_mode
        self.btn_line_mode.config(text='✂ 划线中…' if self._line_mode else '✂ 划线模式')
        self.canvas.config(cursor='crosshair' if self._line_mode else '')
        self._set_status('划线模式：在预览图上单击添加分割线，右键点线删除；再点一次按钮退出'
                         if self._line_mode else '已退出划线模式')

    def _img_geometry(self) -> int:
        """返回预览图在画布上的高度（0 表示无图）。"""
        return self._photo.height() if self._photo else 0

    def _on_canvas_click(self, e):
        if not self._line_mode or not self._photo:
            return
        h = self._img_geometry()
        f = self.canvas.canvasy(e.y) / h
        if 0.0 < f < 1.0:
            lines = self._lines.setdefault(self._page_no, [])
            if not any(abs(f - x) < 0.004 for x in lines):
                lines.append(f)
                self._render_preview()

    def _on_canvas_rightclick(self, e):
        if not self._photo:
            return
        lines = self._lines.get(self._page_no, [])
        if not lines:
            return
        h = self._img_geometry()
        y = self.canvas.canvasy(e.y)
        nearest = min(lines, key=lambda f: abs(f * h - y))
        if abs(nearest * h - y) <= 8:
            lines.remove(nearest)
            if not lines:
                self._lines.pop(self._page_no, None)
            self._render_preview()

    def _undo_line(self):
        lines = self._lines.get(self._page_no, [])
        if lines:
            lines.pop()
            if not lines:
                self._lines.pop(self._page_no, None)
            self._render_preview()

    def _clear_lines(self):
        if self._lines.pop(self._page_no, None) is not None:
            self._render_preview()

    def _even_split(self, n: int):
        self._lines[self._page_no] = [i / n for i in range(1, n)]
        self._render_preview()
        self._set_status(f'第 {self._page_no} 页已均分为 {n} 份；可进入划线模式微调（右键删除后重画）')

    def _copy_lines_to_all(self):
        fracs = self._lines.get(self._page_no)
        if not fracs:
            messagebox.showinfo('提示', '请先在当前页添加分割线')
            return
        total = self._page_count()
        for p in range(1, total + 1):
            self._lines[p] = list(fracs)
        self._set_status(f'已把 {len(fracs)} 条分割线复制到全部 {total} 页')
        self._render_preview()

    def do_split_export(self):
        """按分割线把页面裁切导出为新 PDF。"""
        if self.busy:
            return
        pdf_path = self.src_pdf_path.get().strip()
        if not pdf_path or not os.path.isfile(pdf_path):
            messagebox.showwarning('提示', '请先选择源 PDF 文件')
            return
        all_pages = bool(self.split_all_pages.get())
        if all_pages:
            cur = self._lines.get(self._page_no)
            if not cur:
                messagebox.showwarning('提示', '当前页还没有分割线，无法应用到所有页')
                return
            page_lines = {p: list(cur) for p in range(1, self._page_count() + 1)}
        else:
            page_lines = {p: fs for p, fs in self._lines.items() if fs}
            if not page_lines:
                messagebox.showwarning('提示', '请先在预览图上添加分割线（或点均分按钮）')
                return

        out_dir = self._output_dir()
        if not out_dir:
            messagebox.showwarning('提示', '请指定输出目录')
            return
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception as ex:
            messagebox.showerror('错误', f'无法创建输出目录：\n{ex}')
            return

        base = rp.sanitize_filename(os.path.splitext(os.path.basename(pdf_path))[0], '回单')
        if all_pages:
            name = f'{base}__全文档拆分.pdf'
        elif len(page_lines) == 1:
            name = f'{base}__拆分_第{next(iter(page_lines))}页.pdf'
        else:
            pages = '、'.join(str(p) for p in sorted(page_lines))
            name = f'{base}__拆分_第{pages}页.pdf'
        target = os.path.join(out_dir, name)
        overwrite = not bool(self.avoid_overwrite.get())

        self._set_busy(True)
        self._set_status('正在拆分导出…')

        def work():
            actual, seg = rp.split_pdf_by_lines(pdf_path, page_lines, target, overwrite=overwrite)
            self._q.put(('exported_msg', f'拆分完成：{seg} 段已导出\n\n{actual}'))

        self._run_async(work, keep_busy_until_msg=True)

    def _on_split_done(self, msg):
        self._set_busy(False)
        first = msg.splitlines()[0]
        self._set_status(f'✅ {first}')
        target = msg.strip().splitlines()[-1]
        LOG.info('划线拆分: %s', target)
        if messagebox.askyesno('拆分完成', msg + '\n\n是否打开输出目录？'):
            open_folder(os.path.dirname(target))

    # ── OCR ─────────────────────────────────────────────────────────────────

    def _probe_ocr_async(self):
        """启动时后台探测 OCR 引擎，界面显示可用性（不阻塞窗口出现）。"""
        def work():
            name, desc = rp.ocr_probe()
            if name:
                self._q.put(('ocrinfo', (f'OCR：{desc} 可用', '#2e7d32')))
            else:
                self._q.put(('ocrinfo', ('OCR 未安装（pip install rapidocr_onnxruntime 启用）',
                                         '#b26a00')))
            LOG.info('OCR 探测: %s', name or 'none')

        threading.Thread(target=work, daemon=True).start()

    # ── 异步基础设施 ─────────────────────────────────────────────────────────

    def _poll_queue(self):
        """主线程轮询后台任务消息，保证所有 UI 操作都在主线程执行。"""
        try:
            while True:
                kind, payload = self._q.get_nowait()
                if kind == 'status':
                    self.status_text.set(payload)
                elif kind == 'progress':
                    done, total = payload
                    self.progress['value'] = (done / total * 100) if total else 0
                elif kind == 'ocrinfo':
                    msg, color = payload
                    # ocr_label 绑定了 textvariable，config(text=...) 会被变量覆盖，必须改变量本身
                    self.ocr_info.set(msg)
                    self.ocr_label.config(foreground=color)
                elif kind == 'parsed':
                    self._on_parsed(payload)
                elif kind == 'preview':
                    self._on_preview_ready(payload)
                elif kind == 'exported':
                    self._on_export_done(payload)
                elif kind == 'exported_msg':
                    self._on_split_done(payload)
                elif kind == 'busy':
                    self._set_busy(bool(payload))
                elif kind == 'error':
                    self._set_busy(False)
                    messagebox.showerror('出错了', payload)
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)

    def _run_async(self, fn, keep_busy_until_msg: bool = False):
        """在后台线程执行 fn（fn 内部通过 self._q 汇报进度）。

        keep_busy_until_msg=True 时（如拆分导出），完成后由对应消息处理器解除忙碌；
        出错路径始终会通过 error 消息解除忙碌。
        """
        def runner():
            try:
                fn()
            except rp.PDFEncryptedError as ex:
                self._q.put(('error', f'{ex}\n\n请用 Adobe Acrobat / WPS 等工具移除密码后重试。'))
                LOG.warning('%s', ex)
            except rp.OCRUnavailableError as ex:
                self._q.put(('error', str(ex)))
                LOG.warning('%s', ex)
            except rp.PDFUnreadableError as ex:
                self._q.put(('error', f'{ex}'))
                LOG.error('%s\n%s', ex, traceback.format_exc())
            except Exception as ex:  # noqa: BLE001
                self._q.put(('error', f'发生未预期的错误：\n{ex}'))
                LOG.error('后台任务异常\n%s', traceback.format_exc())
            finally:
                if not keep_busy_until_msg:
                    self._q.put(('busy', False))

        threading.Thread(target=runner, daemon=True).start()

    def _set_busy(self, busy: bool):
        self.busy = busy
        state = 'disabled' if busy else 'normal'
        for b in (self.btn_parse, self.btn_preview, self.btn_export, self.btn_split):
            b.configure(state=state)
        if not busy:
            self.progress['value'] = 0

    # ── 文件选择 ────────────────────────────────────────────────────────────

    def _file_sig(self, path: str):
        try:
            st = os.stat(path)
            return (os.path.abspath(path), st.st_size, int(st.st_mtime))
        except OSError:
            return None

    def _set_source(self, path: str):
        self.src_pdf_path.set(path)
        d = self.cfg.get('last_dir')
        self.save_dir.set(d if d else os.path.dirname(path))
        # 换了文件 → 旧解析结果必须失效，否则会用上一个 PDF 的页码拆分
        if self._parsed_sig and self._parsed_sig != self._file_sig(path):
            self._reset_parse_state()
        self._set_status(f'已选择：{os.path.basename(path)}　（请点「① 解析回单」）')
        LOG.info('选择源文件: %s', path)
        self._lines = {}
        self._show_page(1)

    def _reset_parse_state(self):
        self.pages = []
        self.texts = []
        self._parsed_sig = None
        self._candidates = []
        self._filtered = []
        self.matches = []
        self.parse_info = {'total': 0, 'empty_pages': 0, 'ocr_pages': 0, 'ocr_used': False}
        self.cand_list.delete(0, 'end')
        self.cand_hint.config(text='尚未解析', foreground='#999')
        self._clear_tree()

    def select_pdf(self):
        path = filedialog.askopenfilename(
            title='选择银行回单 PDF',
            filetypes=[('PDF 文件', '*.pdf'), ('所有文件', '*.*')])
        if path:
            self._set_source(path)

    def select_dir(self):
        init = self.save_dir.get()
        if not init or not os.path.isdir(init):
            init = os.path.expanduser('~')
        path = filedialog.askdirectory(title='选择输出目录', initialdir=init)
        if path:
            self.save_dir.set(path)
            self.cfg['last_dir'] = path
            save_config(self.cfg)

    def _setup_dnd(self):
        self.drop_label.drop_target_register('DND_Files')
        self.drop_label.dnd_bind('<<Drop>>', self._on_drop)
        self.drop_label.dnd_bind('<<DragEnter>>', self._on_drag_enter)
        self.drop_label.dnd_bind('<<DragLeave>>', self._on_drag_leave)

    def _on_drag_enter(self, e=None):
        self.drop_label.config(bg='#cfe3fb', fg='#0f4c9e')

    def _on_drag_leave(self, e=None):
        self.drop_label.config(bg='#eef4fc', fg='#5a6b7d')

    def _on_drop(self, e=None):
        self._on_drag_leave()
        raw = getattr(e, 'data', '') or ''
        paths = self._parse_dnd_data(raw)
        if not paths:
            return
        pdfs = [p for p in paths if p.lower().endswith('.pdf') and os.path.isfile(p)]
        if not pdfs:
            messagebox.showwarning('提示', '请拖入 PDF 文件（.pdf）')
            return
        if len(pdfs) > 1:
            messagebox.showinfo('提示', f'检测到 {len(pdfs)} 个文件，将使用第一个：\n{os.path.basename(pdfs[0])}')
        self._set_source(pdfs[0])

    @staticmethod
    def _parse_dnd_data(data: str) -> list:
        """解析拖拽数据（花括号包裹的多路径 / 单路径 / 引号包裹）。"""
        if not data:
            return []
        paths = re.findall(r'\{([^}]+)\}', data)
        if paths:
            return [p.strip() for p in paths if p.strip()]
        return [p.strip().strip('"').strip("'") for p in data.split() if p.strip().strip('"').strip("'")]

    # ── 解析 ────────────────────────────────────────────────────────────────

    def do_parse(self):
        if self.busy:
            return
        pdf_path = self.src_pdf_path.get().strip()
        if not pdf_path or not os.path.isfile(pdf_path):
            messagebox.showwarning('提示', '请先选择源 PDF 文件')
            return
        sig = self._file_sig(pdf_path)
        if sig and sig == self._parsed_sig and self.pages:
            self._set_status(f'已使用缓存结果（共 {len(self.pages)} 页），无需重复解析')
            return

        self._set_busy(True)
        use_ocr = bool(self.ocr_enabled.get())
        self._set_status('正在提取 PDF 文本…' + ('（含扫描件 OCR，耗时较长）' if use_ocr else ''))
        self.progress['value'] = 0

        def work():
            texts, pages, info = rp.parse_document(
                pdf_path, ocr=use_ocr,
                progress=lambda d, t: self._q.put(('progress', (d, t))),
                status=lambda m: self._q.put(('status', m)))
            if info['ocr_used']:
                missed = info['empty_pages'] - info['ocr_pages']
                tip = (f'OCR 完成：识别 {info["ocr_pages"]}/{info["empty_pages"]} 页扫描页。'
                       if info['ocr_pages'] else 'OCR 未识别出文本。')
                if missed:
                    tip += f'（{missed} 页仍未识别，可能清晰度不足）'
                self._q.put(('status', tip))
            if info.get('positional_fixed'):
                self._q.put(('status',
                             f'{info["positional_fixed"]} 页为分离式版式，已通过位置解析识别'))
            self._q.put(('parsed', (pdf_path, sig, texts, pages, info)))

        self._run_async(work)

    def _on_parsed(self, payload):
        _pdf_path, self._parsed_sig, self.texts, self.pages, self.parse_info = payload
        self._set_busy(False)

        if not self.pages:
            self._set_status('PDF 不含任何页面')
            return

        empty = sum(1 for p in self.pages if not p.get('is_text_page'))
        self._candidates = rp.list_all_parties(self.pages)
        self._refresh_candidates()

        if empty == len(self.pages):
            self.cand_hint.config(text='未提取到文本', foreground='#c0392b')
            if self.parse_info.get('ocr_used'):
                tip = ('已勾选 OCR，但仍未从扫描页识别出客户名称。'
                       '可能图片清晰度不足或字号过小。')
                self._set_status(f'⚠ {tip}')
                messagebox.showwarning('OCR 未能识别', tip)
            else:
                tip = ('该 PDF 的页面是扫描图片（没有文本层），无法直接读取客户名称。\n\n'
                       '可勾选「扫描件启用 OCR」后重新解析；\n'
                       '首次使用需安装 OCR 引擎：pip install rapidocr_onnxruntime')
                self._set_status(f'⚠ {len(self.pages)} 页均为扫描件，可勾选 OCR 后重新解析')
                messagebox.showwarning('需要 OCR', tip)
        elif empty:
            self._set_status(f'解析完成：共 {len(self.pages)} 页，识别到 {len(self._candidates)} 个客户名称'
                             f'（其中 {empty} 页无文本，可能被跳过）')
        else:
            self._set_status(f'解析完成：共 {len(self.pages)} 页，识别到 {len(self._candidates)} 个客户名称')
        LOG.info('解析完成: %s 页, %s 个候选', len(self.pages), len(self._candidates))

    # ── 候选客户列表 ────────────────────────────────────────────────────────

    def _refresh_candidates(self):
        kw = (self.filter_text.get() or '').strip()
        if kw:
            k = rp.normalize_name(kw)
            self._filtered = [c for c in self._candidates if k in rp.normalize_name(c)]
        else:
            self._filtered = list(self._candidates)

        self.cand_list.delete(0, 'end')
        for c in self._filtered:
            self.cand_list.insert('end', c)

        total, shown = len(self._candidates), len(self._filtered)
        if total == 0:
            self.cand_hint.config(text='尚未解析', foreground='#999')
        elif shown == total:
            self.cand_hint.config(text=f'共 {total} 个', foreground='#2e7d32')
        else:
            self.cand_hint.config(text=f'匹配 {shown} / {total} 个', foreground='#2e7d32')

    def _selected_candidates(self) -> list:
        return [self._filtered[i] for i in self.cand_list.curselection()]

    def _on_cand_double_click(self, _e=None):
        sel = self._selected_candidates()
        if sel:
            self.query_text.set(sel[0])
            self.do_preview()

    def _use_selected(self):
        sel = self._selected_candidates()
        if not sel:
            messagebox.showinfo('提示', '请先在候选列表中选中客户（可按住 Ctrl / Shift 多选）')
            return
        existing = parse_queries(self.query_text.get())
        merged = existing + [s for s in sel if s not in existing]
        self.query_text.set('；'.join(merged))
        self._set_status(f'已加入 {len(sel)} 个客户到名称框，点「② 预览」查看结果')

    # ── 预览匹配 ────────────────────────────────────────────────────────────

    def _ensure_parsed(self) -> bool:
        """确保使用当前文件的解析结果（避免用旧文件缓存做匹配/导出）。"""
        pdf_path = self.src_pdf_path.get().strip()
        if not pdf_path or not os.path.isfile(pdf_path):
            messagebox.showwarning('提示', '请先选择源 PDF 文件')
            return False
        sig = self._file_sig(pdf_path)
        if self.pages and self._parsed_sig == sig:
            return True
        messagebox.showinfo('提示', '请先点「① 解析回单」完成解析')
        return False

    def do_preview(self):
        if self.busy or not self._ensure_parsed():
            return
        queries = parse_queries(self.query_text.get())
        if not queries:
            messagebox.showwarning('提示', '请输入客户名称或页码')
            return

        loose = bool(self.loose_match.get())
        self._set_busy(True)
        self._set_status('正在匹配…')

        def work():
            results = []
            page_by_num = {p['page_num']: p for p in self.pages}
            for q in queries:
                n = rp.parse_page_token(q)
                if n is not None:
                    # 页码写法：直接定位该页，不走名称匹配
                    p = page_by_num.get(n)
                    if p:
                        results.append({'page': p, 'query': q,
                                        'party': p.get('customer') or '指定页码'})
                    continue
                results.extend(rp.find_matches(self.pages, [q], loose))
            self._q.put(('preview', (queries, results)))

        self._run_async(work)

    def _on_preview_ready(self, payload):
        queries, results = payload
        query_order = {q: i for i, q in enumerate(queries)}
        results.sort(key=lambda r: (query_order.get(r['query'], 999), r['page']['page_num']))
        self.matches = results

        self._clear_tree()
        if not results:
            self._set_busy(False)
            self.tree.insert('', 'end', values=('—', '—', '没有匹配到任何页面', '—', '—'))
            self._set_status(f'未找到匹配：{"、".join(queries)}')
            return

        for item in results:
            page, party = item['page'], item['party']
            self.tree.insert('', 'end', values=(
                f'第{page["page_num"]}页',
                item['query'],
                party,
                page.get('date') or '—',
                page.get('amount') or '—',
            ))

        pages_hit = len({i['page']['page_num'] for i in results})
        counts = {}
        for i in results:
            counts[i['query']] = counts.get(i['query'], 0) + 1
        detail = '，'.join(f'{k} {v} 页' for k, v in counts.items())
        missing = [q for q in queries if q not in counts]
        tip = f'（未命中：{"、".join(missing)}）' if missing else ''
        if missing and any(
                (n := rp.parse_page_token(q)) is not None and not 1 <= n <= len(self.pages)
                for q in missing):
            tip += f'（页码超出范围 1~{len(self.pages)}）'
        self._set_busy(False)
        self._set_status(f'匹配完成：命中 {pages_hit} 页　{detail}　{tip}（双击行可打开源文件）')
        LOG.info('预览: queries=%s hits=%s', queries, counts)

    def _clear_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

    def _open_selected_page(self, _e=None):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], 'values')
        if not vals or not str(vals[0]).startswith('第'):
            return
        m = re.search(r'\d+', str(vals[0]))
        if not m:
            return
        page_no = int(m.group())
        pdf = self.src_pdf_path.get().strip()
        if not os.path.isfile(pdf):
            return
        self._show_page(page_no)
        self._set_status(f'预览已跳转到第 {page_no} 页')
        try:
            if os.name == 'nt':
                os.startfile(pdf)  # noqa: S606
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', pdf])
            else:
                subprocess.Popen(['xdg-open', pdf])
        except Exception as ex:
            LOG.warning('打开源文件失败: %s', ex)

    # ── 导出 ────────────────────────────────────────────────────────────────

    def _output_dir(self) -> str:
        d = self.save_dir.get().strip()
        if not d:
            d = os.path.dirname(self.src_pdf_path.get().strip())
        return d

    def do_export(self):
        if self.busy or not self._ensure_parsed():
            return
        if not self.matches:
            messagebox.showinfo('提示', '请先点「② 预览」确认要导出的页面')
            return
        out_dir = self._output_dir()
        if not out_dir:
            messagebox.showwarning('提示', '请指定输出目录')
            return
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception as ex:
            messagebox.showerror('错误', f'无法创建输出目录：\n{ex}')
            return

        pdf_path = self.src_pdf_path.get().strip()
        base = rp.sanitize_filename(os.path.splitext(os.path.basename(pdf_path))[0], '回单')
        overwrite = not bool(self.avoid_overwrite.get())
        merge = bool(self.merge_one.get())

        # 按客户分组导出。同一页可能同时属于多个客户（A 付款给 B，
        # 该页既是 A 的回单也是 B 的回单），因此允许跨组重复，仅组内去重。
        groups = {}
        for item in self.matches:
            groups.setdefault(item['query'], set()).add(item['page']['page_num'])
        groups = {q: sorted(pns) for q, pns in groups.items()}

        overlap = sum(len(v) for v in groups.values()) - len({p for v in groups.values() for p in v})

        plan = []
        if merge:
            all_idx = sorted({p - 1 for v in groups.values() for p in v})
            if len(groups) > 1:
                name = f'{base}__合并__{len(groups)}个客户.pdf'
            else:
                only = next(iter(groups)) if groups else '提取'
                name = f'{base}__{rp.sanitize_filename(only)}__提取.pdf'
            plan.append((name, all_idx, '、'.join(groups.keys())))
        else:
            for q, pns in groups.items():
                plan.append((f'{base}__{rp.sanitize_filename(q)}__提取.pdf',
                             [p - 1 for p in pns], q))

        self.cfg['last_dir'] = out_dir
        save_config(self.cfg)

        self._set_busy(True)
        self.progress['value'] = 0
        note = f'（{overlap} 页同时属于多个客户，将分别出现在各自 PDF 中）' if overlap else ''
        self._set_status(f'正在导出 {len(plan)} 个文件…{note}')

        def work():
            done_files = []
            errors = []
            total = len(plan)
            for n, (fname, idx, label) in enumerate(plan, 1):
                target = os.path.join(out_dir, fname)
                try:
                    actual = rp.split_pdf(pdf_path, target, idx, overwrite=overwrite)
                    done_files.append((actual, len(idx), label))
                except Exception as ex:  # noqa: BLE001
                    errors.append(f'{fname}：{ex}')
                    LOG.error('导出失败 %s\n%s', fname, traceback.format_exc())
                self._q.put(('progress', (n, total)))
            self._q.put(('exported', (out_dir, done_files, errors)))

        self._run_async(work)

    def _on_export_done(self, payload):
        out_dir, files, errors = payload
        self._set_busy(False)
        if not files and errors:
            messagebox.showerror('导出失败', '\n'.join(errors))
            self._set_status('导出失败')
            return

        lines = '\n'.join(f'· {os.path.basename(p)}（{n} 页，{label}）' for p, n, label in files)
        msg = f'导出完成，共 {len(files)} 个文件：\n\n{lines}\n\n保存位置：\n{out_dir}'
        if errors:
            msg += f'\n\n有 {len(errors)} 个文件失败：\n' + '\n'.join(errors)
        self._set_status(f'✅ 已导出 {len(files)} 个文件到 {out_dir}')
        LOG.info('导出完成: %s 个文件 -> %s', len(files), out_dir)

        if messagebox.askyesno('导出完成', msg + '\n\n是否打开输出目录？'):
            open_folder(out_dir)

    def do_export_csv(self):
        if not self.matches:
            messagebox.showinfo('提示', '请先点「② 预览」生成结果')
            return
        path = filedialog.asksaveasfilename(
            title='导出匹配清单', defaultextension='.csv',
            initialfile=f'回单清单_{datetime.now():%Y%m%d_%H%M%S}.csv',
            initialdir=self._output_dir() or os.path.expanduser('~'),
            filetypes=[('CSV 文件', '*.csv')])
        if not path:
            return
        try:
            # utf-8-sig 保证 Excel 打开不乱码
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                w = csv.writer(f)
                w.writerow(['页码', '归属客户', '命中主体', '日期', '金额', '摘要'])
                for i in self.matches:
                    p = i['page']
                    w.writerow([p['page_num'], i['query'], i['party'],
                                p.get('date', ''), p.get('amount', ''), p.get('summary', '')])
            self._set_status(f'清单已导出：{path}')
            if messagebox.askyesno('完成', f'清单已导出：\n{path}\n\n是否打开所在目录？'):
                open_folder(os.path.dirname(path))
        except Exception as ex:
            messagebox.showerror('错误', f'导出清单失败：\n{ex}')

    def _open_out_dir(self):
        d = self._output_dir()
        if d and os.path.isdir(d):
            open_folder(d)
        else:
            messagebox.showinfo('提示', '输出目录还不存在')

    # ── 其他 ────────────────────────────────────────────────────────────────

    def do_clear(self):
        if self.busy:
            return
        self.src_pdf_path.set('')
        self.query_text.set('')
        self.filter_text.set('')
        self._reset_parse_state()
        self._lines = {}
        self._page_no = 1
        if self._pdoc:
            self._pdoc.close()
            self._pdoc = None
            self._preview_sig = None
        self._render_preview()
        self._set_status('已清空')

    def _on_match_option_changed(self):
        if self.matches:
            self._set_status('匹配选项已改变，请重新点「② 预览」')

    def _set_status(self, msg: str):
        self.status_text.set(msg)

    def _on_close(self):
        self.cfg['last_dir'] = self.save_dir.get().strip() or self.cfg.get('last_dir', '')
        self.cfg['ocr_enabled'] = bool(self.ocr_enabled.get())
        save_config(self.cfg)
        LOG.info('退出')
        if self._pdoc:
            self._pdoc.close()
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
# 入口
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # 打包为 exe 后，子进程会重新执行本文件，freeze_support 负责拦截并转入工作函数
    multiprocessing.freeze_support()
    try:
        app = App()
        app.mainloop()
    except Exception:
        LOG.exception('程序异常退出')
        raise


if __name__ == '__main__':
    main()
