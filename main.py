# -*- coding: utf-8 -*-
"""
银行 PDF 回单拆分工具（GUI）

用法：python main.py，或直接运行打包后的 exe。

流程：选择/拖入源 PDF → ① 解析回单 → 指定客户名称 → ② 预览匹配 → ③ 导出 PDF。
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

# 查询词分隔符
QUERY_SPLIT_RE = re.compile(r'[,，;；、\n\r\t]+')

UI_FONT = '微软雅黑'


# ─────────────────────────────────────────────────────────────────────────────
# 配置 / 日志
# ─────────────────────────────────────────────────────────────────────────────

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


# ─────────────────────────────────────────────────────────────────────────────
# 主窗口
# ─────────────────────────────────────────────────────────────────────────────

_AppBase = TkinterDnD.Tk if _HAS_DND else tk.Tk


class App(_AppBase):

    def __init__(self):
        super().__init__()
        self.cfg = load_config()

        self.title(APP_NAME)
        self.geometry('960x860')
        self.minsize(820, 640)

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

        self.pages: list = []
        self.texts: list = []
        self._parsed_sig = None         # 已解析文件的签名，用于缓存失效判断
        self._candidates: list = []
        self._filtered: list = []
        self.matches: list = []
        self.busy = False
        self.parse_info = {'total': 0, 'empty_pages': 0, 'ocr_pages': 0, 'ocr_used': False}

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
        style.configure('TCombobox', font=(UI_FONT, 10))

        # 顶部标题条
        head = tk.Frame(self, bg='#0f4c9e', height=54)
        head.pack(fill='x')
        head.pack_propagate(False)
        tk.Label(head, text=APP_NAME, font=(UI_FONT, 16, 'bold'),
                 fg='white', bg='#0f4c9e').pack(side='left', padx=20)
        tk.Label(head, text='按客户名称拆分回单', font=(UI_FONT, 9),
                 fg='#bcd6f5', bg='#0f4c9e').pack(side='left')

        body = ttk.Frame(self, padding=(16, 12, 16, 8))
        body.pack(fill='both', expand=True)
        body.columnconfigure(0, weight=1)

        # ── 源文件 ──
        box1 = ttk.LabelFrame(body, text=' 1. 选择源 PDF ', padding=10)
        box1.grid(row=0, column=0, sticky='ew')
        box1.columnconfigure(0, weight=1)

        ttk.Entry(box1, textvariable=self.src_pdf_path).grid(row=0, column=0, sticky='ew')
        ttk.Button(box1, text='选择文件', width=10, command=self.select_pdf)\
            .grid(row=0, column=1, padx=(8, 0))

        self.drop_label = tk.Label(
            box1, text='⬇  把 PDF 文件拖到这里（或拖到上方输入框）',
            font=(UI_FONT, 9), fg='#5a6b7d', bg='#eef4fc',
            relief='solid', borderwidth=1, pady=8)
        self.drop_label.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(8, 0))

        # ── 输出与匹配选项 ──
        box2 = ttk.LabelFrame(body, text=' 2. 输出位置与匹配选项 ', padding=10)
        box2.grid(row=1, column=0, sticky='ew', pady=(10, 0))
        box2.columnconfigure(0, weight=1)

        ttk.Entry(box2, textvariable=self.save_dir).grid(row=0, column=0, sticky='ew')
        ttk.Button(box2, text='选择目录', width=10, command=self.select_dir)\
            .grid(row=0, column=1, padx=(8, 0))

        opts = ttk.Frame(box2)
        opts.grid(row=1, column=0, columnspan=2, sticky='w', pady=(8, 0))
        ttk.Checkbutton(opts, text='宽松匹配（忽略括号内容）', variable=self.loose_match,
                        command=self._on_match_option_changed).pack(side='left')
        ttk.Checkbutton(opts, text='不覆盖已存在文件（自动加序号）',
                        variable=self.avoid_overwrite).pack(side='left', padx=(16, 0))
        ttk.Checkbutton(opts, text='所有客户合并为一个 PDF', variable=self.merge_one).pack(side='left', padx=(16, 0))

        opts2 = ttk.Frame(box2)
        opts2.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(4, 0))
        ttk.Checkbutton(opts2, text='扫描件启用 OCR（图片页自动识别，较慢）',
                        variable=self.ocr_enabled).pack(side='left')
        self.ocr_label = ttk.Label(opts2, textvariable=self.ocr_info, foreground='#999')
        self.ocr_label.pack(side='left', padx=(12, 0))

        # ── 客户名称 ──
        box3 = ttk.LabelFrame(body, text=' 3. 指定客户名称（可多个，用逗号/分号/顿号/换行分隔） ', padding=10)
        box3.grid(row=2, column=0, sticky='ew', pady=(10, 0))
        box3.columnconfigure(0, weight=1)

        ttk.Entry(box3, textvariable=self.query_text).grid(row=0, column=0, sticky='ew')
        ttk.Button(box3, text='预览匹配', width=10, command=self.do_preview)\
            .grid(row=0, column=1, padx=(8, 0))

        # 候选客户区
        cand = ttk.Frame(box3)
        cand.grid(row=1, column=0, columnspan=2, sticky='nsew', pady=(8, 0))
        box3.rowconfigure(1, weight=1)
        cand.columnconfigure(0, weight=1)

        bar = ttk.Frame(cand)
        bar.grid(row=0, column=0, columnspan=2, sticky='ew')
        ttk.Label(bar, text='候选客户（解析后自动识别，可多选）：', foreground='#444')\
            .pack(side='left')
        self.cand_hint = ttk.Label(bar, text='尚未解析', foreground='#999')
        self.cand_hint.pack(side='left', padx=(6, 0))
        ttk.Entry(bar, textvariable=self.filter_text, width=16)\
            .pack(side='right')
        ttk.Label(bar, text='筛选：', foreground='#444').pack(side='right')
        self.filter_text.trace_add('write', lambda *_: self._refresh_candidates())

        lf = ttk.Frame(cand)
        lf.grid(row=1, column=0, columnspan=2, sticky='nsew', pady=(4, 0))
        lf.columnconfigure(0, weight=1)
        lf.rowconfigure(0, weight=1)

        self.cand_list = tk.Listbox(lf, selectmode='extended', height=4,
                                    font=(UI_FONT, 10), activestyle='none',
                                    exportselection=False,
                                    selectbackground='#cfe3fb', selectforeground='#123')
        self.cand_list.grid(row=0, column=0, sticky='nsew')
        csb = ttk.Scrollbar(lf, orient='vertical', command=self.cand_list.yview)
        self.cand_list.configure(yscrollcommand=csb.set)
        csb.grid(row=0, column=1, sticky='ns')
        self.cand_list.bind('<Double-Button-1>', self._on_cand_double_click)

        btn_row = ttk.Frame(cand)
        btn_row.grid(row=2, column=0, columnspan=2, sticky='w', pady=(6, 0))
        ttk.Button(btn_row, text='加入名称框', width=12, command=self._use_selected)\
            .pack(side='left')
        ttk.Button(btn_row, text='全选', width=7, command=lambda: self.cand_list.select_set(0, 'end'))\
            .pack(side='left', padx=(6, 0))
        ttk.Button(btn_row, text='清除选择', width=9,
                   command=lambda: self.cand_list.selection_clear(0, 'end'))\
            .pack(side='left', padx=(6, 0))

        # ── 预览 ──
        box4 = ttk.LabelFrame(body, text=' 4. 匹配预览 ', padding=10)
        box4.grid(row=3, column=0, sticky='nsew', pady=(10, 0))
        box4.columnconfigure(0, weight=1)
        box4.rowconfigure(0, weight=1)
        body.rowconfigure(3, weight=1)

        cols = ('page', 'query', 'party', 'date', 'amount', 'summary')
        self.tree = ttk.Treeview(box4, columns=cols, show='headings', height=7)
        for cid, text, w, anchor in (
            ('page', '页码', 70, 'center'),
            ('query', '归属客户', 150, 'w'),
            ('party', '命中主体', 210, 'w'),
            ('date', '日期', 130, 'center'),
            ('amount', '金额', 120, 'e'),
            ('summary', '摘要', 170, 'w'),
        ):
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=w, anchor=anchor, stretch=(cid in ('party', 'summary')))
        self.tree.grid(row=0, column=0, sticky='nsew')

        tsb = ttk.Scrollbar(box4, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=tsb.set)
        tsb.grid(row=0, column=1, sticky='ns')

        self.tree.bind('<Double-Button-1>', self._open_selected_page)

        # ── 操作区 ──
        act = ttk.Frame(body)
        act.grid(row=4, column=0, sticky='ew', pady=(10, 0))

        self.btn_parse = ttk.Button(act, text='① 解析回单', width=13, command=self.do_parse)
        self.btn_parse.pack(side='left')
        self.btn_preview = ttk.Button(act, text='② 预览匹配', width=13, command=self.do_preview)
        self.btn_preview.pack(side='left', padx=(8, 0))
        self.btn_export = ttk.Button(act, text='③ 导出 PDF', width=13, command=self.do_export)
        self.btn_export.pack(side='left', padx=(8, 0))

        ttk.Button(act, text='导出清单', width=10, command=self.do_export_csv).pack(side='left', padx=(8, 0))
        ttk.Button(act, text='打开目录', width=10, command=self._open_out_dir).pack(side='left', padx=(8, 0))
        ttk.Button(act, text='清空', width=8, command=self.do_clear).pack(side='right')

        # ── 进度 / 状态 ──
        foot = ttk.Frame(body)
        foot.grid(row=5, column=0, sticky='ew', pady=(8, 0))
        foot.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(foot, mode='determinate', maximum=100)
        self.progress.grid(row=0, column=0, sticky='ew')
        ttk.Label(foot, textvariable=self.status_text, foreground='#333')\
            .grid(row=1, column=0, sticky='w', pady=(4, 0))

    def _setup_dnd(self):
        for widget in (self.drop_label,):
            widget.drop_target_register('DND_Files')
            widget.dnd_bind('<<Drop>>', self._on_drop)
            widget.dnd_bind('<<DragEnter>>', self._on_drag_enter)
            widget.dnd_bind('<<DragLeave>>', self._on_drag_leave)

    # ── OCR ─────────────────────────────────────────────────────────────────

    def _probe_ocr_async(self):
        """启动时后台探测 OCR 引擎，界面显示可用性（不阻塞窗口出现）。"""
        def work():
            name, desc = rp.ocr_probe()
            if name:
                self._q.put(('ocrinfo', (f'OCR：{desc} 可用', '#2e7d32')))
            else:
                self._q.put(('ocrinfo', ('OCR：未安装（勾选后可 pip install rapidocr_onnxruntime 启用）',
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
                elif kind == 'busy':
                    self._set_busy(bool(payload))
                elif kind == 'error':
                    self._set_busy(False)
                    messagebox.showerror('出错了', payload)
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)

    def _run_async(self, fn):
        """在后台线程执行 fn（fn 内部通过 self._q 汇报进度）。"""
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
                self._q.put(('busy', False))

        threading.Thread(target=runner, daemon=True).start()

    def _set_busy(self, busy: bool):
        self.busy = busy
        state = 'disabled' if busy else 'normal'
        for b in (self.btn_parse, self.btn_preview, self.btn_export):
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
            texts, info = rp.build_texts(
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
            self._q.put(('status', f'共 {info["total"]} 页，正在解析字段…'))
            pages = rp.parse_pages(
                texts,
                progress=lambda d, t: self._q.put(('progress', (d, t))))
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
                       '可勾选上方「扫描件启用 OCR」后重新解析；\n'
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
            messagebox.showinfo('提示', '请先在下方列表中选中客户（可按住 Ctrl / Shift 多选）')
            return
        existing = parse_queries(self.query_text.get())
        merged = existing + [s for s in sel if s not in existing]
        self.query_text.set('；'.join(merged))
        self._set_status(f'已加入 {len(sel)} 个客户到名称框，点「② 预览匹配」查看结果')

    # ── 预览 ────────────────────────────────────────────────────────────────

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
            messagebox.showwarning('提示', '请输入或选择客户名称')
            return

        loose = bool(self.loose_match.get())
        self._set_busy(True)
        self._set_status('正在匹配…')

        def work():
            self._q.put(('preview', (queries, rp.find_matches(self.pages, queries, loose))))

        self._run_async(work)

    def _on_preview_ready(self, payload):
        queries, results = payload
        query_order = {q: i for i, q in enumerate(queries)}
        results.sort(key=lambda r: (query_order.get(r['query'], 999), r['page']['page_num']))
        self.matches = results

        self._clear_tree()
        if not results:
            self._set_busy(False)
            self.tree.insert('', 'end', values=('—', '—', '没有匹配到任何页面', '—', '—', '—'))
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
                (page.get('summary') or '—')[:24],
            ))

        pages_hit = len({i['page']['page_num'] for i in results})
        counts = {}
        for i in results:
            counts[i['query']] = counts.get(i['query'], 0) + 1
        detail = '，'.join(f'{k} {v} 页' for k, v in counts.items())
        missing = [q for q in queries if q not in counts]
        tip = f'（未命中：{"、".join(missing)}）' if missing else ''
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
        self._set_status(f'正在打开源文件…（请在阅读器中跳转到第 {page_no} 页）')
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
            messagebox.showinfo('提示', '请先点「② 预览匹配」确认要导出的页面')
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
            messagebox.showinfo('提示', '请先点「② 预览匹配」生成结果')
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
        self._set_status('已清空')

    def _on_match_option_changed(self):
        if self.matches:
            self._set_status('匹配选项已改变，请重新点「② 预览匹配」')

    def _set_status(self, msg: str):
        self.status_text.set(msg)

    def _on_close(self):
        self.cfg['last_dir'] = self.save_dir.get().strip() or self.cfg.get('last_dir', '')
        self.cfg['ocr_enabled'] = bool(self.ocr_enabled.get())
        save_config(self.cfg)
        LOG.info('退出')
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
