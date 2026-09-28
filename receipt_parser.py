# -*- coding: utf-8 -*-
"""
银行 PDF 回单解析核心模块。

- 无 GUI 依赖，可独立单测、可被多进程池 pickle
- 一页提取全部主体名称（付款人 / 收款人 / 户名 …），避免目标客户是付款方时漏单
- 名称归一化（NFKC、去空白、剥标点）后匹配，空值显式短路
- 文本提取：PyMuPDF 为主，空文本页回退 pdfplumber
- 页数达到阈值自动多进程并行
"""

from __future__ import annotations

import os
import re
import unicodedata
from concurrent.futures import ProcessPoolExecutor, as_completed

# ─────────────────────────────────────────────────────────────────────────────
# 常量与配置
# ─────────────────────────────────────────────────────────────────────────────

# 超过该页数启用多进程解析
PARALLEL_THRESHOLD = 20
# 多进程最大工作进程数
MAX_WORKERS = min(8, max(2, (os.cpu_count() or 4)))

# 主体字段标签，按「越具体越靠前」排序（长标签优先，避免被短标签截胡）
CUSTOMER_LABELS = [
    '客户名称', '客户全称', '客户姓名', '账户名称', '账户名',
    '对方户名', '对方账户名', '我方户名', '户名', '客户',
]
PAYER_LABELS = [
    '付款人名称', '付款方名称', '付款账户名称', '付款人全称',
    '付款单位名称', '付款人', '付款方', '付款单位', '付方名称', '汇款人名称', '汇款人',
]
PAYEE_LABELS = [
    '收款人名称', '收款方名称', '收款账户名称', '收款人全称',
    '收款单位名称', '收款人', '收款方', '收款单位', '收方名称',
]
ALL_PARTY_LABELS = CUSTOMER_LABELS + PAYER_LABELS + PAYEE_LABELS

# 金额字段
AMOUNT_LABELS = ['交易金额', '发生额', '小写金额', '金额合计', '金额', '合计', '人民币']
# 大写金额（单独存放，不参与数值列比对）
UPPER_AMOUNT_LABELS = ['大写金额', '金额大写', '大写']

# 摘要字段
SUMMARY_LABELS = ['用途', '摘要', '附言', '交易摘要', '备注', '业务种类', '款项用途']

# 日期与时间分两步取（合并成单个正则时可选分组会抑制回溯，导致时间丢失）
_DATE_PART = re.compile(r'(\d{4}\s*[-/年.]\s*\d{1,2}\s*[-/月.]\s*\d{1,2})\s*(日)?')
_TIME_AFTER = re.compile(r'\s{0,3}(\d{1,2}:\d{2}(?::\d{2})?)(?![\d:])')

# 账号（用于「标签 + 名称 + 长数字账号」兜底匹配）
_ACCOUNT_RE = r'\d{8,}'

# 名称清洗：剥离首尾的标点（中英文）
_NAME_STRIP = ' \t\r\n：:、,，。.;；/\\|-—_()（）[]【】'


# ─────────────────────────────────────────────────────────────────────────────
# 基础工具
# ─────────────────────────────────────────────────────────────────────────────

def _basic_norm(s: str) -> str:
    """基础归一：NFKC 全角转半角 → 压缩所有空白。"""
    if not s:
        return ''
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', s))


def normalize_name(s: str) -> str:
    """名称归一化：基础归一 → 剥离首尾标点。

    归一化是匹配正确性的基础：银行回单 PDF 常因排版在户名中插入空格，
    或使用全角字母/数字，直接比对会漏匹配。
    """
    return _basic_norm(s).strip(_NAME_STRIP)


def normalize_name_loose(s: str) -> str:
    """宽松归一化：先摘掉括号及其内容，再剥离首尾标点。

    用于「XX有限公司(基本户)」与「XX有限公司」的互相匹配。
    注意顺序：必须先摘括号再 strip，否则 `(基本户)` 的右括号会被
    strip 掉、留下不闭合的 `(基本户`，导致括号正则无法命中。
    """
    s = _basic_norm(s)
    if not s:
        return ''
    s = re.sub(r'[（(][^）)]*[）)]?', '', s)
    s = re.sub(r'[【\[][^】\]]*[】\]]?', '', s)
    return s.strip(_NAME_STRIP)


def _cut_at_line_end(s: str) -> str:
    """取首行，避免跨行captured到下一字段。"""
    return re.split(r'[\r\n]', s, maxsplit=1)[0]


def clean_captured(raw: str) -> str:
    """对正则捕获到的原始文本做收尾清理。"""
    if not raw:
        return ''
    s = _cut_at_line_end(raw)
    # 若捕获内容里混入了后续字段的标签，则截断到标签之前
    s = re.split(r'(?:账号|账户|开户行|开户|金额|日期|时间|摘要|用途|附言|币种|状态|流水|凭证|序号)', s, maxsplit=1)[0]
    return normalize_name(s)


def dedupe_keep_order(items) -> list:
    """去重且保持出现顺序。"""
    seen = set()
    out = []
    for it in items:
        if it and it not in seen:
            seen.add(it)
            out.append(it)
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 单页字段提取
# ─────────────────────────────────────────────────────────────────────────────

def _capture_by_labels(text: str, labels, min_len: int = 2, max_len: int = 60) -> list:
    """按标签提取字段值，支持两种写法：
      ① 带分隔符：客户名称：XX有限公司
      ② 无分隔符但后接长账号：付款人 XX有限公司 6222xxxxxxxx
    返回已清洗、已去重的值列表。
    """
    results = []
    for label in labels:
        esc = re.escape(label)
        pat_colon = re.compile(esc + r'\s*[：:]\s*([^\r\n：:]{' + str(min_len) + r',' + str(max_len) + r'})')
        for m in pat_colon.finditer(text):
            v = clean_captured(m.group(1))
            if len(v) >= min_len:
                results.append(v)
        pat_nocolon = re.compile(esc + r'[ \t]*([^ \t\r\n：:0-9]{' + str(min_len) + r',' + str(max_len) + r'})[ \t]*' + _ACCOUNT_RE)
        for m in pat_nocolon.finditer(text):
            v = clean_captured(m.group(1))
            if len(v) >= min_len:
                results.append(v)
    return dedupe_keep_order(results)


def extract_parties(text: str) -> list:
    """提取该页所有主体名称（客户名称 / 户名 / 付款人 / 收款人 …）。"""
    return _capture_by_labels(text, ALL_PARTY_LABELS)


def extract_customer_name(text: str):
    """向后兼容：返回首个主体名称（优先客户名称/户名类标签）。"""
    parties = extract_parties(text)
    return parties[0] if parties else None


def extract_amount(text: str) -> str:
    """提取小写金额。"""
    for label in AMOUNT_LABELS:
        pat = re.compile(re.escape(label) + r'\s*[：:]?\s*[¥￥]?\s*([0-9][0-9,]*\.?[0-9]*)')
        m = pat.search(text)
        if m:
            return m.group(1).strip()
    return ''


def extract_upper_amount(text: str) -> str:
    """提取大写金额。"""
    for label in UPPER_AMOUNT_LABELS:
        pat = re.compile(re.escape(label) + r'\s*[：:]?\s*([^\r\n：:]{2,40})')
        m = pat.search(text)
        if m:
            v = _cut_at_line_end(m.group(1)).strip()
            v = re.split(r'(?:小写|金额|账号|日期)', v, maxsplit=1)[0].strip()
            if v:
                return v
    return ''


def extract_date(text: str) -> str:
    """提取交易日期；若紧接着还有时间则一并保留（如 2024-03-15 10:22:31）。"""
    text = text or ''
    m = _DATE_PART.search(text)
    if not m:
        return ''
    day = re.sub(r'\s+', '', m.group(1)) + (m.group(2) or '')
    tm = _TIME_AFTER.match(text, m.end(), min(len(text), m.end() + 16))
    if tm:
        return f'{day} {tm.group(1)}'
    return day


def extract_summary(text: str) -> str:
    """提取摘要 / 用途。"""
    for label in SUMMARY_LABELS:
        pat = re.compile(re.escape(label) + r'\s*[：:]?\s*([^\r\n：:]{1,60})')
        m = pat.search(text)
        if m:
            v = _cut_at_line_end(m.group(1)).strip()
            v = re.sub(r'\s+', ' ', v)
            if v:
                return v[:40]
    return ''


def parse_page_text(text: str) -> dict:
    """解析单页文本，返回结构化字段。纯函数，可多进程调用。"""
    text = text or ''
    parties = extract_parties(text)
    return {
        'parties': parties,                       # 该页全部主体名称
        'customer': parties[0] if parties else None,  # 主名称（兼容旧字段）
        'date': extract_date(text),
        'amount': extract_amount(text),
        'upper_amount': extract_upper_amount(text),
        'summary': extract_summary(text),
        'text_len': len(text),
        'is_text_page': bool(text.strip()),       # False 说明是图片页，需要 OCR
    }


def parse_page_job(job):
    """多进程工作单元：入参 (页码从0开始, 文本)，返回带 page_num 的结构化结果。"""
    idx, text = job
    info = parse_page_text(text)
    info['page_num'] = idx + 1
    return info


# ─────────────────────────────────────────────────────────────────────────────
# PDF 文本提取
# ─────────────────────────────────────────────────────────────────────────────

class PDFEncryptedError(Exception):
    """PDF 已加密，需要密码。"""


class PDFUnreadableError(Exception):
    """PDF 无法解析（损坏 / 非 PDF）。"""


class OCRUnavailableError(Exception):
    """未安装可用的 OCR 引擎。"""


# OCR 渲染分辨率。回单字号小，200 DPI 在速度与识别率间较平衡。
OCR_DPI = 200

OCR_INSTALL_HINT = (
    '未检测到可用的 OCR 引擎。\n\n'
    '推荐方式（纯 pip 安装，自带中文模型，无需其他软件）：\n'
    '    pip install rapidocr_onnxruntime\n\n'
    '备选方式：安装 Tesseract OCR（需含中文包 chi_sim）：\n'
    '    https://github.com/UB-Mannheim/tesseract/wiki'
)

# 引擎懒加载缓存
_OCR_CACHE = {'probed': False, 'engine': None}


def ocr_probe():
    """探测可用 OCR 引擎，仅做导入级检查（不加载模型，速度快）。

    返回 (引擎名, 描述)；无可用引擎时返回 (None, '')。
    RapidOCR 为纯 pip 依赖且自带中文模型，优先；Tesseract 需单独装二进制。
    """
    try:
        import rapidocr_onnxruntime  # noqa: F401
        return 'rapidocr', 'RapidOCR'
    except ImportError:
        pass
    try:
        import rapidocr  # noqa: F401  # 新版统一包名
        return 'rapidocr', 'RapidOCR'
    except ImportError:
        pass
    try:
        import pytesseract  # noqa: F401
        pytesseract.get_tesseract_version()
        return 'tesseract', 'Tesseract'
    except Exception:
        return None, ''


def get_ocr_engine():
    """加载 OCR 引擎（模型加载较慢，只做一次并缓存）。

    返回 (kind, engine)；不可用时抛 OCRUnavailableError。
    """
    if _OCR_CACHE['probed']:
        if _OCR_CACHE['engine'] is None:
            raise OCRUnavailableError(OCR_INSTALL_HINT)
        return _OCR_CACHE['engine']

    kind = None
    engine = None
    # RapidOCR（两种包名兼容）
    for mod_name in ('rapidocr_onnxruntime', 'rapidocr'):
        try:
            mod = __import__(mod_name, fromlist=['RapidOCR'])
            engine = mod.RapidOCR()
            kind = 'rapidocr'
            break
        except ImportError:
            continue
        except Exception:
            # 包存在但加载失败（如缺 onnxruntime），继续尝试下一个
            continue
    # Tesseract 兜底
    if kind is None:
        try:
            import pytesseract
            pytesseract.get_tesseract_version()
            engine = pytesseract
            kind = 'tesseract'
        except Exception:
            engine = None

    _OCR_CACHE['probed'] = True
    _OCR_CACHE['engine'] = (kind, engine) if kind else None
    if _OCR_CACHE['engine'] is None:
        raise OCRUnavailableError(OCR_INSTALL_HINT)
    return _OCR_CACHE['engine']


def _run_ocr(kind, engine, pix) -> str:
    """对单个页面像素执行 OCR，返回按阅读顺序拼接的文本。"""
    import numpy as np

    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        from PIL import Image
        arr = np.ascontiguousarray(arr[:, :, :3])

    if kind == 'rapidocr':
        result, _ = engine(arr)
        if not result:
            return ''
        return '\n'.join(str(item[1]) for item in result)
    # tesseract
    from PIL import Image
    img = Image.fromarray(arr)
    return engine.image_to_string(img, lang='chi_sim+eng')


def ocr_pages(pdf_path: str, indices, dpi: int = OCR_DPI, progress=None) -> dict:
    """对指定页（0 基索引）执行 OCR，返回 {页索引: 文本}。"""
    import pymupdf

    kind, engine = get_ocr_engine()
    out = {}
    doc = pymupdf.open(pdf_path)
    try:
        mat = pymupdf.Matrix(dpi / 72.0, dpi / 72.0)
        total = len(indices)
        for n, i in enumerate(indices):
            try:
                pix = doc[i].get_pixmap(matrix=mat, alpha=False)
                out[i] = _run_ocr(kind, engine, pix)
            except Exception:
                out[i] = ''
            if progress:
                progress(n + 1, total)
    finally:
        doc.close()
    return out


def extract_texts_pymupdf(pdf_path: str) -> list:
    """用 PyMuPDF 提取每页文本（快，首选）。"""
    import pymupdf  # 延迟导入，便于本模块被轻量复用

    try:
        doc = pymupdf.open(pdf_path)
    except Exception as ex:
        raise PDFUnreadableError(f'无法打开文件：{ex}') from ex

    try:
        if doc.needs_pass:
            raise PDFEncryptedError('该 PDF 已加密（需要打开密码），请先解除加密再处理')
        texts = []
        for page in doc:
            try:
                texts.append(page.get_text('text') or '')
            except Exception:
                texts.append('')
        return texts
    finally:
        doc.close()


def extract_texts_pdfplumber(pdf_path: str, only_indices=None) -> dict:
    """用 pdfplumber 提取指定页文本（慢，作为兜底）。返回 {页码索引: 文本}。"""
    import pdfplumber

    result = {}
    indices = set(only_indices) if only_indices is not None else None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) == 0:
                raise PDFUnreadableError('PDF 不含任何页面')
            for i, page in enumerate(pdf.pages):
                if indices is not None and i not in indices:
                    continue
                try:
                    result[i] = page.extract_text() or ''
                except Exception:
                    result[i] = ''
    except (PDFUnreadableError,):
        raise
    except Exception as ex:
        # pdfplumber 对加密文件的报错信息不统一，统一归类
        msg = str(ex)
        if 'password' in msg.lower() or 'encrypt' in msg.lower():
            raise PDFEncryptedError('该 PDF 已加密（需要打开密码），请先解除加密再处理') from ex
        raise PDFUnreadableError(f'PDF 解析失败：{ex}') from ex
    return result


def extract_all_texts(pdf_path: str, use_pdfplumber_fallback: bool = True) -> list:
    """提取 PDF 全部页面文本。

    策略：PyMuPDF 全量提取 → 对「文本为空」的页面用 pdfplumber 重试一次。
    这样既拿到 PyMuPDF 的速度，又保留 pdfplumber 对特殊版面的兼容性。
    """
    texts = extract_texts_pymupdf(pdf_path)
    if not texts:
        raise PDFUnreadableError('PDF 不含任何页面')

    if use_pdfplumber_fallback:
        empty_idx = [i for i, t in enumerate(texts) if not t.strip()]
        if empty_idx:
            try:
                fallback = extract_texts_pdfplumber(pdf_path, only_indices=empty_idx)
                for i, t in fallback.items():
                    if t and t.strip():
                        texts[i] = t
            except Exception:
                # 兜底失败不影响主流程
                pass
    return texts


def build_texts(pdf_path: str, ocr: bool = False, progress=None, status=None):
    """提取全部页面文本；ocr=True 时对无文本页追加 OCR。

    返回 (texts, info)，info 含：
      total       总页数
      empty_pages 无文本页数（扫描页）
      ocr_pages   OCR 成功产出文本的页数
      ocr_used    是否实际执行了 OCR
    progress(done, total) 仅统计 OCR 阶段；status(msg) 用于阶段提示。
    """
    texts = extract_all_texts(pdf_path)
    empty_idx = [i for i, t in enumerate(texts) if not t.strip()]
    info = {'total': len(texts), 'empty_pages': len(empty_idx),
            'ocr_pages': 0, 'ocr_used': False}

    if empty_idx and ocr:
        if status:
            status(f'检测到 {len(empty_idx)} 页扫描页，OCR 识别中…')
        got = ocr_pages(pdf_path, empty_idx, progress=progress)
        for i, t in got.items():
            if t and t.strip():
                texts[i] = t
                info['ocr_pages'] += 1
        info['ocr_used'] = True
    return texts, info


# ─────────────────────────────────────────────────────────────────────────────
# 批量解析（支持多进程）
# ─────────────────────────────────────────────────────────────────────────────

def parse_pages(texts: list, progress=None, use_parallel: bool = True) -> list:
    """解析全部页面文本。

    texts      : 每页文本列表
    progress   : 可选回调 progress(done, total)
    use_parallel: 页数 >= PARALLEL_THRESHOLD 时启用多进程
    """
    n = len(texts)
    if n == 0:
        return []

    jobs = list(enumerate(texts))

    if (not use_parallel) or n < PARALLEL_THRESHOLD:
        out = []
        for i, job in enumerate(jobs):
            out.append(parse_page_job(job))
            if progress and (i % 10 == 0 or i == n - 1):
                progress(i + 1, n)
        return out

    results = [None] * n
    done = 0
    try:
        with ProcessPoolExecutor(max_workers=MAX_WORKERS) as ex:
            future_map = {ex.submit(parse_page_job, job): job[0] for job in jobs}
            for fut in as_completed(future_map):
                idx = future_map[fut]
                try:
                    results[idx] = fut.result()
                except Exception:
                    info = parse_page_text(texts[idx])
                    info['page_num'] = idx + 1
                    results[idx] = info
                done += 1
                if progress:
                    progress(done, n)
    except Exception:
        # 多进程不可用（如打包环境限制）时降级为串行，保证功能可用
        results = []
        for i, job in enumerate(jobs):
            results.append(parse_page_job(job))
            if progress:
                progress(i + 1, n)

    # 兜底：任何仍为 None 的槽位补齐
    for i in range(n):
        if results[i] is None:
            info = parse_page_text(texts[i])
            info['page_num'] = i + 1
            results[i] = info
    return results


# ─────────────────────────────────────────────────────────────────────────────
# 匹配
# ─────────────────────────────────────────────────────────────────────────────

def name_matches(party: str, query: str, loose: bool = True) -> bool:
    """判断单个主体名称是否命中查询词。

    规则（双方均先归一化，空值直接判否）：
      - 相等 → 命中
      - 互相包含 → 命中（支持「茅台」命中「贵州茅台酒业有限公司」）
      - loose=True 时忽略括号内容后再比一次
    """
    q = normalize_name(query)
    if not q or not party:
        return False

    p = normalize_name(party)
    if not p:
        return False

    if p == q or q in p or p in q:
        return True

    if loose:
        pl = normalize_name_loose(party)
        ql = normalize_name_loose(query)
        if pl and ql and (pl == ql or ql in pl or pl in ql):
            return True
    return False


def _parties_of(page: dict) -> list:
    """取该页主体列表，兼容仅有 customer 字段的旧结构。"""
    parties = page.get('parties')
    if parties:
        return parties
    return [page['customer']] if page.get('customer') else []


def match_pages(pages: list, query: str, loose: bool = True) -> list:
    """单查询词匹配，返回命中的页面对象列表。"""
    if not query or not query.strip():
        return []
    out = []
    for page in pages:
        if any(name_matches(p, query, loose=loose) for p in _parties_of(page)):
            out.append(page)
    return out


def find_matches(pages: list, queries, loose: bool = True) -> list:
    """多查询词匹配，返回明细列表。

    每项形如 {'page': 页面dict, 'query': 查询词, 'party': 命中主体}。
    一页可被多个查询词分别命中（由调用方决定如何去重导出）。
    """
    out = []
    for q in queries:
        if not q or not q.strip():
            continue
        for page in pages:
            for p in _parties_of(page):
                if name_matches(p, q, loose=loose):
                    out.append({'page': page, 'query': q, 'party': p})
                    break
    return out


# 页码写法：第3页 / p5 / #7（前缀支持全角大写）
_PAGE_TOKEN_RE = re.compile(r'^(?:第\s*(\d{1,4})\s*页|[pPｐＰ#＃](\d{1,4})|(\d{1,4})\s*页)$')


def parse_page_token(q: str):
    """把「第3页 / p5 / #7」类页码写法解析为 1 基页码；非页码写法返回 None。"""
    if not q:
        return None
    m = _PAGE_TOKEN_RE.match(q.strip())
    if not m:
        return None
    return int(next(g for g in m.groups() if g))


def list_all_parties(pages: list) -> list:
    """汇总所有页面出现的主体名称（去重、按出现次数降序）。"""
    counter = {}
    for page in pages:
        for p in _parties_of(page):
            counter[p] = counter.get(p, 0) + 1
    return [name for name, _ in sorted(counter.items(), key=lambda kv: (-kv[1], kv[0]))]


# ─────────────────────────────────────────────────────────────────────────────
# PDF 拆分 / 导出
# ─────────────────────────────────────────────────────────────────────────────

def unique_path(path: str) -> str:
    """若目标文件已存在，自动追加 (1)(2)… 避免覆盖用户文件。"""
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    i = 1
    while True:
        cand = f'{root}({i}){ext}'
        if not os.path.exists(cand):
            return cand
        i += 1


def sanitize_filename(name: str, fallback: str = '未命名') -> str:
    """清洗文件名中的非法字符，并处理 Windows 保留名。"""
    s = normalize_name(name) or fallback
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', '_', s)
    s = re.sub(r'_+', '_', s).strip('_ .')
    if not s:
        s = fallback
    reserved = {'CON', 'PRN', 'AUX', 'NUL'} | {f'COM{i}' for i in range(1, 10)} | {f'LPT{i}' for i in range(1, 10)}
    if s.upper() in reserved:
        s = f'_{s}'
    return s[:120]


def split_pdf(pdf_path: str, output_path: str, page_indices, overwrite: bool = True) -> str:
    """把指定页（0 基索引）抽取合并为新 PDF。

    返回实际写入的路径（可能因避让重名而带序号）。
    overwrite=False 时自动避让已存在文件。
    """
    import pymupdf

    if not page_indices:
        raise ValueError('没有可导出的页面')

    if not os.path.isabs(output_path):
        output_path = os.path.abspath(output_path)

    target = output_path if overwrite else unique_path(output_path)

    src = pymupdf.open(pdf_path)
    out = pymupdf.open()
    try:
        total = src.page_count
        for idx in page_indices:
            if idx < 0 or idx >= total:
                continue
            out.insert_pdf(src, from_page=idx, to_page=idx)
        if out.page_count == 0:
            raise ValueError('所有目标页码均越界，未导出任何页面')
        os.makedirs(os.path.dirname(target) or '.', exist_ok=True)
        try:
            out.save(target, garbage=3, deflate=True)
        except Exception as ex:
            # 常见原因：目标文件正被 PDF 阅读器占用
            raise IOError(f'写入失败（文件可能正被其他程序占用）：{ex}') from ex
    finally:
        out.close()
        src.close()
    return target
