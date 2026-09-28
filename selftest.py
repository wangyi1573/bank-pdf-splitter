# -*- coding: utf-8 -*-
"""
贵州银行回单拆分工具 - 全链路自测

覆盖：
  A. 名称归一化与匹配规则（含「空串误匹配」BUG 回归）
  B. 多主体提取（付款人 / 收款人 / 户名）
  C. 单页字段提取（日期 / 金额 / 摘要）
  D. 端到端：生成中文回单 PDF → 提取文本 → 解析 → 匹配 → 拆分导出
  E. 多进程解析路径
  F. 异常：加密 PDF / 扫描件（无文本层）
  G. 输出文件重名避让 / 文件名清洗

运行：
    python selftest.py
"""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
import traceback

sys.stdout.reconfigure(encoding='utf-8')

import pymupdf  # noqa: E402
import receipt_parser as rp  # noqa: E402

PASS, FAIL = [], []


def check(name, cond, extra=''):
    if cond:
        PASS.append(name)
        print(f'  [PASS] {name}')
    else:
        FAIL.append((name, extra))
        print(f'  [FAIL] {name}  {extra}')


def section(title):
    print()
    print('=' * 72)
    print(title)
    print('=' * 72)


# ─────────────────────────────────────────────────────────────────────────────
# 造样本 PDF
# ─────────────────────────────────────────────────────────────────────────────

FONT_CANDIDATES = [
    'C:/Windows/Fonts/msyh.ttc',
    'C:/Windows/Fonts/simsun.ttc',
    'C:/Windows/Fonts/simhei.ttf',
]

# (付款人, 付款账号, 收款人, 收款账号, 日期, 金额, 用途)
RECORDS = [
    ('贵州茅台酒业有限公司', '6222021234567890123', '遵义粮油贸易有限公司', '6222029876543210987',
     '2024-03-15 10:22:31', '500,000.00', '货款'),
    ('五粮液集团有限公司', '6222025555666677778', '贵州茅台酒业有限公司', '6222021234567890123',
     '2024-03-16 14:05:00', '300,000.00', '原料款'),
    ('贵州茅台酒业有限公司', '6222021234567890123', '成都商贸有限责任公司', '6222024444555566667',
     '2024-03-20 09:30:12', '800,000.00', '预付货款（2024年一季度）'),
    ('贵阳建材有限公司', '6222023333444455556', '重庆物流有限公司', '6222022222333344445',
     '2024-03-21 16:48:00', '12,345.67', '运费'),
]

BLANK_PAGE_COUNT = 1  # 末尾插入一页无文本页，用于验证「空页不误匹配」


def build_sample_pdf(path: str) -> bool:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    font_name = 'Helvetica'
    for fp in FONT_CANDIDATES:
        if not os.path.exists(fp):
            continue
        try:
            pdfmetrics.registerFont(TTFont('cn', fp))
            font_name = 'cn'
            break
        except Exception:
            continue
    if font_name == 'Helvetica':
        print('  (警告) 未找到中文字体，中文将无法正确渲染，测试结果仅供参考')

    c = canvas.Canvas(path, pagesize=A4)
    w, h = A4
    for i, (payer, pacct, payee, racc, date, amount, summary) in enumerate(RECORDS):
        if i:
            c.showPage()
        y = h - 45 * mm
        c.setFont(font_name, 15)
        c.drawString(45 * mm, y, '贵州银行电子回单')
        c.setFont(font_name, 10)
        rows = [
            f'交易日期：{date}',
            f'付款人名称：{payer}',
            f'付款人账号：{pacct}',
            f'收款人名称：{payee}',
            f'收款人账号：{racc}',
            f'金额：{amount}',
            f'用途：{summary}',
            '凭证号：GZ20240315000%d' % (i + 1),
        ]
        for j, row in enumerate(rows):
            c.drawString(45 * mm, y - (12 + j * 8) * mm, row)
    c.save()

    # 追加无文本页
    if BLANK_PAGE_COUNT:
        tmp2 = path + '.tmp.pdf'
        doc = pymupdf.open(path)
        for _ in range(BLANK_PAGE_COUNT):
            doc.new_page()
        doc.save(tmp2)
        doc.close()
        os.replace(tmp2, path)
    return font_name != 'Helvetica'


# ─────────────────────────────────────────────────────────────────────────────
# A. 归一化与匹配
# ─────────────────────────────────────────────────────────────────────────────

def test_normalize():
    section('A. 名称归一化与匹配规则')
    check('全角字母转半角', rp.normalize_name('ＡＢＣ１２３') == 'ABC123')
    check('去除所有空白', rp.normalize_name('贵州 茅台  酒业') == '贵州茅台酒业')
    check('剥离首尾标点', rp.normalize_name('：贵州茅台，') == '贵州茅台')
    check('空值返回空串', rp.normalize_name('') == '' and rp.normalize_name(None) == '')
    check('宽松归一化忽略括号', rp.normalize_name_loose('贵州茅台酒业有限公司(基本户)') == '贵州茅台酒业有限公司')

    # 核心规则
    check('完全相同 → 命中', rp.name_matches('贵州茅台酒业有限公司', '贵州茅台酒业有限公司'))
    check('查询词是子串 → 命中', rp.name_matches('贵州茅台酒业有限公司', '茅台'))
    check('带空格仍命中', rp.name_matches('贵州茅台 酒业有限公司', '贵州茅台酒业'))

    # ★ BUG 回归：旧版 `c.lower() in customer_lower` 在 c 为空串时恒为 True
    check('★空主体 不匹配任何查询词', not rp.name_matches('', '贵州茅台'))
    check('★None 主体 不匹配', not rp.name_matches(None, '贵州茅台'))
    check('★空查询词 不匹配任何主体', not rp.name_matches('贵州茅台', ''))
    check('★空白查询词 不匹配', not rp.name_matches('贵州茅台', '   '))

    # match_pages 层面的回归
    pages = [{'page_num': 1, 'parties': [], 'customer': None},
             {'page_num': 2, 'parties': ['遵义粮油贸易有限公司'], 'customer': '遵义粮油贸易有限公司'}]
    hit = rp.match_pages(pages, '贵州茅台')
    check('★无主体的页面不会被误判命中', len(hit) == 0, f'实际命中 {len(hit)} 页')

    hit2 = rp.find_matches(pages, ['遵义'], loose=True)
    check('find_matches 明细结构正确',
          len(hit2) == 1 and hit2[0]['query'] == '遵义' and hit2[0]['party'] == '遵义粮油贸易有限公司')


# ─────────────────────────────────────────────────────────────────────────────
# B / C. 单页提取
# ─────────────────────────────────────────────────────────────────────────────

SAMPLE_PAGE = """贵州银行电子回单
交易日期：2024-03-15 10:22:31
付款人名称：贵州茅台酒业有限公司
付款人账号：6222021234567890123
收款人名称：遵义粮油贸易有限公司
收款人账号：6222029876543210987
金额：500,000.00
用途：货款
凭证号：GZ202403150001
"""


def test_extract():
    section('B/C. 单页字段提取')
    info = rp.parse_page_text(SAMPLE_PAGE)

    check('提取到两个主体', len(info['parties']) == 2, str(info['parties']))
    check('付款人被提取', '贵州茅台酒业有限公司' in info['parties'], str(info['parties']))
    check('收款人被提取', '遵义粮油贸易有限公司' in info['parties'], str(info['parties']))
    check('★付款人为目标客户也能匹配（旧版漏单场景）',
          len(rp.match_pages([info], '贵州茅台')) == 1)
    check('日期提取（含时分秒）', info['date'] == '2024-03-15 10:22:31', repr(info['date']))
    check('金额提取', info['amount'] == '500,000.00', repr(info['amount']))
    check('摘要提取', info['summary'] == '货款', repr(info['summary']))
    check('文本页标记为 True', info['is_text_page'] is True)

    # 无冒号 + 长账号的写法
    t2 = '付款人 贵阳建材有限公司 6222023333444455556\n收款人 重庆物流有限公司 6222022222333344445\n'
    i2 = rp.parse_page_text(t2)
    check('无冒号写法可提取', set(i2['parties']) == {'贵阳建材有限公司', '重庆物流有限公司'}, str(i2['parties']))

    # 户名写法
    i3 = rp.parse_page_text('户名：贵州茅台酒业有限公司\n账号：6222021234567890123\n金额：1,000.00\n')
    check('户名写法可提取', i3['parties'] == ['贵州茅台酒业有限公司'], str(i3['parties']))

    # 空页
    i4 = rp.parse_page_text('')
    check('★空页不产生主体', i4['parties'] == [] and i4['is_text_page'] is False)

    # 日期各种写法（回归：贪婪 \s* 吃掉分隔空格导致时间抓不到）
    date_cases = [
        ('交易日期：2024-03-15 10:22:31\n付款人名称：X', '2024-03-15 10:22:31'),
        ('日期 2024-03-15', '2024-03-15'),
        ('2024年03月15日', '2024年03月15日'),
        ('2024/3/5 9:05', '2024/3/5 9:05'),
        ('交易日期：2024-03-15\n付款人名称：A\n账号：123', '2024-03-15'),
        ('无日期', ''),
    ]
    for src, want in date_cases:
        got = rp.extract_date(src)
        check(f'★日期提取 {src[:18]!r} → {want!r}', got == want, f'实际 {got!r}')


# ─────────────────────────────────────────────────────────────────────────────
# D/E. 端到端
# ─────────────────────────────────────────────────────────────────────────────

def test_end_to_end(pdf_path: str, parallel: bool, outdir: str, tag: str):
    section(f'{"E" if parallel else "D"}. 端到端（{"多进程" if parallel else "串行"}解析）')

    texts = rp.extract_all_texts(pdf_path)
    expected_pages = len(RECORDS) + BLANK_PAGE_COUNT
    check(f'[{tag}] 提取到全部页面文本', len(texts) == expected_pages, f'{len(texts)} vs {expected_pages}')

    pages = rp.parse_pages(texts, use_parallel=parallel)
    check(f'[{tag}] 解析结果页数一致', len(pages) == expected_pages)
    check(f'[{tag}] 页码连续 1..N', [p['page_num'] for p in pages] == list(range(1, expected_pages + 1)))

    cands = rp.list_all_parties(pages)
    check(f'[{tag}] 候选名称含 6 个主体（4页 × 付/收双方去重）', len(cands) == 6, str(cands))
    check(f'[{tag}] 候选按出现频次降序（茅台 3 次居首）', cands[0] == '贵州茅台酒业有限公司', str(cands))

    # 茅台：第1页付款人、第2页收款人、第3页付款人 → 命中 3 页
    m = rp.find_matches(pages, ['贵州茅台酒业有限公司'], loose=True)
    hit_pages = sorted({x['page']['page_num'] for x in m})
    check(f'[{tag}] ★茅台命中 3 页（付款/收款身份都能命中）', hit_pages == [1, 2, 3], str(hit_pages))

    # 模糊短词
    m2 = rp.find_matches(pages, ['茅台'], loose=True)
    check(f'[{tag}] 短词「茅台」命中 3 页',
          sorted({x['page']['page_num'] for x in m2}) == [1, 2, 3])

    # 无命中
    m3 = rp.find_matches(pages, ['不存在的公司'], loose=True)
    check(f'[{tag}] ★无主体页/无匹配时结果为空', m3 == [])

    # 导出
    idx = [p - 1 for p in hit_pages]
    out = os.path.join(outdir, f'{tag}__茅台__提取.pdf')
    actual = rp.split_pdf(pdf_path, out, idx, overwrite=True)
    check(f'[{tag}] 导出文件已生成', os.path.exists(actual), actual)

    d = pymupdf.open(actual)
    check(f'[{tag}] 导出页数与命中数一致', d.page_count == len(idx), f'{d.page_count} vs {len(idx)}')
    txt = '\n'.join(p.get_text() for p in d)
    d.close()
    check(f'[{tag}] 导出内容含目标客户', '贵州茅台酒业有限公司' in txt)
    check(f'[{tag}] 导出内容不含无关客户', '重庆物流有限公司' not in txt)
    return hit_pages


def test_multi_query(pdf_path: str, outdir: str):
    section('D2. 多客户批量导出与跨客户页归属')
    texts = rp.extract_all_texts(pdf_path)
    pages = rp.parse_pages(texts, use_parallel=False)

    queries = ['贵州茅台酒业有限公司', '五粮液集团有限公司']
    matches = rp.find_matches(pages, queries)

    groups = {}
    for it in matches:
        groups.setdefault(it['query'], set()).add(it['page']['page_num'])
    groups = {q: sorted(v) for q, v in groups.items()}

    check('茅台归入 3 页', groups.get('贵州茅台酒业有限公司') == [1, 2, 3],
          str(groups.get('贵州茅台酒业有限公司')))
    check('五粮液归入 1 页（第2页付款人）', groups.get('五粮液集团有限公司') == [2],
          str(groups.get('五粮液集团有限公司')))
    # 第 2 页既是茅台的（收款人）也是五粮液的（付款人）→ 两个客户的文件里都该有
    check('★跨客户共有的第2页，两个客户各自都保留（不互相吞并）',
          2 in groups.get('贵州茅台酒业有限公司', []) and 2 in groups.get('五粮液集团有限公司', []))

    total_files = 0
    for q, pns in groups.items():
        fn = os.path.join(outdir, f'batch__{rp.sanitize_filename(q)}.pdf')
        rp.split_pdf(pdf_path, fn, [p - 1 for p in pns], overwrite=True)
        ok = os.path.exists(fn)
        total_files += 1 if ok else 0
        check(f'批量导出 {q} 成功', ok)

    # 合并模式：整体去重，页数 = 并集大小
    union = sorted({p for v in groups.values() for p in v})
    merged = os.path.join(outdir, 'batch__合并.pdf')
    rp.split_pdf(pdf_path, merged, [p - 1 for p in union], overwrite=True)
    d = pymupdf.open(merged)
    check('合并模式整体去重（3 页而非 4 页）', d.page_count == 3, f'{d.page_count}')
    d.close()


def test_rename_and_sanitize(outdir: str):
    section('G. 文件名与重名处理')
    p = os.path.join(outdir, 'dup.pdf')
    # 直接造空文件验证 unique_path
    open(p, 'w').close()
    p1 = rp.unique_path(p)
    check('重名自动加 (1)', p1.endswith('dup(1).pdf'), p1)
    open(p1, 'w').close()
    p2 = rp.unique_path(p)
    check('重名继续加 (2)', p2.endswith('dup(2).pdf'), p2)

    check('非法字符被替换', rp.sanitize_filename('贵州/茅台:酒业*有限?公司') == '贵州_茅台_酒业_有限_公司',
          rp.sanitize_filename('贵州/茅台:酒业*有限?公司'))
    check('空名称有兜底', rp.sanitize_filename('') == '未命名')
    check('Windows 保留名被转义', rp.sanitize_filename('CON').startswith('_'), rp.sanitize_filename('CON'))


def test_errors(tmpdir: str):
    section('F. 异常处理')
    # 加密 PDF
    enc = os.path.join(tmpdir, 'enc.pdf')
    doc = pymupdf.open()
    doc.new_page().insert_text((72, 100), 'hello')
    doc.save(enc, encryption=pymupdf.PDF_ENCRYPT_AES_256, owner_pw='o', user_pw='u')
    doc.close()

    try:
        rp.extract_all_texts(enc)
        check('加密 PDF 应抛 PDFEncryptedError', False, '未抛出异常')
    except rp.PDFEncryptedError:
        check('加密 PDF 抛 PDFEncryptedError', True)
    except Exception as ex:
        check('加密 PDF 抛 PDFEncryptedError', False, f'抛出了 {type(ex).__name__}: {ex}')

    # 损坏文件
    bad = os.path.join(tmpdir, 'bad.pdf')
    with open(bad, 'wb') as f:
        f.write(b'%PDF-1.4 this is not a real pdf')
    try:
        rp.extract_all_texts(bad)
        check('损坏 PDF 应抛 PDFUnreadableError', False, '未抛出异常')
    except (rp.PDFUnreadableError, rp.PDFEncryptedError):
        check('损坏 PDF 抛 PDFUnreadableError', True)
    except Exception as ex:
        check('损坏 PDF 抛 PDFUnreadableError', False, f'抛出了 {type(ex).__name__}: {ex}')

    # 空页导出应报错
    try:
        rp.split_pdf(enc, os.path.join(tmpdir, 'x.pdf'), [], overwrite=True)
        check('空页码导出应抛 ValueError', False, '未抛出异常')
    except Exception as ex:
        check('空页码导出抛 ValueError', isinstance(ex, ValueError), type(ex).__name__)


def test_scan_detect(pdf_path: str):
    section('F2. 扫描件（无文本层）识别')
    doc = pymupdf.open()
    for _ in range(3):
        doc.new_page()
    scan = os.path.join(tempfile.gettempdir(), 'scan_sample.pdf')
    doc.save(scan)
    doc.close()

    pages = rp.parse_pages(rp.extract_all_texts(scan), use_parallel=False)
    check('扫描件每页 is_text_page=False', all(not p['is_text_page'] for p in pages))
    check('扫描件无候选名称', rp.list_all_parties(pages) == [])
    check('扫描件不出现在匹配结果中', rp.find_matches(pages, ['任何公司']) == [])
    try:
        os.remove(scan)
    except OSError:
        pass


def build_scan_pdf(src_pdf: str, out_path: str, dpi: int = 200):
    """把带文本层的样本渲染成图片再嵌入新 PDF —— 模拟银行扫描件（无文本层）。"""
    src = pymupdf.open(src_pdf)
    doc = pymupdf.open()
    mat = pymupdf.Matrix(dpi / 72.0, dpi / 72.0)
    for page in src:
        pix = page.get_pixmap(matrix=mat, alpha=False)
        p = doc.new_page(width=page.rect.width, height=page.rect.height)
        p.insert_image(p.rect, stream=pix.tobytes('png'))
    doc.save(out_path)
    doc.close()
    src.close()


def test_ocr(pdf_path: str, outdir: str):
    section('H. OCR（可选功能，未安装引擎则跳过）')
    name, desc = rp.ocr_probe()
    if name is None:
        print('  [SKIP] 未安装 OCR 引擎（pip install rapidocr_onnxruntime 可启用）')
        return

    check(f'OCR 引擎探测: {desc}', name in ('rapidocr', 'tesseract'))

    scan = os.path.join(outdir, '扫描件_样本.pdf')
    build_scan_pdf(pdf_path, scan)

    # 关闭 OCR：扫描页保持无文本
    texts_off, info_off = rp.build_texts(scan, ocr=False)
    check('不勾选 OCR 时扫描页仍无文本',
          info_off['empty_pages'] == info_off['total'] and info_off['ocr_used'] is False,
          str(info_off))

    # 开启 OCR：全链路（4 页回单应全部识别；第 5 页是刻意留的空白页，无内容属正常）
    texts, info = rp.build_texts(scan, ocr=True)
    check('勾选 OCR 后 4 页回单扫描页全部产出文本',
          info['ocr_pages'] == 4 and info['total'] == 5, str(info))

    pages = rp.parse_pages(texts, use_parallel=False)
    cands = rp.list_all_parties(pages)
    check('OCR 文本能识别出客户名称（≥3 个主体）', len(cands) >= 3, str(cands))
    check('OCR 文本含目标客户「茅台」', any('茅台' in c for c in cands), str(cands))

    m = rp.find_matches(pages, ['茅台'], loose=True)
    hit = sorted({x['page']['page_num'] for x in m})
    check('★OCR 后匹配茅台命中 3 页', hit == [1, 2, 3], str(hit))

    # OCR 结果可用于拆分导出
    idx = [p - 1 for p in hit]
    out = os.path.join(outdir, 'OCR__茅台__提取.pdf')
    actual = rp.split_pdf(scan, out, idx, overwrite=True)
    check('OCR 结果可正常拆分导出', os.path.exists(actual))
    d = pymupdf.open(actual)
    check('OCR 导出 PDF 为 3 页（扫描原图）', d.page_count == 3, str(d.page_count))
    d.close()

    # get_ocr_engine 二次调用走缓存（不再重复加载）
    k1, _ = rp.get_ocr_engine()
    k2, _ = rp.get_ocr_engine()
    check('引擎懒加载缓存生效', k1 == k2 == name)


def test_parallel_actually_used():
    section('E2. 多进程确实生效')
    texts = [SAMPLE_PAGE] * (rp.PARALLEL_THRESHOLD + 5)
    called = {'n': 0}

    def prog(d, t):
        called['n'] += 1

    pages = rp.parse_pages(texts, progress=prog, use_parallel=True)
    check('多进程返回全部页', len(pages) == len(texts))
    check('多进程结果正确（首尾页主体一致）',
          pages[0]['parties'] == ['贵州茅台酒业有限公司', '遵义粮油贸易有限公司']
          and pages[-1]['parties'] == pages[0]['parties'])
    check('进度回调被调用', called['n'] > 0, str(called['n']))
    check('页数低于阈值时也走串行（不报错）',
          len(rp.parse_pages([SAMPLE_PAGE], use_parallel=True)) == 1)


# ─────────────────────────────────────────────────────────────────────────────

def main():
    tmpdir = tempfile.mkdtemp(prefix='gzbank_selftest_')
    outdir = os.path.join(tmpdir, 'out')
    os.makedirs(outdir, exist_ok=True)
    pdf_path = os.path.join(tmpdir, '贵州银行回单_样本.pdf')

    print(f'\n临时目录: {tmpdir}')
    print(f'阈值 PARALLEL_THRESHOLD = {rp.PARALLEL_THRESHOLD}, MAX_WORKERS = {rp.MAX_WORKERS}')

    try:
        ok_font = build_sample_pdf(pdf_path)
        print(f'样本 PDF 已生成: {pdf_path}（中文字体: {"可用" if ok_font else "不可用"}）')

        test_normalize()
        test_extract()
        test_end_to_end(pdf_path, parallel=False, outdir=outdir, tag='串行')
        test_end_to_end(pdf_path, parallel=True, outdir=outdir, tag='并行')
        test_multi_query(pdf_path, outdir)
        test_rename_and_sanitize(outdir)
        test_errors(tmpdir)
        test_scan_detect(pdf_path)
        test_parallel_actually_used()
        test_ocr(pdf_path, outdir)
    except Exception:
        print('\n!!! 测试过程发生异常 !!!')
        traceback.print_exc()
        FAIL.append(('测试框架异常', traceback.format_exc()))
    finally:
        print()
        print('=' * 72)
        print(f'结果：通过 {len(PASS)} 项，失败 {len(FAIL)} 项')
        if FAIL:
            print('-' * 72)
            for name, extra in FAIL:
                print(f'  失败: {name}  {extra}')
        print('=' * 72)

        # 保留一次输出样本供人工查看
        keep = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'selftest_output')
        try:
            if os.path.isdir(keep):
                shutil.rmtree(keep)
            shutil.copytree(outdir, keep)
            shutil.copy2(pdf_path, os.path.join(keep, '源样本.pdf'))
            print(f'\n样本与导出结果已保留到: {keep}')
        except Exception as ex:
            print(f'保留样本失败: {ex}')

        shutil.rmtree(tmpdir, ignore_errors=True)

    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main())
