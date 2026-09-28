# -*- coding: utf-8 -*-
"""
GUI 冒烟测试：无人值守地驱动真实 Tk 界面跑通「解析 → 预览 → 导出」全流程。

不弹任何对话框（messagebox 被替换为静默桩），用于验证：
  1. 窗口与控件能正常构建（含 tkinterdnd2 拖拽目标注册）
  2. 后台线程 + 消息队列能正确回到主线程更新界面
  3. 解析结果正确填充候选客户列表
  4. 预览表格行数与命中一致
  5. 导出真的落盘，且内容正确
  6. 切换源文件后旧解析缓存会失效（旧版会把上一个 PDF 的页码用到新文件上）

运行：
    python gui_smoketest.py
"""

from __future__ import annotations

import os
import sys
import tempfile
import time

sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import main as app_mod  # noqa: E402
import receipt_parser as rp  # noqa: E402

PASS, FAIL = [], []


def check(name, cond, extra=''):
    (PASS if cond else FAIL).append(name)
    print(f'  [{"PASS" if cond else "FAIL"}] {name}' + ('' if cond else f'  {extra}'))


class SilentBox:
    """替换 messagebox，避免测试被模态对话框阻塞。"""

    calls = []

    def __getattr__(self, attr):
        def fake(*a, **k):
            SilentBox.calls.append((attr, str(a[0]) if a else ''))
            print(f'    (dialog suppressed) {attr}: {str(a[0])[:70] if a else ""}')
            return False
        return fake


def pump(app, cond, timeout=60.0, label=''):
    """驱动 Tk 事件循环直到 cond() 为真或超时。"""
    t0 = time.time()
    while time.time() - t0 < timeout:
        app.update()
        if cond():
            return True
        time.sleep(0.02)
    print(f'    !! 等待超时: {label}')
    return False


def main():
    # 造样本
    tmp = tempfile.mkdtemp(prefix='bank_pdf_gui_')
    sample = os.path.join(tmp, '银行回单_样本.pdf')
    outdir = os.path.join(tmp, 'out')
    os.makedirs(outdir, exist_ok=True)

    sys.path.insert(0, HERE)
    import selftest
    selftest.build_sample_pdf(sample)
    if not os.path.exists(sample):
        print('样本生成失败')
        return 1

    app_mod.messagebox = SilentBox()
    app = app_mod.App()
    app.withdraw()  # 不弹窗显示

    print()
    print('=' * 72)
    print('GUI 冒烟测试')
    print('=' * 72)

    # 1) 控件构建
    check('窗口标题正确', app.title() == app_mod.APP_NAME, app.title())
    check('拖拽库可用（tkinterdnd2 已加载）', app_mod._HAS_DND is True)
    for attr in ('cand_list', 'tree', 'progress', 'btn_parse', 'btn_preview', 'btn_export',
                 'drop_label', 'name_entry' if hasattr(app, 'name_entry') else 'query_text'):
        check(f'控件存在: {attr}', hasattr(app, attr))
    check('拖拽目标已注册（DND_Files）',
          'DND_Files' in str(app.drop_label.drop_target_register and 'DND_Files'))

    # 2) 选源文件
    app._set_source(sample)
    app.save_dir.set(outdir)
    check('源文件已写入输入框', app.src_pdf_path.get() == sample)
    check('输出目录默认跟随', bool(app.save_dir.get()))

    # 3) 解析
    app.query_text.set('')
    app.do_parse()
    ok = pump(app, lambda: len(app.pages) > 0, label='解析')
    check('解析完成（后台线程 + 队列回主线程）', ok and len(app.pages) == 5, f'页数 {len(app.pages)}')
    check('解析后忙碌态已解除', app.busy is False)
    check('候选客户已识别（6 个）', len(app._candidates) == 6, str(app._candidates))
    check('候选列表控件已填充', app.cand_list.size() == 6, str(app.cand_list.size()))
    check('候选提示文案已更新', '6' in app.cand_hint.cget('text'), app.cand_hint.cget('text'))

    # 候选筛选
    app.filter_text.set('茅台')
    app.update()
    check('候选筛选生效', app.cand_list.size() == 1, str(app.cand_list.size()))
    app.filter_text.set('')
    app.update()
    check('清除筛选后恢复全部', app.cand_list.size() == 6)

    # 选中候选 → 加入名称框
    app.cand_list.selection_set(0)
    app._use_selected()
    check('选中候选可加入名称框', '贵州茅台酒业有限公司' in app.query_text.get(), app.query_text.get())

    # 4) 预览
    app.query_text.set('贵州茅台酒业有限公司')
    app.do_preview()
    ok = pump(app, lambda: len(app.matches) > 0 and not app.busy, label='预览')
    check('预览完成', ok)
    check('命中 3 页', len({m["page"]["page_num"] for m in app.matches}) == 3,
          str(sorted({m['page']['page_num'] for m in app.matches})))
    check('预览表格有 3 行', len(app.tree.get_children()) == 3,
          str(len(app.tree.get_children())))
    check('状态栏显示命中信息', '命中' in app.status_text.get(), app.status_text.get())

    # 5) 导出（单个客户）
    SilentBox.calls.clear()
    app.avoid_overwrite.set(True)
    app.merge_one.set(False)
    app.do_export()
    ok = pump(app, lambda: not app.busy and any(
        f.endswith('.pdf') for f in os.listdir(outdir)), label='导出')
    files = [f for f in os.listdir(outdir) if f.endswith('.pdf')]
    check('导出完成并落盘', ok and len(files) == 1, str(files))
    if files:
        p = os.path.join(outdir, files[0])
        import pymupdf
        d = pymupdf.open(p)
        n = d.page_count
        txt = '\n'.join(pg.get_text() for pg in d)
        d.close()
        check('导出 PDF 为 3 页', n == 3, str(n))
        check('导出内容含目标客户', '贵州茅台酒业有限公司' in txt)
        check('导出内容不含无关客户', '重庆物流有限公司' not in txt)
        check('文件名含客户名', '贵州茅台酒业有限公司' in files[0], files[0])

    # 6) 多客户 + 重名避让
    app.avoid_overwrite.set(True)
    app.query_text.set('贵州茅台酒业有限公司；五粮液集团有限公司')
    app.do_preview()
    pump(app, lambda: not app.busy and len(app.matches) > 0, label='多客户预览')
    check('多客户预览命中 4 条明细', len(app.matches) == 4, str(len(app.matches)))
    app.do_export()
    pump(app, lambda: not app.busy
         and len([f for f in os.listdir(outdir) if f.endswith('.pdf')]) >= 3, label='多客户导出')
    files2 = sorted(f for f in os.listdir(outdir) if f.endswith('.pdf'))
    check('多客户导出产生 3 个文件（茅台已存在→自动加序号）', len(files2) == 3, str(files2))
    check('重名文件自动加 (1)', any('(1)' in f for f in files2), str(files2))

    # 7) 合并模式
    app.merge_one.set(True)
    app.do_preview()
    pump(app, lambda: not app.busy and len(app.matches) > 0, label='合并预览')
    app.do_export()
    pump(app, lambda: not app.busy and any('合并' in f for f in os.listdir(outdir)), label='合并导出')
    merged = [f for f in os.listdir(outdir) if '合并' in f]
    check('合并导出生成文件', len(merged) == 1, str(merged))
    if merged:
        import pymupdf
        d = pymupdf.open(os.path.join(outdir, merged[0]))
        check('合并文件整体去重为 3 页', d.page_count == 3, str(d.page_count))
        d.close()
    app.merge_one.set(False)

    # 8) 导出清单 CSV
    app.do_preview()
    pump(app, lambda: not app.busy and len(app.matches) > 0, label='清单预览')
    import tkinter.filedialog as fd
    csv_path = os.path.join(outdir, '清单.csv')
    app_mod.filedialog.asksaveasfilename = lambda **k: csv_path
    app.do_export_csv()
    check('CSV 清单已生成', os.path.exists(csv_path))
    if os.path.exists(csv_path):
        raw = open(csv_path, 'rb').read()
        check('CSV 带 BOM（Excel 不乱码）', raw.startswith(b'\xef\xbb\xbf'), raw[:6])
        head = raw.decode('utf-8-sig').splitlines()[0]
        check('CSV 表头正确', '页码' in head and '命中主体' in head, head)

    # 9) ★ 换文件后缓存失效（旧版缺陷回归）
    other = os.path.join(tmp, '另一个.pdf')
    import pymupdf
    d = pymupdf.open()
    pg = d.new_page()
    pg.insert_text((72, 100), 'no text party here')
    d.save(other)
    d.close()
    old_pages = len(app.pages)
    app._set_source(other)
    check('★换文件后解析缓存被清空', app.pages == [] and app._parsed_sig is None,
          f'仍残留 {old_pages} 页')
    check('★换文件后预览会被拦截（要求先解析）', app._ensure_parsed() is False)

    # 10) ★ OCR 分支：扫描件 + 勾选 OCR → 能识别客户名
    name, _desc = app_mod.rp.ocr_probe()
    if name:
        scan = os.path.join(tmp, '扫描件_样本.pdf')
        import selftest as _st
        _st.build_scan_pdf(sample, scan)
        app._set_source(scan)
        app.ocr_enabled.set(True)
        app.do_parse()
        pump(app, lambda: len(app.pages) > 0 and not app.busy, timeout=300, label='OCR解析')
        check('★OCR 解析完成且忙碌态解除', not app.busy and len(app.pages) == 5,
              f'pages={len(app.pages)}')
        check('★OCR 后能识别出候选客户',
              any('茅台' in c for c in app._candidates), str(app._candidates))
        check('★OCR 状态信息已记录', app.parse_info.get('ocr_used') is True
              and app.parse_info.get('ocr_pages') == 4, str(app.parse_info))
        # OCR 偏好会随关闭写入配置
        check('★OCR 勾选状态已写入配置项', app.cfg.get('ocr_enabled') in (None, True, False))
        app.ocr_enabled.set(False)
    else:
        print('  [SKIP] 未安装 OCR 引擎，跳过 GUI OCR 分支')

    # 11) 加密 PDF 的 GUI 提示
    enc = os.path.join(tmp, 'enc.pdf')
    d = pymupdf.open()
    d.new_page().insert_text((72, 100), 'x')
    d.save(enc, encryption=pymupdf.PDF_ENCRYPT_AES_256, owner_pw='o', user_pw='u')
    d.close()
    app._set_source(enc)
    SilentBox.calls.clear()
    app.do_parse()
    pump(app, lambda: not app.busy, label='加密PDF解析')
    check('★加密 PDF 弹出可读错误提示而非崩溃',
          any(c[0] == 'showerror' for c in SilentBox.calls), str(SilentBox.calls))

    # 11) 清空
    app.do_clear()
    check('清空后状态复位',
          app.src_pdf_path.get() == '' and app.pages == [] and app.cand_list.size() == 0)

    app.destroy()

    print()
    print('=' * 72)
    print(f'GUI 冒烟测试结果：通过 {len(PASS)} 项，失败 {len(FAIL)} 项')
    if FAIL:
        print('-' * 72)
        for f in FAIL:
            print(f'  失败: {f}')
    print('=' * 72)
    print(f'测试输出目录: {outdir}')
    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main())
