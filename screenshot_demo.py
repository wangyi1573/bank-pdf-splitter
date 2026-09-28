# -*- coding: utf-8 -*-
"""渲染真实界面并截图，用于留档 / 展示。"""

from __future__ import annotations

import os
import sys
import time

sys.stdout.reconfigure(encoding='utf-8')
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import main as app_mod  # noqa: E402
import selftest  # noqa: E402


class SilentBox:
    def __getattr__(self, attr):
        def fake(*a, **k):
            return False
        return fake


def pump(app, seconds=1.0):
    t0 = time.time()
    while time.time() - t0 < seconds:
        app.update()
        time.sleep(0.03)


def main():
    tmp = os.path.join(HERE, 'selftest_output')
    os.makedirs(tmp, exist_ok=True)
    sample = os.path.join(tmp, 'source_sample.pdf')
    if not os.path.exists(sample):
        selftest.build_sample_pdf(sample)

    app_mod.messagebox = SilentBox()
    app = app_mod.App()
    app.geometry('960x860+60+20')
    app.attributes('-topmost', True)
    app.update()
    pump(app, 0.5)

    app._set_source(sample)
    app.save_dir.set(os.path.join(tmp, 'exports'))
    os.makedirs(app.save_dir.get(), exist_ok=True)

    app.do_parse()
    t0 = time.time()
    while time.time() - t0 < 40 and not app.pages:
        app.update()
        time.sleep(0.03)

    app.query_text.set('贵州茅台酒业有限公司；五粮液集团有限公司')
    app.do_preview()
    t0 = time.time()
    while time.time() - t0 < 20 and (app.busy or not app.matches):
        app.update()
        time.sleep(0.03)

    app.cand_list.selection_set(0, 'end')
    pump(app, 1.2)

    # 等 OCR 探测结果显示
    t0 = time.time()
    while time.time() - t0 < 15 and not app.ocr_info.get().startswith('OCR：'):
        app.update()
        time.sleep(0.05)
    pump(app, 0.5)

    shot = os.path.join(HERE, 'ui_preview.png')
    try:
        from PIL import ImageGrab
        x, y = app.winfo_rootx(), app.winfo_rooty()
        w, h = app.winfo_width(), app.winfo_height()
        img = ImageGrab.grab(bbox=(x, y, x + w, y + h), all_screens=True)
        img.save(shot)
        print(f'截图已保存: {shot}  ({img.size[0]}x{img.size[1]})')
    except Exception as ex:
        print(f'截图失败: {ex}')
    finally:
        app.destroy()
    return 0


if __name__ == '__main__':
    sys.exit(main())
