# -*- coding: utf-8 -*-
"""打开带草稿的浏览器（.xhs-profile），直接停在草稿箱页面。

草稿存在这个配置目录的本地存储里，用别的浏览器看不到。
关掉窗口即结束。
"""
import sys
import xhs_draft as D
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    ctx = D.open_ctx(p, headless=False)
    pg = ctx.pages[0] if ctx.pages else ctx.new_page()
    D.goto(pg, D.PUBLISH, 10000)
    D.click_text(pg, "草稿箱")
    pg.wait_for_timeout(2500)
    D.click_text(pg, "图文笔记")
    pg.wait_for_timeout(2000)
    pg.bring_to_front()
    print("草稿箱已打开，关窗口结束")
    try:
        while True:
            pg.wait_for_timeout(1000)
    except KeyboardInterrupt:
        pass
    finally:
        ctx.close()
