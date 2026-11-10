# -*- coding: utf-8 -*-
"""小红书创作中心 · 自动存草稿（不发布，只进草稿箱）。

为什么只存草稿：发布是不可逆动作，机器人点发布风险高；草稿由人最后检查一遍再发。

用法:
    py xhs_draft.py login                 # 扫码登录创作中心
    py xhs_draft.py draft note.json       # 把 json 里的笔记存成草稿
    py xhs_draft.py probe                 # 登录后排查页面结构

note.json 格式:
    {
      "title": "麂皮鞋脏了别扔，3分钟救回来",
      "body": "正文……",
      "tags": ["麂皮清洁", "洗鞋"],
      "images": ["麂皮清洁剂-图/01_text.png"]
    }

注意:
    - 创作中心和主站是两个域，cookie 不通用，必须单独登录一次
    - 图片是本地文件，上传后要等缩略图出现再点存草稿
"""
import argparse
import json
import os
import sys

from playwright.sync_api import sync_playwright

PROFILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".xhs-profile")
PUBLISH = "https://creator.xiaohongshu.com/publish/publish?source=official"
LOGIN = "https://creator.xiaohongshu.com/login?source=official"
ARGS = ["--disable-blink-features=AutomationControlled"]


def open_ctx(p, headless=False):
    os.makedirs(PROFILE, exist_ok=True)
    return p.chromium.launch_persistent_context(
        PROFILE, channel="chrome", headless=headless,
        viewport={"width": 1440, "height": 900}, args=ARGS)


def goto(pg, url, wait=8000):
    pg.goto(url, timeout=60000, wait_until="commit")
    pg.wait_for_timeout(wait)


def on_login_page(pg):
    try:
        return "/login" in pg.url
    except Exception:
        return True


QR_ICON = (1263, 305)  # 登录卡片右上角的二维码图标（1440x900 视口下实测）


def ensure_qr(pg):
    """默认是短信登录，要先切到扫码。"""
    try:
        if "扫码" in pg.inner_text("body"):
            return True
    except Exception:
        pass
    pg.mouse.click(*QR_ICON)
    pg.wait_for_timeout(3000)
    try:
        return "扫码" in pg.inner_text("body")
    except Exception:
        return False


def wait_login(pg, minutes=10):
    """等扫码，以离开 /login 为准。"""
    pg.bring_to_front()
    print("=" * 58)
    print(" 创作中心登录窗口已打开（独立 Chrome 窗口）")
    print(f" 请扫码登录，最多等 {minutes} 分钟")
    print("=" * 58)
    ticks = 0
    while ticks < minutes * 60:
        pg.wait_for_timeout(5000)
        ticks += 5
        if not on_login_page(pg):
            print(f"✅ 已登录（{ticks}s）")
            return True
        if ticks % 30 == 0:
            pg.bring_to_front()
            print(f"   …还在等扫码（{ticks}s）")
    print("⚠️  超时未登录")
    return False


def find_editor(pg):
    """标题输入框 + 正文编辑区。小红书改版频繁，按 placeholder / contenteditable 兜底。"""
    title = (pg.query_selector("input[placeholder*='标题']")
             or pg.query_selector("input[placeholder*='填写标题']"))
    body = (pg.query_selector("div[contenteditable='true']")
            or pg.query_selector("textarea[placeholder*='正文']"))
    return title, body


def click_text(pg, text):
    for sel in (f"button:has-text('{text}')", f"div:has-text('{text}')"):
        try:
            el = pg.query_selector(sel)
            if el and el.is_visible():
                el.click()
                return True
        except Exception:
            continue
    return False


def cmd_login(a):
    with sync_playwright() as p:
        ctx = open_ctx(p, a.headless)
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        goto(pg, LOGIN, 6000)
        ensure_qr(pg)
        wait_login(pg, a.minutes)
        ctx.close()


def cmd_draft(a):
    with open(a.file, encoding="utf-8") as f:
        note = json.load(f)
    imgs = [os.path.abspath(p) for p in note.get("images") or [] if os.path.exists(p)]
    if not imgs:
        sys.exit("note.json 里没有可用的图片路径")
    body = note.get("body", "").strip()
    tags = note.get("tags") or []
    if tags:
        body += "\n\n" + " ".join(f"#{t}" for t in tags)

    with sync_playwright() as p:
        ctx = open_ctx(p, a.headless)
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        goto(pg, PUBLISH, 9000)
        if on_login_page(pg):
            goto(pg, LOGIN, 5000)
            ensure_qr(pg)
            if not wait_login(pg, a.minutes):
                ctx.close()
                return
            goto(pg, PUBLISH, 9000)

        # 1. 选「上传图文」，再塞图片
        click_text(pg, "上传图文")
        pg.wait_for_timeout(1500)
        box = pg.query_selector("input[type=file]")
        if not box:
            print("❗ 没找到上传框。跑 `py xhs_draft.py probe` 看结构")
            ctx.close()
            return
        box.set_input_files(imgs)
        print(f"上传 {len(imgs)} 张图，等待缩略图…")
        pg.wait_for_timeout(6000 + 1500 * len(imgs))

        # 2. 标题 + 正文
        title_el, body_el = find_editor(pg)
        if not title_el or not body_el:
            print("❗ 没找到标题或正文输入框。跑 probe")
            ctx.close()
            return
        title_el.click()
        title_el.fill(note.get("title", "")[:20])
        body_el.click()
        body_el.fill(body)
        pg.wait_for_timeout(1500)

        # 3. 存草稿（不点发布）
        if not click_text(pg, "存草稿") and not click_text(pg, "保存草稿"):
            print("❗ 没找到「存草稿」按钮。跑 probe")
            ctx.close()
            return
        pg.wait_for_timeout(4000)
        print("✅ 已存草稿，去创作中心 → 草稿箱 检查")
        if a.keep:
            try:
                while True:
                    pg.wait_for_timeout(1000)
            except KeyboardInterrupt:
                pass
        ctx.close()


def cmd_probe(a):
    with sync_playwright() as p:
        ctx = open_ctx(p, a.headless)
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        goto(pg, PUBLISH, 10000)
        print("url:", pg.url[:90])
        print("file inputs:", len(pg.query_selector_all("input[type=file]")))
        for sel in ("input[placeholder*='标题']", "div[contenteditable='true']",
                    "textarea", "button"):
            els = pg.query_selector_all(sel)
            print(f"{sel:<32} {len(els)}", [ (e.inner_text() or '')[:12] for e in els[:4]])
        print("body:", pg.inner_text("body")[:400].replace("\n", " | "))
        ctx.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["login", "draft", "probe"])
    ap.add_argument("file", nargs="?")
    ap.add_argument("--minutes", type=int, default=10)
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--keep", action="store_true")
    a = ap.parse_args()
    if a.cmd == "draft" and not a.file:
        sys.exit("要带 note.json")
    {"login": cmd_login, "draft": cmd_draft, "probe": cmd_probe}[a.cmd](a)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
