# -*- coding: utf-8 -*-
"""小红书创作中心 · 自动填笔记并存草稿 / 发布。

默认只存草稿；加 --publish 才真的发布（不可逆）。
重要：小红书网页版的草稿存在「当前浏览器本地」，不同步到 App，也不上云。

用法:
    py xhs_draft.py login                 # 扫码登录创作中心
    py xhs_draft.py draft note.json       # 把 json 里的笔记存成草稿（默认）
    py xhs_draft.py draft note.json --publish   # 填完直接发布（不可逆）
    py xhs_draft.py probe                 # 登录后排查页面结构

note.json 格式:
    {
      "title": "麂皮鞋脏了别扔，3分钟救回来",
      "body": "正文……",
      "tags": ["麂皮清洁", "洗鞋"],
      "images": ["麂皮清洁剂-图/01_text.png"]
    }

注意:
    - 草稿是本地的：App 里看不到，换浏览器也看不到，只有 .xhs-profile 这个配置能看到
    - 创作中心和主站是两个域，cookie 不通用，必须单独登录一次
    - 图片是本地文件，上传后要等缩略图出现再点存草稿
"""
import argparse
import json
import os
import re
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


def safe_click(pg, x, y):
    """点之前先确认坐标命中发布组件，避免改版后点错。"""
    tag = pg.evaluate("([x, y]) => document.elementsFromPoint(x, y)[0]?.tagName", [x, y])
    if tag != "XHS-PUBLISH-BTN":
        print(f"⚠️  坐标 ({x},{y}) 命中 {tag}，不是发布组件，已中止")
        return False
    pg.mouse.click(x, y)
    return True


def footer_click(pg, ratio):
    """点底部按钮。y 按组件实际位置算（页面高度不同会移动），x 用比例。"""
    pg.wait_for_timeout(1500)
    el = pg.query_selector("xhs-publish-btn")
    if not el:
        print("  ❗ 没找到发布组件")
        return False
    box = el.bounding_box()
    if not box:
        return False
    x = int(box["x"] + box["width"] * ratio)
    y = int(box["y"] + box["height"] / 2)
    tag = pg.evaluate("([x, y]) => document.elementsFromPoint(x, y)[0]?.tagName", [x, y])
    if tag != "XHS-PUBLISH-BTN":
        print(f"  ⚠️  ({x},{y}) 命中 {tag}，滚动后重试")
        pg.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        pg.wait_for_timeout(1500)
        box = el.bounding_box() or box
        x = int(box["x"] + box["width"] * ratio)
        y = int(box["y"] + box["height"] / 2)
    pg.mouse.click(x, y)
    return True


def draft_count(pg):
    m = re.search(r"草稿箱\((\d+)\)", pg.inner_text("body"))
    return int(m.group(1)) if m else -1


def find_editor(pg):
    """标题输入框 + 正文编辑区。小红书改版频繁，按 placeholder / contenteditable 兜底。"""
    title = (pg.query_selector("input[placeholder*='标题']")
             or pg.query_selector("input[placeholder*='填写标题']"))
    body = (pg.query_selector("div[contenteditable='true']")
            or pg.query_selector("textarea[placeholder*='正文']"))
    return title, body


def click_text(pg, text):
    """点可见文字。普通点击失败就退回 JS 点击（弹窗遮挡时也能点）。"""
    for sel in (f"text={text}", f"button:has-text('{text}')",
                f"[class*=tab]:has-text('{text}')", f"div:has-text('{text}')"):
        try:
            el = pg.query_selector(sel)
            if not el:
                continue
            try:
                el.click(timeout=2500)
            except Exception:
                el.evaluate("e => e.click()")
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


def to_jpeg(paths):
    """转成 JPEG 再传，PNG 一张 1.5MB，4 张传不完会导致「草稿保存失败」。"""
    from PIL import Image
    out = []
    os.makedirs("草稿/_tmp", exist_ok=True)
    for pth in paths:
        if pth.lower().endswith((".jpg", ".jpeg")):
            out.append(pth)
            continue
        dst = os.path.join("草稿/_tmp", os.path.splitext(os.path.basename(pth))[0] + ".jpg")
        Image.open(pth).convert("RGB").save(dst, quality=88, optimize=True)
        out.append(dst)
    return out


def fill_note(ctx, pg, note, a, verify=True):
    """填一篇笔记并提交。返回是否成功。"""
    imgs = [os.path.abspath(x) for x in note.get("images") or [] if os.path.exists(x)]
    imgs = to_jpeg(imgs)
    if not imgs:
        print("  ❗ 没有可用的图片路径")
        return False
    body = note.get("body", "").strip()
    tags = note.get("tags") or []
    if tags:
        body += chr(10) + chr(10) + " ".join(f"#{t}" for t in tags)

    goto(pg, PUBLISH, 9000)
    before = draft_count(pg)
    pg.wait_for_timeout(3000)
    click_text(pg, "上传图文")
    pg.wait_for_timeout(2500)
    try:
        with pg.expect_file_chooser(timeout=15000) as fc:
            if not click_text(pg, "上传图片"):
                raise RuntimeError("没点到「上传图片」")
        fc.value.set_files(imgs)
    except Exception as e:
        print("  ❗ 上传失败:", str(e)[:120])
        return False

    title_el = body_el = None
    for _ in range(40):
        pg.wait_for_timeout(1000)
        title_el, body_el = find_editor(pg)
        if title_el and body_el:
            break
    click_text(pg, "我知道了")
    pg.wait_for_timeout(800)
    if not title_el or not body_el:
        print("  ❗ 没找到标题或正文输入框")
        return False

    title_el.click()
    title_el.fill(note.get("title", "")[:20])
    body_el.click()
    body_el.fill(body)
    pg.wait_for_timeout(1500)

    if a.publish:
        print("  ⚠️  3 秒后点「发布」（不可逆）…")
        pg.wait_for_timeout(3000)
        if not footer_click(pg, 0.60):  # 右：发布
            return False
        pg.wait_for_timeout(6000)
        if click_text(pg, "确认发布") or click_text(pg, "继续发布"):
            pg.wait_for_timeout(5000)
        page = pg.inner_text("body")
        ok = ("发布成功" in page) or ("发布完成" in page) or ("/success" in pg.url)
        print("  ✅ 已发布" if ok else "  ⚠️  没看到成功提示，去笔记管理确认")
        return ok

    pg.wait_for_timeout(18000)  # 等图片传完，传的过程中「暂存离开」点了没反应
    for attempt in range(3):
        if not footer_click(pg, 0.39):  # 左：暂存离开
            return False
        pg.wait_for_timeout(4000)
        if "草稿保存失败" not in pg.inner_text("body"):
            break
        print("  ⚠️  草稿保存失败，等 10 秒重试…")
        pg.wait_for_timeout(10000)
    pg.wait_for_timeout(3000)
    if not verify:
        return True
    print(f"  已提交（保存前草稿箱 {before}）")
    return True


def cmd_draft(a):
    with open(a.file, encoding="utf-8") as f:
        data = json.load(f)
    notes = data if isinstance(data, list) else [data]

    ok = 0
    for i, note in enumerate(notes, 1):
        print(f"[{i}/{len(notes)}] {note.get('title','')[:20]}  图 {len(note.get('images') or [])} 张")
        # 每篇用全新上下文：小红书会把上一篇草稿载进编辑器，复用页面会找不到上传按钮
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
            if fill_note(ctx, pg, note, a):
                ok += 1
            ctx.close()
    print(f"完成 {ok}/{len(notes)}")


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
    ap.add_argument("--publish", action="store_true",
                    help="填完直接点「发布」（不可逆）。不加这个参数只存草稿")
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
