# -*- coding: utf-8 -*-
"""小红书浏览器自动化 —— 用于虚拟品选品的竞品调研。

用持久化用户目录（.xhs-profile/），登录一次后续一直有效。

用法:
    py xhs_browser.py login                        # 打开浏览器，扫码登录，登录后自动退出
    py xhs_browser.py search "期末评语"            # 搜商品，抓结果
    py xhs_browser.py search "期末评语" --notes    # 搜笔记（看内容形式）
    py xhs_browser.py probe "期末评语"             # 打印页面结构，用于排查

注意:
    - 小红书是 SPA，必须用 wait_until="commit"，等 domcontentloaded 会超时
    - 登录态存在 .xhs-profile/，已被 .gitignore 挡住
"""
import argparse
import json
import os
import sys
import time

from playwright.sync_api import sync_playwright

PROFILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".xhs-profile")
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")


def open_ctx(p, headless=False):
    os.makedirs(PROFILE, exist_ok=True)
    return p.chromium.launch_persistent_context(
        PROFILE, channel="chrome", headless=headless,
        viewport={"width": 1440, "height": 900},
        args=["--disable-blink-features=AutomationControlled"],
    )


def goto(pg, url, wait=3000):
    pg.goto(url, timeout=60000, wait_until="commit")
    pg.wait_for_timeout(wait)


def logged_in(pg):
    """页面上没有「登录」按钮，且有头像/消息入口，就算登录了。"""
    try:
        html = pg.content()
    except Exception:
        return False
    return ("登录" not in html) or ("扫码登录" not in html and "me" in pg.url)


def is_logged_in(ctx, pg):
    """只看 cookie：有 web_session 才是真登录。

    别用页面文案判断 —— 未登录时页面不会写「扫码登录」，会误报。下次踩到的就是这条。
    """
    try:
        names = {c["name"] for c in ctx.cookies()}
    except Exception:
        return False
    return any(n in names for n in ("web_session", "customerClientId",
                                     "galaxy_creator_session_id"))


def cmd_login(a):
    with sync_playwright() as p:
        ctx = open_ctx(p, a.headless)
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        goto(pg, "https://www.xiaohongshu.com/explore", 5000)
        print("=" * 62)
        print(" 浏览器窗口已打开（去任务栏找 Chrome）")
        print(" 点右上角「登录」→ 用小红书 App 扫码")
        print(" 登录成功脚本会自动保存并退出，最多等 4 分钟")
        print("=" * 62)
        ok = False
        for i in range(48):
            pg.wait_for_timeout(5000)
            if is_logged_in(ctx, pg):
                print(f"✅ 已登录（{i*5+5}s），web_session 已拿到")
                ok = True
                break
            if i % 6 == 5:
                print(f"   …还在等扫码（{i*5+5}s）")
        if not ok:
            print("⚠️  4 分钟没登录成功，仍会保存当前进度")
        ctx.close()
    print(f"登录态目录: {PROFILE}")


def cards(pg, limit=60):
    """抓搜索结果卡片：链接 + 文本。"""
    out = []
    for sel in ("a[href*='/search_result/']", "a[href*='/goods-detail/']",
                "a[href*='/user/profile/']", "section.note-item", "div.note-item"):
        try:
            els = pg.query_selector_all(sel)
        except Exception:
            continue
        for e in els:
            try:
                href = e.get_attribute("href") or ""
                txt = (e.inner_text() or "").strip().replace("\n", " ")
            except Exception:
                continue
            if txt and len(txt) > 4:
                out.append({"sel": sel, "href": href, "text": txt[:300]})
        if out:
            break
    seen, uniq = set(), []
    for o in out:
        k = o["href"] or o["text"][:60]
        if k in seen:
            continue
        seen.add(k)
        uniq.append(o)
    return uniq[:limit]


def cmd_search(a):
    kw = a.keyword
    url = f"https://www.xiaohongshu.com/search_result?keyword={kw}&source=web_explore_feed"
    with sync_playwright() as p:
        ctx = open_ctx(p, a.headless)
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        if not is_logged_in(ctx, pg):
            print("❌ 未登录（没有 web_session）。先跑: py xhs_browser.py login")
            ctx.close()
            return
        goto(pg, url, 6000)
        pg.wait_for_timeout(3000)
        print(f"标题: {pg.title()}")
        print(f"URL : {pg.url}")
        body = pg.inner_text("body")[:1200]
        print("--- 页面可见文字（前 1200 字）---")
        print(body)
        print("--- 卡片 ---")
        cs = cards(pg, a.limit)
        print(f"抓到 {len(cs)} 个卡片")
        for c in cs[: a.limit]:
            print(f"  [{c['sel'][:28]}] {c['text'][:110]}")
            if c["href"]:
                print(f"      {c['href'][:110]}")
        if a.json:
            with open(a.json, "w", encoding="utf-8") as f:
                json.dump(cs, f, ensure_ascii=False, indent=1)
            print(f"已写出: {a.json}")
        if not a.keep:
            ctx.close()
        else:
            print("浏览器保持打开，按 Ctrl+C 结束")
            try:
                while True:
                    pg.wait_for_timeout(1000)
            except KeyboardInterrupt:
                ctx.close()


def cmd_probe(a):
    with sync_playwright() as p:
        ctx = open_ctx(p, a.headless)
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        goto(pg, f"https://www.xiaohongshu.com/search_result?keyword={a.keyword}", 6000)
        print("=== 所有 tab / 筛选按钮 ===")
        for sel in ("div.channel", "div.tab", "span.channel", "[class*=channel]",
                    "[class*=tab]", "button", "a[href*='type=']"):
            try:
                for e in pg.query_selector_all(sel)[:25]:
                    t = (e.inner_text() or "").strip().replace("\n", "/")
                    if t and len(t) < 30:
                        print(f"  {sel:<22} {t}")
            except Exception:
                pass
        print("\n=== 页面 HTML 长度 ===", len(pg.content()))
        print("=== 是否有登录弹窗 ===", "登录" in pg.content())
        ctx.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["login", "search", "probe"])
    ap.add_argument("keyword", nargs="?")
    ap.add_argument("--notes", action="store_true", help="搜笔记而不是商品")
    ap.add_argument("--limit", type=int, default=30)
    ap.add_argument("--json")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--keep", action="store_true", help="跑完不关浏览器")
    a = ap.parse_args()
    if a.cmd in ("search", "probe") and not a.keyword:
        sys.exit("要带关键词，例如: py xhs_browser.py search \"期末评语\"")
    {"login": cmd_login, "search": cmd_search, "probe": cmd_probe}[a.cmd](a)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
