# -*- coding: utf-8 -*-
"""1688 找货源 —— 登录态复用 + 以图搜同款。

为什么要登录：1688 搜索结果页未登录会被跳到登录页，且以图搜款也一样。
为什么要 --no-proxy-server：本机系统代理（127.0.0.1:10809）的出口被 1688 判成
cloud_ip_bl，直接拒绝；绕开代理反而能正常访问。踩过一次，别再删。

用法:
    py ali1688_browser.py login                      # 扫码登录，登录后自动退出
    py ali1688_browser.py search "爆炸盐"            # 关键词搜同款
    py ali1688_browser.py image _cover.jpg           # 以图搜同款
    py ali1688_browser.py image --cover pi_xxx       # 直接从银河拉封面图再搜
    py ali1688_browser.py probe "爆炸盐"             # 页面结构排查

    --json out.json   结果存 JSON
    --keep            跑完不关浏览器（手动过滑块用）
"""
import argparse
import json
import os
import re
import sys
import urllib.parse
import urllib.request

from playwright.sync_api import sync_playwright

PROFILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".ali1688-profile")
COOKIES = os.path.join(PROFILE, "cookies.json")  # 1688 的 cookie2 是 session cookie，Chrome 关掉就丢，必须自己存
ARGS = ["--disable-blink-features=AutomationControlled", "--no-proxy-server"]
SEARCH_URL = "https://s.1688.com/selloffer/offer_search.htm?charset=utf8&keywords="
IMAGE_URL = "https://air.1688.com/kapp/1688-search/pc-image-search/?tab=imageSearch"
COVER_BASE = "https://aiyinhe.art"

PRICE_RE = re.compile(r"¥(\d+(?:\.\d{1,2})?)")  # 先去掉所有空白再匹配，1688 的价格 DOM 会被拆成 ¥ / 3 / .90
MOQ_RE = re.compile(r"(\d+)\s*\+?\s*(?:件|个|套|包|盒|支|条|双|只|台|瓶|张|箱|袋)")
SHOP_RE = re.compile(r"(商行|有限公司|公司|工厂|厂|店|商贸|经营部)")
SALES_RE = re.compile(r"(?:已售|成交|月销)\s*([\d.]+万?)")
PROV_RE = re.compile(r"(广东|浙江|江苏|福建|山东|河北|河南|湖北|湖南|安徽|江西|四川|重庆|"
                     r"上海|北京|天津|广西|云南|贵州|陕西|山西|辽宁|吉林|黑龙江|"
                     r"内蒙古|新疆|甘肃|宁夏|青海|海南|西藏|香港|澳门|台湾)")
STRENGTH = ("实力商家", "超级工厂", "诚信通", "源头工厂")


def open_ctx(p, headless=False):
    os.makedirs(PROFILE, exist_ok=True)
    ctx = p.chromium.launch_persistent_context(
        PROFILE, channel="chrome", headless=headless,
        viewport={"width": 1440, "height": 900}, args=ARGS)
    if os.path.exists(COOKIES):
        try:
            with open(COOKIES, encoding="utf-8") as f:
                ctx.add_cookies(json.load(f))
        except Exception as e:
            print(f"⚠️  登录态文件读取失败，忽略: {e}")
    return ctx


def save_cookies(ctx):
    try:
        with open(COOKIES, "w", encoding="utf-8") as f:
            json.dump(ctx.cookies(), f, ensure_ascii=False)
    except Exception as e:
        print(f"⚠️  登录态保存失败: {e}")


def logged_in(ctx):
    """cookie2 是阿里系登录态的核心 cookie，有它才算登录。"""
    try:
        names = {c["name"] for c in ctx.cookies()}
    except Exception:
        return False
    return "cookie2" in names  # tfstk 长期有效但登出后还在，用它判断会误报


def goto(pg, url, wait=8000):
    pg.goto(url, timeout=60000, wait_until="commit")
    pg.wait_for_timeout(wait)


def parse_card(text):
    """从一张卡的可见文字里抠出字段。1688 的价格在 DOM 里是拆开的多段，先拼起来。"""
    flat = re.sub(r"\s+", " ", text).strip()
    m = PRICE_RE.search(re.sub(r"\s+", "", text))
    price = float(m.group(1)) if m else None
    moq = MOQ_RE.search(flat)
    sales = SALES_RE.search(flat)
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    shop = next((ln for ln in lines if SHOP_RE.search(ln)), "")
    title = next((ln for ln in lines if len(ln) > 4 and "¥" not in ln
                  and not ln.isdigit() and ln != shop), "")
    prov = PROV_RE.search(flat)
    return {
        "title": title[:60],
        "price": price,
        "moq": int(moq.group(1)) if moq else None,
        "sales": sales.group(1) if sales else "",
        "shop": shop[:40],
        "province": prov.group(1) if prov else "",
        "strength": next((x for x in STRENGTH if x in flat), ""),
    }


SALES_ONLY = re.compile(r"^[\d.]+\s*万?\+?\s*件$")


def clean_rows(rows):
    """去掉空标题、销量行、重复卡（同一商品会有主卡 + 找相似 两张）。"""
    out, seen = [], set()
    for r in rows:
        t = (r.get("title") or "").strip()
        if not t or SALES_ONLY.match(t):
            continue
        r["title"] = t
        r["shop"] = (r.get("shop") or "").replace("真实工厂认证", "").replace("找相似", "").strip()
        key = t[:24]
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def collect(pg, limit=40):
    """抓结果卡：优先 a[href*=detail.1688.com]，没有就用含 ¥ 的 offer/card 容器。"""
    out, seen = [], set()
    anchors = pg.query_selector_all("a[href*='detail.1688.com']")
    for a in anchors:
        try:
            node = a.evaluate_handle("x => x.closest('[class*=offer], [class*=card], li')")
            el = node.as_element() or a
            txt = (el.inner_text() or "").strip()
            href = a.get_attribute("href") or ""
        except Exception:
            continue
        if "¥" not in txt or href in seen:
            continue
        seen.add(href)
        out.append(dict(parse_card(txt), url=href if href.startswith("http") else "https:" + href))
        if len(out) >= limit:
            return clean_rows(out)[:limit]
    if out:
        return clean_rows(out)[:limit]
    for sel in ("[class*=offer]", "[class*=card]", "[class*=item]"):
        for el in pg.query_selector_all(sel):
            try:
                txt = (el.inner_text() or "").strip()
            except Exception:
                continue
            if "¥" not in txt or len(txt) < 20 or txt[:40] in seen:
                continue
            seen.add(txt[:40])
            out.append(dict(parse_card(txt), url=""))
            if len(out) >= limit:
                return clean_rows(out)[:limit]
        if out:
            return clean_rows(out)[:limit]
    return clean_rows(out)


def report(rows):
    if not rows:
        print("没抓到结果。先跑 login，再用 --keep 看页面是不是卡在登录/滑块。")
        return
    print(f"{'价':>7} {'起批':>5} {'已售':>6} {'省份':<4} {'标':<6} 标题 / 店铺")
    for r in rows:
        print(f"{(r['price'] if r['price'] is not None else '-'):>7} "
              f"{(r['moq'] or '-'):>5} {(r['sales'] or '-'):>6} {r.get('province',''):<4} "
              f"{r.get('strength',''):<6} {r['title'][:38]}")
        if r["shop"]:
            print(f"{'':>30} └ {r['shop']}")


def dump(rows, path):
    if not path:
        return
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)
    print(f"已写出: {path}")


def on_login_page(pg):
    """当前页是不是登录页（cookie 看着在、服务端已失效时也会出现）。"""
    try:
        return "login." in pg.url or "扫码登录" in pg.inner_text("body")[:3000]
    except Exception:
        return True


def wait_login(pg, ctx, minutes, interval=5, probe_url=None):
    """等扫码。以「页面不再是登录页」为准，光看 cookie 会误判。"""
    pg.bring_to_front()
    print("=" * 58)
    print(" 1688 登录窗口已打开 —— 是一个独立的 Chrome 窗口，")
    print(" 不是你现在用的那个浏览器，去任务栏找新出现的那个。")
    print(f" 点右上角「登录」→ 扫码，最多等 {minutes} 分钟")
    print("=" * 58)
    ticks = 0
    while ticks < minutes * 60:
        pg.wait_for_timeout(interval * 1000)
        ticks += interval
        if logged_in(ctx):
            goto(pg, probe_url or "https://www.1688.com/", 4000)
            if not on_login_page(pg):
                save_cookies(ctx)
                print(f"✅ 已登录（{ticks}s），登录态已存到 {COOKIES}")
                return True
        if ticks % 30 == 0:
            pg.bring_to_front()
            print(f"   …还在等扫码（{ticks}s）")
    print("⚠️  超时没登录成功，当前进度仍会保存")
    return False


def cmd_login(a):
    with sync_playwright() as p:
        ctx = open_ctx(p, a.headless)
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        goto(pg, "https://login.1688.com/member/signin.htm", 5000)
        wait_login(pg, ctx, a.minutes,
                   probe_url=SEARCH_URL + urllib.parse.quote("保温杯"))
        ctx.close()


def click_text(pg, text):
    """按可见文字点按钮。1688 的按钮 class 是随机串，只能靠文字。"""
    for sel in (f"text={text}", f"button:has-text('{text}')", f"[class*=btn]:has-text('{text}')"):
        try:
            el = pg.query_selector(sel)
            if el and el.is_visible():
                el.click()
                return True
        except Exception:
            continue
    return False


def image_search(pg, img, limit=20):
    box = pg.query_selector("input[type=file]")
    if not box:
        print("❗ 没找到上传框（可能改版）")
        return []
    box.set_input_files(os.path.abspath(img))
    pg.wait_for_timeout(2500)
    if not click_text(pg, "搜索图片"):
        print("⚠️  没点到「搜索图片」，直接等结果")
    pg.wait_for_timeout(10000)
    for _ in range(2):
        pg.mouse.wheel(0, 1800)
        pg.wait_for_timeout(2500)
    return collect(pg, limit)


def cmd_search(a):
    with sync_playwright() as p:
        ctx = open_ctx(p, a.headless)
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        if not logged_in(ctx):
            print("❌ 未登录（没有 cookie2）。先跑: py ali1688_browser.py login")
            ctx.close()
            return
        goto(pg, SEARCH_URL + urllib.parse.quote(a.keyword))
        print(f"标题: {pg.title()[:50]}")
        rows = collect(pg, a.limit)
        report(rows)
        dump(rows, a.json)
        save_cookies(ctx)
        if a.keep:
            print("浏览器保持打开，Ctrl+C 结束")
            try:
                while True:
                    pg.wait_for_timeout(1000)
            except KeyboardInterrupt:
                pass
        ctx.close()


def fetch_cover(plaza_id):
    """从银河图床拉商品封面（免登录），返回本地文件路径。"""
    pid = plaza_id if plaza_id.startswith("pi_") else "pi_" + plaza_id
    url = f"{COVER_BASE}/api/proxy/plaza-cover/item/{pid}"
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"_cover_{pid[-8:]}.webp")
    with urllib.request.urlopen(url, timeout=30) as r, open(path, "wb") as f:
        f.write(r.read())
    print(f"封面已下载: {path}")
    return path


def cmd_image(a):
    img = a.image or (fetch_cover(a.cover) if a.cover else None)
    if not img or not os.path.exists(img):
        sys.exit("要给图片: py ali1688_browser.py image 图片.jpg  或 --cover pi_xxx")
    with sync_playwright() as p:
        ctx = open_ctx(p, a.headless)
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        if not logged_in(ctx):
            print("❌ 未登录（没有 cookie2）。先跑: py ali1688_browser.py login")
            ctx.close()
            return
        goto(pg, IMAGE_URL)
        print("已上传，等结果…")
        rows = image_search(pg, img, a.limit)
        report(rows)
        dump(rows, a.json)
        save_cookies(ctx)
        if a.keep:
            print("浏览器保持打开，Ctrl+C 结束")
            try:
                while True:
                    pg.wait_for_timeout(1000)
            except KeyboardInterrupt:
                pass
        ctx.close()


def cmd_batch(a):
    """一次登录里跑完所有任务：关键词搜 + 以图搜款。避免 cookie 落盘再读的坑。"""
    kws = [k.strip() for k in (a.keywords or "").split(",") if k.strip()]
    covers = [c.strip() for c in (a.covers or "").split(",") if c.strip()]
    if not kws and not covers:
        sys.exit("要带 --keywords 或 --covers")
    out = {}
    with sync_playwright() as p:
        ctx = open_ctx(p, a.headless)
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        goto(pg, "https://www.1688.com/", 6000)
        probe = SEARCH_URL + urllib.parse.quote(kws[0]) if kws else None
        if probe:
            goto(pg, probe, 6000)
        if on_login_page(pg) or not logged_in(ctx):
            goto(pg, "https://login.1688.com/member/signin.htm", 5000)
            wait_login(pg, ctx, a.minutes, probe_url=probe)
        for kw in kws:
            print("\n===== 关键词: " + kw + " =====")
            goto(pg, SEARCH_URL + urllib.parse.quote(kw))
            if on_login_page(pg):
                print("❗ 又被踢到登录页，跳过")
                out[kw] = []
                continue
            out[kw] = collect(pg, a.limit)
            report(out[kw])
        for pid in covers:
            print("\n===== 以图搜款: " + pid + " =====")
            img = fetch_cover(pid)
            goto(pg, IMAGE_URL)
            out[pid] = image_search(pg, img, a.limit)
            report(out[pid])
        save_cookies(ctx)
        dump(out, a.json)
        if a.keep:
            print("浏览器保持打开，Ctrl+C 结束")
            try:
                while True:
                    pg.wait_for_timeout(1000)
            except KeyboardInterrupt:
                pass
        ctx.close()


def offer_info(pg):
    """商品详情页里跟代发有关的字段。1688 的 DOM 会变，全部用文字兜底。"""
    t = pg.inner_text("body")
    flat = re.sub(r"\s+", " ", t)

    def grab(pattern, cast=str, default=""):
        m = re.search(pattern, flat)
        return cast(m.group(1)) if m else default

    return {
        "title": re.sub(r" - 阿里巴巴$", "", pg.title())[:50],
        "one_price": grab(r"1件价格 ?￥ ?([\d.]+)", float, None),
        "dropship": any(k in t for k in ("代发下单", "一件代发", "立即铺货", "密文代发")),
        "stock": grab(r"库存(\d+)", int, None),
        "moq": grab(r"(\d+) ?(?:瓶|袋|件|个|盒|包|支|套)起批", int, None),
        "origin": grab(r"产地 ?([一-龥]{2,4})"),
        "ship": grab(r"发货地 ?([一-龥]{2,4})"),
        "dropship_30d": grab(r"近30天代发数量 ?(\d+)", int, None),
        "shop": grab(r"^([^ ]{2,24}(?:公司|商行|厂|店|商社|营业部))"),
    }


def cmd_detail(a):
    urls = [u.strip() for u in (a.url or "").split(",") if u.strip()]
    if not urls:
        sys.exit("要给链接: py ali1688_browser.py detail \"https://detail.1688.com/offer/xxx.html\"")
    out = {}
    with sync_playwright() as p:
        ctx = open_ctx(p, a.headless)
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        for i, u in enumerate(urls):
            if i:
                pg.wait_for_timeout(a.pause * 1000)  # 连续开详情页会触发验证码
            goto(pg, u if u.startswith("http") else "https://" + u, 8000)
            info = offer_info(pg)
            if "验证码" in info["title"]:
                print("⚠️  被验证码拦了")
                if a.keep:
                    pg.bring_to_front()
                    print("   请在浏览器里过一下验证码，最多等 3 分钟")
                    for _ in range(36):
                        pg.wait_for_timeout(5000)
                        if "验证码" not in pg.title():
                            break
                    info = offer_info(pg)
            out[u] = info
            print("\n" + u.split('/offer/')[-1].split('.')[0])
            for k, v in info.items():
                print(f"  {k:<14} {v}")
        save_cookies(ctx)
        dump(out, a.json)
        ctx.close()


def cmd_probe(a):
    with sync_playwright() as p:
        ctx = open_ctx(p, a.headless)
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        print("登录:", logged_in(ctx))
        goto(pg, SEARCH_URL + urllib.parse.quote(a.keyword))
        print("url:", pg.url[:90])
        print("title:", pg.title()[:40])
        print("file inputs:", len(pg.query_selector_all("input[type=file]")))
        for sel in ("a[href*='detail.1688.com']", "[class*=offer]", "[class*=card]"):
            print(f"{sel:<34} {len(pg.query_selector_all(sel))}")
        print("body:", pg.inner_text("body")[:300].replace("\n", " | "))
        if a.json:
            with open(a.json, "w", encoding="utf-8") as f:
                f.write(pg.content())
            print("HTML 已存:", a.json)
        ctx.close()


def selfcheck():
    """解析器的自检：一段真实的卡片文字应当抠出价格/起批/店铺。"""
    sample = ("秋冬爆炸盐洗衣去污粉\n¥\n3\n.90\n100件起批\n已售 2.3万\n"
              "退货包运费\n义乌市某某日用品有限公司")
    r = parse_card(sample)
    assert r["price"] == 3.90, r
    assert r["moq"] == 100, r
    assert r["sales"] == "2.3万", r
    assert "日用品" in r["shop"], r
    assert r["title"].startswith("秋冬爆炸盐"), r
    print("selfcheck ok")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["login", "search", "image", "detail", "probe", "batch", "selfcheck"])
    ap.add_argument("keyword", nargs="?")
    ap.add_argument("image", nargs="?", help="以图搜款的本地图片路径")
    ap.add_argument("--cover", help="银河 plaza id（pi_xxx），自动下载封面图再搜")
    ap.add_argument("--limit", type=int, default=40)
    ap.add_argument("--json")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--keep", action="store_true", help="跑完不关浏览器")
    ap.add_argument("--minutes", type=int, default=10, help="login 等待扫码的分钟数")
    ap.add_argument("--keywords", help="batch: 逗号分隔的关键词")
    ap.add_argument("--covers", help="batch: 逗号分隔的银河 plaza id")
    ap.add_argument("--url", help="detail: 逗号分隔的 1688 商品链接")
    ap.add_argument("--pause", type=int, default=12, help="detail: 每页之间的间隔秒数")
    a = ap.parse_args()
    if a.cmd == "selfcheck":
        selfcheck()
        return
    if a.cmd == "search" and not a.keyword:
        sys.exit("要带关键词，例如: py ali1688_browser.py search \"爆炸盐\"")
    {"login": cmd_login, "search": cmd_search, "image": cmd_image,
     "detail": cmd_detail, "probe": cmd_probe, "batch": cmd_batch}[a.cmd](a)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
