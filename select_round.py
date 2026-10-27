# -*- coding: utf-8 -*-
"""一轮完整选品：银河池子 → SOP 分档 → 竞品密度 → 候选清单。

用法:
    py select_round.py pool-20261022.json
    py select_round.py pool-20261022.json --max-rivals 5 --min-price 30 -o 结果.md

数据从哪来:
    set MC_TOKEN=ses.xxx
    py -c "import fetch_plaza as f,json,os;json.dump(f.fetch(os.environ['MC_TOKEN'],{'sort':'add_cart_24h','order':'DESC','min_cart':101},pause=1.0),open('pool.json','w'),ensure_ascii=False)"
"""
import argparse
import json
import sys
from collections import Counter, defaultdict

from fetch_plaza import classify
from niche_scan import keywords, tier


def build_buckets(pool):
    kw = {id(i): keywords(i["name"]) for i in pool}
    bucket = defaultdict(list)
    for i in pool:
        for k in kw[id(i)]:
            bucket[k].append(i)
    return kw, bucket


def best_niche(item, kw, bucket, pool_size):
    """最具体的赛道词：商品数最少、但 >=2 且不泛滥的那个。"""
    best = None
    for k in kw[id(item)]:
        b = bucket[k]
        if len(b) < 2 or len(b) > pool_size * 0.06:
            continue
        if best is None or len(b) < best[1]:
            t = Counter(tier(x.get("shop_followers")) for x in b)
            best = (k, len(b), t["big"], t["mid"], t["small"], b)
    return best


def pick(pool, max_rivals=5, max_big=1, min_price=30, min_cart=401, top=15):
    kw, bucket = build_buckets(pool)
    items = [i for i in pool if (i.get("add_cart_24h") or 0) >= min_cart]
    cand = [i for i in items
            if classify(i.get("shop_followers"), i.get("add_cart_24h")) in ("低粉爆款", "常规爆款")
            and (i.get("price") or 0) >= min_price
            and (i.get("name") or "") != "精选商品"]
    rows = []
    for i in cand:
        bn = best_niche(i, kw, bucket, len(pool))
        if not bn:
            continue
        tag, n, big, mid, small, _ = bn
        if n > max_rivals or big > max_big or small < 1:
            continue
        sales = i.get("sales") or 0
        rows.append({
            "name": i["name"], "tag": tag, "n": n, "big": big, "mid": mid, "small": small,
            "price": i.get("price"), "cart": i["add_cart_24h"], "sales": sales,
            "fans": i.get("shop_followers"),
            "ratio": round((i["add_cart_24h"] or 0) / sales, 3) if sales else 0,
            "cat": i.get("first_category_name"), "url": i.get("url"),
        })
    rows.sort(key=lambda r: (-(r["cart"] or 0), -r["ratio"]))
    return items, cand, rows[:top]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pool")
    ap.add_argument("--max-rivals", type=int, default=5)
    ap.add_argument("--max-big", type=int, default=1)
    ap.add_argument("--min-price", type=float, default=30)
    ap.add_argument("--min-cart", type=int, default=401)
    ap.add_argument("--top", type=int, default=15)
    ap.add_argument("-o", "--output")
    a = ap.parse_args()

    pool = json.load(open(a.pool, encoding="utf-8"))
    items, cand, rows = pick(pool, a.max_rivals, a.max_big, a.min_price, a.min_cart, a.top)

    low_price = sum(1 for i in pool if (i.get("price") or 0) < 30)
    print(f"池子 {len(pool)} 条 | 加购>={a.min_cart} {len(items)} 条 | "
          f"低粉/常规+价>={a.min_price:.0f} {len(cand)} 条 | 过闸门 {len(rows)} 条")
    print(f"到手价<30 占比 {low_price/len(pool):.0%}\n")
    print(f"{'加购':>5} {'粉丝':>6} {'销量':>7} {'新鲜':>5} {'价':>6} {'竞品':>8} 赛道 / 产品")
    for r in rows:
        print(f"{r['cart']:>5} {r['fans'] or '?':>6} {r['sales'] or 0:>7} {r['ratio']:>5.2f} "
              f"{r['price']:>6} {r['n']}(高{r['big']}/低{r['small']}) "
              f"{r['tag']:<8} {r['name'][:32]}")

    if a.output:
        lines = ["# 选品结果", "",
                 f"池子 {len(pool)} 条（加购≥{a.min_cart and 101}），过闸门 {len(rows)} 条。", "",
                 "| 赛道 | 产品 | 价 | 加购 | 销量 | 新鲜度 | 粉丝 | 竞品(高/低) | 类目 | 链接 |",
                 "|---|---|---|---|---|---|---|---|---|---|"]
        for r in rows:
            lines.append(f"| {r['tag']} | {r['name'][:30]} | {r['price']} | {r['cart']} | "
                         f"{r['sales']} | {r['ratio']} | {r['fans']} | {r['n']}({r['big']}/{r['small']}) | "
                         f"{r['cat']} | {r['url']} |")
        open(a.output, "w", encoding="utf-8").write("\n".join(lines) + "\n")
        print(f"\n→ {a.output}")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
