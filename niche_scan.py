# -*- coding: utf-8 -*-
"""赛道筛选：从全站数据里找出「高粉竞品少、低粉竞品活跃」的细分赛道。

背景（SOP P22）：入场的硬性要求是「高粉竞品 + 低粉竞品 ≤ 5 家」。
从银河池子里按「需求新鲜」挑品会持续撞这堵墙 —— 银河只收录已经有加购的品，
有加购就说明大店已经进来了。所以反过来：先找大店还没占住的细分，再挑品。

  高粉竞品 = 粉丝 > 2000（或该产品笔记 > 1000）
  中间地带 = 粉丝 500 ~ 2000
  低粉竞品 = 粉丝 < 500   ← 和你同一起跑线，它做成说明小号也能做

用法:
    py niche_scan.py all2.json
    py niche_scan.py all2.json --max-big 0 --min-small 3
"""
import argparse
import json
import re
import sys
from collections import Counter, defaultdict

import jieba.posseg as pseg

# 保留的词性：名词 / 专名 / 动名词 / 名形词 / 英文缩写
KEEP_FLAG = {"n", "nz", "vn", "an", "ns", "nr", "eng", "l", "j"}

NOISE = re.compile(
    r"包邮|正品|同款|现货|特价|限时|清仓|亏本|福利|秒杀|批发|官方|旗舰|专柜|进口|"
    r"新款|升级|加厚|加大|大容量|高颜值|网红|神器|推荐|爆款|新品|优惠|促销|"
    r"家用|室内|多功能|便携|超轻|高档|高级|原创|独家|夏季|秋冬|春夏|春秋|"
    r"学生|宿舍|女士|男士|情侣|儿童|宝宝|孕妇|礼物|"
    r"其他|OTHER|other|本店|新款|这款|那款"
)


def tier(f):
    if f is None:
        return "?"
    if f > 2000:
        return "big"
    if f < 500:
        return "small"
    return "mid"


def keywords(name):
    s = str(name)
    out = set()
    for w, flag in pseg.cut(s):
        w = w.strip()
        if len(w) < 2:
            continue
        if flag not in KEEP_FLAG:
            continue
        if NOISE.search(w):
            continue
        if not re.search(r"[\u4e00-\u9fa5a-zA-Z]{2}", w):
            continue
        out.add(w)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("data")
    ap.add_argument("--min-items", type=int, default=3)
    ap.add_argument("--max-big", type=int, default=1)
    ap.add_argument("--min-small", type=int, default=2)
    ap.add_argument("--max-share", type=float, default=0.06,
                    help="关键词出现率超过这个比例就当太泛，丢掉")
    ap.add_argument("--overlap", type=float, default=0.6,
                    help="和已选赛道商品重合超过这个比例就合并掉")
    ap.add_argument("--min-price", type=float, default=0.0,
                    help="低粉商品里至少有一个到手价 >= 这个数（留下有利润空间的赛道）")
    ap.add_argument("--max-price", type=float, default=0.0,
                    help="低粉商品到手价上限，用来排掉低价赛道（如水果/手机壳）")
    ap.add_argument("-o", "--output", default="_niche.txt")
    a = ap.parse_args()

    rows = json.load(open(a.data, encoding="utf-8"))
    print(f"载入 {len(rows)} 条，分词中…", file=sys.stderr)

    bucket = defaultdict(list)
    for i, o in enumerate(rows):
        for k in keywords(o["name"]):
            bucket[k].append(i)

    cands = []
    for tag, idxs in bucket.items():
        if len(idxs) < a.min_items or len(idxs) > len(rows) * a.max_share:
            continue
        items = [rows[i] for i in idxs]
        c = Counter(tier(o["shop_followers"]) for o in items)
        if c["big"] > a.max_big or c["small"] < a.min_small:
            continue
        sc = [o["add_cart_24h"] or 0 for o in items if tier(o["shop_followers"]) == "small"]
        sp = [o["price"] or 0 for o in items if tier(o["shop_followers"]) == "small"]
        if a.min_price and (not sp or max(sp) < a.min_price):
            continue
        if a.max_price and sp and min(sp) > a.max_price:
            continue
        cands.append({"tag": tag, "idx": set(idxs), "items": items,
                      "big": c["big"], "mid": c["mid"], "small": c["small"],
                      "small_max_cart": max(sc) if sc else 0,
                      "small_max_price": max(sp) if sp else 0})

    # 贪心去重：低粉加购高的优先，和已选赛道重合太多就丢
    cands.sort(key=lambda x: (-x["small_max_cart"], x["big"], -x["small"]))
    kept = []
    for c in cands:
        if any(len(c["idx"] & k["idx"]) >= a.overlap * len(c["idx"]) for k in kept):
            continue
        kept.append(c)

    with open(a.output, "w", encoding="utf-8") as f:
        f.write(f"数据 {len(rows)} 条 | 赛道 {len(kept)} 个"
                f"（高粉<={a.max_big}，低粉>={a.min_small}，商品>={a.min_items}）\n" + "=" * 112 + "\n\n")
        for c in kept[:70]:
            f.write(f"【{c['tag']}】 商品{len(c['items'])} | 高粉{c['big']} 中间{c['mid']} 低粉{c['small']}"
                    f" | 低粉最高加购 {c['small_max_cart']}\n")
            for o in sorted(c["items"], key=lambda x: -(x["add_cart_24h"] or 0))[:4]:
                t = {"big": "高粉", "small": "低粉", "mid": "中间", "?": "??"}[tier(o["shop_followers"])]
                f.write(f"    [{t}] 粉{o['shop_followers'] or '?'} 加购{o['add_cart_24h']} "
                        f"销{o['sales']} ¥{o['price']} {o['name'][:38]}\n")
            f.write("\n")

    print(f"→ {a.output}（{len(kept)} 个赛道）")
    print(f"  {'赛道':<12} {'商品':>4} {'高粉':>4} {'低粉':>4} {'低粉最高加购':>12} {'低粉最高价':>10}")
    for c in kept[:25]:
        print(f"  {c['tag']:<12} {len(c['items']):>4} {c['big']:>4} {c['small']:>4} "
              f"{c['small_max_cart']:>12} ¥{c['small_max_price']:<9}")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
