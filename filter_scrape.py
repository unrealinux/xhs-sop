# -*- coding: utf-8 -*-
"""
爬词表筛选器 —— 按 SOP 分三档并上色

SOP 分档规则（来自《小红书电商SOP实操流程》）:
  低粉爆款   : 粉丝数 < 500                     -> 绿
  大爆款     : 24小时加购 > 800                 -> 红
  常规爆款   : 500 <= 粉丝数 < 2000 且 加购 > 400 -> 黄
  未达标     : 其余                              -> 灰

用法:
    python 筛选爬词表.py 4.8.xlsx
    python 筛选爬词表.py 4.8.xlsx --sheet Sheet1 -o 结果.xlsx
    python 筛选爬词表.py 4.8.csv  --map 粉丝=粉丝数,加购=24小时加购,店铺=店铺名
"""
import argparse
import csv
import os
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

sys.stdout.reconfigure(encoding="utf-8")

FILLS = {
    "低粉爆款": PatternFill("solid", fgColor="C6EFCE"),
    "大爆款": PatternFill("solid", fgColor="FFC7CE"),
    "常规爆款": PatternFill("solid", fgColor="FFEB9C"),
    "未达标": PatternFill("solid", fgColor="F2F2F2"),
}
KEYS = {
    "粉丝": ["粉丝", "粉丝数", "fans", "followers"],
    "加购": ["加购", "24小时加购", "加购数", "cart", "24h"],
    "店铺": ["店铺", "店铺名", "店名", "shop", "nickname", "账号"],
    "产品": ["产品", "商品", "标题", "名称", "title"],
    "链接": ["链接", "url", "link", "地址"],
    "销量": ["销量", "月销", "sold"],
}
REQUIRED = ["粉丝", "加购"]


def to_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("+", "")
    mult = 1
    for suf, m in (("万", 10000), ("w", 10000), ("W", 10000), ("k", 1000), ("K", 1000)):
        if s.endswith(suf):
            s, mult = s[:-1], m
            break
    try:
        return float(s) * mult
    except ValueError:
        return None


def load_rows(path, sheet=None):
    if path.lower().endswith(".csv"):
        with open(path, encoding="utf-8-sig", newline="") as f:
            return list(csv.reader(f))
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def find_header(rows):
    """返回 (表头行下标, {字段: 列下标})；先找同时命中『粉丝』和『加购』的行。"""
    best = (None, {})
    for i, row in enumerate(rows[:30]):
        cells = [str(c).strip().lower() if c is not None else "" for c in row]
        hit = {}
        for field, words in KEYS.items():
            for ci, cell in enumerate(cells):
                if any(w.lower() in cell for w in words) and field not in hit:
                    hit[field] = ci
        if all(k in hit for k in REQUIRED) and len(hit) > len(best[1]):
            best = (i, hit)
    return best


def classify(fans, cart):
    if fans is None or cart is None:
        return "未达标"
    if fans < 500:
        return "低粉爆款"
    if cart > 800:
        return "大爆款"
    if fans < 2000 and cart > 400:
        return "常规爆款"
    return "未达标"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("-o", "--output")
    ap.add_argument("--sheet")
    ap.add_argument("--map", help="手动指定列，如 粉丝=粉丝数,加购=24小时加购")
    a = ap.parse_args()

    if not os.path.isfile(a.input):
        sys.exit(f"找不到文件: {a.input}")

    rows = load_rows(a.input, a.sheet)
    hi, colmap = find_header(rows)
    if a.map:
        colmap = dict(colmap)
        for pair in a.map.split(","):
            k, _, v = pair.partition("=")
            header = rows[hi] if hi is not None else rows[0]
            for ci, cell in enumerate(header):
                if str(cell).strip() == v.strip():
                    colmap[k.strip()] = ci
                    break
    if not all(k in colmap for k in REQUIRED):
        print("❌ 没自动认出『粉丝数』和『24小时加购』两列。")
        print(f"   表头前 3 行: {rows[0][:20] if rows else '空'} / {rows[1][:20] if len(rows)>1 else ''}")
        print("   请手动指定，例如: --map 粉丝=粉丝数,加购=24小时加购")
        sys.exit(1)
    if hi is None:
        hi = 0

    header = rows[hi]
    order = [c for c in ("店铺", "产品", "链接", "粉丝", "加购", "销量") if c in colmap]
    out_cols = [colmap[k] for k in order] + [None, None]
    title = [KEYS[k][1] if k == "加购" else KEYS[k][0] for k in order] + ["档位", "建议"]

    data = []
    for row in rows[hi + 1:]:
        if not any(str(c).strip() for c in row if c is not None):
            continue
        fans = to_num(row[colmap["粉丝"]]) if colmap["粉丝"] < len(row) else None
        cart = to_num(row[colmap["加购"]]) if colmap["加购"] < len(row) else None
        lvl = classify(fans, cart)
        rec = [row[c] if c is not None and c < len(row) else "" for c in out_cols[:len(order)]]
        tip = {"低粉爆款": "优先看，产品本身能打",
               "大爆款": "看天花板，需差异化",
               "常规爆款": "需求稳定，可跟",
               "未达标": ""}[lvl]
        data.append(rec + [lvl, tip])

    order_rank = {"低粉爆款": 0, "常规爆款": 1, "大爆款": 2, "未达标": 3}
    data.sort(key=lambda r: (order_rank[r[-2]], -(to_num(r[order.index("加购")]) or 0)))

    out = a.output or os.path.splitext(a.input)[0] + "_筛选结果.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "筛选结果"
    ws.append(title)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E79")
        c.alignment = Alignment(horizontal="center", vertical="center")
    for rec in data:
        ws.append(rec)
        fill = FILLS[rec[-2]]
        for c in ws[ws.max_row]:
            c.fill = fill
    widths = [24, 30, 42, 10, 12, 12, 12, 30]
    for i, w in enumerate(widths[:len(title)], start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    if "链接" in order:
        for r in ws.iter_rows(min_row=2):
            cell = r[order.index("链接")]
            if cell.value:
                cell.hyperlink = str(cell.value)
    wb.save(out)

    from collections import Counter
    cnt = Counter(r[-2] for r in data)
    cnt = {k: cnt.get(k, 0) for k in FILLS}
    print(f"读入 {len(data)} 条 → {os.path.abspath(out)}")
    print("  低粉爆款 {低粉爆款} | 常规爆款 {常规爆款} | 大爆款 {大爆款} | 未达标 {未达标}".format(**cnt))
    print("\n低粉爆款 TOP 10（最值得先看）:")
    top = [r for r in data if r[-2] == "低粉爆款"][:10]
    for r in top:
        shop = r[order.index("店铺")] if "店铺" in order else ""
        print(f"  {str(shop)[:26]:<28} 加购 {r[order.index('加购')]:>6}")


if __name__ == "__main__":
    main()
