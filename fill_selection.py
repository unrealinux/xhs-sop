# -*- coding: utf-8 -*-
"""把 fetch_plaza.py 的输出灌进 SOP 模板的「初选表」。

重生成一份干净模板（保留公式/下拉/批注），再把候选品逐行填进初选表前 N 行。
「档位(自动)」列不写值，留给公式自己算。

用法:
    py fill_selection.py 家居-加购400以上.xlsx
    py fill_selection.py 家居-低粉埋伏.xlsx -o 我的表.xlsx
    py fill_selection.py 选品广场-SOP分档.xlsx --source 银河选品广场 --limit 60
"""
import argparse
import os
import shutil
import sys

from openpyxl import load_workbook
from openpyxl.styles import Font

import build_template

# fetch_plaza.py 输出的列序（0-based）
SRC = {"档位": 0, "产品名": 1, "到手价": 2, "加购": 3, "日增": 4, "粉丝": 5,
       "销量": 6, "比值": 7, "类目": 8, "银河分层": 9, "红蓝海": 10,
       "价格标记": 11, "可买性": 12, "链接": 13}

# 初选表 1-based 列号
COL = {"日期": 1, "来源": 2, "店铺名": 3, "产品名": 4, "竞品链接": 5,
       "24小时加购": 6, "粉丝数": 7, "档位(自动)": 8, "是否感兴趣": 9, "备注": 10}


def num(v):
    return v if isinstance(v, (int, float)) else None


def note_of(r):
    """备注 = 到手价 + 销量 + 银河分层 + 低价预警。"""
    parts = []
    price = num(r[SRC["到手价"]])
    if price is not None:
        parts.append(f"¥{price:g}")
    sales = num(r[SRC["销量"]])
    if sales is not None:
        parts.append(f"销{sales:g}")
    seg = r[SRC["银河分层"]]
    if seg:
        parts.append(str(seg))
    ocean = r[SRC["红蓝海"]]
    if ocean and ocean != "红海":
        parts.append(str(ocean))
    if price is not None and price < 30:
        parts.append("⚠低价")
    return " ".join(parts)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", help="fetch_plaza.py 输出的 xlsx")
    ap.add_argument("-o", "--output", default="选品SOP-4.0模板.xlsx")
    ap.add_argument("--source", default="银河选品广场", help="初选表「来源」列填什么")
    ap.add_argument("--limit", type=int, default=0,
                    help="只填前 N 条（默认填满 60 行）。源文件已按低粉→常规→大爆→未达标排好序")
    ap.add_argument("--sheet", default="选品广场", help="源文件的工作表名")
    a = ap.parse_args()

    if not os.path.isfile(a.input):
        sys.exit(f"找不到文件: {a.input}")

    src = load_workbook(a.input, data_only=True)
    if a.sheet not in src.sheetnames:
        sys.exit(f"源文件里没有「{a.sheet}」表，实际有: {src.sheetnames}")
    sws = src[a.sheet]
    rows = [r for r in sws.iter_rows(min_row=2, values_only=True)
            if r and r[SRC["产品名"]]]
    if not rows:
        sys.exit("源文件没有数据行")

    # 输出文件被占用就早点报错
    if os.path.exists(a.output):
        try:
            with open(a.output, "a+b"):
                pass
        except PermissionError:
            sys.exit(f"输出文件被占用，写不进去: {a.output}\n关掉 Excel 再跑。")

    # 重新生成干净模板（公式 / 下拉 / 批注都在），再往里填
    tmp = "_fresh_template.xlsx"
    build_template.build(tmp)

    wb = load_workbook(tmp)
    ws = wb["初选表"]
    cap = build_template.ROWS

    filled = rows[:a.limit] if a.limit else rows[:cap]
    if len(filled) > cap:
        sys.exit(f"--limit {a.limit} 超过表容量 {cap}，改小一点")
    for i, r in enumerate(filled, start=2):
        ws.cell(row=i, column=COL["来源"], value=a.source)
        ws.cell(row=i, column=COL["产品名"], value=r[SRC["产品名"]])
        c = ws.cell(row=i, column=COL["竞品链接"], value=r[SRC["链接"]])
        if r[SRC["链接"]]:
            c.hyperlink = str(r[SRC["链接"]])
            c.font = Font(color="0563C1", underline="single")
        ws.cell(row=i, column=COL["24小时加购"], value=num(r[SRC["加购"]]))
        ws.cell(row=i, column=COL["粉丝数"], value=num(r[SRC["粉丝"]]))
        ws.cell(row=i, column=COL["备注"], value=note_of(r))
        # 第 8 列「档位(自动)」故意不写，保留模板里的公式

    wb.save(a.output)
    if os.path.exists(tmp):
        os.remove(tmp)

    from collections import Counter
    lv = Counter(r[SRC["档位"]] for r in filled)
    dropped = rows[cap:]
    print(f"初选表已填 {len(filled)} 行 → {os.path.abspath(a.output)}")
    print("  低粉爆款 {} | 常规爆款 {} | 大爆款 {} | 未达标 {}".format(
        *[lv.get(k, 0) for k in ("低粉爆款", "常规爆款", "大爆款", "未达标")]))
    if dropped:
        dv = Counter(r[SRC["档位"]] for r in dropped)
        if a.limit:
            print(f"  按 --limit {a.limit} 只填前 {len(filled)} 条，其余 {len(dropped)} 条未纳入: {dict(dv)}")
        else:
            print(f"  ⚠️ 表只有 {cap} 行，{len(dropped)} 条没装下: {dict(dv)}")
    print("  日期 / 店铺名 / 是否感兴趣 留空 —— 店铺名银河接口不返回，得点链接进去看")
    print("  第 8 列「档位(自动)」是公式，打开就会自己算出来")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
