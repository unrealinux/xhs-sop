# -*- coding: utf-8 -*-
"""生成测品复盘表：笔记数据 + 商品数据 + 自动判断。"""
import argparse, sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HDR = PatternFill("solid", fgColor="1F4E79")
F = PatternFill("solid", fgColor="FFF2CC")
GREEN = PatternFill("solid", fgColor="C6EFCE")
RED = PatternFill("solid", fgColor="FFC7CE")

NOTES = [
    "麂皮鞋脏了别扔，3分钟救回来",
    "翻毛皮清洁，90%的人第一步就错了",
    "勃肯鞋穿一夏天，这样洗不伤绒面",
    "雪地靴去污，不用送洗鞋店",
    "德训鞋发黄发灰，救回来只用了这一瓶",
    "麂皮护理顺序：先清洁还是先补色",
    "我试了5种方法，这个最省事",
]


def head(ws, cols):
    for i, (name, w) in enumerate(cols, 1):
        c = ws.cell(row=1, column=i, value=name)
        c.fill, c.font = HDR, Font(color="FFFFFF", bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def build(path, product):
    wb = Workbook()
    ws = wb.active
    ws.title = "笔记数据"
    head(ws, [("序号", 6), ("标题", 40), ("发布日", 11), ("曝光", 9), ("点击", 9),
              ("点赞", 8), ("收藏", 8), ("评论", 8), ("点击率(自动)", 12), ("投流(元)", 9)])
    for i, t in enumerate(NOTES, 1):
        r = i + 1
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=t)
        ws.cell(row=r, column=9, value=f"=IF(D{r}=\"\",\"\",E{r}/D{r})").number_format = "0.0%"
        ws.cell(row=r, column=9).fill = F

    ws2 = wb.create_sheet("商品数据")
    head(ws2, [("日期", 12), ("访客", 9), ("加购", 9), ("订单", 9),
               ("销售额", 10), ("花费", 9)])
    for r in range(2, 16):
        for c in range(1, 7):
            ws2.cell(row=r, column=c).border = None
    ws2["H1"] = "汇总（自动）"
    ws2["H1"].font = Font(bold=True)
    wb_ = wb["商品数据"]
    wb_["H2"], wb_["I2"] = "总访客", "=SUM(B2:B15)"
    wb_["H3"], wb_["I3"] = "总订单", "=SUM(D2:D15)"
    wb_["H4"], wb_["I4"] = "总花费", "=SUM(F2:F15)"
    wb_["H5"], wb_["I5"] = "天数", "=COUNT(B2:B15)"
    wb_["H6"], wb_["I6"] = "日均访客", "=IF(I5=0,\"\",I2/I5)"
    wb_["H7"], wb_["I7"] = "结论", "=IF(I5=0,\"待填\",IF(OR(I6>40,I3>0),\"✅ 进第二轮\",\"❌ 淘汰\"))"
    for k in ("I2", "I3", "I4", "I5", "I6"):
        wb_[k].fill = F
    wb_["I7"].fill = F
    wb_["I7"].font = Font(bold=True, size=12)

    ws3 = wb.create_sheet("怎么用")
    ws3.column_dimensions["B"].width = 96
    lines = [
        f"{product} · 测品复盘",
        "",
        "【笔记数据】",
        "  创作中心 → 笔记管理 → 点开每篇看「曝光 / 点击 / 赞藏评」",
        "  点击率自动算，低于 5% 的说明首图或标题不行",
        "",
        "【商品数据】每天一行，填 15 天",
        "  店铺后台 → 数据 → 商品分析：访客 / 加购 / 订单 / 销售额",
        "  投流花费填在最后一列",
        "",
        "【判断标准（SOP P37）】",
        "  第一轮：日均访客 > 40 或 出单 → 进第二轮；否则淘汰",
        "  第二轮：每天稳定出单 → 进店铺运营",
        "",
        "【投流】",
        "  只投点击率最高的 1~2 篇，每篇 100 元起，总预算 200~300",
        "  停手线：单篇花费 = 单件利润（麂皮 ≈ ¥40）还没出单，停",
        "  保本 ROI = 售价 ÷ 单件利润 = 50.8 ÷ 39.6 ≈ 1.28（<4，可投）",
    ]
    for i, t in enumerate(lines, 1):
        c = ws3.cell(row=i, column=2, value=t)
        if t.startswith("【"):
            c.font = Font(bold=True, color="1F4E79", size=12)
    wb.save(path)
    print("已生成:", path)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--output", default="麂皮-复盘表.xlsx")
    ap.add_argument("--product", default="麂皮清洁剂")
    a = ap.parse_args()
    build(a.output, a.product)
