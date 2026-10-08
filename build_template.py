# -*- coding: utf-8 -*-
"""
生成《小红书选品 SOP 4.0》配套表格：初选表 / 二选表 / 调研表

所有表头批注（单元格提示）直接引自《小红书电商SOP实操流程.pdf》原文，
页码标注在注释里，方便回去核对。

用法:
    python 生成模板.py                     # 输出到 ./选品SOP-4.0模板.xlsx
    python 生成模板.py -o 我的表格.xlsx
"""
import argparse
import os
import sys

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
HEAD_FONT = Font(color="FFFFFF", bold=True, size=11)
FORMULA_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
ROWS = 60

初选_HEADERS = [
    ("日期", 12), ("来源", 14), ("店铺名", 22), ("产品名", 20), ("竞品链接", 40),
    ("24小时加购", 12), ("粉丝数", 10), ("档位(自动)", 14), ("是否感兴趣", 12), ("备注", 24),
]
二选_HEADERS = [
    ("店铺名", 22), ("产品", 20), ("笔记总数", 11), ("链接评价数", 12),
    ("账号权重(自动)", 15), ("拍摄能力能对齐吗", 17), ("我能超越吗", 13), ("结论(自动)", 13),
]
调研_HEADERS = [
    ("产品图/截图", 16), ("竞品链接", 34), ("竞品售价", 11), ("1688批发价", 12),
    ("运费", 9), ("包装", 9), ("平台扣点率", 11), ("单件利润(自动)", 14),
    ("24小时加购", 12), ("每日利润(自动)", 14), ("梯队(自动)", 13),
    ("同款竞品店铺数", 14), ("高粉竞品数", 11), ("低粉竞品数", 11), ("竞品总数(自动)", 14),
    ("首篇爆文距今(天)", 15), ("入场判断(自动)", 15),
    ("货源1", 12), ("货源2", 12), ("货源3", 12), ("货源4", 12), ("货源5", 12),
    ("结论", 20),
]

# 表头批注：列号 -> 说明文字。全部引自 PDF 原文（P19~P22）。
初选_提示 = {
    2: "爬词表 / 千帆推广 / 店铺榜单，一次只填一个。每周一轮选品，约半天。",
    5: "一次点开 20 个链接，看完一个叉掉一个。（P14）",
    7: "爬词表分档靠这两个数：粉丝数 + 24小时加购。",
    8: "粉丝<500 → 低粉爆款（最值得关注，产品本身能打）\n"
       "加购>800 → 大爆款（看天花板，需要差异化运营）\n"
       "粉丝500~2000 且 加购>400 → 常规爆款（市场腰部力量）\n"
       "判断顺序：低粉优先，先判粉丝再判加购。（P13）",
}
二选_提示 = {
    3: "看竞品账号的笔记总数。和链接评价数一起判账号权重。（P19）",
    4: "看竞品商品链接的评价数。和笔记总数一起判账号权重。（P19）",
    5: "笔记总数>1000 且 链接评价数>1000 → 高权重账号。\n"
       "这种对手要用差异化运营去打，不能硬碰硬。（P19）",
    6: "【问题1】他们的拍摄能力，我能对齐吗？\n"
       "这是最基本的。如果他们的图拍得比我好太多，那我就算跟了也很难起来。\n"
       "对不齐的出路：要么提升拍摄能力，要么找差异化场景。（P19）",
    7: "【问题3】我能超越他们吗？\n"
       "素材能对齐，运营能力能超越，就进入调研阶段；\n"
       "如果连素材都对不齐，那就放弃，找下一个。\n"
       "这个标准的核心不是这个产品好不好，而是我能不能打得过。（P19）",
    8: "F、G 都「是」→ 进调研，否则放弃。\n"
       "F（拍摄能力能对齐）填「否」，G 就不用纠绘了，结论已锁死「放弃」。\n"
       "只填了 F、G 还没填完时结论留空（不会误报「放弃」）。",
}
调研_提示 = {
    4: "1688 / 拼多多 以图搜图。优先同发货省份的货源；看有标的商家\n"
       "（实力商家、超级工厂、旗舰店）。SKU 规格差太多基本不是同货源。\n"
       "两个平台都找不到，就不再花精力。（P21）",
    11: "≥10000（每天1万利润）→ 第一梯队\n"
        "≥5000 → 第二梯队　　≥1000 → 第三梯队\n"
        "新手主要看第二、第三梯队；老手主要看第一、第二梯队。\n"
        "第一梯队如果不是最近 15 天刚爆的，不适合新手。（P20）",
    12: "用竞品主图在小红书以图搜图，数有多少家店在卖同款（只记有 24 小时加购的链接）。（P22）",
    13: "【高粉竞品】粉丝>2000，或该产品笔记数>1000。\n"
        "这类竞品有账号权重，需要差异化手段。（P22）",
    14: "【低粉竞品】粉丝<500，且笔记数<500。\n"
        "这类竞品和你处于同一起跑线，是你可以超越的对象。（P22）",
    15: "高粉 + 低粉。硬性要求：总和不能大于 5 家。\n"
        "已经有五六家在卖，说明中层生态位快占满了；\n"
        "只有两三家甚至一家，那这个市场还有空间，值得冲。（P22）",
    16: "竞品第一篇产品爆文发布至今多少天。填「0」表示至今没爆过。",
    17: "竞品总数>5 → 放弃·生态位占满\n"
        "首篇爆文<15天 → 不管有无低粉都可入场\n"
        "首篇爆文≥30天 且 有低粉起来 且 总数≤5 → 可以入场\n"
        "首篇爆文≥30天 且 没低粉起来 → 谨慎入场（P22）",
    18: "按 5 种各挑一家对比：同省份 / 最低价 / 同SKU / 实力商家 / 销量最多。（P21）",
}


def add_hints(ws, hints, title="SOP 4.0"):
    """给表头单元格加批注（鼠标悬停可见）。"""
    for col, text in hints.items():
        cm = Comment(text, title)
        cm.width, cm.height = 340, 170
        ws.cell(row=1, column=col).comment = cm



def style_header(ws, headers):
    for idx, (name, width) in enumerate(headers, start=1):
        c = ws.cell(row=1, column=idx, value=name)
        c.fill, c.font = HEAD_FILL, HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ROWS + 1}"


def body_borders(ws, ncols):
    for r in range(2, ROWS + 2):
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).border = BORDER


def add_dv(ws, formula, col, ncols, prompt=None):
    dv = DataValidation(type="list", formula1=f'"{formula}"', allow_blank=True, showDropDown=False)
    if prompt:
        dv.prompt, dv.promptTitle = prompt, "提示"
        dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(f"{col}2:{col}{ROWS + 1}")


def _f_cx_level(r):
    """初选表 档位：粉丝+加购都必须是数字才算。"""
    return (f'=IF(OR(NOT(ISNUMBER(F{r})),NOT(ISNUMBER(G{r}))),"",'
            f'IF(G{r}<500,"低粉爆款",IF(F{r}>800,"大爆款",'
            f'IF(AND(G{r}<2000,F{r}>400),"常规爆款","未达标"))))')


def _f_ex_weight(r):
    """二选表 账号权重：笔记数<=1000 就能直接判低权重；要判高权重必须两个数都是数字。"""
    return (f'=IF(NOT(ISNUMBER(C{r})),"",'
            f'IF(C{r}<=1000,"低权重",'
            f'IF(NOT(ISNUMBER(D{r})),"",IF(D{r}>1000,"高权重","低权重"))))')


def _f_ex_verdict(r):
    """二选表 结论：F=否立即放弃；填一半时留空，不误报。"""
    return (f'=IF(F{r}="否","放弃",'
            f'IF(AND(F{r}="是",G{r}="是"),"进调研",'
            f'IF(OR(F{r}="",G{r}=""),"","放弃")))')


def _f_dy_unit(r):
    """调研表 单件利润。售价 + 1688批发价 都必须是数字才算。

    只靠 SUM 容错不够：运费/包装空着当 0 可以，但批发价空着当 0 会算出
    「货不要钱」的假利潦，把梯队抬到第一梯队。所以批发价必填。
    """
    return (f'=IF(OR(NOT(ISNUMBER(C{r})),NOT(ISNUMBER(D{r}))),"",'
            f'C{r}-SUM(D{r},E{r},F{r})-C{r}*G{r})')


def _f_dy_daily(r):
    return (f'=IF(OR(NOT(ISNUMBER(H{r})),NOT(ISNUMBER(I{r}))),"",H{r}*I{r})')


def _f_dy_tier(r):
    return (f'=IF(NOT(ISNUMBER(J{r})),"",IF(J{r}>=10000,"第一梯队",'
            f'IF(J{r}>=5000,"第二梯队",IF(J{r}>=1000,"第三梯队","不达标"))))')


def _f_dy_total(r):
    """调研表 竞品总数：至少一个数才计算，文本不算。"""
    return f'=IF(COUNT(M{r},N{r})=0,"",SUM(M{r},N{r}))'


def _f_dy_entry(r):
    """调研表 入场判断：SUM(N)>0 保证文本不被当成 >0。"""
    return (f'=IF(NOT(ISNUMBER(O{r})),"",'
            f'IF(O{r}>5,"放弃·生态位占满",'
            f'IF(NOT(ISNUMBER(P{r})),"待查首篇爆文时间",'
            f'IF(P{r}<15,"可以入场",'
            f'IF(P{r}>=30,IF(SUM(N{r})>0,"可以入场","谨慎入场"),'
            f'"待判断·爆文15~30天")))))')


# 表 -> [(列号, 公式生成函数)]。build() 和 patch 都用这一份定义，不会跑偏。
FORMULAS = {
    "初选表": [(8, _f_cx_level)],
    "二选表": [(5, _f_ex_weight), (8, _f_ex_verdict)],
    "调研表": [(8, _f_dy_unit), (10, _f_dy_daily), (11, _f_dy_tier),
             (15, _f_dy_total), (17, _f_dy_entry)],
}
FORMULA_COLS = {"初选表": [8], "二选表": [5, 8],
                "调研表": [7, 8, 10, 11, 15, 17]}   # 含要预设 5% 扣点率的第 7 列


def apply_formulas(wb):
    """把公式写进已有工作簿（不改动任何手填数据）。"""
    for sheet, specs in FORMULAS.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for r in range(2, ROWS + 2):
            for col, fn in specs:
                ws.cell(row=r, column=col).value = fn(r)
                ws.cell(row=r, column=col).fill = FORMULA_FILL


def build(path):
    wb = Workbook()

    # ---------- 初选表 ----------
    ws = wb.active
    ws.title = "初选表"
    style_header(ws, 初选_HEADERS)
    for r in range(2, ROWS + 2):
        ws.cell(row=r, column=8).value = _f_cx_level(r)
        ws.cell(row=r, column=8).fill = FORMULA_FILL
    add_dv(ws, "爬词表,千帆推广,店铺榜单,银河选品广场,其他", "B", 9, "一次只填一个来源")
    add_dv(ws, "是,否", "I", 9)
    add_hints(ws, 初选_提示)
    body_borders(ws, len(初选_HEADERS))

    # ---------- 二选表 ----------
    ws = wb.create_sheet("二选表")
    style_header(ws, 二选_HEADERS)
    for r in range(2, ROWS + 2):
        ws.cell(row=r, column=5).value = _f_ex_weight(r)
        ws.cell(row=r, column=5).fill = FORMULA_FILL
        ws.cell(row=r, column=8).value = _f_ex_verdict(r)
        ws.cell(row=r, column=8).fill = FORMULA_FILL
    add_dv(ws, "是,否", "F", 7, "他们的图拍得比我好太多，跟了也难起来。"
           "对不齐 → 要么提升拍摄能力，要么找差异化场景")
    add_dv(ws, "是,否", "G", 7, "素材能对齐 + 运营能力能超越 = 进调研。"
           "核心不是产品好不好，是我能不能打得过")
    add_hints(ws, 二选_提示)
    body_borders(ws, len(二选_HEADERS))

    # ---------- 调研表 ----------
    ws = wb.create_sheet("调研表")
    style_header(ws, 调研_HEADERS)
    for r in range(2, ROWS + 2):
        ws.cell(row=r, column=7, value=0.05).number_format = "0%"
        ws.cell(row=r, column=8).value = _f_dy_unit(r)
        ws.cell(row=r, column=8).fill = FORMULA_FILL
        ws.cell(row=r, column=10).value = _f_dy_daily(r)
        ws.cell(row=r, column=10).fill = FORMULA_FILL
        ws.cell(row=r, column=11).value = _f_dy_tier(r)
        ws.cell(row=r, column=11).fill = FORMULA_FILL
        ws.cell(row=r, column=15).value = _f_dy_total(r)
        ws.cell(row=r, column=15).fill = FORMULA_FILL
        ws.cell(row=r, column=17).value = _f_dy_entry(r)
        ws.cell(row=r, column=17).fill = FORMULA_FILL
    add_hints(ws, 调研_提示)
    body_borders(ws, len(调研_HEADERS))

    wb.save(path)
    print(f"已生成: {os.path.abspath(path)}")
    print(f"  初选表 {len(初选_HEADERS)} 列 / 二选表 {len(二选_HEADERS)} 列 / 调研表 {len(调研_HEADERS)} 列，各 {ROWS} 行")
    print("  黄色单元格 = 自动公式，别手改；「档位/权重/结论/梯队/竞品总数/入场判断」会自动算")
    print(f"  表头有批注（悬停可见）：初选 {len(初选_提示)} 列 / 二选 {len(二选_提示)} 列 / "
          f"调研 {len(调研_提示)} 列，内容引自 SOP 原文")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--output", default="选品SOP-4.0模板.xlsx")
    build(ap.parse_args().output)
