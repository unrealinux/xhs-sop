# -*- coding: utf-8 -*-
"""生成《虚拟品采集表》—— 在小红书 App 里手工采集用。

为什么要手工：小红书 web 版搜索没有「商品」tab（只有 笔记/用户），
商品数和销量是 App 独占的数据，浏览器自动化拿不到。

用法:
    py build_virtual_sheet.py                    # 生成 虚拟品采集表.xlsx
    py build_virtual_sheet.py -o 我的表.xlsx
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
FILL_DO = PatternFill("solid", fgColor="C6EFCE")     # 值得做
FILL_MAYBE = PatternFill("solid", fgColor="FFEB9C")  # 可测
FILL_NO = PatternFill("solid", fgColor="F2F2F2")     # 不做
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# (职业, 方向, 建议搜索词)
ROWS = [
    # 教师 —— 人群最大，细分最多
    ("教师", "期末评语模板", "小学生期末评语 / 评语大全 / 评语模板100条"),
    ("教师", "家长会PPT+发言稿", "家长会ppt / 家长会发言稿 / 期中家长会"),
    ("教师", "主题班会课件", "班会ppt / 主题班会课件 / 开学第一课"),
    ("教师", "教师资格证资料", "教资笔试资料 / 教资面试逐字稿"),
    ("教师", "教案模板（分学科）", "语文教案模板 / 小学数学教案"),
    ("教师", "班级管理表格", "班级积分表 / 座位表 / 考勤表"),
    ("教师", "手抄报模板", "手抄报模板 / 国庆手抄报 / 半成品手抄报"),
    ("教师", "教师简历/求职", "教师简历模板 / 教师招聘面试"),
    ("教师", "公开课说课稿", "说课稿模板 / 公开课逐字稿"),
    ("教师", "奖状/表彰模板", "奖状模板可编辑 / 学生表彰"),
    ("教师", "黑板报素材", "黑板报设计 / 黑板报模板"),
    ("教师", "职称评审材料", "教师职称材料 / 职称评审模板"),
    ("教师", "小学试卷模板", "试卷模板word / 单元测试卷"),
    ("教师", "教师继续教育", "继续教育学时 / 培训心得"),
    ("教师", "幼儿园环创素材", "幼儿园环境创设 / 主题墙"),
    # 护士
    ("护士", "护士资格证资料", "护考资料 / 护士资格证题库 / 护考冲刺"),
    ("护士", "护理查房记录模板", "护理查房记录 / 查房模板"),
    ("护士", "交班报告模板", "护理交班报告 / 交班模板"),
    ("护士", "三基三严题库", "三基题库 / 护理三基"),
    ("护士", "护理操作流程", "护理操作流程图 / 50项操作"),
    ("护士", "护士简历模板", "护士简历 / 护理简历模板"),
    ("护士", "护理岗位竞聘PPT", "护士竞聘ppt / 竞聘演讲稿"),
    ("护士", "实习手册/总结", "护士实习手册 / 实习总结"),
    ("护士", "排班表模板", "护士排班表 / 排班模板"),
    ("护士", "护理专升本资料", "护理专升本 / 专升本资料"),
    # 文员 / 行政
    ("文员行政", "Excel表格模板包", "行政表格 / 考勤表 / 报销单模板"),
    ("文员行政", "公文写作模板", "公文格式 / 通知模板 / 请示报告"),
    ("文员行政", "会议纪要模板", "会议纪要模板 / 会议记录格式"),
    ("文员行政", "PPT汇报模板", "工作汇报ppt / 年终总结ppt"),
    ("文员行政", "行政管理制度汇编", "行政管理制度 / 公司制度模板"),
    ("文员行政", "员工手册模板", "员工手册范本 / 入职手册"),
    ("文员行政", "年会策划方案", "年会策划 / 年会方案模板"),
    ("文员行政", "劳动合同/协议模板", "劳动合同范本 / 协议模板"),
    ("文员行政", "招聘面试题库", "面试题库 / 招聘话术"),
    ("文员行政", "办公快捷键/技巧", "excel技巧 / word技巧"),
    # 前台 / 客服
    ("前台客服", "客服话术大全", "客服话术 / 前台话术 / 催单话术"),
    ("前台客服", "投诉处理话术", "投诉处理话术 / 差评回复"),
    ("前台客服", "前台工作手册", "前台工作流程 / 前台手册"),
    ("前台客服", "电话接听规范", "电话接听话术 / 前台接电话"),
    ("前台客服", "客户接待流程", "接待流程 / 商务接待"),
    ("前台客服", "五星好评引导话术", "好评话术 / 引导评价"),
    # 会计 / 出纳
    ("会计出纳", "Excel财务模板", "财务表格 / 记账模板 / 台账"),
    ("会计出纳", "会计实操资料", "会计实操 / 零基础会计"),
    ("会计出纳", "初级会计资料", "初级会计资料 / 初级题库"),
    ("会计出纳", "财务报表模板", "财务报表模板 / 资产负债表"),
    ("会计出纳", "出纳工作手册", "出纳工作内容 / 出纳流程"),
    ("会计出纳", "税务申报流程", "报税流程 / 税务申报"),
    # 横切（跨职业，一次做多职业复用）
    ("横切", "简历模板", "简历模板 / 应届生简历"),
    ("横切", "PPT模板", "ppt模板 / 年终总结ppt"),
    ("横切", "述职/竞聘PPT", "述职报告ppt / 竞聘ppt"),
    ("横切", "手抄报", "手抄报模板 / 节日手抄报"),
]

HDR1 = [
    ("职业", 11), ("方向", 22), ("建议搜索词", 34),
    ("笔记总数", 11), ("在卖商品数", 12), ("头部已售", 11),
    ("有低粉卖家吗", 13), ("判断(自动)", 16), ("备注", 26),
]
HDR2 = [
    ("关键词", 20), ("商品标题", 38), ("售价", 9),
    ("已售", 10), ("卖家粉丝数", 12), ("卖家笔记数", 12), ("链接", 46),
]

HINT1 = {
    4: "App 搜关键词后显示的笔记总数（找不到就留空，不影响判断）。经验值：<1万太小 / 1~10万可看 / >10万需求已验证",
    5: "数一下有多少个商品链接在卖。1~3家可能没需求，3~10家好，20家以上红海",
    6: "排在最前面的那个商品，它的「已售」数",
    7: "往下翻，有没有粉丝<500 但销量不错的卖家。有 = 产品本身能打，你也能进",
    8: "自动算。D/E 填了、G 填了才有结果",
}
HINT2 = {
    3: "到手价。虚拟品常见 6 / 9.9 / 19.9 / 29.9 / 49.9 几档",
    4: "已售件数",
    5: "卖家主页的粉丝数。<500 算低粉",
    6: "卖家主页的笔记数。笔记多但粉丝少 = 铺货号",
}


def style_header(ws, headers):
    for i, (name, w) in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=name)
        c.fill, c.font = HEAD_FILL, HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"


def add_hints(ws, hints):
    for col, txt in hints.items():
        cm = Comment(txt, "虚拟品选品")
        cm.width, cm.height = 330, 150
        ws.cell(row=1, column=col).comment = cm


def build(path):
    if os.path.exists(path):
        try:
            with open(path, "a+b"):
                pass
        except PermissionError:
            sys.exit(f"输出文件被占用，写不进去: {path}\n"
                     "大概率还开在 Excel 里，关掉再跑（或用 -o 换个名字）。")
    wb = Workbook()

    # ---------- 表1：方向粗筛 ----------
    ws = wb.active
    ws.title = "方向粗筛"
    style_header(ws, HDR1)
    for r, (job, direction, kw) in enumerate(ROWS, start=2):
        ws.cell(row=r, column=1, value=job)
        ws.cell(row=r, column=2, value=direction)
        ws.cell(row=r, column=3, value=kw)
        ws.cell(row=r, column=8, value=(
            f'=IF(AND(D{r}<>"",D{r}<10000),"需求太小",'
            f'IF(E{r}="","待填商品数",'
            f'IF(E{r}>20,"红海·别进",'
            f'IF(AND(G{r}="是",F{r}>=1000),"✅ 值得做",'
            f'IF(G{r}="是","🟡 可测·量小","❌ 全是高粉")))))'
        )).fill = FORMULA_FILL
        for c in range(1, 10):
            ws.cell(row=r, column=c).border = BORDER
    last = len(ROWS) + 1
    ws.auto_filter.ref = f"A1:I{last}"
    add_hints(ws, HINT1)
    dv = DataValidation(type="list", formula1='"是,否"', allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv)
    dv.add(f"G2:G{last}")

    # ---------- 表2：商品明细 ----------
    ws2 = wb.create_sheet("商品明细")
    style_header(ws2, HDR2)
    for r in range(2, 202):
        for c in range(1, len(HDR2) + 1):
            ws2.cell(row=r, column=c).border = BORDER
    ws2.auto_filter.ref = f"A1:G201"
    add_hints(ws2, HINT2)

    # ---------- 表3：怎么填 ----------
    ws3 = wb.create_sheet("怎么填")
    ws3.column_dimensions["A"].width = 4
    ws3.column_dimensions["B"].width = 100
    steps = [
        "在小红书 App（不是网页版）里做，虚拟品的商品数据只有 App 有",
        "",
        "【方向粗筛】—— 一个方向约 1 分钟",
        "  1. App 搜索栏输入「建议搜索词」里的词",
        "  2. 顶部会显示笔记总数 → 填 D 列",
        "  3. 点「商品」tab，数一下有多少个商品链接 → 填 E 列",
        "  4. 记下排第一那个商品的「已售」→ 填 F 列",
        "  5. 往下翻 1~2 屏，有没有粉丝<500 的卖家 → 填 G 列",
        "  6. H 列自动出判断，不用管",
        "",
        "【商品明细】—— 只对 H 列显示「✅ 值得做」的方向做，约 10 分钟",
        "  把前 5~10 个商品逐个抄进「商品明细」表",
        "  抄完发我，我按 SOP 判据帮你排优先级、给定价建议",
        "",
        "【判断标准从哪来】",
        "  笔记<1万 → 需求太小（生财冬青：单品产出低的不是好产品）",
        "  商品>20家 → 红海，别进",
        "  有低粉卖家 + 头部已售≥1000 → 产品本身能打，不依赖账号权重，值得做",
        "  全是高粉卖家 → 你打不过，跳过",
        "",
        "【版权红线】",
        "  ✅ 格式/模板类（评语、查房记录、Excel、话术）—— 格式不受版权保护，AI 重新生成即原创",
        "  🟡 知识整理类（题库、考证资料）—— 必须自己重新编排",
        "  ❌ 他人作品（小说、影视、原书扫描）—— 高危，别做",
    ]
    for i, t in enumerate(steps, start=1):
        c = ws3.cell(row=i, column=2, value=t)
        if t.startswith("【"):
            c.font = Font(bold=True, color="1F4E79", size=12)
        elif t.startswith("  "):
            c.font = Font(size=10)
    wb.save(path)
    print(f"已生成: {os.path.abspath(path)}")
    print(f"  表1 方向粗筛 {len(ROWS)} 个方向（预填职业/方向/搜索词，你只填 4 列）")
    print(f"  表2 商品明细 200 行")
    print(f"  表3 怎么填")
    from collections import Counter
    c = Counter(j for j, _, _ in ROWS)
    print("  职业分布:", dict(c))


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--output", default="虚拟品采集表.xlsx")
    build(ap.parse_args().output)
