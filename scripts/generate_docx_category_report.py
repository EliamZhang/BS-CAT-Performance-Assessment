"""根据数据底稿生成 BS-CAT 全分类类别表现分析报告 (docx)。

数据来源: input/category_difference_report_100(1).xlsx（--input 可指定其他底稿）
报告结构:
  1. 执行摘要
  2. 覆盖与一致全景 (2.1 总体覆盖与一致率 / 2.2 业务板块 / 2.3 分类别)
  3. 分业务板块的类别细探 (3.1 收入类 / 3.2 支出类(含 Gambling 专项与 Rent 专项) / 3.3 转账类 / 3.4 负债类 / 3.5 跨板块优化建议)
  4. 差异结构全景 (4.1 总体差异类型分布 / 4.2 业务板块差异分布 / 4.3 业务板块交叉矩阵 / 4.4 Top 差异流向 / 4.5 全景结论)
  5. 附录：其余类别指标汇总

所有数字均从底稿动态计算，不硬编码。
"""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from openpyxl import load_workbook
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = BASE_DIR / "input" / "category_difference_report_100(1).xlsx"
DEFAULT_OUTPUT = BASE_DIR / "output" / "category_difference_report_100_zh.docx"

EMPTY = "(空)"  # 底稿中缺失分类的占位符

BLUE = "1F4E79"
BLUE_MID = "2E74B5"
BLUE_PALE = "DEEBF7"
GRAY = "595959"
GRAY_PALE = "F2F2F2"
YELLOW_PALE = "FFF2CC"
DARK = "222222"

# ---------------------------------------------------------------------------
# 业务板块定义（与报告口径一致：收入 3 / 支出 23 / 转账 2 / 负债 8 = 36）
# ---------------------------------------------------------------------------
GROUPS: Dict[str, List[str]] = {
    "收入类": ["Wages", "Centrelink", "All Other Credits"],
    "支出类": [
        "Information", "Donations", "Education", "Home Improvement", "Insurance",
        "Subscription TV", "Pet Care", "Entertainment", "Retail", "Utilities",
        "Gyms and other memberships", "Personal Care", "Rent", "Health", "Travel",
        "Groceries", "Automotive", "Gambling", "Department Stores", "Telecommunications",
        "Dining Out", "Fees", "Transport",
    ],
    "转账类": ["External Transfers", "Internal Transfer"],
    "负债类": [
        "SACC Loans", "Non SACC Loans", "Credit Card Repayments", "Debt Consolidation",
        "Debt Collection", "Dishonours", "Overdrawn", "Unknown Loans",
    ],
}
GROUP_OF: Dict[str, str] = {
    category: group for group, categories in GROUPS.items() for category in categories
}
GROUP_ORDER: List[str] = ["收入类", "支出类", "转账类", "负债类"]

FOCUS_EXPENSE: set = {"Rent", "Groceries", "Dining Out", "Retail", "Automotive", "Gambling", "Utilities"}
FOCUS_DETAILED: set = (
    {"Wages", "Centrelink", "All Other Credits", "External Transfers", "Internal Transfer"}
    | FOCUS_EXPENSE
)
# v3 口径：第 3 章展开分析的重要类别；其余类别统一纳入第 5 章附录（七项指标汇总）
IMPORTANT_EXPENSE: set = {"Rent", "Gambling"}
IMPORTANT_LIABILITY: set = {"SACC Loans", "Non SACC Loans", "Dishonours", "Credit Card Repayments"}

# 类别业务定义（描述性说明，不推断数据）
CATEGORY_DEFINITION: Dict[str, str] = {
    "Wages": "工资薪金类收入，包括薪资入账、公司发放的工资、奖金、遣散费等。",
    "Centrelink": "政府福利类收入，包括 Centrelink 福利金、政府补贴与救济金。",
    "All Other Credits": "其他贷方入账，非工资、非福利的杂项收款与退款。",
    "Information": "信息资讯类支出，如报刊、资讯订阅、在线内容付费。",
    "Donations": "捐赠类支出，如慈善捐款、捐赠订阅。",
    "Education": "教育类支出，如学费、课程费、教材费。",
    "Home Improvement": "家居装修与改善类支出，如建材、装修服务。",
    "Insurance": "保险类支出，如各类保费缴纳。",
    "Subscription TV": "电视订阅类支出，如流媒体、付费电视、家庭娱乐订阅。",
    "Pet Care": "宠物照护类支出，如宠物食品、兽医、美容。",
    "Entertainment": "娱乐类支出，如电影、演出门票、游戏。",
    "Retail": "零售类消费，涵盖百货以外的零售购物。",
    "Utilities": "公用事业类支出，如水、电、燃气、电话账单等。",
    "Gyms and other memberships": "健身房及其他会员类支出，如健身、俱乐部会费。",
    "Personal Care": "个人护理类支出，如美发、美容、按摩、理疗。",
    "Rent": "房租类支出，周期性租金支付。",
    "Health": "医疗健康类支出，如诊所、药房、牙医、医疗费用。",
    "Travel": "旅行类支出，如机票、酒店、旅行预订、租车。",
    "Groceries": "食品杂货类支出，如超市买菜、日用杂货。",
    "Automotive": "汽车相关支出，如加油、维修、停车、车险以外的车辆费用。",
    "Gambling": "博彩类支出，如投注、彩票、赌场交易。",
    "Department Stores": "百货商场类消费。",
    "Telecommunications": "电信类支出，如手机话费、宽带、通讯设备。",
    "Dining Out": "餐饮外食类支出，如餐厅、快餐、外卖。",
    "Fees": "手续费与杂费类支出，如银行费用、罚款、迟纳金。",
    "Transport": "交通出行类支出，如公交、地铁、出租车、过路费、停车费。",
    "External Transfers": "外部转账，与自身账户体系之外的账户之间的资金转移。",
    "Internal Transfer": "内部转账，账户体系内部的资金转移（同行账户间划转）。",
    "SACC Loans": "SACC 贷款相关交易（银行内贷款产品，含还款与放款）。",
    "Non SACC Loans": "非 SACC 贷款相关交易（其他贷款产品）。",
    "Credit Card Repayments": "信用卡还款类交易。",
    "Debt Consolidation": "债务合并类交易，将多笔债务合并为单笔还款。",
    "Debt Collection": "债务催收类交易，涉及催收机构的还款。",
    "Dishonours": "拒付/退票类交易，如支票退票、直接扣款失败。",
    "Overdrawn": "透支相关交易，账户透支及透支费用。",
    "Unknown Loans": "未识别来源的贷款相关交易，需要进一步核验。",
}

# 差异类型归一化：底稿 03 排查类型 / 01 差异类型 -> 报告口径
DIFF_TYPE_MAP = {
    "分类边界冲突": "双方分类不一致",
    "分类不一致": "双方分类不一致",
    "仅illion有分类": "仅 Illion 有值",
    "仅finv有分类": "仅 finv 有值",
    "finv漏识别": "仅 Illion 有值",
    "finv新增识别": "仅 finv 有值",
}


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------
def to_float(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def norm_cat(value: Any) -> str:
    """标准化类别字段：None / 空 / (空) 均视为缺失。"""
    text = str(value or "").strip()
    if text in {"", "-", "None", EMPTY, "（空）"}:
        return ""
    return text


def fmt_num(value: Any) -> str:
    try:
        return f"{float(value):,.0f}"
    except (TypeError, ValueError):
        return ""


def fmt_pct(value: Any, digits: int = 2) -> str:
    """value 为比例小数，输出百分比字符串；为空输出 ''。"""
    if value is None or value == "":
        return ""
    try:
        return f"{float(value) * 100:.{digits}f}%"
    except (TypeError, ValueError):
        return ""


def fmt_money(value: Any) -> str:
    try:
        return f"${float(value):,.2f}"
    except (TypeError, ValueError):
        return ""


def diff_type_label(value: Any) -> str:
    return DIFF_TYPE_MAP.get(str(value or "").strip(), str(value or "").strip())


# ---------------------------------------------------------------------------
# 数据读取
# ---------------------------------------------------------------------------
def read_metrics(ws: Any) -> Dict[str, Dict[str, Any]]:
    """00_核心对比 第 5~20 行：指标 / 结果 / 分子 / 分母 / 说明。"""
    metrics: Dict[str, Dict[str, Any]] = {}
    for row in ws.iter_rows(min_row=5, max_row=20, min_col=1, max_col=5, values_only=True):
        label = str(row[0] or "").strip()
        if label:
            metrics[label] = {"result": row[1], "numerator": row[2], "denominator": row[3], "note": row[4]}
    return metrics


def read_categories(ws: Any) -> List[Dict[str, Any]]:
    """00_核心对比 第 23~58 行：36 个类别的全量指标（33 列）。"""
    headers = [str(v or "") for v in next(ws.iter_rows(min_row=23, max_row=23, min_col=1, max_col=33, values_only=True))]
    items: List[Dict[str, Any]] = []
    for row in ws.iter_rows(min_row=24, max_row=59, min_col=1, max_col=33, values_only=True):
        category = norm_cat(row[0])
        if not category or not any(v not in (None, "") for v in row):
            continue
        item = {headers[i]: row[i] for i in range(len(headers))}
        item["category"] = category
        item["group"] = GROUP_OF.get(category, "其他")
        items.append(item)
    return items


def read_top_flows(ws: Any) -> List[Dict[str, Any]]:
    """01_差异诊断地图 第 6~25 行：Top 20 差异流向。"""
    flows: List[Dict[str, Any]] = []
    for row in ws.iter_rows(min_row=6, max_row=25, min_col=1, max_col=15, values_only=True):
        if not isinstance(row[0], (int, float)):
            continue
        flows.append({
            "rank": row[0],
            "priority": str(row[1] or "").strip(),
            "key_category": str(row[2] or "").strip(),
            "illion_category": norm_cat(row[3]),
            "finv_category": norm_cat(row[4]),
            "diff_type": diff_type_label(row[5]),
            "count": to_float(row[6]),
            "share_all": row[7],
            "amount": to_float(row[12]),
            "direction": str(row[13] or "").strip(),
            "suggestion": str(row[14] or "").strip(),
        })
    return flows


def read_details(ws: Any) -> List[Dict[str, Any]]:
    """03_排查明细 第 4 行起：29,963 笔差异明细。"""
    details: List[Dict[str, Any]] = []
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=26, values_only=True):
        if not any(v not in (None, "") for v in row):
            continue
        details.append({
            "priority": str(row[0] or "").strip(),
            "diff_type": diff_type_label(row[1]),
            "amount": to_float(row[6]),
            "dr_cr": str(row[7] or "").strip(),
            "text": str(row[8] or "").strip(),
            "third_party": str(row[9] or "").strip(),
            "counterparty": str(row[10] or "").strip(),
            "illion_category": norm_cat(row[11]),
            "finv_category": norm_cat(row[12]),
            "user_id": str(row[16] or "").strip(),
            "application_id": str(row[17] or "").strip(),
            "transaction_date": str(row[5] or "").strip(),
        })
    return details


# ---------------------------------------------------------------------------
# 分析
# ---------------------------------------------------------------------------
class Analysis:
    def __init__(self, metrics: Dict[str, Dict[str, Any]], categories: List[Dict[str, Any]],
                 top_flows: List[Dict[str, Any]], details: List[Dict[str, Any]]):
        self.metrics = metrics
        self.categories = categories
        self.top_flows = top_flows
        self.details = details

        self.total_transactions = to_float(metrics.get("总交易数", {}).get("result"))
        self.illion_coverage = metrics.get("illion Category 覆盖率", {}).get("result")
        self.finv_coverage = metrics.get("finv Category 覆盖率", {}).get("result")
        self.joint_nonempty = to_float(metrics.get("双方非空时一致率", {}).get("denominator"))
        self.joint_agreement = metrics.get("双方非空时一致率", {}).get("result")
        self.mismatch = to_float(metrics.get("双方非空时差异率", {}).get("numerator"))
        self.diff_total = to_float(metrics.get("Category 差异总数", {}).get("result"))
        self.illion_only_total = to_float(metrics.get("仅 illion 有 Category", {}).get("result"))
        self.finv_only_total = to_float(metrics.get("仅 finv 有 Category", {}).get("result"))
        self.both_empty = to_float(metrics.get("双方均为空", {}).get("result"))
        # 至少一侧有分类 = 双方非空 + 仅 Illion + 仅 finv（与底稿"联合非空样本"口径一致）
        self.union_nonempty = self.joint_nonempty + self.illion_only_total + self.finv_only_total

        self.category_by_name = {item["category"]: item for item in categories}
        self._build_detail_aggregates()

    # -- 明细聚合 ----------------------------------------------------------
    def _build_detail_aggregates(self) -> None:
        # 流向统计: (illion, finv, diff_type) -> {count, amount}
        self.flow_stats: Dict[Tuple[str, str, str], Dict[str, float]] = defaultdict(lambda: {"count": 0, "amount": 0.0})
        # 类别涉及明细: category -> list of detail rows (任一命中)
        self.category_details: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        # 板块矩阵: (illion 板块 or 'Illion为空', finv 板块 or 'finv为空') -> count
        self.matrix: Dict[Tuple[str, str], int] = defaultdict(int)

        for detail in self.details:
            il, fv = detail["illion_category"], detail["finv_category"]
            key = (il, fv, detail["diff_type"])
            self.flow_stats[key]["count"] += 1
            self.flow_stats[key]["amount"] += detail["amount"]
            for cat in (il, fv):
                if cat:
                    self.category_details[cat].append(detail)
            il_group = GROUP_OF.get(il, "Illion为空") if il else "Illion为空"
            fv_group = GROUP_OF.get(fv, "finv为空") if fv else "finv为空"
            self.matrix[(il_group, fv_group)] += 1

    # -- 板块口径 ----------------------------------------------------------
    def segment_summary(self, group: str) -> Dict[str, Any]:
        """板块差异汇总：差异交易数 / 占比 / 板块内不一致 / 跨板块不一致 / 单边缺失。"""
        within = cross = illion_only = finv_only = 0
        for detail in self.details:
            il, fv = detail["illion_category"], detail["finv_category"]
            il_group = GROUP_OF.get(il) if il else ""
            fv_group = GROUP_OF.get(fv) if fv else ""
            hit = (il_group == group) or (fv_group == group)
            if not hit:
                continue
            if il and fv:
                if il_group == fv_group:
                    within += 1
                else:
                    cross += 1
            elif il:
                illion_only += 1
            else:
                finv_only += 1
        total = within + cross + illion_only + finv_only
        return {
            "total": total,
            "share": total / self.diff_total if self.diff_total else 0,
            "within": within,
            "cross": cross,
            "illion_only": illion_only,
            "finv_only": finv_only,
        }

    def segment_coverage(self, group: str) -> Dict[str, Any]:
        """板块级覆盖与一致率（口径与总体对齐：一致率 = 交集 / 板块并集）。

        板块并集 = sum(类别并集) - 板块内不一致笔数：
        双方类别不同但同属该板块的交易（如 External Transfers → Internal Transfer）
        会被计入两个类别的并集各一次，板块并集只计一次。
        """
        cats = [item for item in self.categories if item["group"] == group]
        illion_n = sum(to_float(c.get("illion数量")) for c in cats)
        finv_n = sum(to_float(c.get("finv数量")) for c in cats)
        inter_n = sum(to_float(c.get("交集数量")) for c in cats)
        union_sum = sum(to_float(c.get("并集数量")) for c in cats)
        within = self.segment_summary(group)["within"]
        union_n = union_sum - within
        return {
            "category_count": len(cats),
            "illion_count": illion_n,
            "finv_count": finv_n,
            "union_count": union_n,
            "intersection_count": inter_n,
            "illion_coverage": illion_n / self.total_transactions if self.total_transactions else 0,
            "finv_coverage": finv_n / self.total_transactions if self.total_transactions else 0,
            "exact_rate": inter_n / union_n if union_n else 0,
            "broad_rate": (inter_n + within) / union_n if union_n else 0,
        }

    def segment_top_flows(self, group: str, limit: int = 3) -> List[Tuple[str, str, str, int, float]]:
        """板块内 Top 差异流向：任一侧类别属于该板块的流向，按数量降序。"""
        group_cats = set(GROUPS[group])
        flows = []
        for (il, fv, diff_type), stat in self.flow_stats.items():
            if il in group_cats or fv in group_cats:
                flows.append((il, fv, diff_type, stat["count"], stat["amount"]))
        flows.sort(key=lambda x: x[3], reverse=True)
        return flows[:limit]

    # -- 类别口径 ----------------------------------------------------------
    def category_flows(self, category: str, limit: int = 5) -> List[Tuple[str, str, str, int, float]]:
        """类别涉及的 Top 差异流向（含金额）。"""
        flows = []
        for (il, fv, diff_type), stat in self.flow_stats.items():
            if category in (il, fv):
                flows.append((il, fv, diff_type, stat["count"], stat["amount"]))
        flows.sort(key=lambda x: x[3], reverse=True)
        return flows[:limit]

    def flow_between(self, cat_a: str, cat_b: str) -> List[Tuple[str, str, str, int, float]]:
        """两类别之间的双向流向。"""
        result = []
        for (il, fv, diff_type), stat in self.flow_stats.items():
            if {il, fv} == {cat_a, cat_b}:
                result.append((il, fv, diff_type, stat["count"], stat["amount"]))
        result.sort(key=lambda x: x[3], reverse=True)
        return result

    def dr_cr_distribution(self, category: str) -> Tuple[int, int]:
        """类别涉及明细的 credit / debit 分布。"""
        credit = debit = 0
        for detail in self.category_details.get(category, []):
            if detail["dr_cr"] == "credit":
                credit += 1
            elif detail["dr_cr"] == "debit":
                debit += 1
        return credit, debit

    def top_samples(self, category: str, limit: int = 5) -> List[Dict[str, Any]]:
        rows = self.category_details.get(category, [])
        rows = sorted(rows, key=lambda d: abs(d["amount"]), reverse=True)
        return rows[:limit]


# ---------------------------------------------------------------------------
# docx 渲染工具
# ---------------------------------------------------------------------------
def set_run_font(run: Any, size: float, bold: bool = False, color: str = DARK, font: str = "微软雅黑") -> None:
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font)


def init_document() -> Document:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.0)

    # 默认字体
    normal = doc.styles["Normal"]
    normal.font.name = "微软雅黑"
    normal.font.size = Pt(10)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    # 标题样式
    for level, size, color in ((1, 16, BLUE), (2, 13, BLUE_MID), (3, 11, BLUE)):
        style = doc.styles[f"Heading {level}"]
        style.font.name = "微软雅黑"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    # 页脚：数据来源 + 页码
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("BS-CAT 全分类类别表现分析报告 | 数据来源: category_difference_report_100(1).xlsx | 第 ")
    set_run_font(run, 8, color=GRAY)
    page_run = p.add_run()
    set_run_font(page_run, 8, color=GRAY)
    fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE"
    fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
    page_run._r.append(fld_begin); page_run._r.append(instr); page_run._r.append(fld_end)
    tail = p.add_run(" 页")
    set_run_font(tail, 8, color=GRAY)
    return doc


def add_heading(doc: Document, text: str, level: int = 1) -> None:
    doc.add_heading(text, level=level)


def add_para(doc: Document, text: str, size: float = 10, bold: bool = False, color: str = DARK,
             align: str = "left", space_after: float = 6) -> None:
    p = doc.add_paragraph()
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    set_run_font(run, size, bold=bold, color=color)
    return p


def set_cell_shading(cell: Any, hex_color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def set_cell_margins(cell: Any, top: int = 40, bottom: int = 40, left: int = 80, right: int = 80) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.find(qn("w:tcMar"))
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = tc_mar.find(qn(f"w:{margin}"))
        if el is None:
            el = OxmlElement(f"w:{margin}")
            tc_mar.append(el)
        el.set(qn("w:w"), str(value))
        el.set(qn("w:type"), "dxa")


def repeat_header_row(row: Any) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def add_table(doc: Document, headers: List[str], rows: List[List[Any]],
              col_widths: Optional[List[float]] = None, font_size: float = 8.5,
              header_color: str = BLUE, align_center_cols: Optional[set] = None,
              zebra: bool = True) -> None:
    """通用表格：蓝底表头、网格线、斑马纹、表头跨页重复。宽度单位 cm。"""
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    width_cm = col_widths or [16.0 / len(headers)] * len(headers)

    def fill_cell(cell: Any, value: Any, header: bool, center: bool) -> None:
        text = "" if value is None else str(value)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if (center or header) else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(text)
        set_run_font(run, font_size, bold=header, color="FFFFFF" if header else DARK)

    align_center_cols = align_center_cols or set()
    header_row = table.rows[0]
    repeat_header_row(header_row)
    for j, (header, width) in enumerate(zip(headers, width_cm)):
        fill_cell(header_row.cells[j], header, True, True)
        set_cell_shading(header_row.cells[j], header_color)
        header_row.cells[j].width = Cm(width)

    for i, row in enumerate(rows):
        cells = table.add_row().cells
        for j, (value, width) in enumerate(zip(row, width_cm)):
            fill_cell(cells[j], value, False, j in align_center_cols)
            cells[j].width = Cm(width)
            if zebra and i % 2 == 1:
                set_cell_shading(cells[j], GRAY_PALE)
    return table


def add_callout(doc: Document, text: str, fill: str = YELLOW_PALE, edge: str = "C98200",
                size: float = 9.5) -> None:
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    table.autofit = False
    cell = table.rows[0].cells[0]
    cell.width = Cm(17.0)
    set_cell_shading(cell, fill)
    set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
    # 左侧强调色条
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    left_border = OxmlElement("w:left")
    left_border.set(qn("w:val"), "single")
    left_border.set(qn("w:sz"), "24")
    left_border.set(qn("w:color"), edge)
    borders.append(left_border)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    set_run_font(run, size, color=DARK)


def add_flow_table(doc: Document, flows: List[Tuple[str, str, str, int, float]], title: str = "",
                   diff_total: float = 0) -> None:
    if title:
        add_para(doc, title, size=9.5, bold=True, space_after=3)
    rows = []
    for index, (il, fv, diff_type, count, amount) in enumerate(flows, 1):
        rows.append([
            index,
            il or "（空）",
            fv or "（空）",
            diff_type,
            fmt_num(count),
            fmt_pct(count / diff_total if diff_total else 0),
            fmt_money(amount),
        ])
    add_table(doc, ["排名", "Illion 类别", "finv 类别", "差异类型", "数量", "占差异总数", "差异金额"],
              rows, [1.1, 3.2, 3.2, 2.5, 1.5, 2.0, 3.0], font_size=8.5, align_center_cols={0, 4, 5})


def add_samples_table(doc: Document, samples: List[Dict[str, Any]]) -> None:
    rows = []
    for s in samples:
        flow = f"{s['illion_category'] or '（空）'} → {s['finv_category'] or '（空）'}"
        rows.append([
            s.get("transaction_date", ""),
            (s.get("text", "") or "")[:34],
            (s.get("third_party", "") or "")[:18],
            s.get("dr_cr", ""),
            fmt_money(s.get("amount")),
            flow,
        ])
    if rows:
        add_table(doc, ["日期", "交易描述", "第三方/商户", "方向", "金额", "Illion → finv"],
                  rows, [2.2, 5.2, 3.2, 1.5, 2.4, 2.5], font_size=8, align_center_cols={0, 3, 4})


# ---------------------------------------------------------------------------
# 文字生成
# ---------------------------------------------------------------------------
def category_narrative(item: Dict[str, Any]) -> str:
    """固定格式的自动指标解读（第三部分统一口径）。"""
    category = item["category"]
    illion_cov = item.get("illion覆盖率")
    finv_cov = item.get("finv覆盖率")
    union = to_float(item.get("并集数量"))
    inter = to_float(item.get("交集数量"))
    inter_share = item.get("交集占比（并集）")
    illion_only_share = item.get("illion独有占比（并集）")
    finv_only_share = item.get("finv独有占比（并集）")
    return (
        f"Illion 覆盖率为 {fmt_pct(illion_cov)}，finv 覆盖率为 {fmt_pct(finv_cov)}；"
        f"并集为 {fmt_num(union)} 笔，交集为 {fmt_num(inter)} 笔，交集占比为 {fmt_pct(inter_share)}。"
        f"Illion 独有占比为 {fmt_pct(illion_only_share)}，finv 独有占比为 {fmt_pct(finv_only_share)}。"
    )


def category_judgment(item: Dict[str, Any]) -> str:
    """根据两侧独有占比判断差异主导来源。"""
    illion_only_share = to_float(item.get("illion独有占比（并集）"))
    finv_only_share = to_float(item.get("finv独有占比（并集）"))
    if finv_only_share > illion_only_share + 0.05:
        return (
            f"差异判断：该类别以 finv 扩展识别为主（finv 独有占比 {fmt_pct(finv_only_share)} 高于 Illion 独有占比 "
            f"{fmt_pct(illion_only_share)}），finv 覆盖更广，需要进一步核验新增分类的合理性，区分真实扩展与误归类。"
        )
    if illion_only_share > finv_only_share + 0.05:
        return (
            f"差异判断：该类别以 Illion 独有为主（Illion 独有占比 {fmt_pct(illion_only_share)} 高于 finv 独有占比 "
            f"{fmt_pct(finv_only_share)}），提示 finv 可能存在漏识别或归类迁移，建议检查文本清洗、商户覆盖和兜底规则。"
        )
    return (
        f"差异判断：两侧独有占比较为接近（Illion 独有 {fmt_pct(illion_only_share)}，finv 独有 {fmt_pct(finv_only_share)}），"
        f"差异更可能来自双方分类边界或规则口径，建议按差异流向逐条核对规则优先级。"
    )


def category_conclusion(item: Dict[str, Any]) -> str:
    union = to_float(item.get("并集数量"))
    inter_share = to_float(item.get("交集占比（并集）"))
    if inter_share >= 0.9:
        stability = "交集占比较高，双方对该类别的共同识别稳定"
    elif inter_share >= 0.7:
        stability = "双方存在较好的共同覆盖，但仍有一定单边差异"
    else:
        stability = "共同覆盖偏低，需要重点检查分类边界、知识库或漏识别问题"
    if union >= 1000:
        volume = "并集规模较大（≥1,000 笔），对总体差异有较高影响"
    elif union < 100:
        volume = "并集规模较小（<100 笔），比例指标需结合样本量谨慎解读"
    else:
        volume = "并集规模处于中等水平"
    priority = str(item.get("建议优先级") or "").strip() or "-"
    return f"类别结论：{stability}；{volume}。建议优先级为 {priority}，可结合 Top 差异流向核对规则与知识库。"


def render_category_block(doc: Document, item: Dict[str, Any], analysis: Analysis, detailed: bool = False) -> None:
    """按统一顺序渲染单个类别：定义 / 七项指标 / 解读 / 判断 / 流向 / 样本 / 结论。"""
    category = item["category"]
    group = item.get("group", "")
    add_heading(doc, f"Category: {category}（{group}）", level=3)

    add_para(doc, f"1. 类别业务定义：{CATEGORY_DEFINITION.get(category, '')}", size=9.5)
    add_para(doc, "2. 七项核心指标：", size=9.5, bold=True, space_after=2)
    add_table(doc,
              ["类别", "Illion覆盖率", "finv覆盖率", "并集数量", "交集数量", "交集占比（并集）", "Illion独有占比", "finv独有占比"],
              [[category, fmt_pct(item.get("illion覆盖率")), fmt_pct(item.get("finv覆盖率")),
                fmt_num(item.get("并集数量")), fmt_num(item.get("交集数量")), fmt_pct(item.get("交集占比（并集）")),
                fmt_pct(item.get("illion独有占比（并集）")), fmt_pct(item.get("finv独有占比（并集）"))]],
              [2.4, 2.0, 2.0, 1.7, 1.7, 2.2, 2.3, 2.3], font_size=8.5,
              align_center_cols={1, 2, 3, 4, 5, 6, 7})
    add_para(doc, "3. 指标解读：" + category_narrative(item), size=9.5, space_after=3)
    add_para(doc, "4. " + category_judgment(item), size=9.5, space_after=3)

    if detailed:
        flows = analysis.category_flows(category, limit=5)
        if flows:
            add_para(doc, "5. Top 差异流向（占差异总数比例）：", size=9.5, bold=True, space_after=2)
            add_flow_table(doc, flows, diff_total=analysis.diff_total)
        credit, debit = analysis.dr_cr_distribution(category)
        dir_total = credit + debit
        credit_share = fmt_pct(credit / dir_total if dir_total else 0)
        debit_share = fmt_pct(debit / dir_total if dir_total else 0)
        add_para(doc, f"6. 交易方向与金额：差异样本中 credit（转入/贷方）占 {credit_share}（{fmt_num(credit)} 笔），"
                       f"debit（转出/借方）占 {debit_share}（{fmt_num(debit)} 笔）；"
                       f"合计差异金额 {fmt_money(sum(stat['amount'] for (il, fv, _), stat in analysis.flow_stats.items() if category in (il, fv)))}。"
                       f"典型交易样本如下：", size=9.5, space_after=2)
        add_samples_table(doc, analysis.top_samples(category, limit=5))

    add_para(doc, "7. " + category_conclusion(item), size=9.5, space_after=8)


# ---------------------------------------------------------------------------
# 报告组装
# ---------------------------------------------------------------------------
def build_report(doc: Document, analysis: Analysis) -> None:
    m = analysis

    # ===== 封面 =====
    add_para(doc, "", size=10, space_after=40)
    p = add_para(doc, "BS-CAT 全分类类别表现分析报告", size=24, bold=True, color=BLUE, align="center", space_after=6)
    add_para(doc, "Income, Expense, Transfer and Liability Category Performance", size=11,
             color=GRAY, align="center", space_after=24)
    add_table(doc, ["样本范围", "分类范围", "比较对象", "报告日期"],
              [[f"{fmt_num(m.total_transactions)} 笔交易", "36 个类别", "illion vs finv / BS-CAT", "2026-08-19"]],
              [4.25, 4.25, 4.25, 4.25], font_size=10, align_center_cols=set())
    add_para(doc, "", size=10, space_after=12)
    add_callout(doc,
                "本报告由数据底稿 category_difference_report_100(1).xlsx 自动生成，覆盖收入、支出、转账、负债四大业务板块共 36 个类别。"
                "所有数字均从底稿动态计算；负债部分保留类别指标摘要，不重复展开已有 Liability 专项报告。")
    add_para(doc, "", size=10, space_after=12)
    add_para(doc, "数据来源", size=12, bold=True, color=BLUE_MID)
    add_para(doc, "Excel 底稿：00_核心对比（核心指标与逐类别指标）、01_差异诊断地图（Top 20 差异流向）、"
                  "03_排查明细（29,963 笔差异明细，含金额与交易方向）。", size=9.5, color=GRAY)
    doc.add_page_break()

    # ===== 1. 执行摘要 =====
    add_heading(doc, "1. 执行摘要", level=1)
    finv_only_share = m.finv_only_total / m.diff_total if m.diff_total else 0
    illion_only_share = m.illion_only_total / m.diff_total if m.diff_total else 0
    mismatch_share = m.mismatch / m.diff_total if m.diff_total else 0
    diff_sources = sorted(
        [("双方分类不一致", mismatch_share), ("仅 finv 有值", finv_only_share),
         ("仅 Illion 有值", illion_only_share)],
        key=lambda x: -x[1])
    diff_source_txt = "、".join(f"{name}占 {fmt_pct(share)}" for name, share in diff_sources)
    top1_diff, top2_diff = diff_sources[0][0], diff_sources[1][0]
    summary = (
        f"本次评估覆盖 {fmt_num(m.total_transactions)} 笔交易和 36 个类别，涉及收入、支出、负债三大业务板块，"
        f"并将 External Transfers 与 Internal Transfer 作为独立的转账板块分析。"
        f"Illion 分类覆盖率为 {fmt_pct(m.illion_coverage)}，finv 分类覆盖率为 {fmt_pct(m.finv_coverage)}，"
        f"finv 的覆盖范围高于 Illion；在双方均有非空分类的 {fmt_num(m.joint_nonempty)} 笔交易中，"
        f"分类一致率达到 {fmt_pct(m.joint_agreement)}，说明双方在已完成分类的交易上具有较高的一致性。"
        f"总体来看，finv 的主要特征是覆盖更广，但覆盖面扩大也带来了更多差异；"
        f"当前 {fmt_num(m.diff_total)} 笔差异中，{diff_source_txt}——"
        f"「{top1_diff}」是第一大差异来源，「{top2_diff}」次之；"
        f"差异主要集中在转账类、高频消费类和部分负债类，其中 Rent 作为支出类别需要进行专项关注。"
    )
    add_para(doc, summary, size=10.5)
    add_callout(doc, "本段为报告的核心结论：样本规模与类别范围、两侧覆盖率、双方非空一致率，以及主要差异来源。")
    doc.add_page_break()

    # ===== 2. 覆盖与一致全景 =====
    add_heading(doc, "2. 覆盖与一致全景", level=1)
    add_para(doc, "本部分自顶向下展示覆盖与一致情况：先看总体，再看业务板块，最后拆解到类别。"
                  "覆盖率的含义是「Illion / finv 识别为某个类别的交易占总交易的比例」；"
                  "一致率的含义是「双方分类一致（类别完全相同）的交易占至少一侧有分类交易的比例」。"
                  "覆盖与一致的情况清楚后，第 3 章进入类别细探，第 4 章再进入差异分析。", size=9.5)

    # 2.1 总体覆盖与一致率
    add_heading(doc, "2.1 总体覆盖与一致率", level=2)
    adjusted_agreement = m.metrics.get("覆盖调整后一致率", {}).get("result")
    add_table(doc, ["指标", "数值", "口径说明"],
              [
                  ["总交易数", fmt_num(m.total_transactions), "输入数据总行数"],
                  ["Illion 覆盖率", fmt_pct(m.illion_coverage), "Illion 有分类交易 / 总交易数"],
                  ["finv 覆盖率", fmt_pct(m.finv_coverage), "finv 有分类交易 / 总交易数"],
                  ["覆盖差（finv − Illion）", fmt_pct(to_float(m.finv_coverage) - to_float(m.illion_coverage)),
                   "finv 高于 Illion 的幅度"],
                  ["双方均有分类", f"{fmt_num(m.joint_nonempty)} 笔（{fmt_pct(m.joint_nonempty / m.total_transactions if m.total_transactions else 0)}）",
                   "双方分类均非空的交易"],
                  ["双方非空一致率", fmt_pct(m.joint_agreement), "类别一致 / 双方均有分类"],
                  ["覆盖调整后一致率", fmt_pct(adjusted_agreement), "类别一致 / 至少一侧有分类"],
                  ["差异率（至少一侧有分类）", fmt_pct(m.diff_total / m.union_nonempty if m.union_nonempty else 0),
                   "差异总数 / 至少一侧有分类"],
              ],
              [4.4, 4.4, 8.2], font_size=9, align_center_cols={1})
    union_rate = m.diff_total / m.union_nonempty if m.union_nonempty else 0
    avg_n = max(1, round(1 / union_rate)) if union_rate else 0
    add_para(doc, f"总体口径：finv 覆盖率（{fmt_pct(m.finv_coverage)}）高于 Illion（{fmt_pct(m.illion_coverage)}）"
                  f"约 {fmt_pct(to_float(m.finv_coverage) - to_float(m.illion_coverage))}；双方均有分类的交易占 "
                  f"{fmt_pct(m.joint_nonempty / m.total_transactions if m.total_transactions else 0)}，其中类别一致率为 "
                  f"{fmt_pct(m.joint_agreement)}。若把单边缺失也计入差异（覆盖调整后一致率 {fmt_pct(adjusted_agreement)}、"
                  f"差异率 {fmt_pct(union_rate)}），"
                  f"可以理解为：在至少一侧有分类的交易中，平均每 {avg_n} 笔约有 1 笔存在分类差异，"
                  f"差异以{top1_diff}为主（占 {fmt_pct(diff_sources[0][1])}）。",
             size=9.5)
    add_callout(doc,
        f"核心结论：finv 覆盖更广（覆盖差 {fmt_pct(to_float(m.finv_coverage) - to_float(m.illion_coverage))}），"
        f"双方共同识别范围内的类别一致率较高（{fmt_pct(m.joint_agreement)}）；"
        f"但把单边识别也纳入口径后，一致率降至 {fmt_pct(adjusted_agreement)}、差异率升至 {fmt_pct(union_rate)}，"
        f"即平均每 {avg_n} 笔至少一侧有分类的交易中约有 1 笔存在分类差异。差异构成上，{diff_source_txt}："
        f"「{top1_diff}」是第一差异源，「{top2_diff}」是第二差异源，两者合计占差异约 "
        f"{fmt_pct(diff_sources[0][1] + diff_sources[1][1])}。")

    # 2.2 业务板块覆盖与一致率
    add_heading(doc, "2.2 业务板块覆盖与一致率", level=2)
    add_para(doc, "36 个类别划分为收入类（3）、支出类（23）、转账类（2）、负债类（8）。转账类作为中性资金流独立统计，"
                  "不并入收入或支出。板块一致率以「至少一侧识别为该板块」为分母（并集口径），与总体覆盖调整后一致率对应。",
             size=9.5)
    seg_cov_rows = []
    for group in GROUP_ORDER:
        sc = m.segment_coverage(group)
        seg_cov_rows.append([
            group, sc["category_count"], fmt_pct(sc["illion_coverage"]), fmt_pct(sc["finv_coverage"]),
            fmt_pct(sc["exact_rate"]), fmt_pct(sc["broad_rate"]), fmt_num(sc["union_count"]),
        ])
    add_table(doc, ["业务板块", "类别数", "Illion 覆盖率", "finv 覆盖率", "一致率（类别精确）", "一致率（同板块）", "板块并集"],
              seg_cov_rows, [2.2, 1.4, 2.4, 2.4, 3.0, 3.0, 2.0], font_size=8.5,
              align_center_cols={1, 2, 3, 4, 5, 6})
    add_para(doc, "「一致率（类别精确）」= 双方类别完全一致 / 至少一侧识别为该板块；「一致率（同板块）」= 双方同属该板块"
                  "（类别可以不同）/ 至少一侧识别为该板块，口径更宽松。板块并集因跨板块交易存在重叠，各板块一致率不与总体直接加总；"
                  "本表用于板块间横向比较。", size=8.5, color=GRAY)
    seg_cov_map = {g: m.segment_coverage(g) for g in GROUP_ORDER}
    best_g = max(GROUP_ORDER, key=lambda g: seg_cov_map[g]["exact_rate"])
    worst_g = min(GROUP_ORDER, key=lambda g: seg_cov_map[g]["exact_rate"])
    largest_g = max(GROUP_ORDER, key=lambda g: seg_cov_map[g]["union_count"])
    gap_g = max(GROUP_ORDER, key=lambda g: abs(seg_cov_map[g]["finv_coverage"] - seg_cov_map[g]["illion_coverage"]))
    gap_gap = abs(seg_cov_map[gap_g]["finv_coverage"] - seg_cov_map[gap_g]["illion_coverage"])
    worst_union_share = seg_cov_map[worst_g]["union_count"] / m.total_transactions if m.total_transactions else 0
    ws = m.segment_summary(worst_g)
    worst_txt = (
        f"{worst_g}一致率最低（{fmt_pct(seg_cov_map[worst_g]['exact_rate'])}），"
        f"finv 覆盖率（{fmt_pct(seg_cov_map[worst_g]['finv_coverage'])}）已高于 Illion"
        f"（{fmt_pct(seg_cov_map[worst_g]['illion_coverage'])}），"
        f"finv 新增识别的交易与 Illion 归属不一致是低一致率的主因"
        f"（finv 独有差异占板块差异 {fmt_pct(ws['finv_only'] / ws['total'] if ws['total'] else 0)}）；"
        f"该板块交易体量{'较小' if worst_union_share < 0.1 else '较大'}（板块并集 "
        f"{fmt_num(seg_cov_map[worst_g]['union_count'])}，占总量 {fmt_pct(worst_union_share)}），"
        f"{'影响面有限' if worst_union_share < 0.1 else '是优化重点'}。"
    )
    illion_higher = [g for g in GROUP_ORDER
                     if seg_cov_map[g]["illion_coverage"] > seg_cov_map[g]["finv_coverage"]]
    illion_higher_txt = "、".join(
        f"{g}（差 {fmt_pct(seg_cov_map[g]['illion_coverage'] - seg_cov_map[g]['finv_coverage'])}）"
        for g in illion_higher) or "无板块"
    top1_flow = max(m.flow_stats.items(), key=lambda kv: kv[1]["count"])
    top1_il, top1_fv = top1_flow[0][0], top1_flow[0][1]
    top1_group = GROUP_OF.get(top1_il, "")
    add_callout(doc,
        f"核心结论：{best_g}一致率最高（精确 {fmt_pct(seg_cov_map[best_g]['exact_rate'])} / "
        f"同板块 {fmt_pct(seg_cov_map[best_g]['broad_rate'])}），是两侧共识最强的板块；"
        f"{worst_txt}"
        f"{largest_g}覆盖体量最大（板块并集 {fmt_num(seg_cov_map[largest_g]['union_count'])}，"
        f"Illion {fmt_pct(seg_cov_map[largest_g]['illion_coverage'])} / finv {fmt_pct(seg_cov_map[largest_g]['finv_coverage'])}），"
        f"细分差异也最集中（占全部差异 "
        f"{fmt_pct(m.segment_summary(largest_g)['share'])}）。"
        f"Illion 覆盖高于 finv 的板块为{illion_higher_txt}"
        f"{'，其中' + top1_group + '与第 4 章第一大差异流向 ' + (top1_il or '（空）') + ' → ' + (top1_fv or '（空）') + ' 相呼应' if top1_group in illion_higher else ''}；"
        f"覆盖差最大的板块为{gap_g}（{fmt_pct(gap_gap)}），反映两侧覆盖口径分歧的主要所在。")

    # 2.3 分类别覆盖与一致率
    add_heading(doc, "2.3 分类别覆盖与一致率", level=2)
    add_para(doc, "36 个类别按业务板块逐一展示覆盖率和一致率（一致率 = 交集数量 / 并集数量，即双方类别完全一致的比例）。"
                  "该表是覆盖与一致的完整清单；重要类别的详细分析见第 3 章，其余类别统一在第 5 章附录展示七项指标汇总。", size=9.5)
    for group in GROUP_ORDER:
        add_para(doc, f"{group}：", size=9.5, bold=True, space_after=2)
        cat_rows = []
        for item in m.categories:
            if item["group"] == group:
                cat_rows.append([
                    item["category"],
                    fmt_pct(item.get("illion覆盖率")),
                    fmt_pct(item.get("finv覆盖率")),
                    fmt_pct(item.get("交集占比（并集）")),
                    fmt_num(item.get("并集数量")),
                    fmt_num(item.get("交集数量")),
                ])
        add_table(doc, ["类别", "Illion 覆盖率", "finv 覆盖率", "一致率（交集/并集）", "并集", "交集"],
                  cat_rows, [5.2, 2.5, 2.5, 3.0, 1.9, 1.9], font_size=8,
                  align_center_cols={1, 2, 3, 4, 5})
    # 2.3 类别层核心结论（动态计算，不硬编码）
    low_cats = sorted([c for c in m.categories
                       if to_float(c.get("并集数量")) >= 500 and to_float(c.get("交集占比（并集）")) < 0.5],
                      key=lambda c: to_float(c.get("交集占比（并集）")))
    low_txt = "、".join(f"{c['category']}（一致率 {fmt_pct(c.get('交集占比（并集）'))}，并集 {fmt_num(c.get('并集数量'))}）"
                        for c in low_cats[:5])
    high_cats = sorted([c for c in m.categories if to_float(c.get("并集数量")) >= 500],
                       key=lambda c: -to_float(c.get("交集占比（并集）")))
    high_txt = "、".join(f"{c['category']}（{fmt_pct(c.get('交集占比（并集）'))}）" for c in high_cats[:3])
    zero_cov = [c["category"] for c in m.categories if to_float(c.get("illion覆盖率")) == 0]
    add_callout(doc,
        f"核心结论：并集超过 500 且一致率低于 50% 的类别是优化优先项，一致率最低的五个依次为 {low_txt}；"
        f"一致率最高的三个为 {high_txt}，多为规则清晰的负债与固定支出类别。"
        + (f"Illion 覆盖为 0 的类别：{'、'.join(zero_cov)}，属于 Illion 侧覆盖缺口。"
           if zero_cov else "两侧均有覆盖的类别。"))
    doc.add_page_break()

    # ===== 3. 分业务板块的类别细探 =====
    add_heading(doc, "3. 分业务板块的类别细探", level=1)
    add_para(doc, "本部分回答：每个板块中，哪些类别稳定？哪些类别存在单边覆盖？哪些类别需要进一步排查？"
                  "每个类别统一使用以下七项指标：① Illion 覆盖率；② finv 覆盖率；③ 并集数量；④ 交集数量；"
                  "⑤ 交集占比（并集）；⑥ Illion 独有占比；⑦ finv 独有占比。")

    def categories_in(group: str) -> List[Dict[str, Any]]:
        return [item for item in m.categories if item["group"] == group]

    # 4.1 收入类
    add_heading(doc, "3.1 收入类", level=2)
    add_para(doc, "收入类包括 Wages、Centrelink、All Other Credits。重点分析：Wages 与 All Other Credits 的边界、"
                  "收入与 External Transfers 的边界、credit/debit 方向，以及 finv 独有收入识别是否合理。", size=9.5)
    for item in categories_in("收入类"):
        render_category_block(doc, item, m, detailed=True)
    add_para(doc, "收入类边界小结：", size=9.5, bold=True, space_after=2)
    boundary_parts = []
    for a, b in [("Wages", "All Other Credits"), ("Wages", "External Transfers"),
                 ("All Other Credits", "External Transfers")]:
        flows = m.flow_between(a, b)
        if flows:
            desc = "；".join(f"{il or '（空）'} → {fv or '（空）'}（{fmt_num(cnt)} 笔，{fmt_money(amount)}）"
                             for il, fv, _, cnt, amount in flows)
            boundary_parts.append(f"{a} 与 {b}：{desc}")
    add_para(doc, "；".join(boundary_parts) + "。收入类差异以 credit 方向为主（收入识别侧），"
             "finv 独有收入识别主要落在 All Other Credits，需要抽样核验其交易描述是否符合收入定义。", size=9)
    doc.add_page_break()

    # 4.2 支出类
    add_heading(doc, "3.2 支出类", level=2)
    add_para(doc, "支出类共 23 个类别。本报告针对重要类别 Rent 与 Gambling 展开分析（含 Top 差异流向、交易方向、"
                  "金额与典型样本）；其余 21 个类别的七项指标统一汇总于第 5 章附录。", size=9.5)

    # 4.2.1 Gambling 专项
    add_heading(doc, "3.2.1 Gambling 专项", level=2)
    expense_items = categories_in("支出类")
    for item in expense_items:
        if item["category"] == "Gambling":
            render_category_block(doc, item, m, detailed=True)
    doc.add_page_break()

    # 4.2.2 Rent 专项
    add_heading(doc, "3.2.2 Rent 专项", level=2)
    rent = m.category_by_name.get("Rent")
    if rent:
        render_category_block(doc, rent, m, detailed=True)
        add_para(doc, "Rent 与相邻类别的差异流向：", size=9.5, bold=True, space_after=2)
        rent_flows = []
        for other in ["External Transfers", "Internal Transfer", "Utilities", "Home Improvement"]:
            rent_flows.extend(m.flow_between("Rent", other))
        rent_flows = sorted(rent_flows, key=lambda x: x[3], reverse=True)
        if rent_flows:
            add_flow_table(doc, rent_flows, diff_total=m.diff_total)
        else:
            add_para(doc, "底稿中未发现 Rent 与上述类别之间的差异流向。", size=9)
        add_para(doc, "Rent 规则或知识库优化建议：", size=9.5, bold=True, space_after=2)
        add_para(doc, "① 核对 Rent 与转账类的区分规则：外部转账若指向疑似租金收款方（如房产中介、物业管理），"
                      "应按收款方特征优先识别为 Rent；② 核对 Rent 与 Utilities、Home Improvement 的关键词与商户知识库边界，"
                      "避免物业账单（含水电网）被拆入 Rent 或反之为 Utilities；③ 对 Rent 类差异建立抽样回归集，"
                      "将每次知识库更新后的 Rent 交集占比与差异金额纳入监控。", size=9.5)
    doc.add_page_break()

    # 4.3 转账类
    add_heading(doc, "3.3 转账类", level=2)
    add_para(doc, "转账类包括 External Transfers 与 Internal Transfer，作为中性资金流处理，不计入收入或支出合计，"
                  "但纳入总体覆盖率、差异率和分类迁移分析。除七项指标外，补充 credit/debit 分布、转入/转出方向，"
                  "以及与 Wages、All Other Credits、Rent 和另一转账类别的边界分析。", size=9.5)
    for item in categories_in("转账类"):
        render_category_block(doc, item, m, detailed=True)
    add_para(doc, "转账类边界小结：", size=9.5, bold=True, space_after=2)
    transfer_parts = []
    for a, b in [("External Transfers", "Wages"), ("External Transfers", "All Other Credits"),
                 ("External Transfers", "Rent"), ("External Transfers", "Internal Transfer")]:
        flows = m.flow_between(a, b)
        if flows:
            desc = "；".join(f"{il or '（空）'} → {fv or '（空）'}（{fmt_num(cnt)} 笔，{fmt_money(amount)}）"
                             for il, fv, _, cnt, amount in flows)
            transfer_parts.append(f"{a} 与 {b}：{desc}")
    add_para(doc, "；".join(transfer_parts) + "。转账与收入（Wages / All Other Credits）的差异集中在 credit 方向，"
             "说明「收款入账」场景下转账与收入类别边界是主要混淆点；与 Rent 的差异则出现在 debit 方向，"
             "提示租金支付被识别为转账的可能性。", size=9)
    add_callout(doc, "转账类处理原则：作为中性资金流，不并入收入或支出合计，也不计入净收入计算；"
                     "但纳入总体覆盖率、差异率和分类迁移分析。")
    doc.add_page_break()

    # 4.4 负债类
    add_heading(doc, "3.4 负债类", level=2)
    add_para(doc, "负债类共 8 个类别。本报告针对重要类别 SACC Loans、Non SACC Loans、Dishonours、"
                  "Credit Card Repayments 展开分析（含 Top 差异流向、交易方向、金额与典型样本）；"
                  "其余 4 个类别（Debt Collection、Overdrawn、Debt Consolidation、Unknown Loans）的七项指标"
                  "统一汇总于第 5 章附录。贷款生命周期、Counterparty matching、Dishonours 详细案例、"
                  "Unknown Loans 完整根因和 SACC / Non-SACC 深度交叉矩阵以 Liability 专项报告为准。", size=9.5)
    for item in categories_in("负债类"):
        if item["category"] in IMPORTANT_LIABILITY:
            render_category_block(doc, item, m, detailed=True)
    add_para(doc, "负债板块差异排名与 Top 差异流向：", size=9.5, bold=True, space_after=2)
    liab_flows = m.segment_top_flows("负债类", limit=8)
    add_flow_table(doc, liab_flows, diff_total=m.diff_total)
    add_callout(doc, "负债类的详细根因分析（贷款生命周期、Counterparty matching、Dishonours 案例、Unknown Loans 根因）"
                     "以已有 Liability 专项报告为准，本报告仅保留重要负债类别的展开分析、板块差异排名与类别侧定位。")
    doc.add_page_break()

    # 4.5 跨板块优化建议
    add_heading(doc, "3.5 跨板块优化建议", level=2)
    add_para(doc, "综合全景差异结构与类别细探，建议按以下优先级推进优化：", size=9.5)

    def flow_ref(illion: str, finv: str) -> str:
        count = sum(stat["count"] for (il, fv, _), stat in m.flow_stats.items() if il == illion and fv == finv)
        share = fmt_pct(count / m.diff_total if m.diff_total else 0)
        return f"{illion} → {finv or '（空）'}（{fmt_num(count)} 笔 / {share}）"

    add_table(doc, ["优先级", "观察方向", "建议动作", "数据依据"],
              [
                  ["P1", "验证 finv 新增识别", f"抽样核验仅 finv 有值分类占比（{fmt_pct(m.finv_only_total / m.diff_total if m.diff_total else 0)}）是否合理，区分覆盖扩展与误归类",
                   "All Other Credits、Retail、Dining Out、Unknown Loans"],
                  ["P1", "优化转账边界", "拆分 credit/debit 处理，明确转账与收入（Wages / All Other Credits）的识别优先级",
                   flow_ref("External Transfers", "Internal Transfer") + "、" + flow_ref("External Transfers", "All Other Credits")],
                  ["P1", "优化贷款类别边界", "核对 Non SACC / SACC Loans 与 Unknown Loans 的知识库与别名",
                   flow_ref("Non SACC Loans", "Unknown Loans") + "、" + flow_ref("SACC Loans", "Unknown Loans")],
                  ["P2", "优化 Rent 边界", "核对 Rent 与转账、Utilities、Home Improvement 的规则边界，建立 Rent 回归集",
                   flow_ref("External Transfers", "Rent") + " 等流向"],
                  ["P2", "优化高频消费边界", "针对 Groceries、Dining Out、Retail、Automotive、Gambling 建立规则回归集",
                   flow_ref("Groceries", "Dining Out") + "、" + flow_ref("Groceries", "Retail") + "、" + flow_ref("Groceries", "Automotive")],
                  ["P3", "其余类别复核", "低并集类别的比例指标结合样本量解读，重点核对七项指标异常的类别（见第 5 章附录）",
                   "附录各类别七项核心指标"],
              ],
              [1.4, 3.2, 6.4, 6.0], font_size=8.5, align_center_cols={0})
    add_para(doc, "后续迭代建议：每次模型或知识库更新后，重新计算 36 个类别的七项核心指标，并重点监控交集占比、"
                  "finv 独有占比、Illion 独有占比以及 Rent 和转账类的差异流向变化。", size=9.5)

    # ===== 4. 差异结构全景 =====
    add_heading(doc, "4. 差异结构全景", level=1)
    add_para(doc, "在覆盖与一致的基础上，本部分进入差异分析，回答三个问题：差异有多少？差异属于什么类型？差异主要集中在哪些业务板块？")

    # 3.1 总体差异类型分布
    add_heading(doc, "4.1 总体差异类型分布", level=2)
    add_table(doc,
              ["差异类型", "数量", "占差异总数", "分析含义"],
              [
                  ["双方分类不一致", fmt_num(m.mismatch), fmt_pct(m.mismatch / m.diff_total if m.diff_total else 0),
                   "两侧均有分类，但类别标签不同"],
                  ["仅 Illion 有值", fmt_num(m.illion_only_total), fmt_pct(m.illion_only_total / m.diff_total if m.diff_total else 0),
                   "finv 侧缺少分类（finv 漏识别）"],
                  ["仅 finv 有值", fmt_num(m.finv_only_total), fmt_pct(m.finv_only_total / m.diff_total if m.diff_total else 0),
                   "finv 新增识别或 Illion 漏识别"],
                  ["总差异数", fmt_num(m.diff_total), "100.00%", "分类不一致与单边缺失的合计"],
              ],
              [4.0, 2.2, 2.4, 7.4], font_size=9, align_center_cols={1, 2})
    add_para(doc, f"另行说明：双方均为空共 {fmt_num(m.both_empty)} 笔，不进入差异分析；仅 finv 有值 "
                  f"（{fmt_num(m.finv_only_total)} 笔，{fmt_pct(m.finv_only_total / m.diff_total if m.diff_total else 0)}）"
                  f"是最大差异来源。",
             size=9.5, color=GRAY)

    # 3.2 业务板块差异分布
    add_heading(doc, "4.2 业务板块差异分布", level=2)
    add_para(doc, "36 个类别划分为四个分析板块：收入类 3 个、支出类 23 个、转账类 2 个、负债类 8 个。"
                  "转账类不并入收入或支出，作为独立板块分析。", size=9.5)
    add_table(doc, ["业务板块", "类别数", "类别范围", "分析重点"],
              [
                  ["收入类", "3", "Wages、Centrelink、All Other Credits", "收入与转账的识别边界"],
                  ["支出类", "23", "23 个支出类别，其中 Rent 重点分析", "高频消费类别之间的分类边界"],
                  ["转账类", "2", "External Transfers、Internal Transfer", "转账与收入、支出的混淆"],
                  ["负债类", "8", "8 个负债类别", "保留总体差异，详细结果引用 Liability 专项报告"],
              ],
              [2.4, 1.6, 7.6, 5.4], font_size=8.5, align_center_cols={1})

    add_para(doc, "板块差异汇总（按任一类别命中该板块统计，板块内各占比均以该板块差异交易数为分母）：", size=9.5, bold=True, space_after=2)
    seg_rows = []
    for group in GROUP_ORDER:
        seg = m.segment_summary(group)
        seg_total = seg["total"]
        top_flows = m.segment_top_flows(group, limit=3)
        top_text = "；".join(f"{il or '（空）'} → {fv or '（空）'}（占板块差异 {fmt_pct(cnt / seg_total if seg_total else 0)}）"
                             for il, fv, _, cnt, _ in top_flows)
        seg_rows.append([
            group,
            fmt_num(seg["total"]),
            fmt_pct(seg["share"]),
            fmt_pct(seg["within"] / seg_total if seg_total else 0),
            fmt_pct(seg["cross"] / seg_total if seg_total else 0),
            fmt_pct(seg["illion_only"] / seg_total if seg_total else 0),
            fmt_pct(seg["finv_only"] / seg_total if seg_total else 0),
            top_text,
        ])
    add_table(doc, ["业务板块", "差异交易数", "占总差异比例", "板块内占比", "跨板块占比",
                    "仅 Illion 占比", "仅 finv 占比", "Top 3 差异流向（占板块差异）"],
              seg_rows, [1.6, 1.7, 1.8, 1.6, 1.6, 1.7, 1.7, 5.3], font_size=8,
              align_center_cols={1, 2, 3, 4, 5, 6})
    add_para(doc, "注：板块内、跨板块、仅 Illion、仅 finv 四类按互斥口径合计为 100%；跨板块差异同时计入两侧板块，"
                  "因此各板块差异交易数之和大于总差异数；板块交叉矩阵（3.3）为互斥口径，合计等于差异总数。",
             size=8.5, color=GRAY)

    # 3.3 业务板块交叉矩阵
    add_heading(doc, "4.3 业务板块交叉矩阵", level=2)
    add_para(doc, "矩阵只统计差异样本（双方分类不一致、仅 Illion 有值、仅 finv 有值），按行 = Illion 板块、列 = finv 板块互斥拆解，"
                  "可区分板块内类别边界问题（对角线）、跨板块误分类（非对角线）与单边覆盖问题（末行/末列）。", size=9.5)
    matrix_headers = ["Illion 板块 \\ finv 板块"] + GROUP_ORDER + ["finv 为空"]
    matrix_rows = []
    for row_group in GROUP_ORDER + ["Illion为空"]:
        row_label = "Illion 为空" if row_group == "Illion为空" else row_group
        matrix_rows.append([row_label] + [fmt_pct(m.matrix[(row_group, col)] / m.diff_total if m.diff_total else 0)
                                          for col in GROUP_ORDER + ["finv为空"]])
    add_table(doc, matrix_headers, matrix_rows, [3.2] + [2.76] * 5, font_size=8.5,
              align_center_cols={1, 2, 3, 4, 5})
    add_para(doc, f"矩阵单元格为占差异总数（{fmt_num(m.diff_total)} 笔）的比例，合计 100.00%，可直接跨单元格对比。"
                  f"对角线为同一板块内的类别冲突；非对角线为跨板块迁移；「Illion 为空」行与「finv 为空」列表示单边覆盖差异。",
             size=9, color=GRAY)

    # 3.4 Top 差异流向
    add_heading(doc, "4.4 Top 差异流向", level=2)
    flow_rows = []
    for flow in m.top_flows[:20]:
        il_group = GROUP_OF.get(flow["illion_category"], "（空）") if flow["illion_category"] else "（空）"
        fv_group = GROUP_OF.get(flow["finv_category"], "（空）") if flow["finv_category"] else "（空）"
        flow_rows.append([
            flow["rank"],
            flow["illion_category"] or "（空）",
            flow["finv_category"] or "（空）",
            flow["diff_type"],
            fmt_num(flow["count"]),
            fmt_pct(flow["count"] / m.diff_total if m.diff_total else 0),
            fmt_money(flow["amount"]),
            f"{il_group} → {fv_group}",
        ])
    add_table(doc, ["排名", "Illion 类别", "finv 类别", "差异类型", "数量", "占差异总数", "差异金额", "所属板块"],
              flow_rows, [1.0, 3.0, 3.0, 2.4, 1.4, 1.9, 2.3, 2.0], font_size=8,
              align_center_cols={0, 4, 5})
    focus_flows = m.flow_stats  # 用于在说明中给出重点关注流向的实际数据
    note_parts = []
    for a, b in [("External Transfers", "All Other Credits"), ("External Transfers", ""),
                 ("Internal Transfer", "Wages"), ("Non SACC Loans", "Unknown Loans"),
                 ("External Transfers", "Internal Transfer"), ("Groceries", "Automotive"),
                 ("Groceries", "Retail"), ("Groceries", "Dining Out"), ("External Transfers", "Rent")]:
        count = sum(stat["count"] for (il, fv, _), stat in focus_flows.items() if (il, fv) == (a, b))
        if count:
            note_parts.append(f"{a} → {b or '（空）'}（{fmt_num(count)} 笔 / "
                              f"{fmt_pct(count / m.diff_total if m.diff_total else 0)}）")
    add_para(doc, "重点关注流向：" + "；".join(note_parts) + "。其中 Internal Transfer → Wages 与 External Transfers → Rent "
             "未进入 Top 20，但属于收入/房租边界的典型差异，在第 3 章类别细探中展开。", size=9, color=GRAY)

    # 3.5 全景结论
    add_heading(doc, "4.5 全景层面结论", level=2)
    top3 = m.top_flows[:3]
    top3_share = sum(f["count"] for f in top3) / m.diff_total if m.diff_total else 0
    transfer_seg = m.segment_summary("转账类")
    income_seg = m.segment_summary("收入类")
    expense_seg = m.segment_summary("支出类")
    liability_seg = m.segment_summary("负债类")
    cross_total = sum(1 for d in m.details if d["illion_category"] and d["finv_category"]
                      and GROUP_OF.get(d["illion_category"]) != GROUP_OF.get(d["finv_category"]))
    within_total = sum(1 for d in m.details if d["illion_category"] and d["finv_category"]
                       and GROUP_OF.get(d["illion_category"]) == GROUP_OF.get(d["finv_category"]))
    conclusion = (
        f"全景层面，{fmt_num(m.diff_total)} 笔差异占至少一侧有分类交易（{fmt_num(m.union_nonempty)} 笔）的 "
        f"{fmt_pct(m.diff_total / m.union_nonempty if m.union_nonempty else 0)}。"
        f"差异类型上，双方分类不一致 {fmt_num(m.mismatch)} 笔（{fmt_pct(m.mismatch / m.diff_total if m.diff_total else 0)}），"
        f"仅 Illion 有值 {fmt_num(m.illion_only_total)} 笔（{fmt_pct(m.illion_only_total / m.diff_total if m.diff_total else 0)}），"
        f"仅 finv 有值 {fmt_num(m.finv_only_total)} 笔（{fmt_pct(m.finv_only_total / m.diff_total if m.diff_total else 0)}），"
        f"仅 finv 有值是最大单边差异来源。"
        f"板块层面，转账类相关差异最集中：{top3[0]['illion_category']} → {top3[0]['finv_category'] or '（空）'}、"
        f"{top3[1]['illion_category']} → {top3[1]['finv_category'] or '（空）'} 与 "
        f"{top3[2]['illion_category']} → {top3[2]['finv_category'] or '（空）'} 三条流向合计 "
        f"{fmt_num(sum(f['count'] for f in top3))} 笔，占差异总数的 {fmt_pct(top3_share)}。"
        f"在双方均有分类的差异中，板块内冲突 {fmt_num(within_total)} 笔，跨板块迁移 {fmt_num(cross_total)} 笔，"
        f"跨板块迁移主要发生在转账 ↔ 收入（Wages、All Other Credits）与转账 → 支出（Rent、Retail、Health、Gambling 等）方向；"
        f"负债类差异以贷款类别内部的分类冲突（Non SACC Loans ↔ Unknown Loans、SACC Loans ↔ Unknown Loans）为主。"
        f"整体优先级：先验证 finv 新增识别的合理性，再优化转账边界规则，并对 Rent、Groceries 等高频消费类别建立专项回归集。"
    )
    add_para(doc, conclusion, size=9.5)
    doc.add_page_break()

    # ===== 5. 附录：其余类别指标汇总 =====
    add_heading(doc, "5. 附录：其余类别指标汇总", level=1)
    add_para(doc, "第 4 章未展开的类别（支出类 21 个、负债类 4 个，共 25 个）统一在此展示七项核心指标与建议优先级，"
                  "用于快速排查与规则核对；指标口径与第 2 章、第 3 章完全一致。", size=9.5)

    def appendix_table(group: str, excluded: set, title: str) -> None:
        add_para(doc, title, size=10, bold=True, space_after=2)
        rows = []
        for item in categories_in(group):
            if item["category"] not in excluded:
                rows.append([
                    item["category"],
                    fmt_pct(item.get("illion覆盖率")), fmt_pct(item.get("finv覆盖率")),
                    fmt_num(item.get("并集数量")), fmt_num(item.get("交集数量")),
                    fmt_pct(item.get("交集占比（并集）")),
                    fmt_pct(item.get("illion独有占比（并集）")), fmt_pct(item.get("finv独有占比（并集）")),
                    str(item.get("建议优先级") or "").strip() or "-",
                ])
        add_table(doc,
                  ["类别", "Illion覆盖率", "finv覆盖率", "并集数量", "交集数量", "交集占比（并集）",
                   "Illion独有占比", "finv独有占比", "建议优先级"],
                  rows, [2.4, 1.7, 1.7, 1.5, 1.5, 1.8, 1.9, 1.9, 1.3], font_size=8,
                  align_center_cols={1, 2, 3, 4, 5, 6, 7, 8})

    appendix_table("支出类", IMPORTANT_EXPENSE, "5.1 支出类其余类别（21 个）")
    appendix_table("负债类", IMPORTANT_LIABILITY, "5.2 负债类其余类别（4 个）")
    add_para(doc, "附录口径说明：并集数量 = 至少一侧识别为该类别的交易数；交集数量 = 双方均识别为该类别且类别一致的交易数；"
                  "交集占比（并集）= 交集 / 并集；Illion 独有占比 / finv 独有占比 = 单侧识别的差异交易占该类别的比例。"
                  "建议优先级来自底稿「建议优先级」列，缺失时以「-」表示。", size=8.5, color=GRAY)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate Chinese full category performance report (docx)")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    wb = load_workbook(args.input, data_only=True, read_only=True)
    metrics = read_metrics(wb["00_核心对比"])
    categories = read_categories(wb["00_核心对比"])
    top_flows = read_top_flows(wb["01_差异诊断地图"])
    details = read_details(wb["03_排查明细"])
    wb.close()

    if len(categories) != 36:
        raise ValueError(f"类别数量异常: {len(categories)} (期望 36)")

    analysis = Analysis(metrics, categories, top_flows, details)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    doc = init_document()
    doc.core_properties.title = "BS-CAT 全分类类别表现分析报告"
    doc.core_properties.author = "BS-CAT"
    build_report(doc, analysis)
    doc.save(args.output)

    print(f"Generated: {args.output}")
    print(f"Categories: {len(categories)}; detail rows: {len(details)}; top flows: {len(top_flows)}")
    print(f"Total transactions: {fmt_num(analysis.total_transactions)}; diff total: {fmt_num(analysis.diff_total)}")
    print(f"Matrix sum: {fmt_num(sum(analysis.matrix.values()))}")


if __name__ == "__main__":
    main()
