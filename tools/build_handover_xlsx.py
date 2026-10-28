#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成《值班交接单模板.xlsx》

输出：templates/值班交接单模板.xlsx
包含 5 个工作表：
    使用说明 / 交接单 / 示例 / 交接记录 / 风险清单

设计原则：
    1. 交接单的每个区块，对应 8 类交接风险中的一类（见风险清单 sheet）
    2. 最容易被漏掉的两块（临时规避措施、外部依赖状态）用深红色标注
    3. 状态、级别等字段全部做成下拉，避免"恢复了"这种模糊表述
    4. 表格中的示例数据全部为虚构的通用场景，不含任何真实信息

运行：
    python build_handover_xlsx.py
"""

import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "templates")
OUT_FILE = os.path.join(OUT_DIR, "值班交接单模板.xlsx")

# ---------------------------------------------------------------- 调色板
C_DARK = "1F3864"       # 深蓝：标题栏
C_SECTION = "2F5597"    # 中蓝：普通区块
C_DANGER = "9C0006"     # 深红：最容易漏的区块
C_COLHDR = "D9E1F2"     # 浅蓝灰：列头
C_LABEL = "EDEDED"      # 浅灰：标签
C_INPUT = "FFFFFF"      # 白：待填
C_RED_FILL = "FFC7CE"
C_RED_TEXT = "9C0006"
C_YELLOW = "FFEB9C"

THIN = Side(style="thin", color="B4B4B4")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FONT_BASE = Font(name="微软雅黑", size=9)


# ---------------------------------------------------------------- 工具函数
def style_range(ws, rng, *, font=None, fill=None, align=None, border=True):
    """给一个区域批量套样式。rng 形如 'A1:H1'。"""
    for row in ws[rng]:
        for cell in row:
            if font is not None:
                cell.font = font
            if fill is not None:
                cell.fill = fill
            if align is not None:
                cell.alignment = align
            if border:
                cell.border = BORDER


def fill_of(color):
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def banner(ws, r, ncols, text, color=C_DARK, height=24, size=11):
    """整行横幅：用于标题和区块标题。"""
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    cell = ws.cell(row=r, column=1, value=text)
    cell.font = Font(name="微软雅黑", size=size, bold=True, color="FFFFFF")
    cell.fill = fill_of(color)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for c in range(2, ncols + 1):
        ws.cell(row=r, column=c).fill = fill_of(color)
    ws.row_dimensions[r].height = height
    return r + 1


def table_header(ws, r, cols):
    for i, name in enumerate(cols, start=1):
        cell = ws.cell(row=r, column=i, value=name)
        cell.font = Font(name="微软雅黑", size=9, bold=True, color="1F3864")
        cell.fill = fill_of(C_COLHDR)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[r].height = 30
    return r + 1


def data_rows(ws, r, n, ncols, height=26):
    for _ in range(n):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill_of(C_INPUT)
            cell.border = BORDER
            cell.font = FONT_BASE
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws.row_dimensions[r].height = height
        r += 1
    return r


def spacer(ws, r, ncols, height=6):
    for c in range(1, ncols + 1):
        ws.cell(row=r, column=c).fill = fill_of("FFFFFF")
    ws.row_dimensions[r].height = height
    return r + 1


# ---------------------------------------------------------------- 区块定义
# 每个区块 = 一类交接风险。cols 固定 8 列，保证表格右边缘整齐。
SECTIONS = [
    {
        "key": "events",
        "title": "【一】进行中的事件（未闭环）  ★ 必填 · 对应风险 1、8",
        "danger": False,
        "cols": ["事件编号", "首报时间", "现象（只写看到的事实）", "严重级别",
                 "当前状态", "已做的处置", "下一步 / 责任人", "必须跟进时间"],
        "rows": 6,
        "dvs": {3: "severity", 4: "event_status"},
        "cf": [
            {"col": "E", "value": "未闭环", "fill": C_RED_FILL},
            {"col": "D", "value": "P1", "fill": C_RED_FILL, "bold": True},
        ],
    },
    {
        "key": "observing",
        "title": "【二】已恢复，但需要观察  · 对应风险 1、8",
        "danger": False,
        "cols": ["事件编号", "首报时间", "恢复时间", "现象",
                 "观察什么", "观察截止", "责任人", "再次发生的迹象"],
        "rows": 3,
        "dvs": {},
        "cf": [],
    },
    {
        "key": "changes",
        "title": "【三】进行中的变更  ★ 必填 · 对应风险 2",
        "danger": False,
        "cols": ["变更单号", "窗口开始", "窗口结束", "影响范围",
                 "回滚方案", "执行人", "联系方式", "是否已通知业务方"],
        "rows": 3,
        "dvs": {7: "yesno"},
        "cf": [],
    },
    {
        "key": "workarounds",
        "title": "【四】临时规避措施  ★★ 最容易被漏掉 · 对应风险 3",
        "danger": True,
        "cols": ["规避的是什么问题", "临时做法", "生效时间", "到期 / 失效时间",
                 "失效会发生什么", "责任人", "转正计划", "到期提醒方式"],
        "rows": 4,
        "dvs": {},
        "cf": [],
    },
    {
        "key": "external",
        "title": "【五】外部依赖状态（运营商 / 云服务商 / 第三方）  ★★ 对应风险 4",
        "danger": True,
        "cols": ["供应商 / 线路", "状态", "故障编号 / 公告", "影响范围",
                 "已知信息", "跟进人", "上次确认", "下次确认"],
        "rows": 3,
        "dvs": {1: "dep_status"},
        "cf": [
            {"col": "B", "value": "异常", "fill": C_RED_FILL},
        ],
    },
    {
        "key": "capacity",
        "title": "【六】临界资源预警（还没叫，但正在变坏）  · 对应风险 5",
        "danger": False,
        "cols": ["对象", "当前值", "阈值", "预计到达时间",
                 "责任人", "处理计划", "是否已提单", "备注"],
        "rows": 3,
        "dvs": {6: "yesno"},
        "cf": [],
    },
    {
        "key": "nearmiss",
        "title": "【七】惊险瞬间 Near Miss（差点出事但没出事）  ★ 对应风险 6",
        "danger": False,
        "cols": ["时间", "差点发生什么", "为什么没成灾", "当时的判断",
                 "需要跟进", "责任人", "截止时间", "备注"],
        "rows": 3,
        "dvs": {},
        "cf": [],
    },
    {
        "key": "todos",
        "title": "【八】待跟进事项  ★ 必填 · 对应风险 7",
        "danger": False,
        "cols": ["事项", "责任人", "截止时间", "状态",
                 "来源（哪个事件 / 变更）", "优先级", "阻塞点", "备注"],
        "rows": 4,
        "dvs": {3: "task_status"},
        "cf": [
            {"col": "D", "value": "进行中", "fill": C_YELLOW},
        ],
    },
]

SELFCHECKS = [
    "已列出全部未闭环事件，并写明当前状态（未闭环 / 已恢复待观察 / 已闭环）",
    "已写明进行中的变更，以及它的时间窗和回滚方案",
    "已写明所有临时规避措施，并标注到期时间和失效后果",
    "已写明外部依赖（运营商 / 云服务商 / 第三方）的状态与下次确认时间",
    "已写明临界资源（磁盘 / 证书 / 容量）预警",
    "每个事件都区分了「现象」和「判断」，没有把猜测写成事实",
    "已写明「我还不确定什么」（未知项）—— 这一栏比已知项更重要",
    "每件待办都有具体责任人和截止时间（没有「待跟进」这种写法）",
    "「恢复」已明确到哪一层：告警消失 / 业务可用 / 根因消除",
    "接班人已当场提出疑问并得到答复，双方确认签名",
]

RISKS = [
    (1, "未闭环事件丢失", "★★★★★",
     "以为处理完了其实没有；事件被转走了没提；多个事件只传最严重的一个",
     "故障在交接后复发，被当成新问题重新排查",
     "【一】进行中的事件", "上一班走的时候，还有哪几件事没画上句号？"),
    (2, "进行中变更的时间窗", "★★★★★",
     "只说了「今晚有变更」，没说窗口和影响；回滚方案没传",
     "把变更引起的预期告警当故障查；在窗口内做冲突操作",
     "【三】进行中的变更", "现在有变更吗？几点结束？回滚怎么回？"),
    (3, "临时规避措施遗忘", "★★★★★",
     "临时改了配置/加了白名单/切了旁路，交接没说；没有到期时间",
     "接班人在「以为安全」的时候把临时方案拆了 → 二次故障",
     "【四】临时规避措施", "系统现在靠什么撑着？什么时候到期？"),
    (4, "外部依赖状态未传递", "★★★★☆",
     "运营商/云服务商有故障公告，上一班看过，交接没提",
     "在自己这侧排查一个根本不在自己这里的问题",
     "【五】外部依赖状态", "咱们依赖的外部服务，现在正常吗？"),
    (5, "临界资源预警", "★★★☆☆",
     "磁盘 85%、证书还剩 20 天、队列接近满，都没说",
     "在接班人的班上爆掉，但他完全没准备",
     "【六】临界资源预警", "有什么东西正在慢慢变坏，只是还没叫？"),
    (6, "惊险瞬间（Near Miss）未记录", "★★★★☆",
     "差点造成故障但及时回退；告警响一下又自己好了，没追",
     "同样的坑下一个人再踩一次；组织失去最便宜的学习机会",
     "【七】惊险瞬间", "这一班有没有什么差点就出事了？"),
    (7, "责任人与时限模糊", "★★★★☆",
     "「待跟进」「已通知」「等对方回复」「后续优化」—— 没名字没时间",
     "事项在交接链条里无限漂流，永远没人做",
     "【八】待跟进事项", "每件待办有没有一个具体的人名和一个具体的时间？"),
    (8, "状态口径不一致", "★★★★☆",
     "「恢复了」：告警消失？业务可用？根因消除？三个人三种理解",
     "双方都以为对齐了，实际理解不同；事后各说各话",
     "【一】当前状态的下拉框", "你说的「恢复了」，具体指哪一层？"),
]


# ---------------------------------------------------------------- 数据验证
def build_validations(wb):
    defs = {
        "shift": ("白班,夜班,中班,其他", "班次"),
        "severity": ("P1,P2,P3,P4", "严重级别"),
        "event_status": ("未闭环,已恢复待观察,已闭环", "事件状态"),
        "yesno": ("是,否", "是/否"),
        "task_status": ("未开始,进行中,已完成,已取消", "待办状态"),
        "dep_status": ("正常,异常,未知", "依赖状态"),
    }
    out = {}
    for key, (items, prompt) in defs.items():
        dv = DataValidation(
            type="list",
            formula1=f'"{items}"',
            allow_blank=True,
            showDropDown=False,
        )
        dv.error = "请从下拉列表中选择"
        dv.errorTitle = "取值不在允许范围内"
        dv.prompt = prompt
        dv.promptTitle = prompt
        out[key] = dv
    return out


# ---------------------------------------------------------------- 交接单（可与示例共用）
def build_form(ws, ncols, subtitle, data=None):
    """
    data: dict {section_key: [ [8 个值], ... ]}；为 None 时生成空白表单。
    """
    data = data or {}
    widths = [18, 18, 34, 13, 22, 26, 20, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    r = 1
    r = banner(ws, r, ncols, "值 班 交 接 单", C_DARK, height=34, size=16)

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    c = ws.cell(row=r, column=1, value=subtitle)
    c.font = Font(name="微软雅黑", size=9, italic=True, color="595959")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    c.fill = fill_of(C_LABEL)
    for cc in range(2, ncols + 1):
        ws.cell(row=r, column=cc).fill = fill_of(C_LABEL)
    ws.row_dimensions[r].height = 22
    r += 1

    # ---- 基础信息
    r = spacer(ws, r, ncols)
    r = banner(ws, r, ncols, "基础信息", C_SECTION, height=20, size=10)

    basics = [
        (("交接日期", ""), ("班次", ""), ("交班人", ""), ("接班人", "")),
        (("交接时间", ""), ("接班时间", ""), ("填表人", ""), ("本班告警量", "")),
    ]
    for row_def in basics:
        for idx, (label, value) in enumerate(row_def):
            label_col = idx * 2 + 1
            value_col = label_col + 1
            lc = ws.cell(row=r, column=label_col, value=label)
            lc.font = Font(name="微软雅黑", size=9, bold=True, color="1F3864")
            lc.fill = fill_of(C_LABEL)
            lc.alignment = Alignment(horizontal="center", vertical="center")
            lc.border = BORDER
            vc = ws.cell(row=r, column=value_col, value=value)
            vc.font = FONT_BASE
            vc.fill = fill_of(C_INPUT)
            vc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            vc.border = BORDER
        ws.row_dimensions[r].height = 24
        r += 1
    basic_shift_cell = "D6"

    # ---- 八个区块
    dv_ranges = {}
    cf_specs = []
    section_ranges = {}

    for sec in SECTIONS:
        r = spacer(ws, r, ncols)
        r = banner(ws, r, ncols, sec["title"],
                   C_DANGER if sec["danger"] else C_SECTION, height=22, size=10)
        r = table_header(ws, r, sec["cols"])
        first = r
        rows = data.get(sec["key"], [])
        for i in range(sec["rows"]):
            vals = rows[i] if i < len(rows) else [""] * ncols
            for ci in range(1, ncols + 1):
                cell = ws.cell(row=r, column=ci, value=vals[ci - 1] if ci - 1 < len(vals) else "")
                cell.fill = fill_of(C_INPUT)
                cell.border = BORDER
                cell.font = FONT_BASE
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            ws.row_dimensions[r].height = 30
            r += 1
        last = r - 1
        section_ranges[sec["key"]] = (first, last)

        for col_idx, dv_key in sec["dvs"].items():
            col_letter = chr(64 + col_idx)
            dv_ranges.setdefault(dv_key, []).append(f"{col_letter}{first}:{col_letter}{last}")

        for rule in sec["cf"]:
            cf_specs.append((rule, f"{rule['col']}{first}:{rule['col']}{last}"))

    # ---- 自检清单
    r = spacer(ws, r, ncols)
    r = banner(ws, r, ncols, "【九】交接自检（10 项）  ★ 接班人逐条确认，任一项「否」必须当场问清楚",
               C_SECTION, height=22, size=10)
    hdr = ws.cell(row=r, column=1, value="#")
    hdr.font = Font(name="微软雅黑", size=9, bold=True, color="1F3864")
    hdr.fill = fill_of(C_COLHDR)
    hdr.alignment = Alignment(horizontal="center", vertical="center")
    hdr.border = BORDER

    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    t = ws.cell(row=r, column=2, value="检查项")
    t.font = Font(name="微软雅黑", size=9, bold=True, color="1F3864")
    t.alignment = Alignment(horizontal="center", vertical="center")
    for cc in range(2, 8):
        ws.cell(row=r, column=cc).fill = fill_of(C_COLHDR)
        ws.cell(row=r, column=cc).border = BORDER

    y = ws.cell(row=r, column=8, value="是 / 否")
    y.font = Font(name="微软雅黑", size=9, bold=True, color="1F3864")
    y.fill = fill_of(C_COLHDR)
    y.alignment = Alignment(horizontal="center", vertical="center")
    y.border = BORDER
    ws.row_dimensions[r].height = 26
    r += 1

    check_first = r
    for i, text in enumerate(SELFCHECKS, start=1):
        nc = ws.cell(row=r, column=1, value=i)
        nc.font = Font(name="微软雅黑", size=9, bold=True)
        nc.fill = fill_of(C_LABEL)
        nc.alignment = Alignment(horizontal="center", vertical="center")
        nc.border = BORDER

        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        tc = ws.cell(row=r, column=2, value=text)
        tc.font = FONT_BASE
        tc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        for cc in range(2, 8):
            ws.cell(row=r, column=cc).border = BORDER
            ws.cell(row=r, column=cc).fill = fill_of(C_INPUT)

        yc = ws.cell(row=r, column=8, value="")
        yc.font = FONT_BASE
        yc.fill = fill_of(C_INPUT)
        yc.alignment = Alignment(horizontal="center", vertical="center")
        yc.border = BORDER
        ws.row_dimensions[r].height = 30
        r += 1
    check_last = r - 1
    dv_ranges.setdefault("yesno", []).append(f"H{check_first}:H{check_last}")
    cf_specs.append(({"col": "H", "value": "否", "fill": C_YELLOW}, f"H{check_first}:H{check_last}"))

    # ---- 双方确认
    r = spacer(ws, r, ncols)
    r = banner(ws, r, ncols, "【十】双方确认", C_DARK, height=20, size=10)

    confirm = [
        ("交班人确认", "我已经完整说明以上全部信息，包括未确定项。", "签字 / 时间"),
        ("接班人确认", "我已理解以上信息，并已就疑问当场提问。", "签字 / 时间"),
    ]
    for left, text, right in confirm:
        a = ws.cell(row=r, column=1, value=left)
        a.font = Font(name="微软雅黑", size=9, bold=True, color="1F3864")
        a.fill = fill_of(C_LABEL)
        a.alignment = Alignment(horizontal="center", vertical="center")
        a.border = BORDER

        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        t2 = ws.cell(row=r, column=2, value=text)
        t2.font = FONT_BASE
        t2.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
        for cc in range(2, 6):
            ws.cell(row=r, column=cc).border = BORDER
            ws.cell(row=r, column=cc).fill = fill_of(C_INPUT)

        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
        l2 = ws.cell(row=r, column=6, value=right)
        l2.font = Font(name="微软雅黑", size=9, bold=True, color="595959")
        l2.fill = fill_of(C_LABEL)
        l2.alignment = Alignment(horizontal="center", vertical="center")
        for cc in range(6, 8):
            ws.cell(row=r, column=cc).border = BORDER

        sig = ws.cell(row=r, column=8, value="")
        sig.font = FONT_BASE
        sig.fill = fill_of(C_INPUT)
        sig.border = BORDER
        ws.row_dimensions[r].height = 28
        r += 1

    r = spacer(ws, r, ncols)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    note = ws.cell(
        row=r, column=1,
        value="填写原则：现象是事实，判断是假设，未知项比已知项更重要。"
              "没有责任人和截止时间的待办，不是待办，是愿望。",
    )
    note.font = Font(name="微软雅黑", size=9, italic=True, color="9C0006")
    note.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r].height = 22
    r += 1

    return r, dv_ranges, cf_specs, section_ranges, basic_shift_cell


def apply_rules(ws, dv_pool, dv_ranges, cf_specs):
    for dv_key, ranges in dv_ranges.items():
        dv = dv_pool[dv_key]
        ws.add_data_validation(dv)
        for rng in ranges:
            dv.add(rng)

    for rule, rng in cf_specs:
        fill = fill_of(rule["fill"])
        if rule["col"] == "H" and rule["value"] == "否":
            ws.conditional_formatting.add(
                rng, CellIsRule(operator="equal", formula=['"否"'], fill=fill))
        elif rule.get("bold"):
            ws.conditional_formatting.add(
                rng,
                CellIsRule(
                    operator="equal",
                    formula=[f'"{rule["value"]}"'],
                    fill=fill,
                    font=Font(name="微软雅黑", size=9, bold=True, color=C_RED_TEXT),
                ),
            )
        else:
            ws.conditional_formatting.add(
                rng, CellIsRule(operator="equal", formula=[f'"{rule["value"]}"'], fill=fill))


# ---------------------------------------------------------------- 示例数据
EXAMPLE = {
    "events": [
        ["E-001", "02:14",
         "监控显示 5 个应用同时接口超时，连接数降为 0，告警开始大量重复",
         "P1", "未闭环",
         "已通知网络组与值班经理；未重启应用（多业务同时异常，判断为共因）",
         "网络组 张某 / 等待网络侧确认链路状态", "03:00 前必须跟进"],
        ["E-002", "09:30",
         "日志盘使用率 86%，近期每小时上涨约 1%",
         "P3", "未闭环",
         "已提交扩容工单，尚未执行",
         "运维组 李某 / 明日 10:00 前完成扩容", "今日 18:00 复检"],
    ],
    "observing": [
        ["E-003", "01:05", "03:12", "服务不可达告警",
         "连接数与错误率，连续观察 2 小时", "今日 12:00", "王某",
         "错误率再次抬头即升级为 P2"],
    ],
    "changes": [
        ["CHG-0001", "22:00", "23:30", "业务 A 数据库主从切换",
         "回切主从、恢复原配置，预计 10 分钟", "赵某", "138****0000", "是"],
    ],
    "workarounds": [
        ["链路抖动导致业务超时", "业务流量临时切到备用线路", "02:40", "今日 08:00",
         "备用线路容量不足，高峰期可能出现超时", "网络组 张某",
         "线路修复后切回并做验证", "08:00 交接时必须当面确认"],
    ],
    "external": [
        ["某运营商专线", "异常", "TK-0000-0912", "主备链路",
         "人井施工导致中断，预计 04:30 恢复", "网络组 张某", "03:20", "03:50"],
    ],
    "capacity": [
        ["日志盘 /data", "86%", "90%", "约 4 小时",
         "运维组 李某", "已提单扩容", "是", "扩容后需清理旧日志"],
    ],
    "nearmiss": [
        ["02:50", "差点直接重启数据库主库",
         "先查了连接数，发现是连接池耗尽而不是数据库卡死",
         "多业务同时告警 → 优先怀疑共因",
         "把这条判据写进交接单的检查项", "本人", "下次交接前", "已在 E-001 中记录"],
    ],
    "todos": [
        ["确认备用线路容量是否够用到明早", "网络组 张某", "今日 12:00", "进行中",
         "E-001", "高", "等待运营商回复", ""],
        ["清理 /data 下 30 天前日志", "运维组 李某", "今日 18:00", "未开始",
         "E-002", "中", "无", ""],
    ],
}


# ---------------------------------------------------------------- 各工作表
def sheet_usage(wb):
    ws = wb.active
    ws.title = "使用说明"
    ws.sheet_view.showGridLines = False
    widths = [4, 26, 90]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    r = 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    t = ws.cell(row=r, column=1, value="值班交接单 · 使用说明")
    t.font = Font(name="微软雅黑", size=18, bold=True, color="FFFFFF")
    t.fill = fill_of(C_DARK)
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r].height = 40
    r += 2

    blocks = [
        ("它解决什么问题", [
            "值班交接是故障处置链条里唯一完全依赖「人的记性和嘴」的环节，也是事故高发的时间点。",
            "这张表不是流程文件，是一份「上下文传递清单」：让接班的人在 60 秒内知道"
            "现在发生了什么、已经做了什么、还不确定什么。",
        ]),
        ("什么时候填", [
            "交班前 10 分钟开始填，交班时当面过一遍，接班人逐条确认。",
            "填一次约 2 分钟。填不完就先填三栏：未闭环事件、临时规避措施、未知项。",
        ]),
        ("谁填", [
            "交班人填，接班人核对并勾选自检清单。双方各自签字。",
            "两个人对同一行的理解不一致时，以「接班人复述一遍、交班人确认」为准。",
        ]),
        ("三条填写原则", [
            "1. 现象和判断分开写。现象是事实（连接数归零），判断是假设（链路断了）。混在一起，假设就变成了事实。",
            "2. 「恢复」必须说清是哪一层：告警消失 / 业务可用 / 根因消除。三者差别很大。",
            "3. 一定要写「我还不确定什么」。未知项比已知项更重要，因为它决定接班人往哪儿看。",
        ]),
        ("最容易被漏掉的两块", [
            "【四】临时规避措施 —— 系统现在靠什么在撑着？什么时候到期？到期会怎样？",
            "【五】外部依赖状态 —— 运营商 / 云服务商 / 第三方现在有没有异常或公告？",
            "这两块用深红色标注。反复出事的，几乎都是它们。",
        ]),
        ("不要做什么", [
            "不要在这里写「谁的责任」。一旦它变成追责工具，所有人都会开始防御性填写，信息质量立刻归零。",
            "不要用它考核别人。同理。",
            "不要在没人用的时候先写一份《交接管理制度》。制度是结果的固化，不是起点。",
        ]),
        ("工作表说明", [
            "「交接单」—— 空白表单，每班一份（另存为日期命名的文件，或直接打印）。",
            "「示例」—— 一份填好的示范，照着填即可。示例数据全部虚构。",
            "「交接记录」—— 一行一次交接，用来统计完整率和闭环率。",
            "「风险清单」—— 8 类交接风险，解释每个区块为什么存在。",
        ]),
    ]

    for title, lines in blocks:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        h = ws.cell(row=r, column=1, value=title)
        h.font = Font(name="微软雅黑", size=11, bold=True, color="1F3864")
        h.fill = fill_of(C_COLHDR)
        h.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[r].height = 24
        r += 1
        for line in lines:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            c = ws.cell(row=r, column=2, value=line)
            c.font = Font(name="微软雅黑", size=10)
            c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
            ws.row_dimensions[r].height = 32
            r += 1
        r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    foot = ws.cell(row=r, column=1,
                   value="本文件中的全部示例数据均为虚构的通用场景，不含任何真实主机名、IP、业务名称或公司信息。")
    foot.font = Font(name="微软雅黑", size=9, italic=True, color="595959")
    foot.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    return ws


def sheet_form(wb, dv_pool, *, title, subtitle, data):
    ws = wb.create_sheet(title)
    ncols = 8
    _, dv_ranges, cf_specs, _, _ = build_form(ws, ncols, subtitle, data)

    # 班次下拉（基础信息处单独加）
    shift_dv = DataValidation(type="list", formula1='"白班,夜班,中班,其他"',
                              allow_blank=True, showDropDown=False)
    ws.add_data_validation(shift_dv)
    shift_dv.add("D6:D6")

    apply_rules(ws, dv_pool, dv_ranges, cf_specs)

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    return ws


def sheet_log(wb, dv_pool):
    ws = wb.create_sheet("交接记录")
    ws.sheet_view.showGridLines = False
    headers = ["序号", "日期", "班次", "交班人", "接班人", "未闭环事件数",
               "进行中变更数", "规避措施数", "外部依赖异常数", "惊险瞬间数",
               "自检通过数", "自检总数", "自检通过率", "漏传事件数", "备注"]
    widths = [6, 12, 9, 10, 10, 12, 12, 11, 13, 11, 11, 9, 11, 11, 26]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    ncols = len(headers)
    r = 1
    r = banner(ws, r, ncols, "交接记录 · 统计与复盘用", C_DARK, height=28, size=14)

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    d = ws.cell(row=r, column=1,
                value="每次交接填一行。三个月后，这张表就是你说服组长「交接需要规范化」的全部证据。")
    d.font = Font(name="微软雅黑", size=9, italic=True, color="595959")
    d.fill = fill_of(C_LABEL)
    d.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for cc in range(1, ncols + 1):
        ws.cell(row=r, column=cc).fill = fill_of(C_LABEL)
    ws.row_dimensions[r].height = 22
    r += 2

    # 汇总区（表格在上方，数据行位置需要先算准）
    header_row = r + 2
    first_data = header_row + 1
    last_data = first_data + 39

    def rng(letter):
        return f"{letter}{first_data}:{letter}{last_data}"

    summary = [
        ("累计交接次数", f"=COUNTA({rng('B')})"),
        ("平均自检通过率", f"=IFERROR(AVERAGE({rng('M')}),0)"),
        ("漏传事件合计", f"=SUM({rng('N')})"),
        ("平均未闭环事件", f"=IFERROR(AVERAGE({rng('F')}),0)"),
    ]
    c = 1
    for label, formula in summary:
        lc = ws.cell(row=r, column=c, value=label)
        lc.font = Font(name="微软雅黑", size=9, bold=True, color="1F3864")
        lc.fill = fill_of(C_LABEL)
        lc.alignment = Alignment(horizontal="center", vertical="center")
        lc.border = BORDER
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 2)
        for cc in range(c, c + 3):
            ws.cell(row=r, column=cc).fill = fill_of(C_LABEL)
            ws.cell(row=r, column=cc).border = BORDER

        vc = ws.cell(row=r, column=c + 3, value=formula)
        vc.font = Font(name="微软雅黑", size=11, bold=True, color="9C0006")
        vc.fill = fill_of(C_INPUT)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.border = BORDER
        c += 4
        if c > ncols - 3:
            break
    ws.row_dimensions[r].height = 26
    r += 2

    # 表头
    hr = r
    for i, name in enumerate(headers, start=1):
        cell = ws.cell(row=r, column=i, value=name)
        cell.font = Font(name="微软雅黑", size=9, bold=True, color="1F3864")
        cell.fill = fill_of(C_COLHDR)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[r].height = 34
    r += 1

    for i in range(40):
        row = r + i
        ws.cell(row=row, column=1, value=i + 1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row, column=12, value=10).alignment = Alignment(horizontal="center", vertical="center")
        rate = ws.cell(row=row, column=13, value=f"=IFERROR(K{row}/L{row},\"\")")
        rate.number_format = "0%"
        rate.alignment = Alignment(horizontal="center", vertical="center")
        for cc in range(1, ncols + 1):
            cell = ws.cell(row=row, column=cc)
            cell.font = FONT_BASE
            cell.border = BORDER
            cell.fill = fill_of(C_INPUT)
            if cc not in (1, 12, 13):
                cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 20

    # 下拉
    shift_dv = DataValidation(type="list", formula1='"白班,夜班,中班,其他"',
                              allow_blank=True, showDropDown=False)
    ws.add_data_validation(shift_dv)
    shift_dv.add(f"C{first_data}:C{last_data}")

    ws.freeze_panes = f"A{hr + 1}"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    return ws


def sheet_risks(wb):
    ws = wb.create_sheet("风险清单")
    ws.sheet_view.showGridLines = False
    headers = ["#", "风险类别", "危险等级", "典型表现", "会造成的后果", "对应交接单字段", "一句话检查"]
    widths = [5, 22, 10, 44, 40, 22, 34]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    ncols = len(headers)
    r = 1
    r = banner(ws, r, ncols, "交接风险清单 · 8 类「不传就会出事」的信息", C_DARK, height=28, size=14)

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    d = ws.cell(row=r, column=1,
                value="每一行对应交接单里的一个区块。没传 ≠ 丢人，没传且没人问 = 事故。")
    d.font = Font(name="微软雅黑", size=9, italic=True, color="595959")
    d.fill = fill_of(C_LABEL)
    d.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for cc in range(1, ncols + 1):
        ws.cell(row=r, column=cc).fill = fill_of(C_LABEL)
    ws.row_dimensions[r].height = 22
    r += 2

    for i, name in enumerate(headers, start=1):
        cell = ws.cell(row=r, column=i, value=name)
        cell.font = Font(name="微软雅黑", size=9, bold=True, color="1F3864")
        cell.fill = fill_of(C_COLHDR)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[r].height = 30
    r += 1

    for row_data in RISKS:
        for i, value in enumerate(row_data, start=1):
            cell = ws.cell(row=r, column=i, value=value)
            cell.font = FONT_BASE
            cell.border = BORDER
            cell.fill = fill_of(C_INPUT)
            if i in (1, 3):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        # 高危险等级标红
        if row_data[2].startswith("★★★★★"):
            ws.cell(row=r, column=3).font = Font(name="微软雅黑", size=9, bold=True, color=C_RED_TEXT)
        ws.row_dimensions[r].height = 46
        r += 1

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    foot = ws.cell(row=r, column=1,
                   value="详细解释（含每类的典型场景和应对）见 docs/02-交接风险清单.md")
    foot.font = Font(name="微软雅黑", size=9, italic=True, color="595959")
    foot.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    return ws


# ---------------------------------------------------------------- main
def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    wb = Workbook()
    dv_pool = build_validations(wb)

    sheet_usage(wb)
    sheet_form(
        wb, dv_pool,
        title="交接单",
        subtitle="填写原则：现象是事实，判断是假设，未知项比已知项更重要。"
                 " ｜ 交班人填，接班人逐条确认 ｜ 填一次约 2 分钟",
        data=None,
    )
    sheet_form(
        wb, dv_pool,
        title="示例",
        subtitle="这是一份填好的示范（数据全部虚构）。照这个颗粒度填就够用，不用追求写满。",
        data=EXAMPLE,
    )
    sheet_log(wb, dv_pool)
    sheet_risks(wb)

    wb.active = 0  # 默认打开「使用说明」
    wb.save(OUT_FILE)
    print(f"已生成：{OUT_FILE}")
    print(f"工作表：{', '.join(wb.sheetnames)}")


if __name__ == "__main__":
    sys.exit(main())
