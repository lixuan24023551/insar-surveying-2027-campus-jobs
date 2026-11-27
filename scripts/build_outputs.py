#!/usr/bin/env python3
"""Validate recruiting CSV files and rebuild README.md plus a formatted Excel workbook."""
from __future__ import annotations

import csv
import re
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data" / "jobs.csv"
WATCH = ROOT / "data" / "watchlist_and_excluded.csv"
AS_OF = date(2026, 11, 27)
FIELDS = [
    "id", "unit", "unit_type", "recruitment_type", "target", "education",
    "match_tags", "positions", "work_locations", "application", "deadline",
    "deadline_type", "status", "announcement_date", "source_url",
    "source_level", "last_verified", "notes",
]
HEADERS = [
    "序号", "单位名称", "单位类型", "招聘类型", "招聘对象", "招聘学历",
    "匹配方向/专业代码", "相关岗位", "岗位城市/地点", "网申链接/投递方式",
    "招聘截止时间", "截止口径", "当前状态", "公告/发布日期", "信息来源",
    "来源级别", "最近核验", "备注",
]
WIDTHS = [7, 34, 16, 26, 28, 26, 28, 44, 32, 48, 28, 25, 25, 17, 50, 25, 15, 55]
KEY_ROWS = [43, 5, 4, 1, 2, 3, 47, 48, 46, 25, 36, 11, 13, 14, 27]
KEY_LABELS = {43: "公告", 5: "公告", 4: "公告", 1: "官网", 2: "公告", 3: "职位页", 47: "公告", 48: "简章", 46: "官网", 25: "简章", 36: "公告", 11: "官网", 13: "官网", 14: "网申", 27: "公告"}


def load(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        if reader.fieldnames is None:
            raise ValueError(f"{path}: 缺少表头")
        if any(None in row for row in rows):
            raise ValueError(f"{path}: 存在列数错误")
        return rows


def validate(rows: list[dict[str, str]]) -> None:
    if not rows or list(rows[0]) != FIELDS:
        raise ValueError("主表为空或字段顺序不符合18字段约定")
    if len({r["id"] for r in rows}) != len(rows) or len({r["unit"] for r in rows}) != len(rows):
        raise ValueError("序号或单位名称重复")
    for row in rows:
        for key in ("unit", "match_tags", "positions", "deadline", "status", "source_url", "last_verified"):
            if not row[key].strip():
                raise ValueError(f"{row['unit']}: {key}为空")
        if not row["source_url"].startswith(("http://", "https://")):
            raise ValueError(f"{row['unit']}: 来源URL无效")
        date.fromisoformat(row["last_verified"])


def deadline_key(row: dict[str, str]) -> tuple[int, str, str]:
    match = re.match(r"(20\d{2}-\d{2}-\d{2})", row["deadline"])
    if match:
        return 0, match.group(1), row["unit"]
    return (1 if "招满即止" in row["deadline"] else 2), "9999-12-31", row["unit"]


def md(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", "<br>")


def build_readme(rows: list[dict[str, str]], watch: list[dict[str, str]]) -> None:
    by_id = {int(row["id"]): row for row in rows}
    insar = sum("InSAR" in r["match_tags"] or "InSAR" in r["positions"] for r in rows)
    code = sum("085704" in r["match_tags"] for r in rows)
    rolling = sum("招满即止" in r["deadline"] for r in rows)
    unknown = sum("未注明" in r["deadline"] for r in rows)
    lines = [
        "# 2027届 InSAR / 测绘工程 / 资源与环境秋招岗位表", "",
        f"> 最近核验：{AS_OF.isoformat()}（北京时间）｜主表 **{len(rows)} 家单位**｜每天 12:00 核查并更新", "",
        "面向 **InSAR、测绘工程（按专业代码 085704 整理）、资源与环境（0857，限测绘/遥感/地理信息相关方向）** 的全国公开招聘信息。", "",
        "## 数据概览", "",
        f"- 可投递或官方岗位页仍可查：**{len(rows)} 家**（不代表所有名额均未招满）。",
        f"- 明确出现 InSAR：**{insar} 家**；明确标注 085704：**{code} 家**。",
        f"- 招满即止：**{rolling} 家**；未注明截止：**{unknown} 家**。",
        f"- 已截止、待正式公告或非校招备选：**{len(watch)} 条**，见 [`data/watchlist_and_excluded.csv`](data/watchlist_and_excluded.csv)。",
        "- 机器可读完整数据：[`data/jobs.csv`](data/jobs.csv)；运行生成脚本可得到可筛选工作簿 `2027_campus_jobs.xlsx`。", "",
        "## 在招岗位（按明确截止日期优先排序）", "",
        "完整表格以 [`data/jobs.csv`](data/jobs.csv) 为准，包含单位类型、招聘对象、学历、专业代码、岗位、城市、网申入口、截止口径、公告来源和备注等 18 个字段。", "",
        "### 最近截止与重点匹配", "",
        "| 单位 | 匹配方向 | 相关岗位 | 截止时间 | 投递/来源 |",
        "|---|---|---|---|---|",
    ]
    for record_id in KEY_ROWS:
        row = by_id[record_id]
        target = row["source_url"]
        lines.append(f"| {md(row['unit'])} | {md(row['match_tags'])} | {md(row['positions'])} | {md(row['deadline'])} | [{KEY_LABELS[record_id]}]({target}) |")
    lines += [
        "", "## 截止时间与状态口径", "",
        "- **企业/单位明确截止**：按公告原文记录。",
        "- **高校就业网公告截止**：公告由高校就业网发布或转载，仍建议在企业网申页再次确认。",
        "- **页面/职位有效期**：只是高校页面或职位有效期，不等于企业承诺的统一截止日。",
        "- **第三方集团批次口径**：只作提醒，官方岗位页未单列日期时标注“待官方确认”。",
        "- **招满即止 / 未注明**：没有可核验的固定日期，不应理解为长期有效。",
        "- 主表只保留公告、网申页或近期职位记录仍可查的单位；发现明确过期后移入排除表。", "",
        "## 收录边界", "",
        "1. 首版以 2027 届秋招为主，兼顾当前仍可投的相关校园岗位；严格限定测绘、遥感、GIS、InSAR、导航或资源与环境相关方向。",
        "2. “全国所有单位”无法由公开网络一次性证明穷尽。本仓库是**公开渠道可核验合集**，通过每日更新持续补漏，不作绝对穷尽承诺。",
        "3. 用户原始表述中的“测绘工程 087504”暂按常见专业代码 **085704** 校正；若原意确为其他代码，可再调整。",
        "4. 同一集团的独立招聘单位在岗位和投递入口不同的情况下单列。",
        "5. 任何投递前，请点击来源与网申入口核对岗位是否仍开放、专业目录、学历、工作地和材料要求。", "",
        "## 每日更新规则", "",
        "- 每天 12:00（Asia/Shanghai）检查新增公告、截止状态、网申入口及已有记录变化。",
        "- 新增记录必须有来源；明确过期的记录移出主表，并保留排除原因。",
        "- 更新 `data/jobs.csv` 后运行 `python3 scripts/build_outputs.py`，重新生成 README 与 Excel。",
        "- GitHub Actions 在推送及每天 04:00 UTC（北京时间 12:00）执行数据校验和产物重建检查。", "",
        "## 本地重建", "", "```bash", "python3 -m pip install -r requirements.txt",
        "python3 scripts/build_outputs.py", "```", "", "## 免责声明", "",
        "招聘状态和截止时间可能随时调整。本仓库仅做信息整理，不代替招聘单位正式公告，不构成录用承诺。", "",
    ]
    (ROOT / "README.md").write_text("\n".join(lines), encoding="utf-8")


def style_sheet(ws, widths: list[int]) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    ws.sheet_view.showGridLines = False
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin", color="B8C7D9"))
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="微软雅黑", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[row[0].row].height = 54


def build_excel(rows: list[dict[str, str]], watch: list[dict[str, str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "在招岗位"
    ws.append(HEADERS)
    ordered = sorted(rows, key=deadline_key)
    for row in ordered:
        ws.append([row[key] for key in FIELDS])
    style_sheet(ws, WIDTHS)
    for col in (10, 15):
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(r, col)
            if str(cell.value).startswith(("http://", "https://")):
                cell.hyperlink = cell.value
                cell.font = Font(name="Arial", size=9, color="0563C1", underline="single")
    for r, row in enumerate(ordered, 2):
        match = re.match(r"(20\d{2}-\d{2}-\d{2})", row["deadline"])
        if match and 0 <= (date.fromisoformat(match.group(1)) - AS_OF).days <= 7:
            ws.cell(r, 11).fill = PatternFill("solid", fgColor="FCE8E6")
        elif "招满即止" in row["deadline"]:
            ws.cell(r, 11).fill = PatternFill("solid", fgColor="E8F3E8")

    info = wb.create_sheet("使用说明")
    info_rows = [
        ("项目", "说明"), ("最近核验", f"{AS_OF.isoformat()}（北京时间）"),
        ("记录数", f"{len(rows)}家在招/岗位页可查；{len(watch)}条待关注或排除"),
        ("专业代码", "用户原文087504暂按085704测绘工程整理"),
        ("收录边界", "0857仅收录测绘、遥感、GIS、InSAR相关方向，不扩展到所有环境类岗位"),
        ("颜色", "红色=7天内截止；绿色=招满即止"),
        ("完整性", "公开渠道可核验合集，不作绝对穷尽承诺；每天12点持续更新"),
        ("提示", "未注明截止不等于长期有效；投递前请复核官方公告与网申状态"),
    ]
    for row in info_rows:
        info.append(row)
    style_sheet(info, [24, 110])

    watch_ws = wb.create_sheet("待关注与已排除")
    watch_fields = ["unit", "related_position", "category", "deadline", "reason", "source_url", "last_verified"]
    watch_ws.append(["单位", "相关岗位", "分类", "截止/状态", "说明", "来源", "最近核验"])
    for row in watch:
        watch_ws.append([row[key] for key in watch_fields])
    style_sheet(watch_ws, [34, 38, 20, 24, 65, 55, 15])
    for r in range(2, watch_ws.max_row + 1):
        cell = watch_ws.cell(r, 6)
        if str(cell.value).startswith(("http://", "https://")):
            cell.hyperlink = cell.value
            cell.font = Font(name="Arial", size=9, color="0563C1", underline="single")

    output = ROOT / "2027_campus_jobs.xlsx"
    wb.save(output)
    check = load_workbook(output, read_only=True)
    assert check.sheetnames == ["在招岗位", "使用说明", "待关注与已排除"]
    assert check["在招岗位"].max_row == len(rows) + 1
    check.close()


def main() -> None:
    rows, watch = load(DATA), load(WATCH)
    validate(rows)
    build_readme(rows, watch)
    build_excel(rows, watch)
    print(f"OK: {len(rows)} active records; {len(watch)} watch/excluded records")


if __name__ == "__main__":
    main()
