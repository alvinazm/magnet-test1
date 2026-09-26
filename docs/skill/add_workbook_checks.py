"""
交付物工作簿自查(默认只读校验)+ 可选检查工作表注入
====================================================

默认模式(verify):**不修改交付物**,按 Query 与评估表的口径复算关键派生值、枚举取值与对象行数,
打印结果并可写出自检报告;交付物因此保持"提交即用"(不新增辅助工作表)。

    .venv/bin/python docs/skill/add_workbook_checks.py                      # 校验 output/ 下交付物
    .venv/bin/python docs/skill/add_workbook_checks.py --report docs/reviews/交付物自检.md
    .venv/bin/python docs/skill/add_workbook_checks.py --mode validations   # 仅刷新枚举下拉(不新增工作表)
    .venv/bin/python docs/skill/add_workbook_checks.py --mode inject        # 注入 Check 工作表(工作版用)

口径来源:Query 六 的交付物字段/Sheet/行数要求、《渠道与门店管理策略.docx》第 2.1–2.3、4.6、7.1 条。
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent.parent
OUTPUT = ROOT / "output"
HEAD_FILL = PatternFill("solid", fgColor="D9D9D9")
OK_FONT = Font(bold=True)

REQUIRED_SHEETS = {
    "门店分级与调整建议.xlsx": ["门店清单"],
    "渠道资源再配置建议.xlsx": ["渠道清单"],
    "促销组合优化建议.xlsx": ["活动评估", "H2 节奏建议"],
    "异常与待核清单.xlsx": ["异常清单"],
}
DV_SPEC = {
    "门店分级与调整建议.xlsx": [("门店清单", "K", '"是,否"'), ("门店清单", "L", '"A,B,C,D"'),
                                 ("门店清单", "M", '"维持,整改,转前置仓或自提点,闭店评估"'),
                                 ("门店清单", "O", '"已确认,待补"')],
    "渠道资源再配置建议.xlsx": [("渠道清单", "C", '"自营,平台,即时零售"'),
                                 ("渠道清单", "L", '"增投,维持,收缩"'), ("渠道清单", "P", '"已确认,待补"')],
    "促销组合优化建议.xlsx": [("活动评估", "I", '"保留,调整,取消"')],
    "异常与待核清单.xlsx": [("异常清单", "B", '"数据冲突,口径不清,缺失,越界,时效性"'),
                            ("异常清单", "G", '"高,中,低"')],
}
ENUM_OK = {
    "门店分级与调整建议.xlsx": [("分级", {"A", "B", "C", "D"}),
                                 ("调整建议", {"维持", "整改", "转前置仓或自提点", "闭店评估"}),
                                 ("数据状态", {"已确认", "待补"}), ("面积是否调整", {"是", "否"})],
    "渠道资源再配置建议.xlsx": [("渠道类型", {"自营", "平台", "即时零售"}),
                                 ("H2资源调整方向", {"增投", "维持", "收缩"}), ("数据状态", {"已确认", "待补"})],
    "促销组合优化建议.xlsx": [("建议动作", {"保留", "调整", "取消"})],
    "异常与待核清单.xlsx": [("异常类别", {"数据冲突", "口径不清", "缺失", "越界", "时效性"}),
                            ("严重度", {"高", "中", "低"})],
}


def _rows(ws):
    hdr = [str(c.value) if c.value is not None else "" for c in ws[1]]
    return hdr, list(ws.iter_rows(min_row=2, values_only=True))


def verify_workbook(path: Path) -> list[tuple[str, str, bool, str]]:
    """返回 [(Sheet, 校验项, 是否通过, 实测值描述)]"""
    out: list[tuple[str, str, bool, str]] = []
    wb = openpyxl.load_workbook(path, data_only=True)
    name = path.name
    for sheet in REQUIRED_SHEETS.get(name, []):
        if sheet not in wb.sheetnames:
            out.append((sheet, "Sheet 存在", False, f"缺少 Sheet「{sheet}」"))
    if name == "门店分级与调整建议.xlsx":
        hdr, rows = _rows(wb["门店清单"])
        i = {n: k for k, n in enumerate(hdr)}
        bad_ping = max((abs(r[i["到店坪效(元/㎡/月)"]] - round(r[i["2026H1累计到店销售额"]] / 6 / r[i["营业面积"]], 0))
                        for r in rows), default=0)
        bad_grade = sum(1 for r in rows if r[i["分级"]] != (
            "D" if r[i["租售比"]] > 0.35 else "A" if r[i["到店坪效(元/㎡/月)"]] >= 5000 else
            "B" if r[i["到店坪效(元/㎡/月)"]] >= 2000 else "C" if r[i["到店坪效(元/㎡/月)"]] >= 1000 else "D"))
        bad_adj = sum(1 for r in rows if r[i["面积是否调整"]] == "是" and r[i["数据状态"]] != "待补")
        out += [("门店清单", "对象行数 = 42", len(rows) == 42, f"{len(rows)} 行"),
                ("门店清单", "坪效复算差异 ≤ 1", bad_ping <= 1, f"最大差异 {bad_ping:.4f}"),
                ("门店清单", "分级与坪效/租售比阈值一致", bad_grade == 0, f"不一致 {bad_grade} 家"),
                ("门店清单", "面积调整门店标「待补」", bad_adj == 0, f"未标 {bad_adj} 家")]
    if name == "渠道资源再配置建议.xlsx":
        hdr, rows = _rows(wb["渠道清单"])
        i = {n: k for k, n in enumerate(hdr)}
        bad_gm = max((abs(r[i["渠道毛利率"]] - r[i["毛利(元)"]] / r[i["2026H1净收入(元)"]]) for r in rows), default=0)
        out += [("渠道清单", "对象行数 = 6", len(rows) == 6, f"{len(rows)} 行"),
                ("渠道清单", "毛利率 = 毛利 ÷ 净收入(≤0.0005)", bad_gm <= 0.0005, f"最大差异 {bad_gm:.6f}")]
    if name == "促销组合优化建议.xlsx":
        hdr, rows = _rows(wb["活动评估"])
        i = {n: k for k, n in enumerate(hdr)}
        bad_roi = max((abs(r[i["ROI"]] - round(r[i["增量毛利(元)"]] / r[i["实际费用(元)"]], 2)) for r in rows), default=0)
        h2 = wb["H2 节奏建议"] if "H2 节奏建议" in wb.sheetnames else None
        out += [("活动评估", "活动行数 = 50", len(rows) == 50, f"{len(rows)} 行"),
                ("活动评估", "ROI = 增量毛利 ÷ 实际费用(按 2 位呈现,≤0.0005)", bad_roi <= 0.0005, f"最大差异 {bad_roi:.6f}"),
                ("H2 节奏建议", "月份行数 = 6", (h2 is not None and h2.max_row - 1 == 6),
                 f"{h2.max_row - 1 if h2 is not None else 0} 行")]
    if name == "异常与待核清单.xlsx":
        hdr, rows = _rows(wb["异常清单"])
        i = {n: k for k, n in enumerate(hdr)}
        ids = [r[i["异常编号"]] for r in rows]
        out += [("异常清单", "条目数 ≥ 4 且编号唯一", len(rows) >= 4 and len(ids) == len(set(ids)),
                 f"{len(rows)} 条,唯一={len(ids) == len(set(ids))}")]
    for field, allowed in ENUM_OK.get(name, []):
        sheet = REQUIRED_SHEETS[name][0]
        hdr, rows = _rows(wb[sheet])
        k = hdr.index(field)
        bad = sum(1 for r in rows if r[k] not in allowed)
        out.append((sheet, f"枚举合规:{field}", bad == 0, f"违规 {bad} 行"))
    wb.close()
    return out


def apply_validations() -> None:
    for fname, specs in DV_SPEC.items():
        path = OUTPUT / fname
        if not path.exists():
            continue
        wb = openpyxl.load_workbook(path)
        for sheet, col, formula in specs:
            ws = wb[sheet]
            ws.data_validations.dataValidation = []
            dv = DataValidation(type="list", formula1=formula, allow_blank=True, showDropDown=False)
            dv.errorTitle = "取值超出规定枚举"
            dv.error = "请使用规定枚举值"
            ws.add_data_validation(dv)
            dv.add(f"{col}2:{col}200")
        wb.save(path)
    print("已刷新枚举数据验证(未新增工作表)")


def inject_check_sheets() -> None:
    """可选:注入 Check 工作表(仅用于工作版,提交版请勿使用)。"""
    for fname in REQUIRED_SHEETS:
        path = OUTPUT / fname
        if not path.exists():
            continue
        wb = openpyxl.load_workbook(path)
        if "Check" in wb.sheetnames:
            del wb["Check"]
        ws = wb.create_sheet("Check")
        ws["A1"] = "交付物自查区(只读)"
        ws["A1"].font = OK_FONT
        ws.append([])
        ws.append(["Sheet", "校验项", "结果"])
        for c in ws[3]:
            c.fill = HEAD_FILL
        for sheet, item, ok, actual in verify_workbook(path):
            ws.append([sheet, item, f"{'✓' if ok else '✗'} {actual}"])
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 42
        ws.column_dimensions["C"].width = 40
        for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=3):
            for cell in row:
                cell.alignment = Alignment(vertical="center")
        wb.save(path)
    print("已注入 Check 工作表(工作版)")


def main(argv: list[str]) -> int:
    mode = "verify"
    report_path = None
    if "--mode" in argv:
        mode = argv[argv.index("--mode") + 1]
    if "--report" in argv:
        report_path = Path(argv[argv.index("--report") + 1])

    if mode == "inject":
        inject_check_sheets()
        return 0

    files = sorted(OUTPUT.glob("*.xlsx"))
    total = passed = 0
    lines = [f"# 交付物自检报告\n", f"> 生成时间:{datetime.now():%Y-%m-%d %H:%M};对象:`output/` 下的 Excel 交付物;"
             "口径来源:Query 六 与《渠道与门店管理策略.docx》。\n"]
    print("=" * 72)
    print(f"交付物自检({mode})")
    print("=" * 72)
    for f in files:
        if f.name not in REQUIRED_SHEETS:
            continue
        print(f"\n{f.name}")
        lines.append(f"\n## {f.name}\n\n| Sheet | 校验项 | 结果 | 实测 |\n| --- | --- | --- | --- |")
        for sheet, item, ok, actual in verify_workbook(f):
            total += 1
            passed += 1 if ok else 0
            print(f"  {'✓' if ok else '✗'} {item}: {actual}")
            lines.append(f"| {sheet} | {item} | {'✓' if ok else '✗'} | {actual} |")
    print(f"\n合计:{passed}/{total} 项通过")
    lines.append(f"\n**合计:{passed}/{total} 项通过**\n")
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"已写出自检报告:{report_path}")
    if mode == "validations":
        apply_validations()
    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
