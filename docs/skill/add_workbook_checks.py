"""
交付物工作簿自查区(Check)与数据验证注入
========================================

用途:给已生成的 Excel 交付物补一个**只读的 Check 工作表**与枚举下拉验证:
  1. Check 表用公式从该工作簿自身的数据列复算关键派生值,并与表内呈现值比对(差异列 + 最大差异 + 违规计数);
  2. 枚举列加数据验证(下拉),取值超出枚举时 Excel 会提示;
  3. Check 区只做校验,任何结论列都不引用它的结果。

用法:
    .venv/bin/python docs/skill/add_workbook_checks.py            # 处理 output/ 下 4 个交付物
    .venv/bin/python docs/skill/add_workbook_checks.py <文件>...   # 处理指定文件

注意:公式写入后由 Excel/WPS 打开时计算;本脚本不缓存计算结果,因此不影响原有静态值。
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent.parent
OUTPUT = ROOT / "output"

HEAD_FILL = PatternFill("solid", fgColor="D9D9D9")
OK_FONT = Font(bold=True)

# 每个工作簿:Sheet 名 -> 校验定义
# rows: (标签, 公式, 期望说明);dv: (列标题, 枚举列表)
SPEC = {
    "门店分级与调整建议.xlsx": {
        "sheet": "门店清单",
        "n_rows": 42,
        "rows": [
            ("对象行数", "=COUNTA('门店清单'!A2:A43)", "应为 42(全量门店)"),
            ("坪效最大差异(元/㎡/月)",
             "=MAX(ABS('门店清单'!I2:I43-ROUND('门店清单'!G2:G43/6/'门店清单'!F2:F43,0)))",
             "复算 = 累计到店销售额 ÷ 6 ÷ 营业面积,取整数;应 ≤ 1"),
            ("分级与坪效/租售比阈值不一致家数",
             "=SUMPRODUCT(--('门店清单'!L2:L43<>IF('门店清单'!J2:J43>0.35,\"D\","
             "IF('门店清单'!I2:I43>=5000,\"A\",IF('门店清单'!I2:I43>=2000,\"B\","
             "IF('门店清单'!I2:I43>=1000,\"C\",\"D\")))))",
             "分级应 = 坪效定级,且租售比 >35% 直接 D 级;应为 0"),
            ("面积调整门店未标待补家数",
             "=SUMPRODUCT(--(('门店清单'!K2:K43=\"是\")*('门店清单'!O2:O43<>\"待补\")))",
             "面积调整门店的数据状态应为「待补」;应为 0"),
            ("枚举违规家数(分级/调整建议/数据状态/面积是否调整)",
             "=SUMPRODUCT(--(ISNA(MATCH('门店清单'!L2:L43,{\"A\";\"B\";\"C\";\"D\"},0))))"
             "+SUMPRODUCT(--(ISNA(MATCH('门店清单'!M2:M43,"
             "{\"维持\";\"整改\";\"转前置仓或自提点\";\"闭店评估\"},0))))"
             "+SUMPRODUCT(--(ISNA(MATCH('门店清单'!O2:O43,{\"已确认\";\"待补\"},0))))"
             "+SUMPRODUCT(--(ISNA(MATCH('门店清单'!K2:K43,{\"是\";\"否\"},0))))",
             "四个枚举列取值应落在规定集合内;应为 0"),
        ],
        "dv": [("K", '"是,否"'), ("L", '"A,B,C,D"'),
               ("M", '"维持,整改,转前置仓或自提点,闭店评估"'), ("O", '"已确认,待补"')],
    },
    "渠道资源再配置建议.xlsx": {
        "sheet": "渠道清单",
        "n_rows": 6,
        "rows": [
            ("对象行数", "=COUNTA('渠道清单'!A2:A7)", "应为 6(全量渠道)"),
            ("毛利率最大差异",
             "=MAX(ABS('渠道清单'!I2:I7-'渠道清单'!E2:E7/'渠道清单'!D2:D7))",
             "渠道毛利率 = 毛利 ÷ 净收入;应 ≤ 0.0005"),
            ("枚举违规行数(渠道类型/方向/数据状态)",
             "=SUMPRODUCT(--(ISNA(MATCH('渠道清单'!C2:C7,{\"自营\";\"平台\";\"即时零售\"},0))))"
             "+SUMPRODUCT(--(ISNA(MATCH('渠道清单'!L2:L7,"
             "{\"增投\";\"维持\";\"收缩\"},0))))"
             "+SUMPRODUCT(--(ISNA(MATCH('渠道清单'!P2:P7,{\"已确认\";\"待补\"},0))))",
             "枚举列取值应落在规定集合内;应为 0"),
            ("提示:销售费用率/渠道净利不可在本表内复算",
             "=1", "本表只含广告费/平台佣金/履约费 3 项费用,缺失其余费用项,故这两个指标须回附件复算"),
        ],
        "dv": [("C", '"自营,平台,即时零售"'), ("L", '"增投,维持,收缩"'), ("P", '"已确认,待补"')],
    },
    "促销组合优化建议.xlsx": {
        "sheet": "活动评估",
        "n_rows": 50,
        "rows": [
            ("活动行数", "=COUNTA('活动评估'!A2:A51)", "应为 50(全量活动)"),
            ("ROI 最大差异(按呈现精度比对)",
             "=MAX(ABS('活动评估'!H2:H51-ROUND('活动评估'!G2:G51/'活动评估'!E2:E51,2)))",
             "ROI = 增量毛利 ÷ 实际费用,按 2 位小数呈现后比对;应 ≤ 0.0005"),
            ("枚举违规行数(建议动作)",
             "=SUMPRODUCT(--(ISNA(MATCH('活动评估'!I2:I51,{\"保留\";\"调整\";\"取消\"},0))))",
             "建议动作取值应落在规定集合内;应为 0"),
            ("H2 节奏建议月份数", "=COUNTA('H2节奏建议'!A2:A7)", "应为 6(2026-07 ~ 2026-12)"),
        ],
        "dv": [("I", '"保留,调整,取消"')],
    },
    "异常与待核清单.xlsx": {
        "sheet": "异常清单",
        "n_rows": 19,
        "rows": [
            ("异常条数", "=COUNTA('异常清单'!A2:A100)", "材料自检范围内的异常条目数"),
            ("异常编号唯一", "=IF(COUNTA('异常清单'!A2:A100)=SUMPRODUCT(1/COUNTIF('异常清单'!A2:A100,'异常清单'!A2:A100)),\"唯一\",\"重复\")",
             "异常编号应唯一;显示「唯一」"),
            ("枚举违规条数(异常类别/严重度)",
             "=SUMPRODUCT(--(ISNA(MATCH('异常清单'!B2:B100,"
             "{\"数据冲突\";\"口径不清\";\"缺失\";\"越界\";\"时效性\"},0))))"
             "+SUMPRODUCT(--(ISNA(MATCH('异常清单'!G2:G100,{\"高\";\"中\";\"低\"},0))))",
             "异常类别与严重度取值应落在规定集合内;应为 0"),
        ],
        "dv": [("B", '"数据冲突,口径不清,缺失,越界,时效性"'), ("G", '"高,中,低"')],
    },
}


def build_check_sheet(wb, fname: str) -> None:
    spec = SPEC[fname]
    sheet_name = spec["sheet"]
    if sheet_name not in wb.sheetnames:
        print(f"  ! {fname} 缺少 Sheet「{sheet_name}」,跳过")
        return
    if "Check" in wb.sheetnames:
        del wb["Check"]
    ws = wb.create_sheet("Check")
    ws["A1"] = f"{sheet_name} 自查区(只读;公式由 Excel 计算)"
    ws["A1"].font = OK_FONT
    ws.append([])
    ws.append(["校验项", "结果", "期望/口径"])
    for c in ws[3]:
        c.fill = HEAD_FILL
    for label, formula, expect in spec["rows"]:
        ws.append([label, formula, expect])
    data_ws = ws.parent[sheet_name]
    # 幂等:先清掉该表已有的数据验证,避免重复注入
    data_ws.data_validations.dataValidation = []
    last = max(spec["n_rows"] + 1, 200)  # 预留扩展行,避免新增对象后验证范围不覆盖
    for col, formula_range in spec["dv"]:
        dv = DataValidation(type="list", formula1=formula_range, allow_blank=True, showDropDown=False)
        dv.errorTitle = "取值超出规定枚举"
        dv.error = "请使用规定枚举值"
        data_ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{last}")
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 70
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=3):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=(cell.column == 3))
    print(f"  ✓ {fname}: 已写入 Check 表({len(spec['rows'])} 项)与 {len(spec['dv'])} 列数据验证")


def main(paths: list[str]) -> int:
    files = [Path(p) for p in paths] if paths else sorted(OUTPUT.glob("*.xlsx"))
    checked = 0
    for f in files:
        if f.name not in SPEC:
            continue
        wb = openpyxl.load_workbook(f)
        build_check_sheet(wb, f.name)
        wb.save(f)
        checked += 1
    print(f"完成:{checked} 个工作簿已补自查区")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
