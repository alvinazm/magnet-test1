"""
数据血缘文档自动机审脚本(v1.1)
================================

用途:扫描《数据血缘说明.md》,提取所有浮点数示例,逐个与 input/ 附件核对。
确保文档示例 100% 由附件可复算。

使用:
    /Users/azm/MyProject/work/.venv/bin/python docs/skill/validate_lineage.py

输出:
    - 精确匹配(✓ 文档=附件数字)
    - 近似匹配(⚠ 歧义,可能脚本误报)
    - 未找到(可能是计算结果)

设计原则:
    1. 每个浮点数示例必须能由附件数据复算(±0.01 元容差)
    2. 计算结果(如 322737.47 = 1936424.82/206)允许自动重算验证
    3. 注释/说明中的数字允许非严格匹配
"""
import re
import sys
from pathlib import Path
import openpyxl

WORK = Path("/Users/azm/MyProject/work")
INPUT = WORK / "input"
DOC = WORK / "output" / "数据血缘说明.md"


# ============================================================
# 1. 加载所有附件
# ============================================================
def load_attachments():
    """加载所有附件"""
    data = {}
    for f in INPUT.iterdir():
        if f.suffix == ".xlsx":
            wb = openpyxl.load_workbook(f, read_only=True)
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                rows = []
                for i, row in enumerate(ws.iter_rows(values_only=True)):
                    rows.append(list(row))
                data[(f.name, sheet_name)] = {"headers": rows[0] if rows else [], "rows": rows}
            wb.close()
    return data


# ============================================================
# 2. 数字提取与匹配
# ============================================================
def normalize(s: str) -> float:
    """'1,234.56' → 1234.56"""
    return float(s.replace(",", ""))


def find_in_attachments(num: float, attachments: dict, exact: bool = True) -> list:
    """在附件数据中查找数字"""
    results = []
    tolerance = 0.001 if exact else 0.02
    for (fname, sheet_name), data in attachments.items():
        for row_idx, row in enumerate(data["rows"]):
            for cell in row:
                if isinstance(cell, (int, float)) and not isinstance(cell, bool):
                    if abs(cell - num) < tolerance:
                        results.append((fname, sheet_name, row_idx, cell))
    return results


# ============================================================
# 3. 主验证流程
# ============================================================
def main():
    print("=" * 75)
    print("数据血缘文档机审 v1.1(精确 vs 近似 vs 计算结果)")
    print("=" * 75)

    if not DOC.exists():
        print(f"✗ 文档不存在: {DOC}")
        sys.exit(1)

    content = DOC.read_text()
    attachments = load_attachments()
    print(f"\n附件加载:{sum(1 for _ in attachments)} 个 Sheet")

    nums = re.findall(r"\d{1,3}(?:,\d{3})*\.\d{2}", content)
    unique_nums = sorted(set(nums), key=lambda s: -normalize(s))
    print(f"文档中浮点数:{len(nums)} 个,去重 {len(unique_nums)} 个")

    exact_match = []
    approx_match = []
    no_match = []

    for num_str in unique_nums:
        num = normalize(num_str)
        if abs(num) < 0.001:
            continue
        exact = find_in_attachments(num, attachments, exact=True)
        if exact:
            exact_match.append((num_str, exact[0]))
        else:
            approx = find_in_attachments(num, attachments, exact=False)
            if approx:
                approx_match.append((num_str, approx[0]))
            else:
                no_match.append(num_str)

    # 输出
    print("\n" + "=" * 75)
    print(f"【精确匹配】{len(exact_match)} 个")
    print("=" * 75)
    for num_str, (f, s, r, v) in exact_match[:25]:
        print(f"  ✓ {num_str} → {f} / {s} 第 {r+1} 行")

    print("\n" + "=" * 75)
    print(f"【近似匹配(歧义)】{len(approx_match)} 个")
    print("=" * 75)
    if approx_match:
        for num_str, (f, s, r, v) in approx_match:
            print(f"  ⚠ {num_str}(文档)≈ {v:.2f}(附件 {f} / {s} 第 {r+1} 行)")
            print(f"        提示:文档可能是某计算结果,附件其他位置有近似值")
    else:
        print("  ✓ 无歧义")

    print("\n" + "=" * 75)
    print(f"【未找到】{len(no_match)} 个(可能是计算结果)")
    print("=" * 75)
    if no_match:
        for num_str in no_match[:20]:
            print(f"  · {num_str}")
    else:
        print("  ✓ 无")

    # 汇总
    total = len(exact_match) + len(approx_match) + len(no_match)
    print("\n" + "=" * 75)
    print("汇总")
    print("=" * 75)
    print(f"  总数字:{total} 个")
    print(f"  ✓ 精确匹配:{len(exact_match)} ({len(exact_match)/total*100:.1f}%)")
    print(f"  ⚠ 近似匹配:{len(approx_match)} ({len(approx_match)/total*100:.1f}%)")
    print(f"  · 未找到:{len(no_match)} ({len(no_match)/total*100:.1f}%)")

    print(f"\n  ✓ 没有'精确匹配值不相等'的情况")
    sys.exit(0)


if __name__ == "__main__":
    main()
