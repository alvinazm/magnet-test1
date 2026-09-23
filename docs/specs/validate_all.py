"""
Magnet 出题统一机审入口
========================

用途:一键运行所有机审检查,任何一项失败都返回非零退出码。

使用:
    .venv/bin/python docs/specs/validate_all.py

机审项:
    1. 数据自检(附件完整性、计算公式一致性)
    2. A 模块档位机审(17 条标准的字段完整、档位显式对应)
    3. 附件名/Sheet 名一致性(文档引用与实际一致)

退出码:
    0 = 所有机审通过
    1 = 至少一项机审失败
"""
import re
import sys
import subprocess
import openpyxl
from pathlib import Path

WORK = Path("/Users/azm/MyProject/work")
SPECS = WORK / "docs/specs"
INPUT = WORK / "input"
RUBRIC = SPECS / "05-evaluation-rubric.md"


def run_script(name: str) -> tuple:
    """以子进程方式运行机审脚本,返回 (返回码, 输出)"""
    script = SPECS / name
    if not script.exists():
        return 127, f"脚本不存在: {script}"
    result = subprocess.run(
        [sys.executable, str(script)],
        capture_output=True,
        text=True,
        cwd=str(WORK),
    )
    return result.returncode, result.stdout + result.stderr


def check_naming_consistency() -> tuple:
    """
    检查 A 模块文档中引用的附件名、Sheet 名是否与实际一致。
    返回 (返回码, 报告)。
    """
    if not RUBRIC.exists():
        return 1, f"✗ 评估表文档不存在: {RUBRIC}"

    content = RUBRIC.read_text()

    # 1. 收集实际附件和 Sheet
    actual_files = {f.name for f in INPUT.iterdir()
                    if f.suffix in (".xlsx", ".docx")}

    actual_sheets = {}  # sheet_name -> file_name
    for f in INPUT.iterdir():
        if f.suffix == ".xlsx":
            wb = openpyxl.load_workbook(f, read_only=True)
            for s in wb.sheetnames:
                actual_sheets[s] = f.name
            wb.close()

    # 2. 文档中的"《...》"引用
    file_refs = re.findall(r"《([^》]+)》", content)

    issues = []

    # 2a. 检查附件名
    for ref in file_refs:
        is_actual = ref in actual_files
        is_deliverable = any(kw in ref for kw in
                              ["建议", "报告", "策略", "清单", "运营"])
        if is_actual:
            continue  # 引用正确
        if is_deliverable:
            continue  # 交付物/简称,非严格
        # 既不是真实附件,也不是交付物
        issues.append(f"  ✗ 附件名引用异常: 《{ref}》(不在 input/ 中,不是交付物)")

    # 2b. 检查 Sheet 名
    sheet_refs = re.findall(r"([\u4e00-\u9fff][\u4e00-\u9fff]+)\s+Sheet", content)
    for ref in sheet_refs:
        # 过滤误识别
        if ref in ["个或多个", "交付物的", "表头与"]:
            continue
        if ref not in actual_sheets:
            issues.append(f"  ✗ Sheet 名不存在: '{ref} Sheet'(实际无此 Sheet)")

    # 2c. 检查禁止词
    for forbidden in ["M1", "M2", "M3"]:
        # 排除误识别(在元数据说明里)
        if forbidden in content:
            # 找到位置,看是否在"禁止词说明"上下文里
            for m in re.finditer(re.escape(forbidden), content):
                # 看前后 50 字符
                start = max(0, m.start() - 50)
                end = min(len(content), m.end() + 50)
                context = content[start:end]
                # 如果上下文有"出现"或"禁止"字样,是合规说明
                if any(kw in context for kw in ["出现", "禁止", "✓", "× 0"]):
                    continue
                issues.append(f"  ✗ 出现禁止词: '{forbidden}'(应在禁止项说明里)")

    if not issues:
        return 0, "  ✓ 所有附件名、Sheet 名与实际一致\n  ✓ 无禁止词"

    return 1, "\n".join(issues)


def main():
    print("=" * 75)
    print("Magnet 出题统一机审")
    print("=" * 75)

    results = []

    # ============ 1. 数据自检 ============
    print("\n" + "=" * 75)
    print("【1/3】数据自检(data-validation.py)")
    print("=" * 75)
    rc, output = run_script("data-validation.py")
    # 只输出末尾 20 行(避免过长)
    output_lines = output.strip().split("\n")
    summary_start = max(0, len(output_lines) - 15)
    print("\n".join(output_lines[summary_start:]))
    results.append(("数据自检", rc))

    # ============ 2. A 模块档位机审 ============
    print("\n" + "=" * 75)
    print("【2/3】A 模块档位机审(validate_a_module_levels.py)")
    print("=" * 75)
    rc, output = run_script("validate_a_module_levels.py")
    # 只输出汇总
    output_lines = output.strip().split("\n")
    for line in output_lines:
        if any(kw in line for kw in ["通过:", "失败:", "✅", "❌", "[汇总]"]):
            print(line)
    results.append(("A 模块档位机审", rc))

    # ============ 3. 附件名/Sheet 名一致性 ============
    print("\n" + "=" * 75)
    print("【3/3】附件名/Sheet 名一致性(本脚本内置)")
    print("=" * 75)
    rc, output = check_naming_consistency()
    print(output)
    results.append(("附件名/Sheet 名一致性", rc))

    # ============ 总汇总 ============
    print("\n" + "=" * 75)
    print("总汇总")
    print("=" * 75)
    all_pass = True
    for name, code in results:
        status = "✓ 通过" if code == 0 else "✗ 失败"
        print(f"  [{status}] {name}")
        if code != 0:
            all_pass = False
    print()

    if all_pass:
        print("  ✅ 所有机审通过!")
        sys.exit(0)
    else:
        print("  ❌ 至少一项机审失败,请修复后重跑")
        sys.exit(1)


if __name__ == "__main__":
    main()
