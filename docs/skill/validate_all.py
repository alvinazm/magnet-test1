"""
Magnet 出题统一机审入口
========================

用途:一键运行所有机审检查,任何一项失败都返回非零退出码。

使用:
    .venv/bin/python docs/skill/validate_all.py

目录约定(见 docs/skill/README.md):
    docs/skill/   出题工具箱:培训文档、工作流 SOP、写作规则、机审与自检脚本(可跨题复用)
    docs/specs/   本题产物:任务说明、Query、评估表
    docs/reviews/ 本题复盘:材料问题清单、M1/M2/M3 结果评估

机审项:
    1. 数据自检(附件完整性、附件包合规、计算公式一致性、题目可解性断言、跨口径反证探针 L5)
    2. A 模块档位机审(字段完整、档位显式对应)
    3. 附件名/Sheet 名一致性(文档引用与实际一致)
    4. 引用完整性(所有《》引用必须写完整文件名 + 扩展名,见项目根目录 AGENTS.md)
    5. 口径登记一致性(validate_calibers.py):每个口径在制度 / Query / 评估表三处是否都有锚点
    6. 材料基线比对(报告性,不参与通过判定):列出材料相对基线的变化,确认"只改了预期的部分"

说明:数据自检中的 L5 风险登记默认不阻断(真实材料允许存在冲突/缺失/异常,但须可披露);
     如需把风险登记一并视为失败,运行数据自检时设置环境变量 MAGNET_STRICT_CHECKS=1。

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
SKILL = WORK / "docs/skill"
INPUT = WORK / "input"
RUBRIC = SPECS / "step3-evaluation-rubric.md"


def run_script(name: str) -> tuple:
    """以子进程方式运行机审脚本,返回 (返回码, 输出)"""
    script = SKILL / name
    if not script.exists():                       # 兼容脚本仍放在 docs/specs/ 的情况
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

    # 模块 B / C 按培训文档 6.5、6.7 节要求必须写明模型编号(三家情况、红线涉及产物);
    # 禁止词只在模块 A(以及前置说明)范围内检查。
    module_b_start = content.find("## 模块 B")
    a_scope = content if module_b_start < 0 else content[:module_b_start]

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
        if forbidden in a_scope:
            # 找到位置,看是否在"禁止词说明"上下文里
            for m in re.finditer(re.escape(forbidden), a_scope):
                # 看前后 50 字符
                start = max(0, m.start() - 50)
                end = min(len(a_scope), m.end() + 50)
                context = a_scope[start:end]
                # 如果上下文有"出现"或"禁止"字样,是合规说明
                if any(kw in context for kw in ["出现", "禁止", "✓", "× 0"]):
                    continue
                issues.append(f"  ✗ 出现禁止词: '{forbidden}'(应在禁止项说明里)")

    if not issues:
        return 0, "  ✓ 所有附件名、Sheet 名与实际一致\n  ✓ 无禁止词"

    return 1, "\n".join(issues)


def check_reference_completeness() -> tuple:
    """
    项目约束检查:所有《...》引用必须写完整文件名 + 扩展名
    (不得用简称、别名,也不得把 Sheet 名当文件名)。

    跳过:反例行(含 ❌ / ✗)、逐字引用(含"引文照录")、脚本中的正则示例。
    """
    targets = [
        SPECS / "step1-task-brief.md",
        SPECS / "step2-query.md",
        SPECS / "step3-evaluation-rubric.md",
        SKILL / "step3-writing-guide.md",
        SKILL / "出题工作流.md",
        WORK / "output/数据血缘说明.md",
    ]
    # 注意:docs/skill/rules1.md 与 rules2.md 是平台培训文档原文,按"逐字引用"处理,不纳入本检查
    reviews_dir = WORK / "docs/reviews"
    if reviews_dir.exists():
        targets += [p for p in reviews_dir.glob("*.md") if p not in targets]

    allowed_ext = (".xlsx", ".docx", ".md", ".py", ".csv", ".json")
    skip_markers = ("❌", "✗", "引文照录", "re.findall", "re.sub", "{ref}")
    issues = []
    scanned = 0

    for path in targets:
        if not path.exists():
            continue
        scanned += 1
        for lineno, line in enumerate(path.read_text().splitlines(), 1):
            if any(m in line for m in skip_markers):
                continue
            for ref in re.findall(r"《([^》]+)》", line):
                if not ref.endswith(allowed_ext):
                    issues.append(
                        f"  ✗ {path.relative_to(WORK)} 第 {lineno} 行:"
                        f"《{ref}》缺少完整文件名或扩展名"
                    )

    if not issues:
        return 0, f"  ✓ 已扫描 {scanned} 份文档:所有《》引用均为完整文件名(含扩展名)"
    return 1, "\n".join(issues[:20])


def main():
    print("=" * 75)
    print("Magnet 出题统一机审")
    print("=" * 75)

    results = []

    # ============ 1. 数据自检 ============
    print("\n" + "=" * 75)
    print("【1/6】数据自检(data-validation.py)")
    print("=" * 75)
    rc, output = run_script("data-validation.py")
    # 只输出末尾 20 行(避免过长)
    output_lines = output.strip().split("\n")
    summary_start = max(0, len(output_lines) - 15)
    print("\n".join(output_lines[summary_start:]))
    results.append(("数据自检", rc))

    # ============ 2. A 模块档位机审 ============
    print("\n" + "=" * 75)
    print("【2/6】A 模块档位机审(validate_a_module_levels.py)")
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
    print("【3/6】附件名/Sheet 名一致性(本脚本内置)")
    print("=" * 75)
    rc, output = check_naming_consistency()
    print(output)
    results.append(("附件名/Sheet 名一致性", rc))

    # ============ 4. 引用完整性(项目约束) ============
    print("\n" + "=" * 75)
    print("【4/6】引用完整性检查(《》必须为完整文件名,见 AGENTS.md)")
    print("=" * 75)
    rc, output = check_reference_completeness()
    print(output)
    results.append(("引用完整性", rc))

    # ============ 5. 口径登记一致性 ============
    print("\n" + "=" * 75)
    print("【5/6】口径登记一致性(validate_calibers.py)")
    print("=" * 75)
    rc, output = run_script("validate_calibers.py")
    print(output.strip())
    results.append(("口径登记一致性", rc))

    # ============ 6. 材料基线比对(报告性,不计入通过判定) ============
    print("\n" + "=" * 75)
    print("【6/6】材料基线比对(validate_baseline.py,仅报告)")
    print("=" * 75)
    rc, output = run_script("validate_baseline.py")
    print(output.strip())
    print("  · 本节只报告材料相对基线的变化,不影响机审结论;")
    print("    确认变化符合预期后,运行 `--update` 更新基线:")
    print("    .venv/bin/python docs/skill/validate_baseline.py --update")

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
