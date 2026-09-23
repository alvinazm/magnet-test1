"""
A 模块档位机审脚本(v1.2)
==========================

用途:自动检查 17 条 A 标准的中间状态档位是否:
    1. 显式对应标准描述的关键对象(培训文档 6.3.2 节)
    2. 不含培训文档 6.4 节禁止的抽象词
    3. 不是半截话(不以"但"或";"结尾)

使用:
    .venv/bin/python docs/specs/validate_a_module_levels.py

输出:
    - 每条 A 标准的检查结果
    - 标记"未覆盖对象"和"抽象词违规"
    - 总体合规报告 + 退出码(0=通过,1=有问题)
"""
import re
import sys
from pathlib import Path

WORK = Path("/Users/azm/MyProject/work")
DOC_PATH = WORK / "docs/specs/05-evaluation-rubric.md"


# ============================================================
# 1. 17 条 A 标准的"关键对象"清单
# ============================================================
# 每个对象必须在档位内容中显式出现
EXPECTED_OBJECTS = {
    "A-1": ["净收入", "毛利", "毛利率", "客单价"],
    "A-2": ["到店销售额", "门店 O2O 履约贡献", "即时配送"],
    "A-3": ["到店坪效", "租售比", "面积已调整"],
    "A-4": ["增量销售", "增量毛利", "ROI"],
    # A-5: 无档位
    "A-6": ["门店", "渠道", "订单", "SKU", "会员", "活动", "口径"],
    "A-7": ["异常", "高严重度", "时效性", "跨表", "缺失", "越界"],
    "A-8": ["文件", "Sheet", "字段", "记录"],
    # A-9: 无档位
    "A-10": ["具体动作", "量化目标", "优先级", "P0/P1/P2"],
    # A-11: 无档位
    "A-12": ["必含字段", "章节"],
    "A-13": ["渠道", "门店", "区域", "会员", "SKU", "活动"],
    "A-14": ["表头", "字段类型", "Sheet"],
    "A-15": ["封面", "摘要", "附录"],
    "A-16": ["数据", "判断", "建议", "绝对词"],
    "A-17": ["时间口径", "版本差异", "6 个月", "同店排除", "未覆盖"],
}


# ============================================================
# 2. 培训文档禁止的抽象词(6.3.2 + 6.4 节)
# ============================================================
# 任何档位描述中包含这些词都视为违规
# 培训文档原话:"不写'基本''较好''少量''多数'"
FORBIDDEN_WORDS = [
    "基本", "较好", "少量", "多数", "大部分", "几乎",
    "大概", "差不多", "个别",
    # 其他常见但应避免的词
    "大量", "一些", "几个", "很多", "少量", "少量",
    "较", "大概", "差不多",
]


# ============================================================
# 3. 文档解析
# ============================================================
def parse_standard(content: str, code: str) -> dict | None:
    """解析某条 A 标准的标准描述 + 档位"""
    pattern = rf"### {code} · .*?(?=### A-\d+|## |\Z)"
    m = re.search(pattern, content, re.DOTALL)
    if not m:
        return None
    section = m.group(0)

    # 提取"标准描述"代码块
    desc_match = re.search(r"```\n(.*?)\n```", section, re.DOTALL)
    desc = desc_match.group(1) if desc_match else ""

    # 提取"检查"行
    check_match = re.search(r"检查 ([^;]+?)(?:;|必要边界|$)", desc)
    check_text = check_match.group(1).strip() if check_match else ""

    # 提取档位表格
    levels = re.findall(r"\| 第 (\d+) 档 \| (\S+) \| (.+?) \|", section)

    return {
        "section": section,
        "desc": desc,
        "check_text": check_text,
        "levels": levels,
    }


# ============================================================
# 4. 关键对象匹配
# ============================================================
def check_object_in_levels(obj: str, levels_text: str) -> bool:
    """检查关键对象是否在档位内容中显式出现(子串匹配)"""
    return obj in levels_text


# ============================================================
# 5. 档位质量检查(抽象词 + 半截话)
# ============================================================
def check_level_quality(levels_text: str) -> list:
    """
    检查档位内容的质量:
    1. 不含培训文档禁止的抽象词
    2. 不是半截话(不以"但"或";"结尾)
    返回:违规列表
    """
    issues = []

    # 1. 抽象词检查
    for word in FORBIDDEN_WORDS:
        if word in levels_text:
            issues.append(f"含禁止抽象词: '{word}'")

    # 2. 半截话检查(以"但"或";"结尾,且不是完整句子)
    # 处理每个档位独立检查
    for num, label, desc in re.findall(
        r"\| 第 (\d+) 档 \| (\S+) \| (.+?) \|", levels_text + "|||"
    ):
        desc_stripped = desc.rstrip()
        # 末尾是"但"或";" 或 "但"在末尾 5 字符内
        if desc_stripped.endswith("但") or desc_stripped.endswith(";"):
            issues.append(f"第 {num} 档:以'但'或';'结尾(半截话)")
        # 检测"但" 在句子中是否接续(没说完)
        if "但" in desc and not re.search(r"但[^,;:。]*?(?:错|漏|缺失|缺少|不正确|未)", desc):
            # "但" 后面没接续(没具体后果),可能是半截话
            # 例:"A 级满足,但 B 级不满足" → OK
            # 例:"在上一档基础上,但" → 不 OK
            tail = desc.split("但")[-1].strip()
            if len(tail) < 5:  # 后面没几个字
                issues.append(f"第 {num} 档:'但' 后未接续(可能半截话)")

    return issues


# ============================================================
# 6. 主检查流程
# ============================================================
def main():
    if not DOC_PATH.exists():
        print(f"✗ 文档不存在: {DOC_PATH}")
        sys.exit(1)

    content = DOC_PATH.read_text()
    print("=" * 75)
    print(f"A 模块档位机审(v1.2 - 显式对应 + 抽象词检查)")
    print(f"文档:{DOC_PATH.name}")
    print("=" * 75)

    # 5 项必填字段检查
    print("\n[基础检查] 5 字段完整性")
    all_standards = re.findall(r"### (A-\d+) ·", content)
    print(f"  发现 A 标准:{len(all_standards)} 条")
    if len(all_standards) < 15:
        print(f"  ✗ 总条数 {len(all_standards)} < 15")

    # 逐条检查
    print("\n" + "=" * 75)
    print("[详细检查] 档位显式对应 + 抽象词/半截话")
    print("=" * 75)

    results = []  # (code, status, missing, covered, quality_issues)

    for code in sorted(EXPECTED_OBJECTS.keys()):
        parsed = parse_standard(content, code)
        if not parsed:
            print(f"\n  {code}: ✗ 未找到该标准")
            results.append((code, "未找到", [], [], []))
            continue

        # 字段完整性
        for field in ["类型", "影响程度", "是否有中间状态"]:
            if field not in parsed["section"]:
                print(f"  {code}: ⚠ 缺字段 '{field}'")

        # 档位检查
        if not parsed["levels"]:
            print(f"\n  {code}: - 无档位(跳过)")
            results.append((code, "无档位", [], [], []))
            continue

        levels_text = " ".join([l[2] for l in parsed["levels"]])

        # 1. 关键对象检查
        expected = EXPECTED_OBJECTS[code]
        missing = [o for o in expected if not check_object_in_levels(o, levels_text)]
        covered = [o for o in expected if check_object_in_levels(o, levels_text)]

        # 2. 档位质量检查(抽象词 + 半截话)
        quality_issues = check_level_quality(levels_text)

        # 状态判定
        if not missing and not quality_issues:
            status = "✓"
        else:
            status = "✗"

        print(f"\n  {code}: {status}")
        print(f"    检查部分:{parsed['check_text'][:80]}{'...' if len(parsed['check_text'])>80 else ''}")
        print(f"    档位数:{len(parsed['levels'])}")
        if missing:
            print(f"    ✗ 未覆盖对象:{missing}")
        else:
            print(f"    ✓ 已覆盖:{covered}")
        if quality_issues:
            print(f"    ✗ 抽象词/半截话问题:{quality_issues}")

        results.append((code, status, missing, covered, quality_issues))

    # 汇总
    print("\n" + "=" * 75)
    print("[汇总]")
    print("=" * 75)

    pass_count = sum(1 for r in results if r[1] == "✓" or r[1] in ["无档位", "未找到"])
    fail_count = sum(1 for r in results if r[1] == "✗")

    print(f"  通过:{pass_count}/{len(results)}")
    print(f"  失败:{fail_count}/{len(results)}")

    if fail_count > 0:
        print(f"\n  失败明细:")
        for code, status, missing, _, quality_issues in results:
            if status == "✗":
                reasons = []
                if missing:
                    reasons.append(f"缺对象 {missing}")
                if quality_issues:
                    reasons.append(f"质量问题 {quality_issues}")
                print(f"    {code}: {'; '.join(reasons)}")
        print()
        print("  ✗ 机审未通过,请修复后重跑")
        sys.exit(1)
    else:
        print("\n  ✅ 所有档位均通过机审(显式对应 + 无抽象词/半截话)")
        sys.exit(0)


if __name__ == "__main__":
    main()
