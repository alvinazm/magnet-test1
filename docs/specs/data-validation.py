"""
Magnet 出题附件数据自检脚本
==========================
用途:每次修改 num.py 后,运行本脚本验证数据完整性和计算一致性。

使用:
    .venv/bin/python docs/specs/data-validation.py

输出:
    - 通过项标记 ✓
    - 问题项标记 ✗,并给出详细信息
    - 最后给出汇总报告
"""
import sys
from pathlib import Path
import pandas as pd
import numpy as np

WORKSPACE = Path(__file__).resolve().parent.parent.parent

# 附件目录:input/ 子目录(模拟真实业务数据导出)
INPUT_DIR = WORKSPACE / "input"

# 预期附件清单
EXPECTED_FILES = {
    "2026H1渠道销售明细.xlsx": "主附件",
    "2026H1门店经营台账.xlsx": "主附件",
    "2026H1线上投放与平台费用.xlsx": "主附件",
    "2026H1会员与跨渠道订单.xlsx": "主附件",
    "2026H1促销活动记录.xlsx": "主附件",
    "华东区门店租约汇总.docx": "主附件",
    "渠道与门店管理策略.docx": "主附件",
    "旧版门店清单_2024.xlsx": "噪音附件",
    "SKU主数据.xlsx": "主附件",
}

issues = []  # 全局问题收集
warnings = []  # 警告(非阻断)


def check(condition, msg_ok, msg_fail):
    """统一检查入口,实时打印,返回 (passed, msg)"""
    if condition:
        print(msg_ok)
        return True, msg_ok
    issues.append(msg_fail)
    print(msg_fail)
    return False, msg_fail


def section(title):
    """打印分隔符"""
    print(f"\n{'=' * 75}")
    print(f"【{title}】")
    print("=" * 75)


def subsection(title):
    print(f"\n[{title}]")


# ============================================================
# 0. 文件存在性
# ============================================================
section("0. 文件存在性检查")
for filename, role in EXPECTED_FILES.items():
    filepath = INPUT_DIR / filename
    if filepath.exists():
        size_kb = filepath.stat().st_size / 1024
        print(f"  ✓ {filename:<35} {size_kb:>8.1f} KB  ({role})")
    else:
        check(False, f"文件缺失:{filename}")
        print(f"  ✗ {filename:<35} 缺失")


# ============================================================
# 1. 渠道销售明细
# ============================================================
section("1. 2026H1渠道销售明细.xlsx")
fp = INPUT_DIR / "2026H1渠道销售明细.xlsx"
if fp.exists():
    xls = pd.read_excel(fp, sheet_name=None)
    orders = xls["渠道销售明细"]
    refunds = xls["渠道退款明细"]
    monthly = xls["渠道月度汇总"]

    subsection("渠道销售明细 Sheet")
    expected_cols = ['订单日期', '订单号', '渠道ID', '渠道名称', '渠道类型', '门店ID',
                     '会员ID', 'SKU', '品类', '数量', '标价', '成交价', '优惠金额',
                     '收入', '商品成本', '退款金额', '退款数量', '订单状态', '发货方式',
                     '渠道归属', '月份', '净收入']
    missing = [c for c in expected_cols if c not in orders.columns]
    check(not missing, f"  ✓ 字段完整({len(orders.columns)} 个)",
          f"  ✗ 缺少字段:{missing}")

    # 主键唯一性
    dup = orders["订单号"].duplicated().sum()
    check(dup == 0, f"  ✓ 订单号唯一({len(orders)} 行)",
          f"  ✗ 订单号有 {dup} 个重复")

    # 计算公式
    orders["_收入_重算"] = orders["成交价"] * orders["数量"]
    err = (abs(orders["收入"] - orders["_收入_重算"]) > 0.01).sum()
    check(err == 0, f"  ✓ 收入 = 成交价×数量",
          f"  ✗ 收入计算错误 {err} 行")

    orders["_优惠_重算"] = (orders["标价"] - orders["成交价"]) * orders["数量"]
    err = (abs(orders["优惠金额"] - orders["_优惠_重算"]) > 0.01).sum()
    check(err == 0, f"  ✓ 优惠金额 = (标价-成交价)×数量",
          f"  ✗ 优惠金额错误 {err} 行")

    orders["_净收入_重算"] = orders["收入"] - orders["退款金额"]
    err = (abs(orders["净收入"] - orders["_净收入_重算"]) > 0.01).sum()
    check(err == 0, f"  ✓ 净收入 = 收入-退款金额",
          f"  ✗ 净收入错误 {err} 行")

    # 退款一致性
    err = (abs(orders["退款金额"] - orders["成交价"] * orders["退款数量"]) > 0.01).sum()
    check(err == 0, f"  ✓ 退款金额 = 成交价×退款数量",
          f"  ✗ 退款金额与成交价×退款数量不符 {err} 行")

    violate = (orders["退款数量"] > orders["数量"]).sum()
    check(violate == 0, f"  ✓ 退款数量 ≤ 数量",
          f"  ✗ 退款数量超过数量 {violate} 行")

    violate = (orders["成交价"] > orders["标价"]).sum()
    check(violate == 0, f"  ✓ 成交价 ≤ 标价",
          f"  ✗ 成交价超过标价 {violate} 行")

    # SKU 主数据交叉验证
    sku_fp = INPUT_DIR / "SKU主数据.xlsx"
    if sku_fp.exists():
        sku = pd.read_excel(sku_fp, sheet_name="SKU主数据")
        sku_cost_map = dict(zip(sku["SKU编码"], sku["单位成本"]))
        sku_cat_map = dict(zip(sku["SKU编码"], sku["品类"]))
        orders["_成本_重算"] = orders["SKU"].map(sku_cost_map) * orders["数量"]
        err = (abs(orders["商品成本"] - orders["_成本_重算"]) > 0.01).sum()
        check(err == 0, f"  ✓ 商品成本 = SKU单位成本×数量",
              f"  ✗ 商品成本与SKU主数据不符 {err} 行")

        orders["_品类_重算"] = orders["SKU"].map(sku_cat_map)
        err = (orders["品类"] != orders["_品类_重算"]).sum()
        check(err == 0, f"  ✓ 品类与SKU主数据一致",
              f"  ✗ 品类与SKU主数据不符 {err} 行")

    # 渠道分布(权重检查)
    subsection("渠道分布检查")
    dist = orders["渠道名称"].value_counts()
    total = dist.sum()
    for name, cnt in dist.items():
        pct = cnt / total * 100
        print(f"    {name:<14} {cnt:>5} 单  {pct:>5.1f}%")
    douyin_tmall = dist.get("抖音小店", 0) + dist.get("天猫旗舰店", 0)
    pct = douyin_tmall / total * 100
    check(pct >= 55, f"  ✓ 抖音+天猫占比 {pct:.1f}% ≥ 55%",
          f"  ✗ 抖音+天猫占比 {pct:.1f}% < 55%(可能不真实)")


# ============================================================
# 2. 门店经营台账
# ============================================================
section("2. 2026H1门店经营台账.xlsx")
fp = INPUT_DIR / "2026H1门店经营台账.xlsx"
if fp.exists():
    base = pd.read_excel(fp, sheet_name="门店基础信息")
    m = pd.read_excel(fp, sheet_name="门店月度经营")
    inv = pd.read_excel(fp, sheet_name="门店库存")
    exp = pd.read_excel(fp, sheet_name="门店费用")

    subsection("门店基础信息")
    check(base["门店ID"].nunique() == len(base),
          f"  ✓ 门店ID 唯一({len(base)} 家)",
          f"  ✗ 门店ID 有重复")
    check(base["租约ID"].nunique() == len(base),
          f"  ✓ 租约ID 唯一",
          f"  ✗ 租约ID 有重复")

    # 区域分布
    region_count = base["区域"].nunique()
    check(region_count >= 2, f"  ✓ 区域分布 {region_count} 个",
          f"  ✗ 区域只有 {region_count} 个,建议 ≥ 2")

    # 同店口径
    non_ss = (base["是否同店口径"] == "否").sum()
    if non_ss > 0:
        print(f"  · 同店口径=否 的门店:{non_ss} 家")

    subsection("门店月度经营")
    # 销售与订单同步
    ab1 = ((m["到店销售额"] > 0) & (m["订单数"] == 0)).sum()
    check(ab1 == 0, f"  ✓ 到店销售>0 但订单=0 的异常行 {ab1}",
          f"  ✗ 到店销售>0 但订单=0 异常 {ab1} 行")

    # 客单价计算
    m["_aov_重算"] = (m["到店销售额"] / m["订单数"].replace(0, np.nan)).round(2)
    err = (abs(m["客单价"].fillna(0) - m["_aov_重算"].fillna(0)) > 0.01).sum()
    check(err == 0, f"  ✓ 客单价 = 到店销售/订单数",
          f"  ✗ 客单价错误 {err} 行")

    # 毛利率范围
    violate = ((m["毛利率"] < 0) | (m["毛利率"] > 1)).sum()
    check(violate == 0, f"  ✓ 毛利率范围(0-1)",
          f"  ✗ 毛利率越界 {violate} 行")

    # 线上履约单量 = 自提 + 门店发货
    err = (m["线上订单履约单量"] - m["自提单量"] - m["门店发货单量"]).abs().gt(0).sum()
    check(err == 0, f"  ✓ 线上履约单量 = 自提+门店发货",
          f"  ✗ 线上履约单量构成错误 {err} 行")

    # O2O 履约贡献与渠道一致性
    if fp.exists():
        df_orders = pd.read_excel(INPUT_DIR / "2026H1渠道销售明细.xlsx",
                                   sheet_name="渠道销售明细")
        o2o = df_orders[df_orders["发货方式"].isin(["门店发货", "到店自提"])]
        o2o_agg = o2o.groupby(["门店ID", "月份"])["净收入"].sum().reset_index()
        o2o_agg.columns = ["门店ID", "月份", "_O2O_重算"]
        mm = m.merge(o2o_agg, on=["门店ID", "月份"], how="left")
        mm["_O2O_重算"] = mm["_O2O_重算"].fillna(0)
        err = (abs(mm["门店O2O履约贡献"] - mm["_O2O_重算"]) > 0.01).sum()
        check(err == 0, f"  ✓ 门店O2O履约贡献 = 渠道O2O聚合",
              f"  ✗ 门店O2O履约贡献与渠道不符 {err} 行")

    subsection("门店库存")
    zero_qty = ((inv["库存数量"] == 0) & (inv["库存金额"] > 0)).sum()
    check(zero_qty == 0, f"  ✓ 库存数量=0 但金额>0 {zero_qty}",
          f"  ✗ 库存数量=0 但金额>0 {zero_qty} 行")

    subsection("门店费用")
    for col in ['租金', '人力', '营销', '水电', '履约包装', '其他', '折旧摊销']:
        neg = (exp[col] < 0).sum()
        check(neg == 0, f"  ✓ {col} 无负值",
              f"  ✗ {col} 有 {neg} 个负值")


# ============================================================
# 3. 线上投放与平台费用
# ============================================================
section("3. 2026H1线上投放与平台费用.xlsx")
fp = INPUT_DIR / "2026H1线上投放与平台费用.xlsx"
if fp.exists():
    fee = pd.read_excel(fp, sheet_name="渠道月度费用")
    ad = pd.read_excel(fp, sheet_name="投放明细")
    rules = pd.read_excel(fp, sheet_name="平台佣金规则")

    subsection("渠道月度费用")
    for col in ['广告费', '平台佣金', '技术服务费', '支付手续费', '履约费',
                '包装费', '退货售后费', '其他费用']:
        neg = (fee[col] < 0).sum()
        check(neg == 0, f"  ✓ {col} 无负值",
              f"  ✗ {col} 有 {neg} 个负值")

    subsection("投放明细")
    # 投放金额量级
    ad_total = ad["投放金额"].sum()
    ad_monthly_total = fee["广告费"].sum()
    ratio = ad_total / ad_monthly_total
    check(0.5 <= ratio <= 2.0,
          f"  ✓ 投放明细/广告费 比值 {ratio:.2%}",
          f"  ⚠ 投放明细/广告费 比值 {ratio:.2%} 偏离(0.5-2.0)")

    subsection("平台佣金规则")
    check(len(rules) == 36,
          f"  ✓ 规则数 {len(rules)} (6渠道×6品类)",
          f"  ✗ 规则数 {len(rules)},预期 36")


# ============================================================
# 4. 会员与跨渠道订单
# ============================================================
section("4. 2026H1会员与跨渠道订单.xlsx")
fp = INPUT_DIR / "2026H1会员与跨渠道订单.xlsx"
if fp.exists():
    mb = pd.read_excel(fp, sheet_name="会员基础")
    mo = pd.read_excel(fp, sheet_name="会员订单")
    cross = pd.read_excel(fp, sheet_name="跨渠道行为")

    subsection("会员基础")
    check(mb["会员ID"].nunique() == len(mb),
          f"  ✓ 会员ID 唯一({len(mb)} 人)",
          f"  ✗ 会员ID 有重复")

    subsection("会员订单")
    dup = mo["订单号"].duplicated().sum()
    check(dup == 0, f"  ✓ 订单号唯一({len(mo)} 行)",
          f"  ✗ 订单号有 {dup} 个重复")

    # 与渠道订单号不撞号
    df_orders = pd.read_excel(INPUT_DIR / "2026H1渠道销售明细.xlsx",
                               sheet_name="渠道销售明细")
    collide = set(mo["订单号"]) & set(df_orders["订单号"])
    check(len(collide) == 0,
          f"  ✓ 会员订单号与渠道不撞号",
          f"  ✗ 撞号 {len(collide)} 个")

    subsection("跨渠道行为")
    # 累计消费应等于该会员在会员订单的金额合计
    mo_total = mo.groupby("会员ID")["金额"].sum().reset_index()
    cross_check = cross.merge(mo_total, on="会员ID", how="left")
    cross_check["金额"] = cross_check["金额"].fillna(0)
    err = (abs(cross_check["累计消费"] - cross_check["金额"]) > 0.01).sum()
    check(err == 0, f"  ✓ 累计消费 = 会员订单金额合计",
          f"  ✗ 累计消费与会员订单不符 {err} 行")


# ============================================================
# 5. 华东区门店租约汇总.docx
# ============================================================
section("5. 华东区门店租约汇总.docx")
fp = INPUT_DIR / "华东区门店租约汇总.docx"
if fp.exists():
    try:
        from docx import Document
        doc = Document(str(fp))
        all_text = "\n".join(p.text for p in doc.paragraphs)
        import re
        sids = re.findall(r'门店ID[:：]\s*(S\d{3})', all_text)
        lease_ids = re.findall(r'租约ID[:：]\s*(L\d{3})', all_text)
        check(len(set(sids)) == 42,
              f"  ✓ 门店ID {len(set(sids))} 个(预期 42)",
              f"  ✗ 门店ID {len(set(sids))} 个,预期 42")
        check(len(set(lease_ids)) == 42,
              f"  ✓ 租约ID {len(set(lease_ids))} 个",
              f"  ✗ 租约ID 数量不符")
        check("华东区门店租约汇总" in all_text or "华东区" in all_text,
              f"  ✓ 文档包含「华东区」字样",
              f"  ✗ 文档标题缺失")
        check("月租金" in all_text,
              f"  ✓ 文档包含月租金字段",
              f"  ✗ 文档缺少月租金字段")
    except ImportError:
        warnings.append("python-docx 未安装,跳过 docx 检查")
        print("  ⚠ python-docx 未安装,跳过")
else:
    check(False, "  ✓ 租约 Word 存在", "  ✗ 租约 Word 不存在")


# ============================================================
# 6. 渠道与门店管理策略.docx
# ============================================================
section("6. 渠道与门店管理策略.docx")
fp = INPUT_DIR / "渠道与门店管理策略.docx"
if fp.exists():
    try:
        from docx import Document
        doc = Document(str(fp))
        all_text = "\n".join(p.text for p in doc.paragraphs)
        chapters = ['渠道分类', '门店分级', 'O2O', '数据来源', '缺失与异常', '闭店']
        for ch in chapters:
            check(ch in all_text,
                  f"  ✓ 包含章节「{ch}」",
                  f"  ✗ 缺少章节「{ch}」")
    except ImportError:
        warnings.append("python-docx 未安装")
else:
    check(False, "  ✓ 策略 Word 存在", "  ✗ 策略 Word 不存在")


# ============================================================
# 7. 促销活动记录
# ============================================================
section("7. 2026H1促销活动记录.xlsx")
fp = INPUT_DIR / "2026H1促销活动记录.xlsx"
if fp.exists():
    act = pd.read_excel(fp, sheet_name="活动清单")
    det = pd.read_excel(fp, sheet_name="活动明细")
    eff = pd.read_excel(fp, sheet_name="活动效果")

    check(act["活动ID"].nunique() == len(act),
          f"  ✓ 活动ID 唯一({len(act)} 个)",
          f"  ✗ 活动ID 有重复")

    check((act["结束日期"] >= act["开始日期"]).all(),
          f"  ✓ 活动结束日期 ≥ 开始日期",
          f"  ✗ 活动结束日期早于开始日期")

    # ROI 校验
    eff_full = eff.merge(act[["活动ID", "实际费用"]], on="活动ID")
    eff_full["_ROI_重算"] = (eff_full["增量毛利"] / eff_full["实际费用"]).round(2)
    err = (abs(eff_full["ROI"] - eff_full["_ROI_重算"]) > 0.01).sum()
    check(err == 0, f"  ✓ ROI = 增量毛利/实际费用",
          f"  ✗ ROI 错误 {err} 行")

    # 增量销售
    eff["_增量_重算"] = (eff["活动期间销售额"] - eff["活动前销售额"]).round(2)
    err = (abs(eff["增量销售额"] - eff["_增量_重算"]) > 0.01).sum()
    check(err == 0, f"  ✓ 增量销售额 = 期间-活动前",
          f"  ✗ 增量销售额错误 {err} 行")

    # ROI 范围合理性(应有大有小,不能全是同一个数)
    roi_unique = eff["ROI"].nunique()
    check(roi_unique > 10,
          f"  ✓ ROI 有 {roi_unique} 个不同值",
          f"  ⚠ ROI 只有 {roi_unique} 个不同值")


# ============================================================
# 8. 旧版门店清单(噪音检查)
# ============================================================
section("8. 旧版门店清单_2024.xlsx(噪音检查)")
fp = INPUT_DIR / "旧版门店清单_2024.xlsx"
if fp.exists():
    old = pd.read_excel(fp, sheet_name="门店清单")
    # 禁用词检查
    bad_words = ['测试', '模拟', '样例', '合成', 'synthetic', 'test', '噪音', '答案']
    bad_found = []
    for word in bad_words:
        for col in old.columns:
            if old[col].dtype == 'O':
                hits = old[col].astype(str).str.contains(word, na=False).sum()
                if hits > 0:
                    bad_found.append(f"{word} 在 {col} 出现 {hits} 次")
    check(not bad_found,
          f"  ✓ 无禁用词",
          f"  ✗ 发现禁用词:{bad_found}")
    print(f"  · 旧版清单 {len(old)} 家门店,状态分布:{old['状态'].value_counts().to_dict()}")


# ============================================================
# 9. SKU 主数据
# ============================================================
section("9. SKU主数据.xlsx")
fp = INPUT_DIR / "SKU主数据.xlsx"
if fp.exists():
    sku = pd.read_excel(fp, sheet_name="SKU主数据")
    check(sku["SKU编码"].nunique() == len(sku),
          f"  ✓ SKU 编码唯一({len(sku)} 个)",
          f"  ✗ SKU 编码有重复")

    violate = (sku["单位成本"] >= sku["标准标价"]).sum()
    check(violate == 0, f"  ✓ 单位成本 < 标准标价",
          f"  ✗ 单位成本 ≥ 标准标价 {violate} 行")

    sku["_最低价_重算"] = (sku["单位成本"] * 1.15).round(2)
    err = (abs(sku["建议最低售价"] - sku["_最低价_重算"]) > 0.01).sum()
    check(err == 0, f"  ✓ 建议最低售价 = 成本×1.15",
          f"  ✗ 建议最低售价错误 {err} 行")


# ============================================================
# 汇总
# ============================================================
section("汇总")
print(f"\n  问题数:{len(issues)}")
print(f"  警告数:{len(warnings)}")

if issues:
    print(f"\n  详细问题:")
    for i, msg in enumerate(issues, 1):
        print(f"    {i}. {msg}")

if warnings:
    print(f"\n  警告:")
    for w in warnings:
        print(f"    - {w}")

if not issues:
    print(f"\n  ✅ 所有硬性检查通过!数据可用。")
    sys.exit(0)
else:
    print(f"\n  ❌ 存在 {len(issues)} 个问题,请修复后重新生成。")
    sys.exit(1)
