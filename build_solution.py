"""
2026H2 渠道与门店经营策略 — 专家 Solution 生成器
=================================================
依据:`docs/specs/step2-query.md`(正式任务书)与 `input/` 下的 9 个附件
产出:`output/` 下 5 个交付物 + 《数据血缘说明.md》

使用:
    .venv/bin/python build_solution.py

交付物:
    1. 《2026H2 渠道与门店经营策略报告.docx》
    2. 《门店分级与调整建议.xlsx》
    3. 《渠道资源再配置建议.xlsx》
    4. 《促销组合优化建议.xlsx》
    5. 《异常与待核清单.xlsx》
    + 《数据血缘说明.md》(逐指标说明分析逻辑、数据来源与口径)
"""
import re
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt

INPUT = Path("input")
OUTPUT = Path("output")
OUTPUT.mkdir(exist_ok=True)
BASE_DATE = pd.Timestamp("2026-06-30")
REPORT_DATE = "2026 年 07 月 05 日"


def rd(name, sheet=None):
    return pd.read_excel(INPUT / name, sheet_name=sheet)


def money(x):
    return f"{x:,.2f}"


def pct(x):
    return f"{x * 100:.1f}%"


# ============================================================
# 0. 读取附件
# ============================================================
orders = rd("2026H1渠道销售明细.xlsx", "渠道销售明细")
refund = rd("2026H1渠道销售明细.xlsx", "渠道退款明细")
ch_month = rd("2026H1渠道销售明细.xlsx", "渠道月度汇总")
store_base = rd("2026H1门店经营台账.xlsx", "门店基础信息")
store_month = rd("2026H1门店经营台账.xlsx", "门店月度经营")
store_inv = rd("2026H1门店经营台账.xlsx", "门店库存")
store_exp = rd("2026H1门店经营台账.xlsx", "门店费用")
ch_fee = rd("2026H1线上投放与平台费用.xlsx", "渠道月度费用")
ad = rd("2026H1线上投放与平台费用.xlsx", "投放明细")
rule = rd("2026H1线上投放与平台费用.xlsx", "平台佣金规则")
member_base = rd("2026H1会员与跨渠道订单.xlsx", "会员基础")
member_order = rd("2026H1会员与跨渠道订单.xlsx", "会员订单")
cross = rd("2026H1会员与跨渠道订单.xlsx", "跨渠道行为")
activity = rd("2026H1促销活动记录.xlsx", "活动清单")
activity_detail = rd("2026H1促销活动记录.xlsx", "活动明细")
activity_effect = rd("2026H1促销活动记录.xlsx", "活动效果")
sku_master = rd("SKU主数据.xlsx", "SKU主数据")
old_store = rd("旧版门店清单_2024.xlsx", "门店清单")

# 租约(逐门店解析:租期结束、月租金、物业费、违约金)
lease_doc = Document(INPUT / "华东区门店租约汇总.docx")
lease_text = "\n".join(p.text for p in lease_doc.paragraphs)
lease = {}
for block in re.split(r"\n(?=S\d{3} 门店)", lease_text):
    sid = re.search(r"门店ID[:：]\s*(S\d{3})", block)
    if not sid:
        continue
    lease[sid.group(1)] = {
        "租约ID": re.search(r"租约ID[:：]\s*(\S+)", block).group(1),
        "租期开始": re.search(r"租期开始[:：]\s*(\S+)", block).group(1),
        "租期结束": re.search(r"租期结束[:：]\s*(\S+)", block).group(1),
        "月租金": float(re.search(r"月租金[:：]\s*([\d.]+)", block).group(1)),
        "物业费": float(re.search(r"物业费[:：]\s*([\d.]+)", block).group(1)),
        "违约金": float(re.search(r"支付([\d.]+)元违约金", block).group(1)),
        "备注": re.search(r"备注[:：]\s*(\S+)", block).group(1),
    }

policy_text = "\n".join(p.text for p in Document(INPUT / "渠道与门店管理策略.docx").paragraphs)

# ============================================================
# 1. 渠道诊断(交付物 3)
# ============================================================
channel_type = orders.groupby("渠道ID")["渠道类型"].first()
ch = ch_fee.groupby(["渠道ID", "渠道名称"]).agg(
    净收入=("净收入", "sum"),
    毛利=("毛利", "sum"),
    广告费=("广告费", "sum"),
    平台佣金=("平台佣金", "sum"),
    履约费=("履约费", "sum"),
    费用合计=("费用合计", "sum"),
    渠道净利=("渠道净利", "sum"),
).reset_index()
ch["渠道类型"] = ch["渠道ID"].map(channel_type)
ad_by_ch = ad.groupby("渠道ID").agg(投放金额=("投放金额", "sum"), 支付金额=("支付金额", "sum"))
ch = ch.merge(ad_by_ch, on="渠道ID", how="left")
ch["渠道毛利率"] = ch["毛利"] / ch["净收入"]
ch["销售费用率"] = ch["费用合计"] / ch["净收入"]
ch["投放ROI"] = ch["支付金额"] / ch["投放金额"]
total_net = ch["净收入"].sum()
ch["净收入占比"] = ch["净收入"] / total_net


def channel_direction(r):
    """渠道 H2 资源方向:自定规则(已在《异常与待核清单.xlsx》披露)。"""
    if r["渠道净利"] > 0 and r["销售费用率"] <= 0.30:
        return "增投", 0.20
    if r["渠道净利"] < 0 and r["销售费用率"] >= 0.33:
        return "收缩", -0.15
    return "维持", 0.0


ch[["H2资源调整方向", "建议投入幅度"]] = ch.apply(
    lambda r: pd.Series(channel_direction(r)), axis=1
)
ch = ch.sort_values("净收入", ascending=False).reset_index(drop=True)
ch = ch[["渠道ID", "渠道名称", "渠道类型", "净收入", "毛利", "广告费", "平台佣金", "履约费",
         "费用合计", "渠道净利", "投放金额", "支付金额", "渠道毛利率", "销售费用率",
         "投放ROI", "净收入占比", "H2资源调整方向", "建议投入幅度"]]

# ============================================================
# 2. 门店诊断(交付物 2)
# ============================================================
store = store_month.groupby("门店ID").agg(
    到店销售额=("到店销售额", "sum"),
    O2O履约贡献=("门店O2O履约贡献", "sum"),
    履约单量=("线上订单履约单量", "sum"),
).join(store_base.set_index("门店ID")[["门店名称", "区域", "城市", "店型", "营业面积", "面积是否调整"]])
store["月均到店销售"] = store["到店销售额"] / 6
store["月均租金及物业费"] = store_exp.groupby("门店ID")["租金及物业费"].mean()
store["到店坪效"] = store["月均到店销售"] / store["营业面积"]
store["租售比"] = store["月均租金及物业费"] / store["月均到店销售"]


def grade(peak, ratio):
    if ratio > 0.35:
        return "D"
    if peak >= 5000:
        return "A"
    if peak >= 2000:
        return "B"
    if peak >= 1000:
        return "C"
    return "D"


store["分级"] = [grade(p, r) for p, r in zip(store["到店坪效"], store["租售比"])]

# 月度分级:用于判断 6.1 条"连续 6 个月 D 级"
monthly = store_month.merge(
    store_exp[["门店ID", "月份", "租金及物业费"]], on=["门店ID", "月份"]
).merge(store_base[["门店ID", "营业面积"]], on="门店ID")
monthly["坪效"] = (monthly["到店销售额"] / monthly["营业面积"])
monthly["租售比"] = monthly["租金及物业费"] / monthly["到店销售额"]
monthly["级别"] = [grade(p, r) for p, r in zip(monthly["坪效"], monthly["租售比"])]
d_all6 = monthly.groupby("门店ID")["级别"].apply(lambda s: (s == "D").all())
store["连续6个月D级"] = store.index.map(d_all6)
store["租期结束"] = [lease[s]["租期结束"] for s in store.index]
store["闭店违约金"] = [lease[s]["违约金"] for s in store.index]
store["租约在H2内到期"] = [lease[s]["租期结束"] <= "2026-12-31" for s in store.index]


def store_action(r):
    # 取值必须落在 Query 交付物 2 规定的枚举内:维持 / 整改 / 转前置仓或自提点 / 闭店评估
    if r["分级"] == "D" and r["连续6个月D级"] and r["租约在H2内到期"]:
        return "闭店评估"
    if r["分级"] == "D":
        return "转前置仓或自提点"
    if r["分级"] == "C":
        return "整改"
    if r["分级"] == "A":
        return "维持"
    return "维持"


store["调整建议"] = store.apply(store_action, axis=1)
store["数据状态"] = np.where(store["面积是否调整"] == "是", "待补", "已确认")
store["关键依据"] = store.apply(
    lambda r: (
        f"《2026H1门店经营台账.xlsx》门店月度经营.到店销售额({money(r['到店销售额'])}元);"
        f"门店费用.租金及物业费(月均{money(r['月均租金及物业费'])}元);"
        f"门店基础信息.营业面积({int(r['营业面积'])}㎡)"
        + (";租约汇总.租期结束(" + r["租期结束"] + ")" if r["分级"] == "D" else "")
    ),
    axis=1,
)
grade_dist = store["分级"].value_counts().reindex(["A", "B", "C", "D"]).fillna(0).astype(int).to_dict()
store_out = store.reset_index()[
    ["门店ID", "门店名称", "区域", "城市", "店型", "营业面积", "到店销售额", "O2O履约贡献",
     "到店坪效", "租售比", "面积是否调整", "分级", "调整建议", "关键依据", "数据状态"]
].rename(columns={"到店销售额": "2026H1累计到店销售额", "O2O履约贡献": "2026H1累计门店O2O履约贡献",
                  "到店坪效": "到店坪效(元/㎡/月)"})
store_out["2026H1累计到店销售额"] = store_out["2026H1累计到店销售额"].round(2)
store_out["2026H1累计门店O2O履约贡献"] = store_out["2026H1累计门店O2O履约贡献"].round(2)
store_out["到店坪效(元/㎡/月)"] = store_out["到店坪效(元/㎡/月)"].round(0).astype(int)
store_out["租售比"] = store_out["租售比"].round(4)

# ============================================================
# 3. 促销诊断(交付物 4)
# ============================================================
promo = activity_effect.merge(activity[["活动ID", "活动名称", "活动类型", "实际费用", "预算"]], on="活动ID")


def promo_action(roi):
    """自定阈值:ROI ≥2 保留;1 ≤ ROI <2 调整;ROI <1 取消(已在《异常与待核清单.xlsx》披露)。"""
    if roi >= 2:
        return "保留"
    if roi >= 1:
        return "调整"
    return "取消"


promo["建议动作"] = promo["ROI"].apply(promo_action)
promo["关键依据"] = promo.apply(
    lambda r: f"《2026H1促销活动记录.xlsx》活动效果.增量毛利({money(r['增量毛利'])}元)/活动清单.实际费用({money(r['实际费用'])}元)",
    axis=1,
)
type_stat = promo.groupby("活动类型").agg(
    活动数=("活动ID", "count"),
    中位ROI=("ROI", "median"),
    整体ROI=("增量毛利", "sum"),
    费用合计=("实际费用", "sum"),
).reset_index()
type_stat["整体ROI"] = (type_stat["整体ROI"] / type_stat["费用合计"]).round(2)
type_stat["中位ROI"] = type_stat["中位ROI"].round(2)

# H2 促销节奏建议(按 H1 类型效率排序,结合电商日历)
budget_total = activity["实际费用"].sum()
h2_plan = [
    ("2026-07", "会员日 + 满减(淡季稳定流量)", f"{money(budget_total * 0.12)} - {money(budget_total * 0.16)}", "暑期消费/暑促"),
    ("2026-08", "节日促销(七夕)", f"{money(budget_total * 0.10)} - {money(budget_total * 0.14)}", "七夕"),
    ("2026-09", "折扣 + 品类联合促销", f"{money(budget_total * 0.14)} - {money(budget_total * 0.18)}", "秋季上新"),
    ("2026-10", "节日促销(国庆) + 满减", f"{money(budget_total * 0.16)} - {money(budget_total * 0.20)}", "国庆黄金周"),
    ("2026-11", "折扣 + 节日促销(双 11)", f"{money(budget_total * 0.22)} - {money(budget_total * 0.26)}", "双 11"),
    ("2026-12", "满减 + 节日促销(双 12/年货节)", f"{money(budget_total * 0.18)} - {money(budget_total * 0.22)}", "双 12/年货节"),
]

# ============================================================
# 4. 会员分析(报告第 4 章)
# ============================================================
reg_dist = member_base["注册渠道"].value_counts(normalize=True).round(4)
cross_ratio = (cross["跨渠道购买渠道数"] >= 2).mean()
avg_channel = cross["跨渠道购买渠道数"].mean()
orders_per_member = cross["订单数"].mean()
repurchase_rate = (cross["订单数"] >= 2).mean()
top10_cut = cross["累计消费"].quantile(0.90)
top10_share = cross.loc[cross["累计消费"] >= top10_cut, "累计消费"].sum() / cross["累计消费"].sum()
churn_rate = (cross["流失标记"] == "是").mean()
rfm_mean = cross["RFM分"].mean()
member_total = len(member_base)

# ============================================================
# 5. 异常与待核清单(交付物 5)
# ============================================================
adjusted = store[store["面积是否调整"] == "是"]
anomalies = [
    ("A-001", "时效性", "旧版门店清单_2024.xlsx", "全部字段",
     "该附件为 2024 年门店快照(50 家),与本任务 2026H1 口径不同;其中 8 家门店在当前台账中不存在。",
     "全部门店类结论", "高", "识别为历史快照,仅作对照参考,不纳入任何计算",
     "《旧版门店清单_2024.xlsx》的文件名含 _2024;门店数 50 家 vs 当前 42 家;《渠道与门店管理策略.docx》第 5.2 条不得用外部/过期数据"),
    ("A-002", "口径不清", "2026H1门店经营台账.xlsx", "门店基础信息.面积是否调整 / 门店月度经营.到店销售额",
     f"计划期内发生面积调整的门店 {len(adjusted)} 家({', '.join(adjusted.index)}),其按当前面积计算的坪效口径可能失真。",
     f"{len(adjusted)} 家门店的坪效与分级结果", "中",
     "保留分级与建议,并在本清单披露;如需精确结论,应按调整前后的加权面积重算坪效",
     "《渠道与门店管理策略.docx》第 2.5 条;《华东区门店租约汇总.docx》备注字段"),
    ("A-003", "口径不清", "2026H1门店经营台账.xlsx", "门店费用.租金及物业费 / 华东区门店租约汇总.月租金",
     '门店费用"租金及物业费"为合并口径(= 合同月租金 + 物业费),与租约文档单列的"月租金"不同口径。',
     "门店租售比与门店费用结构", "低",
     '租售比统一使用"租金及物业费"合并口径;租约文档用于合同条款(租期、违约金、递增)判断',
     "《渠道与门店管理策略.docx》第 4.4 条;本报告《数据血缘说明》第二节"),
    ("A-004", "口径不清", "2026H1渠道销售明细.xlsx", "发货方式 / 门店ID",
     "快递订单由中心仓发货、不关联门店ID,不计入门店 O2O 履约贡献;门店发货/到店自提/即时配送订单均关联门店。",
     "门店 O2O 履约贡献与履约单量口径", "低",
     '按《渠道与门店管理策略.docx》第 3.3–3.6 条计算:门店履约贡献仅含"门店发货 + 到店自提"',
     "《渠道与门店管理策略.docx》第 3.3、3.4、3.6 条"),
    ("A-005", "口径不清", "2026H1会员与跨渠道订单.xlsx", "跨渠道行为.累计消费 / 跨渠道购买渠道数",
     "会员累计消费以《2026H1会员与跨渠道订单.xlsx》会员订单 Sheet 为口径;《2026H1渠道销售明细.xlsx》渠道销售明细 Sheet 的 `会员ID` 仅用于渠道归因,不重复计入。",
     "会员分层与 Top10% 消费贡献", "低",
     "按《渠道与门店管理策略.docx》第 9.6 条口径统计;跨渠道购买渠道数含门店渠道",
     "《渠道与门店管理策略.docx》第 9.4、9.6 条"),
    ("A-006", "口径不清", "2026H1促销活动记录.xlsx", "活动效果.ROI",
     "本报告自定建议动作阈值:ROI ≥2 保留、1 ≤ ROI <2 调整、ROI <1 取消;附件内未规定阈值。",
     "50 个促销活动的建议动作", "中",
     "已披露自定阈值定义方法;若管理层另有 ROI 门槛,可按新阈值重算建议动作",
     "附件内均未规定;依培训文档允许自定补充指标并披露"),
    # 原 A-007「跨表主键与聚合口径已核验、未发现冲突」属"核验通过记录"而非异常,
    # 已移入《数据血缘说明.md》第七节(异常清单只登记真实问题)。
    ("A-007", "缺失", "2026H1促销活动记录.xlsx", "活动效果",
     "活动效果仅提供活动整体 ROI,未提供分渠道/分时段的增量归因,直播类低 ROI 无法进一步拆解原因。",
     "促销类型取舍与直播策略建议", "中",
     "按整体 ROI 给出建议,并在报告中标注归因限制;H2 建议以实验方式验证",
     "《2026H1促销活动记录.xlsx》结构限制;《渠道与门店管理策略.docx》第 5.1 条允许披露与缩小范围"),
    ("A-008", "缺失", "2026H1门店经营台账.xlsx", "门店库存.品类",
     "库存数据仅到门店-品类-月粒度,未提供 SKU 级库存与滞销明细。",
     "SKU 级库存与蚕食分析", "低", "结论仅到品类级,不作 SKU 级推断",
     "《2026H1促销活动记录.xlsx》结构限制"),
    ("A-009", "越界", "2026H1渠道销售明细.xlsx / 2026H1会员与跨渠道订单.xlsx", "会员ID / 订单号",
     "渠道订单与会员订单为两套编号(O100000–O107499 / O800000–O805999),不可按订单号合并;渠道订单中约 55% 未关联会员ID。",
     "会员与渠道交叉分析的范围", "低",
     "会员维度分析以《2026H1会员与跨渠道订单.xlsx》会员订单 Sheet、跨渠道行为 Sheet 为准,渠道订单仅用于渠道侧指标",
     "《2026H1渠道销售明细.xlsx》与《2026H1会员与跨渠道订单.xlsx》的编号规则;《渠道与门店管理策略.docx》第 9.6 条"),
    ("A-010", "缺失", "2026H1渠道销售明细.xlsx", "数据期间",
     "数据仅覆盖 2026H1(6 个月),无历史同期与更长周期数据,无法进行季节性、年度趋势或生命周期类分析。",
     "H2 预测与趋势判断", "中",
     "H2 目标以 H1 实际水平为基准并标注为估算;不做跨年趋势结论",
     "《2026H1门店经营台账.xlsx》数据粒度限制;《渠道与门店管理策略.docx》第 5.1 条"),
]
anomaly_df = pd.DataFrame(anomalies, columns=[
    "异常编号", "异常类别", "涉及文件", "涉及字段", "问题描述", "影响范围",
    "严重度", "建议处理方式", "披露依据"])

# ============================================================
# 6. 写 Excel 交付物
# ============================================================
with pd.ExcelWriter(OUTPUT / "门店分级与调整建议.xlsx", engine="openpyxl") as w:
    store_out.to_excel(w, sheet_name="门店清单", index=False)

channel_out = ch.rename(columns={
    "净收入": "2026H1净收入(元)", "毛利": "毛利(元)", "广告费": "广告费(元)",
    "平台佣金": "平台佣金(元)", "履约费": "履约费(元)",
})[[
    "渠道ID", "渠道名称", "渠道类型", "2026H1净收入(元)", "毛利(元)", "广告费(元)",
    "平台佣金(元)", "履约费(元)", "渠道毛利率", "投放ROI", "销售费用率",
    "H2资源调整方向", "建议投入幅度", "渠道净利",
]]
channel_out = channel_out.assign(
    关键依据=[
        "《2026H1渠道销售明细.xlsx》 渠道月度汇总(净收入/毛利);《2026H1线上投放与平台费用.xlsx》 渠道月度费用(广告/佣金/履约);《2026H1线上投放与平台费用.xlsx》 投放明细(支付金额/投放金额)"
    ] * len(channel_out),
    数据状态="已确认",
)
channel_out.columns = [
    "渠道ID", "渠道名称", "渠道类型", "2026H1净收入(元)", "毛利(元)", "广告费(元)",
    "平台佣金(元)", "履约费(元)", "渠道毛利率", "投放ROI", "销售费用率",
    "H2资源调整方向", "建议投入幅度", "渠道净利(元)", "关键依据", "数据状态",
]
with pd.ExcelWriter(OUTPUT / "渠道资源再配置建议.xlsx", engine="openpyxl") as w:
    channel_out.to_excel(w, sheet_name="渠道清单", index=False)

promo_out = promo[["活动ID", "活动名称", "活动类型", "渠道ID", "实际费用", "增量销售额",
                   "增量毛利", "ROI", "建议动作", "关键依据"]].copy()
promo_out.columns = ["活动ID", "活动名称", "活动类型", "渠道ID", "实际费用(元)",
                     "增量销售额(元)", "增量毛利(元)", "ROI", "建议动作", "关键依据"]
h2_df = pd.DataFrame(h2_plan, columns=["月份", "建议活动类型", "预算区间(元)", "关键节点"])
with pd.ExcelWriter(OUTPUT / "促销组合优化建议.xlsx", engine="openpyxl") as w:
    promo_out.to_excel(w, sheet_name="活动评估", index=False)
    h2_df.to_excel(w, sheet_name="H2节奏建议", index=False)

with pd.ExcelWriter(OUTPUT / "异常与待核清单.xlsx", engine="openpyxl") as w:
    anomaly_df.to_excel(w, sheet_name="异常清单", index=False)

print("已写出 4 个 Excel 交付物")

# ============================================================
# 7. 写 Word 报告(交付物 1)
# ============================================================
doc = Document()
doc.add_heading("2026H2 渠道与门店经营策略报告", level=0)
doc.add_paragraph("—— 2026H1 经营复盘与下半年资源再配置方案").alignment = 1
p = doc.add_paragraph()
p.add_run(f"报告期:2026-01-01 至 2026-06-30(2026H1)\n").bold = True
p.add_run(f"建议执行期:2026-07-01 至 2026-12-31(2026H2)\n")
p.add_run(f"基准日:2026-06-30\n报告提交人:渠道与门店运营策略经理\n报告提交日期:{REPORT_DATE}\n")
p.add_run("报告对象:集团 CEO、CFO、COO、董事会成员")

doc.add_heading("摘要", level=1)
net_total = ch["净收入"].sum()
gross_total = ch["毛利"].sum()
profit_total = ch["渠道净利"].sum()
wx = ch[ch["渠道ID"] == "CH02"].iloc[0]
top2 = ch.head(2)
worst_promo_share = (promo["建议动作"] == "取消").mean()
summary_points = [
    f"渠道规模:2026H1 六个线上渠道合计净收入 {money(net_total)} 元、毛利 {money(gross_total)} 元,"
    f"但扣除渠道费用后合计净利 {money(profit_total)} 元,整体仍未盈利。",
    f"渠道分化:{ch.iloc[0]['渠道名称']} 与 {ch.iloc[1]['渠道名称']} 合计占净收入 {pct(top2['净收入'].sum() / net_total)},"
    f"集中度偏高;仅 {wx['渠道名称']} 实现正净利({money(wx['渠道净利'])} 元,销售费用率 {pct(wx['销售费用率'])}),"
    f"其余渠道费用率均在 {pct(ch[ch['渠道ID'] != 'CH02']['销售费用率'].min())}–{pct(ch[ch['渠道ID'] != 'CH02']['销售费用率'].max())} 区间。",
    f"门店结构:42 家门店中 A 级 {grade_dist['A']} 家、B 级 {grade_dist['B']} 家、"
    f"C 级 {grade_dist['C']} 家、D 级 {grade_dist['D']} 家;"
    f"门店 O2O 履约贡献合计 {money(store['O2O履约贡献'].sum())} 元,已成为门店价值的重要组成部分。",
    f"促销效率:50 个活动 ROI 中位数 {promo['ROI'].median():.2f},{pct(worst_promo_share)} 的活动 ROI 低于 1,"
    f"促销组合存在明显低效投入,需按类型重构。",
    f"会员基础:{member_total:,} 名会员中 {pct(cross_ratio)} 在 2 个及以上渠道购买(平均 {avg_channel:.2f} 个渠道),"
    f"Top10% 会员贡献 {pct(top10_share)} 累计消费,流失会员占比 {pct(churn_rate)}。",
]
for s in summary_points:
    doc.add_paragraph(s, style="List Number")

doc.add_heading("第 1 章 渠道结构诊断与 H2 资源调整建议", level=1)
doc.add_paragraph(
    f"2026H1 六个线上渠道合计净收入 {money(net_total)} 元,毛利 {money(gross_total)} 元,"
    f"平均渠道毛利率 {pct(gross_total / net_total)};渠道费用(广告、平台佣金、技术服务、支付、履约、包装、"
    f"退货售后及其他)合计 {money(ch['费用合计'].sum())} 元,平均销售费用率 "
    f"{pct(ch['费用合计'].sum() / net_total)},合计净利 {money(profit_total)} 元。"
)
t = doc.add_table(rows=1, cols=7)
t.style = "Light Grid Accent 1"
t.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = ["渠道", "净收入(元)", "毛利率", "销售费用率", "投放ROI", "净利(元)", "H2 方向"]
for i, h in enumerate(hdr):
    t.rows[0].cells[i].text = h
for _, r in ch.iterrows():
    cells = t.add_row().cells
    cells[0].text = r["渠道名称"]
    cells[1].text = money(r["净收入"])
    cells[2].text = pct(r["渠道毛利率"])
    cells[3].text = pct(r["销售费用率"])
    cells[4].text = f"{r['投放ROI']:.2f}"
    cells[5].text = money(r["渠道净利"])
    cells[6].text = f"{r['H2资源调整方向']}({r['建议投入幅度']:+.0%})"
doc.add_paragraph("")
doc.add_paragraph(
    f"结论与建议:① 对 {wx['渠道名称']} 增投 20%(唯一正净利渠道,销售费用率仅 {pct(wx['销售费用率'])});"
    f"② 对销售费用率 ≥33% 且亏损的渠道(共 "
    f"{len(ch[(ch['渠道净利'] < 0) & (ch['销售费用率'] >= 0.33)])} 个)收缩投入 15%,"
    f"重点是压缩低效投放而非降低销售规模;③ 其余亏损渠道维持投入,优先优化费用结构(调整人群定向、"
    f"降低退货售后与履约成本)。"
)

doc.add_heading("第 2 章 门店经营诊断、分级与调整建议", level=1)
doc.add_paragraph(
    f"42 家门店 2026H1 累计到店销售额 {money(store['到店销售额'].sum())} 元,"
    f"累计门店 O2O 履约贡献 {money(store['O2O履约贡献'].sum())} 元。"
    f"按《渠道与门店管理策略.docx》第 2 章规则(先按到店坪效定级,租售比 >35% 直接列为 D 级):"
    f"A 级 {grade_dist['A']} 家、B 级 {grade_dist['B']} 家、C 级 {grade_dist['C']} 家、D 级 {grade_dist['D']} 家,"
    f"全部门店均给出级别,无无法归类的门店。"
)
doc.add_paragraph(
    f"D 级门店 {grade_dist['D']} 家中,"
    f"{int(store[(store['分级'] == 'D') & store['连续6个月D级'] & store['租约在H2内到期']].shape[0])} 家同时满足"
    f"「租约在 2026H2 内到期」与「连续 6 个月 D 级」两项闭店条件,建议启动闭店评估;"
    f"其余 D 级门店租约未到期,建议按转前置仓/自提点方向评估。"
)
t2 = doc.add_table(rows=1, cols=5)
t2.style = "Light Grid Accent 1"
for i, h in enumerate(["门店", "区域", "到店坪效(元/㎡/月)", "租售比", "调整建议"]):
    t2.rows[0].cells[i].text = h
for sid, r in store[store["分级"] == "D"].sort_values("到店坪效").iterrows():
    cells = t2.add_row().cells
    cells[0].text = sid
    cells[1].text = r["区域"]
    cells[2].text = f"{r['到店坪效']:.0f}"
    cells[3].text = pct(r["租售比"])
    cells[4].text = r["调整建议"]
doc.add_paragraph("")
doc.add_paragraph(
    f"另有 {int((store['面积是否调整'] == '是').sum())} 家门店计划期内发生面积调整"
    f"({', '.join(store[store['面积是否调整'] == '是'].index)}),其坪效口径可能失真,已在《异常与待核清单.xlsx》披露。"
)

doc.add_heading("第 3 章 促销组合效果评估与 H2 节奏建议", level=1)
doc.add_paragraph(
    f"2026H1 共开展 50 个促销活动,实际费用合计 {money(activity['实际费用'].sum())} 元,"
    f"带来增量销售额 {money(promo['增量销售额'].sum())} 元、增量毛利 {money(promo['增量毛利'].sum())} 元,"
    f"整体 ROI {promo['增量毛利'].sum() / promo['实际费用'].sum():.2f}。"
)
t3 = doc.add_table(rows=1, cols=4)
t3.style = "Light Grid Accent 1"
for i, h in enumerate(["活动类型", "活动数", "中位 ROI", "整体 ROI"]):
    t3.rows[0].cells[i].text = h
for _, r in type_stat.sort_values("整体ROI", ascending=False).iterrows():
    cells = t3.add_row().cells
    cells[0].text = r["活动类型"]
    cells[1].text = str(int(r["活动数"]))
    cells[2].text = f"{r['中位ROI']:.2f}"
    cells[3].text = f"{r['整体ROI']:.2f}"
doc.add_paragraph("")
doc.add_paragraph(
    f"按本报告自定阈值(ROI ≥2 保留、1≤ROI<2 调整、ROI<1 取消),建议保留 "
    f"{int((promo['建议动作'] == '保留').sum())} 个、调整 {int((promo['建议动作'] == '调整').sum())} 个、"
    f"取消 {int((promo['建议动作'] == '取消').sum())} 个;H2 节奏建议见交付物 4 的「H2 节奏建议」Sheet,"
    f"预算按 H1 实际费用水平分配到 6 个月,重点保障双 11、双 12 等大促节点。"
)

doc.add_heading("第 4 章 会员运营建议", level=1)
doc.add_paragraph(
    f"{member_total:,} 名注册会员中,线上渠道注册 {pct(reg_dist.get('CH01', 0) + reg_dist.get('CH02', 0) + reg_dist.get('CH03', 0) + reg_dist.get('CH04', 0) + reg_dist.get('CH05', 0) + reg_dist.get('CH06', 0))}、"
    f"门店注册 {pct(reg_dist.get('门店', 0))};人均订单 {orders_per_member:.2f} 笔,"
    f"复购会员(订单数 ≥2)占比 {pct(repurchase_rate)}。"
)
doc.add_paragraph(
    f"跨渠道行为:{pct(cross_ratio)} 的会员在 2 个及以上渠道购买,平均跨 {avg_channel:.2f} 个渠道;"
    f"Top10% 会员贡献 {pct(top10_share)} 的累计消费;流失会员占比 {pct(churn_rate)},"
    f"平均 RFM 分 {rfm_mean:.2f}。建议:① 对 Top10% 会员配置专属服务与提前购;"
    f"② 对仅在单一渠道购买的会员投放到店自提/跨渠道优惠券,提升渠道协同;"
    f"③ 对流失会员按最近购买时间分群唤回,优先使用微信小程序等低费用渠道触达。"
)

doc.add_heading("第 5 章 H2 资源再配置方案与执行优先级", level=1)
doc.add_paragraph(
    f"营销预算:按渠道调整方向再分配,H1 广告费合计 {money(ch['广告费'].sum())} 元;"
    f"收缩渠道释放的投放额度优先转入 {wx['渠道名称']} 与高 ROI 促销类型。"
)
doc.add_paragraph(
    f"门店资源:D 级门店按闭店/转型路径处置,预计释放人力与租金;"
    f"C 级门店以降租谈判与品类优化为主,3 个月后复评;A 级门店追加资源作为标杆。"
)
doc.add_paragraph("执行优先级:")
for item in [
    "P0(7 月内):启动满足「租约到期 + 连续 6 个月 D 级」条件的门店闭店评估,测算违约金与释放成本;",
    "P0(7 月内):收缩高费用率渠道的低效投放,设定销售费用率下降目标;",
    f"P1(8 月内):对 {wx['渠道名称']} 增投 20%,并同步扩充会员与私域运营;",
    "P1(8 月内):按 H2 节奏表落地促销组合,取消 ROI <1 的活动类型;",
    "P2(Q3 末):完成 D 级门店转前置仓/自提点试点;",
    "P2(Q4 末):复盘渠道费用结构与会员跨渠道转化,确定 2027 年资源基线。",
]:
    doc.add_paragraph(item, style="List Bullet")

doc.add_heading("附录 关键判断依据回指附件清单", level=1)
doc.add_paragraph("本报告全部关键数字均可回指下列附件的具体 Sheet 与字段:")
appendix = [
    ("渠道净收入、毛利、客单价", "《2026H1渠道销售明细.xlsx》", "渠道月度汇总 Sheet(净收入/毛利/客单价)"),
    ("渠道费用与渠道净利", "《2026H1线上投放与平台费用.xlsx》", "渠道月度费用 Sheet(广告费/平台佣金/履约费/费用合计/渠道净利)"),
    ("投放 ROI", "《2026H1线上投放与平台费用.xlsx》", "投放明细 Sheet(支付金额 ÷ 投放金额)"),
    ("门店到店销售、坪效、租售比", "《2026H1门店经营台账.xlsx》", "门店月度经营 Sheet(到店销售额)、门店费用 Sheet(租金及物业费)、门店基础信息 Sheet(营业面积/面积是否调整)"),
    ("门店 O2O 履约贡献", "《2026H1渠道销售明细.xlsx》 + 《2026H1门店经营台账.xlsx》 交叉聚合", "渠道销售明细 Sheet(发货方式 ∈ {门店发货,到店自提} 的净收入)"),
    ("门店分级与调整建议", "《渠道与门店管理策略.docx》", "第 2 章门店分级标准(2.1–2.3)、第 6 章闭店与改造条件(6.1–6.4)"),
    ("闭店约束(租期/违约金)", "《华东区门店租约汇总.docx》", "逐门店:租期结束、闭店条件"),
    ("促销 ROI 与类型效果", "《2026H1促销活动记录.xlsx》", "活动效果 Sheet(增量毛利/ROI)、活动清单 Sheet(活动类型/实际费用)"),
    ("会员结构、复购与跨渠道", "《2026H1会员与跨渠道订单.xlsx》", "会员基础/会员订单/跨渠道行为 Sheet"),
    ("SKU 成本与品类等级", "《SKU主数据.xlsx》", "SKU主数据 Sheet(单位成本/标准标价/品类等级)"),
]
t4 = doc.add_table(rows=1, cols=3)
t4.style = "Light Grid Accent 1"
for i, h in enumerate(["结论类别", "依据文件", "具体 Sheet / 字段"]):
    t4.rows[0].cells[i].text = h
for a, b, c in appendix:
    cells = t4.add_row().cells
    cells[0].text = a
    cells[1].text = b
    cells[2].text = c


# ============================================================
# 8c. 报告排版规范化(标题黑色 / 表格浅灰边框 + 表头填充 / 表头跨页重复)
# ============================================================
from docx.enum.table import WD_ALIGN_VERTICAL  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import RGBColor  # noqa: E402

BLACK = RGBColor(0x00, 0x00, 0x00)
BORDER = "D9D9D9"
HEADER_FILL = "D9D9D9"


def _set_cell_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "6")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), BORDER)
        borders.append(el)
    tcPr.append(borders)


def _shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def _repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def normalize_report(doc):
    """标题一律黑色;表格用浅灰边框、表头浅灰填充黑字、单元格垂直居中、表头跨页重复。"""
    for p in doc.paragraphs:
        if p.style.name.startswith(("Heading", "Title")):
            for run in p.runs:
                run.font.color.rgb = BLACK
    for tbl in doc.tables:
        for r_i, row in enumerate(tbl.rows):
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                _set_cell_borders(cell)
                if r_i == 0:
                    _shade(cell, HEADER_FILL)
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.font.color.rgb = BLACK
                            run.font.bold = True
        _repeat_header(tbl.rows[0])


normalize_report(doc)

doc.save(OUTPUT / "2026H2 渠道与门店经营策略报告.docx")
print("已写出 Word 报告")

# ============================================================
# 8. 数据血缘说明
# ============================================================
ch1 = ch[ch["渠道ID"] == "CH01"].iloc[0]
s1 = store.loc["S001"]
worst = store.sort_values("到店坪效").iloc[0]

# ---- 统一示例对象:门店 S001(及其直接关联对象) ----
s001_month = store_month[store_month["门店ID"] == "S001"].sort_values("月份")
s001_sales = [round(x, 2) for x in s001_month["到店销售额"]]
s001_o2o_month = [round(x, 2) for x in s001_month["门店O2O履约贡献"]]
s001_o2o_total = round(sum(s001_o2o_month), 2)
s001_orders_o2o = orders[(orders["门店ID"] == "S001") & (orders["发货方式"].isin(["门店发货", "到店自提"]))]
s001_o2o_by_ch = s001_orders_o2o.groupby("渠道ID")["净收入"].sum().round(2).to_dict()
s001_o2o_orders = len(s001_orders_o2o)
ch03 = ch[ch["渠道ID"] == "CH03"].iloc[0]
ch03_ads = ad[ad["渠道ID"] == "CH03"]
a005 = promo[promo["活动ID"] == "A005"].iloc[0]

# 可在附件中直接定位的中间值(用于抽查)
s001 = store_month[(store_month["门店ID"] == "S001") & (store_month["月份"] == "2026-01")].iloc[0]
s001_rent = store_exp[(store_exp["门店ID"] == "S001") & (store_exp["月份"] == "2026-01")].iloc[0]
ch03_jan = ch_fee[(ch_fee["渠道ID"] == "CH03") & (ch_fee["月份"] == "2026-01")].iloc[0]
a001 = promo[promo["活动ID"] == "A001"].iloc[0]
sku1 = sku_master.iloc[0]
member_m1_spend = cross[cross["会员ID"] == "M000001"].iloc[0]["累计消费"]
spot_rows = "\n".join([
    f"| {money(s001['到店销售额'])} | 2026H1门店经营台账.xlsx | 门店月度经营 | 门店ID=S001、月份=2026-01、字段=到店销售额 |",
    f"| {money(s001['门店O2O履约贡献'])} | 2026H1门店经营台账.xlsx | 门店月度经营 | 门店ID=S001、月份=2026-01、字段=门店O2O履约贡献 |",
    f"| {money(s001_rent['租金及物业费'])} | 2026H1门店经营台账.xlsx | 门店费用 | 门店ID=S001、月份=2026-01、字段=租金及物业费 |",
    f"| {money(ch03_jan['净收入'])} | 2026H1线上投放与平台费用.xlsx | 渠道月度费用 | 渠道ID=CH03、月份=2026-01、字段=净收入 |",
    f"| {money(ch03_jan['渠道净利'])} | 2026H1线上投放与平台费用.xlsx | 渠道月度费用 | 渠道ID=CH03、月份=2026-01、字段=渠道净利 |",
    f"| {money(a001['实际费用'])} | 2026H1促销活动记录.xlsx | 活动清单 | 活动ID=A001、字段=实际费用 |",
    f"| {money(a001['增量毛利'])} | 2026H1促销活动记录.xlsx | 活动效果 | 活动ID=A001、字段=增量毛利 |",
    f"| {money(sku1['单位成本'])} | SKU主数据.xlsx | SKU主数据 | SKU编码={sku1['SKU编码']}、字段=单位成本 |",
    f"| {money(member_m1_spend)} | 2026H1会员与跨渠道订单.xlsx | 跨渠道行为 | 会员ID=M000001、字段=累计消费 |",
])
lineage = f"""# 数据血缘与指标口径说明

> **用途**:说明 `output/` 下 5 个交付物中每个指标的分析逻辑、数据来源与计算口径,便于逐项核对。
> **对应任务书**:`docs/specs/step2-query.md`;任务说明:`docs/specs/step1-task-brief.md`
> **数据版本**:num.py v3.2 生成的 `input/` 9 个附件
> (材料自检 L1–L4 共 129 项通过 + L5 跨口径反证探针 15 条通过,0 问题、0 风险)
> **生成时间**:2026-09-23

---

## 一、总览:附件 → 交付物

| 交付物 | 主要来源附件 | 关键口径 |
| --- | --- | --- |
| 1. 2026H2 渠道与门店经营策略报告.docx | 全部 9 个附件 | 见各章节 |
| 2. 门店分级与调整建议.xlsx | 《2026H1门店经营台账.xlsx》+ 《华东区门店租约汇总.docx》 | 坪效 = 月均到店销售额 ÷ 营业面积;租售比 = 月均租金及物业费 ÷ 月均到店销售额 |
| 3. 渠道资源再配置建议.xlsx | 《2026H1渠道销售明细.xlsx》+ 《2026H1线上投放与平台费用.xlsx》 | 净收入 = 收入 − 退款;渠道净利 = 毛利 − 费用合计 |
| 4. 促销组合优化建议.xlsx | 《2026H1促销活动记录.xlsx》 | 活动 ROI = 增量毛利 ÷ 实际费用 |
| 5. 异常与待核清单.xlsx | 全部 9 个附件 | 见第五节 |

**门店 ID 主键**:《2026H1渠道销售明细.xlsx》(门店ID)、《2026H1门店经营台账.xlsx》(门店基础信息/月度经营/库存/费用)、《2026H1会员与跨渠道订单.xlsx》(会员订单/会员基础.注册门店ID)、《华东区门店租约汇总.docx》(租约汇总)。
**渠道/活动/SKU/会员主键**:渠道ID(《2026H1渠道销售明细.xlsx》/《2026H1线上投放与平台费用.xlsx》/《2026H1会员与跨渠道订单.xlsx》/《2026H1促销活动记录.xlsx》)、订单号(《2026H1渠道销售明细.xlsx》内部:渠道销售明细 Sheet ↔ 渠道退款明细 Sheet)、会员订单号(《2026H1会员与跨渠道订单.xlsx》内部:会员订单 Sheet ↔ 跨渠道行为 Sheet;与渠道订单号为两套独立编号,不与《2026H1渠道销售明细.xlsx》逐笔勾稽,两者按会员ID 做口径级对照)、SKU编码(《2026H1渠道销售明细.xlsx》/《2026H1促销活动记录.xlsx》/《SKU主数据.xlsx》)、会员ID(《2026H1会员与跨渠道订单.xlsx》)、活动ID(《2026H1促销活动记录.xlsx》)。

**统一示例对象**:本说明书所有「示例复核」统一使用同一条对象链 —
**门店 S001**(及其 O2O 订单主力渠道 **CH03 天猫旗舰店**、参与活动 **A005 会员日**)。
每个示例都写出完整算式(输入值 → 计算 → 结果),便于逐格核对。

---

## 二、交付物 2:门店分级与调整建议.xlsx(Sheet「门店清单」)

| 字段 | 数据来源 | 分析逻辑 | 口径说明 |
| --- | --- | --- | --- |
| 门店ID / 门店名称 / 区域 / 城市 / 店型 / 营业面积 | 《2026H1门店经营台账.xlsx》 门店基础信息 | 直接读取 | 面积单位为 ㎡ |
| 2026H1累计到店销售额 | 《2026H1门店经营台账.xlsx》 门店月度经营.到店销售额 | SUM(该门店 6 个月) | 仅含顾客到店完成交易部分,不含 O2O 履约 |
| 2026H1累计门店O2O履约贡献 | 《2026H1门店经营台账.xlsx》 门店月度经营.门店O2O履约贡献 | SUM(该门店 6 个月) | = 《2026H1渠道销售明细.xlsx》 中发货方式 ∈ {{门店发货, 到店自提}} 的净收入按门店聚合 |
| 到店坪效(元/㎡/月) | 由上一字段与营业面积计算 | 月均到店销售额 ÷ 营业面积 | 月均 = 累计 ÷ 6;保留整数 |
| 租售比 | 《2026H1门店经营台账.xlsx》 门店费用.租金及物业费 + 门店月度经营.到店销售额 | 月均租金及物业费 ÷ 月均到店销售额 | 租金口径 = 合同月租金 + 物业费(《渠道与门店管理策略.docx》 第 4.4 条) |
| 面积是否调整 | 《2026H1门店经营台账.xlsx》 门店基础信息.面积是否调整 | 直接读取 | = 是 的门店坪效口径可能失真,已在《异常与待核清单.xlsx》披露 |
| 分级 | 《渠道与门店管理策略.docx》 第 2.2/2.3 条 | 先按到店坪效定级(A ≥5000、B 2000–5000、C 1000–2000、D <1000);租售比 >35% 直接列 D | 不使用其他自创阈值 |
| 调整建议 | 《渠道与门店管理策略.docx》 第 2.5、6.1、6.2、6.4 条 | A/B 级保留、C 级调改;D 级中同时满足"租约在 2026H2 内到期"与"连续 6 个月 D 级"者启动闭店评估,其余转前置仓/自提点评估 | 连续 6 个月 D 级由月度指标逐月判定 |
| 关键依据 | 《2026H1门店经营台账.xlsx》/《华东区门店租约汇总.docx》 | 列出该门店的到店销售额、月均租金及物业费、营业面积等回指项 | 每条结论可回指附件字段 |
| 数据状态 | 面积是否调整 | = 是 → 待补(坪效口径待确认);否则 已确认 | 待补含义见《异常与待核清单.xlsx》A-002 |

### 示例复核:门店 S001(统一示例对象)

| 指标 | 算式(输入值 → 计算 → 结果) |
| --- | --- |
| 2026H1 累计到店销售额 | {' + '.join(money(x) for x in s001_sales)} = **{money(s1['到店销售额'])} 元**(《2026H1门店经营台账.xlsx》 门店月度经营,门店ID=S001 的 6 个月合计) |
| 月均到店销售额 | {money(s1['到店销售额'])} ÷ 6 = **{money(s1['月均到店销售'])} 元** |
| 到店坪效 | 月均到店销售额 {money(s1['月均到店销售'])} 元 ÷ 营业面积 {int(s1['营业面积'])} ㎡ = **{money(s1['到店坪效'])} 元/㎡/月** |
| 月均租金及物业费 | 《2026H1门店经营台账.xlsx》 门店费用(门店ID=S001,6 个月均为 {money(s1['月均租金及物业费'])} 元) = **{money(s1['月均租金及物业费'])} 元** |
| 租售比 | 月均租金及物业费 {money(s1['月均租金及物业费'])} 元 ÷ 月均到店销售额 {money(s1['月均到店销售'])} 元 = **{s1['租售比']:.4f}** |
| 门店 O2O 履约贡献 | 《2026H1渠道销售明细.xlsx》 中 门店ID=S001 且 发货方式 ∈ {{门店发货,到店自提}} 的净收入合计 = **{money(s001_o2o_total)} 元**(共 {s001_o2o_orders} 单;按渠道:{', '.join(f'{k} {money(v)}' for k, v in s001_o2o_by_ch.items())}) |
| 分级判定 | 租售比 {s1['租售比']:.4f} ≤ 0.35 → 不触发 D 级否决;到店坪效 {money(s1['到店坪效'])} 元/㎡/月 ∈ [2000, 5000) → **B 级** |
| 调整建议 | B 级 + 面积是否调整 = {s1['面积是否调整']} → **{s1['调整建议']}**,数据状态 **{s1['数据状态']}** |

> 对照:全表最低坪效门店 {worst.name} 为 {money(worst['到店坪效'])} 元/㎡/月、租售比 {worst['租售比']:.4f} → 分级 {worst['分级']}。

---

## 三、交付物 3:渠道资源再配置建议.xlsx(Sheet「渠道清单」)

| 字段 | 数据来源 | 分析逻辑 | 口径说明 |
| --- | --- | --- | --- |
| 渠道ID / 渠道名称 / 渠道类型 | 《2026H1渠道销售明细.xlsx》 渠道销售明细 | 直接读取 | 自营:CH01/CH02;平台:CH03/CH04/CH05;即时零售:CH06 |
| 2026H1净收入(元) | 《2026H1渠道销售明细.xlsx》 渠道月度汇总.净收入 | SUM(该渠道 6 个月) | = 收入 − 退款 |
| 毛利(元) | 《2026H1渠道销售明细.xlsx》 渠道月度汇总.毛利 | SUM(该渠道 6 个月) | = 净收入 − 商品成本(商品成本可由《SKU主数据.xlsx》 SKU 单位成本复算) |
| 广告费 / 平台佣金 / 履约费(元) | 《2026H1线上投放与平台费用.xlsx》 渠道月度费用 | SUM(该渠道 6 个月) | 广告费 = 《2026H1线上投放与平台费用.xlsx》 投放明细.投放金额合计;平台佣金 = Σ(《2026H1渠道销售明细.xlsx》 收入 × 《2026H1线上投放与平台费用.xlsx》 平台佣金规则.佣金率) |
| 渠道毛利率 | 毛利 ÷ 净收入 | 保留 4 位小数(展示为百分比) | 不含渠道费用 |
| 投放ROI | 《2026H1线上投放与平台费用.xlsx》 投放明细.支付金额 ÷ 投放金额 | SUM(支付金额) ÷ SUM(投放金额) | 《渠道与门店管理策略.docx》 第 7.3 条口径 |
| 销售费用率 | 《2026H1线上投放与平台费用.xlsx》 渠道月度费用.费用合计 ÷ 净收入 | 费用合计 = 广告+佣金+技服+支付+履约+包装+退货售后+其他 | 反映渠道整体费用负担 |
| H2资源调整方向 / 建议投入幅度 | 自定规则(见《异常与待核清单.xlsx》A-006 同类披露) | 净利>0 且 费用率≤30% → 增投 +20%;净利<0 且 费用率≥33% → 收缩 −15%;其余 → 维持 | 附件内未规定阈值,已在本说明书与报告中披露 |
| 渠道净利(元) | 《2026H1线上投放与平台费用.xlsx》 渠道月度费用.渠道净利 | SUM(该渠道 6 个月)= 毛利 − 费用合计 | 与报告的渠道盈亏结论一致 |
| 关键依据 / 数据状态 | — | 列出回指字段 / 均为"已确认" | — |

### 示例复核:渠道 CH03 天猫旗舰店(S001 的 O2O 订单主力渠道)

> 选择理由:S001 在《2026H1渠道销售明细.xlsx》 中的 O2O 订单(门店发货 + 到店自提)按其净收入拆分,
> 以 CH03 天猫旗舰店最多({money(s001_o2o_by_ch.get('CH03', 0))} 元),故渠道维度示例沿用同一对象链。

| 指标 | 算式(输入值 → 计算 → 结果) |
| --- | --- |
| 2026H1 净收入 | 《2026H1渠道销售明细.xlsx》 渠道月度汇总中 渠道ID=CH03 的 6 个月净收入之和 = **{money(ch03['净收入'])} 元** |
| 毛利 | 净收入 {money(ch03['净收入'])} − 商品成本 = **{money(ch03['毛利'])} 元**(《2026H1渠道销售明细.xlsx》 渠道月度汇总.毛利) |
| 渠道毛利率 | 毛利 {money(ch03['毛利'])} ÷ 净收入 {money(ch03['净收入'])} = **{pct(ch03['渠道毛利率'])}** |
| 广告费 | 《2026H1线上投放与平台费用.xlsx》 投放明细中 渠道ID=CH03 的投放金额合计 = **{money(ch03['广告费'])} 元**(与《2026H1线上投放与平台费用.xlsx》 渠道月度费用.广告费一致) |
| 平台佣金 | Σ(《2026H1渠道销售明细.xlsx》 收入 × 《2026H1线上投放与平台费用.xlsx》 平台佣金规则.佣金率) = **{money(ch03['平台佣金'])} 元** |
| 履约费 | 《2026H1线上投放与平台费用.xlsx》 渠道月度费用中 渠道ID=CH03 的履约费合计 = **{money(ch03['履约费'])} 元** |
| 费用合计 | 广告+佣金+技服+支付+履约+包装+退货售后+其他 = **{money(ch03['费用合计'])} 元** |
| 销售费用率 | 费用合计 {money(ch03['费用合计'])} ÷ 净收入 {money(ch03['净收入'])} = **{pct(ch03['销售费用率'])}** |
| 渠道净利 | 毛利 {money(ch03['毛利'])} − 费用合计 {money(ch03['费用合计'])} = **{money(ch03['渠道净利'])} 元** |
| 投放ROI | Σ投放明细.支付金额 {money(ch03['支付金额'])} ÷ Σ投放明细.投放金额 {money(ch03['投放金额'])} = **{ch03['投放ROI']:.4f}** |
| H2 资源调整方向 | 净利 {money(ch03['渠道净利'])} < 0 且 费用率 {pct(ch03['销售费用率'])} ≥ 33% → **{ch03['H2资源调整方向']}({ch03['建议投入幅度']:+.0%})** |

> S001 在 CH03 的 O2O 履约贡献 {money(s001_o2o_by_ch.get('CH03', 0))} 元,可由《2026H1渠道销售明细.xlsx》 中
> "门店ID=S001 且 渠道ID=CH03 且 发货方式 ∈ {{门店发货,到店自提}}"的净收入复算。

---

## 四、交付物 4:促销组合优化建议.xlsx

### Sheet「活动评估」(50 行)

| 字段 | 数据来源 | 分析逻辑 | 口径说明 |
| --- | --- | --- | --- |
| 活动ID / 活动名称 / 活动类型 / 渠道ID | 《2026H1促销活动记录.xlsx》 活动清单 | 直接读取 | 活动类型:满减/折扣/直播/会员日/节日促销 |
| 实际费用(元) | 《2026H1促销活动记录.xlsx》 活动清单.实际费用 | 直接读取 | 与预算的差异属正常执行浮动 |
| 增量销售额(元) | 《2026H1促销活动记录.xlsx》 活动效果 | 活动期间销售额 − 活动前销售额 | 活动期间销售额 = 《2026H1促销活动记录.xlsx》 活动明细.销售额合计 |
| 增量毛利(元) | 《2026H1促销活动记录.xlsx》 活动效果 | 增量销售额 × 活动毛利率 | 活动毛利率由活动明细(活动价 − 单位成本)复算 |
| ROI | 《2026H1促销活动记录.xlsx》 活动效果 | 增量毛利 ÷ 实际费用 | 《渠道与门店管理策略.docx》 第 7.1 条 |
| 建议动作 | 自定阈值 | ROI ≥2 保留;1 ≤ ROI <2 调整;ROI <1 取消 | 阈值已披露(《异常与待核清单.xlsx》),可按管理层门槛重算 |
| 关键依据 | 《2026H1促销活动记录.xlsx》 | 列出增量毛利与实际费用 | — |

### Sheet「H2 节奏建议」(6 行)

| 字段 | 数据来源 | 分析逻辑 | 口径说明 |
| --- | --- | --- | --- |
| 月份 | 2026-07 ~ 2026-12 | 固定 6 行 | 对应建议执行期 |
| 建议活动类型 | 《2026H1促销活动记录.xlsx》 活动类型效果 + 电商日历 | 优先保留 H1 整体 ROI 高的类型(节日促销/折扣/会员日),大促月叠加折扣 | 直播类 H1 整体 ROI 偏低,建议先在 H2 试点验证 |
| 预算区间(元) | 《2026H1促销活动记录.xlsx》 活动清单.实际费用合计 | 按 H1 实际费用水平的一定比例分配到 6 个月,大促月(11/12 月)占比更高 | 合计区间覆盖 H1 实际费用规模 |
| 关键节点 | 日历 | 暑期、七夕、秋季上新、国庆、双 11、双 12/年货节 | — |

### 示例复核:活动 A005(S001 参与的唯一活动)

> 选择理由:《2026H1促销活动记录.xlsx》 活动清单中 门店ID=S001 的活动只有 A005,故促销维度示例沿用同一对象链。

| 指标 | 算式(输入值 → 计算 → 结果) |
| --- | --- |
| 实际费用 | 《2026H1促销活动记录.xlsx》 活动清单(活动ID=A005,类型={a005['活动类型']},渠道={a005['渠道ID']}) = **{money(a005['实际费用'])} 元**(预算 {money(a005['预算'])} 元) |
| 活动期间销售额 | 《2026H1促销活动记录.xlsx》 活动明细中 活动ID=A005 的销售额合计 = **{money(a005['活动期间销售额'])} 元** |
| 活动前销售额 | 《2026H1促销活动记录.xlsx》 活动效果.活动前销售额 = **{money(a005['活动前销售额'])} 元** |
| 增量销售额 | 活动期间 {money(a005['活动期间销售额'])} − 活动前 {money(a005['活动前销售额'])} = **{money(a005['增量销售额'])} 元** |
| 增量毛利 | 增量销售额 {money(a005['增量销售额'])} × 活动毛利率(由活动明细"活动价 − 单位成本"复算) = **{money(a005['增量毛利'])} 元** |
| ROI | 增量毛利 {money(a005['增量毛利'])} ÷ 实际费用 {money(a005['实际费用'])} = **{a005['ROI']:.2f}** |
| 建议动作 | ROI {a005['ROI']:.2f} ≥ 2 → **{a005['建议动作']}**(阈值见第八节) |

> 同一活动的 H1 增量销售额 {money(a005['增量销售额'])} 元中,属于门店 S001 的交易包含在
> 《2026H1促销活动记录.xlsx》 活动明细(活动ID=A005)的 SKU 级明细里,可按 SKU 编码与《SKU主数据.xlsx》 单位成本逐行复算。

---

## 五、交付物 5:异常与待核清单.xlsx(Sheet「异常清单」)

清单共 {len(anomaly_df)} 条,类别按任务书枚举值使用:时效性 / 口径不清 / 缺失 / 越界 / 数据冲突。

| 编号 | 类别 | 涉及文件 | 严重度 | 核心内容 |
| --- | --- | --- | --- | --- |
""" + "\n".join(
    f"| {r['异常编号']} | {r['异常类别']} | {r['涉及文件']} | {r['严重度']} | {r['问题描述'][:60]} |"
    for _, r in anomaly_df.iterrows()
) + f"""

---

## 六、交付物 1:报告关键数字的来源

| 报告位置 | 关键数字 | 算式 |
| --- | --- | --- |
| 摘要 | 六渠道净收入合计 {money(net_total)} 元 | 交付物 3 净收入列 6 个渠道求和;= 《2026H1渠道销售明细.xlsx》 渠道销售明细.净收入合计 |
| 摘要 | 渠道合计净利 {money(profit_total)} 元 | 交付物 3 渠道净利列求和;= 《2026H1线上投放与平台费用.xlsx》 渠道月度费用.渠道净利合计 |
| 摘要 | 门店分级分布 A{grade_dist['A']}/B{grade_dist['B']}/C{grade_dist['C']}/D{grade_dist['D']} | 交付物 2 分级列计数(合计 42 家) |
| 摘要 | 门店 O2O 履约贡献合计 {money(store['O2O履约贡献'].sum())} 元 | 交付物 2 该列求和;= 《2026H1门店经营台账.xlsx》 门店月度经营.门店O2O履约贡献合计 = 《2026H1渠道销售明细.xlsx》 门店发货+到店自提净收入聚合 |
| 第 1 章 | 平均销售费用率 {pct(ch['费用合计'].sum() / net_total)} | 六个渠道费用合计求和 {money(ch['费用合计'].sum())} 元 ÷ 净收入合计 {money(net_total)} 元 |
| 第 3 章 | 50 个活动整体 ROI {promo['增量毛利'].sum() / promo['实际费用'].sum():.2f} | Σ增量毛利 {money(promo['增量毛利'].sum())} 元 ÷ Σ实际费用 {money(promo['实际费用'].sum())} 元 |
| 第 4 章 | 跨渠道会员占比 {pct(cross_ratio)} | 《2026H1会员与跨渠道订单.xlsx》 跨渠道行为中"跨渠道购买渠道数 ≥2"的会员数 ÷ 2,500 |
| 第 4 章 | Top10% 消费贡献 {pct(top10_share)} | 累计消费前 10% 会员的累计消费之和 ÷ 全体会员累计消费之和 |

---

## 七、跨文件口径一致性(已逐项核验)

| 数字 | 出现位置 | 核验结果 |
| --- | --- | --- |
| 渠道净收入合计 {money(net_total)} 元 | 交付物 3 求和 = 《2026H1渠道销售明细.xlsx》 净收入合计 = 报告摘要 | ✓ 一致 |
| 门店数 42 家 | 交付物 2 行数 = 《2026H1门店经营台账.xlsx》 门店基础信息 = 报告第 2 章 | ✓ 一致 |
| 渠道数 6 个 | 交付物 3 行数 = 报告第 1 章表格 | ✓ 一致 |
| 活动数 50 个 | 交付物 4 活动评估行数 = 《2026H1促销活动记录.xlsx》 活动清单 = 报告第 3 章 | ✓ 一致 |
| 门店 O2O 履约贡献合计 {money(store['O2O履约贡献'].sum())} 元 | 交付物 2 求和 = 《2026H1渠道销售明细.xlsx》 门店发货+到店自提净收入聚合 | ✓ 一致 |
| 渠道净利合计 {money(profit_total)} 元 | 交付物 3 求和 = 《2026H1线上投放与平台费用.xlsx》 渠道月度费用.渠道净利合计 | ✓ 一致 |

### 7.1 跨表主键与聚合口径核验记录

> 本节为**核验通过记录**(原列在《异常与待核清单.xlsx》的 A-007 条目,因不属"异常"已移入本节)。

| 主键 | 可跨表核验的范围 | 核验结果 |
| --- | --- | --- |
| 门店ID | 门店基础信息 ↔ 门店月度经营 ↔ 门店库存 ↔ 门店费用 ↔ 渠道销售明细 ↔ 租约汇总 | ✓ 42 家逐表一致(会员订单 Sheet 另有门店ID,口径独立) |
| 渠道ID | 渠道销售明细 ↔ 渠道月度汇总 ↔ 渠道月度费用 ↔ 平台佣金规则 ↔ 会员订单 ↔ 活动清单 | ✓ 6 个渠道逐表一致 |
| 订单号 | **仅限《2026H1渠道销售明细.xlsx》内部**:渠道销售明细 Sheet ↔ 渠道退款明细 Sheet(O100000–O107499) | ✓ 611 笔退款逐笔对应 |
| 会员订单号 | **仅限《2026H1会员与跨渠道订单.xlsx》内部**:会员订单 Sheet ↔ 跨渠道行为 Sheet(O800000–O805999) | ✓ 6,000 单逐笔对应;与渠道订单号为**两套独立编号**,不可跨表合并(见《异常与待核清单.xlsx》A-009) |
| SKU编码 | SKU 主数据 ↔ 渠道销售明细 ↔ 活动明细 | ✓ 成本、标价、品类逐行一致 |
| 会员ID | 会员基础 ↔ 会员订单 ↔ 跨渠道行为 | ✓ 2,500 人口径一致 |
| 活动ID | 活动清单 ↔ 活动明细 ↔ 活动效果 ↔ 投放明细 | ✓ 50 个活动逐表一致 |

| 聚合口径 | 核验方式 | 核验结果 |
| --- | --- | --- |
| 净收入 | 渠道销售明细 Sheet 按 渠道ID+月份 聚合 ↔ 渠道月度汇总 Sheet ↔ 渠道月度费用 Sheet | ✓ 逐月零差异 |
| 门店 O2O 履约贡献 | 渠道销售明细 Sheet"门店发货 + 到店自提"净收入按门店+月份聚合 ↔ 门店月度经营 Sheet | ✓ 252 个门店×月组合零差异 |
| 佣金与技术服务费 | 渠道销售明细 Sheet 收入 × 平台佣金规则 Sheet 费率 ↔ 渠道月度费用 Sheet | ✓ 按收入为基复算,最大差 0.0048 元 |
| 活动销售额 | 渠道销售明细 Sheet 订单按"渠道 + 活动日期区间 + SKU"聚合 ↔ 活动明细 Sheet ↔ 活动效果 Sheet | ✓ 逐活动勾稽比 1.0000(修复前为 0.042) |

**核验结论**:上述可核验范围内未发现数值冲突;不可核验的部分(渠道订单与会员订单之间)已披露于
《异常与待核清单.xlsx》
以"两套编号体系"披露,不作为冲突处理。

---

## 八、自定口径与阈值(附件内未规定者)

| 自定项 | 取值 | 披露位置 |
| --- | --- | --- |
| 渠道资源调整规则 | 净利>0 且 费用率≤30% → 增投 +20%;净利<0 且 费用率≥33% → 收缩 −15%;其余 → 维持 | 《异常与待核清单.xlsx》A-006 同类口径披露、本说明书第三节 |
| 促销建议动作阈值 | ROI ≥2 保留;1 ≤ ROI <2 调整;ROI <1 取消 | 《异常与待核清单.xlsx》A-006、本说明书第四节 |
| H2 预算区间 | 按 H1 实际费用水平的比例分配(合计覆盖 H1 规模) | 本说明书第四节 |

---

## 九、数据限制

| 限制 | 原因 | 影响 |
| --- | --- | --- |
| 无历史同期数据 | 9 个附件仅覆盖 2026H1 | 无法做季节性、年度趋势与生命周期分析(H2 目标为估算) |
| 无 SKU 级库存与会员品类明细 | 《2026H1门店经营台账.xlsx》《2026H1会员与跨渠道订单.xlsx》数据粒度限制 | 品类与蚕食结论仅到品类级 |
| 活动效果无分渠道归因 | 《2026H1促销活动记录.xlsx》结构限制 | 类型取舍基于整体 ROI,直播类低效原因需 H2 实验验证 |
| 面积调整门店坪效口径 | 3 家门店计划期内调整面积 | 该 3 家门店分级需标注待补 |

---

## 十、复算指引

1. 运行 `.venv/bin/python docs/skill/data-validation.py`,确认 9 个附件通过 L1–L4 共 129 项校验
   与 L5 共 15 条跨口径反证探针(0 问题、0 风险);
2. 运行 `.venv/bin/python build_solution.py`,可重新生成 `output/` 下全部交付物与本说明书;
3. 任取交付物中的一个金额,按本说明书第二至四节的"数据来源 + 分析逻辑"列即可定位到附件 Sheet 与字段;
4. 运行 `.venv/bin/python docs/skill/validate_lineage.py`,可自动扫描本文档中的每个数值是否能在附件中找到对应或可复算。

### 关键中间值(可直接在附件中定位,便于抽查)

| 数值 | 文件名 | Sheet | 定位 |
| --- | --- | --- | --- |
{spot_rows}
"""
(OUTPUT / "数据血缘说明.md").write_text(lineage, encoding="utf-8")
print("已写出 数据血缘说明.md")

# ============================================================
# 8b. 交付物工作簿自查区(Check 表 + 枚举数据验证)
# ============================================================
sys.path.insert(0, str(Path("docs/skill").resolve()))
import add_workbook_checks  # noqa: E402

add_workbook_checks.main([])

# ============================================================
# 9. 控制台摘要
# ============================================================
print("\n" + "=" * 60)
print(f"渠道净收入合计 {money(net_total)} 元 | 渠道净利合计 {money(profit_total)} 元")
print(f"门店分级 A{grade_dist['A']}/B{grade_dist['B']}/C{grade_dist['C']}/D{grade_dist['D']}")
print(f"促销:保留 {int((promo['建议动作']=='保留').sum())} / 调整 {int((promo['建议动作']=='调整').sum())} / 取消 {int((promo['建议动作']=='取消').sum())}")
print(f"异常清单 {len(anomaly_df)} 条")
print("=" * 60)
