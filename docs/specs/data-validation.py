"""
Magnet 出题附件数据自检脚本(v2.0)
==================================
用途:每次修改 num.py 后,运行本脚本验证 9 个附件的数据完整性与跨文件一致性。

使用:
    .venv/bin/python docs/specs/data-validation.py                 # 校验 input/
    .venv/bin/python docs/specs/data-validation.py <目录>           # 校验指定目录(试跑用)

检查分四层:
    L1 文件与字段完整性
    L2 单文件公式自洽(内部可复算)
    L3 跨文件勾稽(主键、聚合口径、规则复算)
    L4 业务合理性(相关性、量级、分级可判定性)

输出:通过项 ✓、问题项 ✗、最后给出汇总;存在问题则以退出码 1 结束。
"""
import os
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

WORKSPACE = Path(__file__).resolve().parent.parent.parent
INPUT_DIR = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
    os.environ.get("MAGNET_INPUT_DIR", WORKSPACE / "input")
)

EXPECTED_FILES = [
    "2026H1渠道销售明细.xlsx",
    "2026H1门店经营台账.xlsx",
    "2026H1线上投放与平台费用.xlsx",
    "2026H1会员与跨渠道订单.xlsx",
    "2026H1促销活动记录.xlsx",
    "华东区门店租约汇总.docx",
    "渠道与门店管理策略.docx",
    "SKU主数据.xlsx",
    "旧版门店清单_2024.xlsx",
]

issues = []
warnings = []


def check(cond, msg_ok, msg_fail):
    if cond:
        print(msg_ok)
    else:
        issues.append(msg_fail)
        print(msg_fail)
    return bool(cond)


def section(title):
    print(f"\n{'=' * 78}\n【{title}】\n{'=' * 78}")


def read_docx_text(path):
    from docx import Document
    doc = Document(str(path))
    parts = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for r in t.rows:
            parts.append(' | '.join(c.text for c in r.cells))
    return '\n'.join(parts)


# ============================================================
# 0. 文件存在性
# ============================================================
section(f"0. 文件存在性检查({INPUT_DIR})")
missing_files = []
for fn in EXPECTED_FILES:
    fp = INPUT_DIR / fn
    if fp.exists():
        print(f"  ✓ {fn:<34} {fp.stat().st_size / 1024:>9.1f} KB")
    else:
        missing_files.append(fn)
        print(f"  ✗ {fn:<34} 缺失")
if missing_files:
    print(f"\n  ❌ 缺少 {len(missing_files)} 个附件,终止校验。")
    sys.exit(1)

orders = pd.read_excel(INPUT_DIR / "2026H1渠道销售明细.xlsx", sheet_name='渠道销售明细')
refund = pd.read_excel(INPUT_DIR / "2026H1渠道销售明细.xlsx", sheet_name='渠道退款明细')
ch_month = pd.read_excel(INPUT_DIR / "2026H1渠道销售明细.xlsx", sheet_name='渠道月度汇总')
store_base = pd.read_excel(INPUT_DIR / "2026H1门店经营台账.xlsx", sheet_name='门店基础信息')
store_month = pd.read_excel(INPUT_DIR / "2026H1门店经营台账.xlsx", sheet_name='门店月度经营')
store_inv = pd.read_excel(INPUT_DIR / "2026H1门店经营台账.xlsx", sheet_name='门店库存')
store_exp = pd.read_excel(INPUT_DIR / "2026H1门店经营台账.xlsx", sheet_name='门店费用')
ch_fee = pd.read_excel(INPUT_DIR / "2026H1线上投放与平台费用.xlsx", sheet_name='渠道月度费用')
ad_detail = pd.read_excel(INPUT_DIR / "2026H1线上投放与平台费用.xlsx", sheet_name='投放明细')
commission_rule = pd.read_excel(INPUT_DIR / "2026H1线上投放与平台费用.xlsx", sheet_name='平台佣金规则')
member_base = pd.read_excel(INPUT_DIR / "2026H1会员与跨渠道订单.xlsx", sheet_name='会员基础')
member_order = pd.read_excel(INPUT_DIR / "2026H1会员与跨渠道订单.xlsx", sheet_name='会员订单')
cross_behavior = pd.read_excel(INPUT_DIR / "2026H1会员与跨渠道订单.xlsx", sheet_name='跨渠道行为')
activity = pd.read_excel(INPUT_DIR / "2026H1促销活动记录.xlsx", sheet_name='活动清单')
activity_detail = pd.read_excel(INPUT_DIR / "2026H1促销活动记录.xlsx", sheet_name='活动明细')
activity_effect = pd.read_excel(INPUT_DIR / "2026H1促销活动记录.xlsx", sheet_name='活动效果')
sku_master = pd.read_excel(INPUT_DIR / "SKU主数据.xlsx", sheet_name='SKU主数据')
old_store = pd.read_excel(INPUT_DIR / "旧版门店清单_2024.xlsx", sheet_name='门店清单')
lease_text = read_docx_text(INPUT_DIR / "华东区门店租约汇总.docx")
policy_text = read_docx_text(INPUT_DIR / "渠道与门店管理策略.docx")

# 门店租约解析
lease_rent, lease_property, lease_end, lease_note, lease_region = {}, {}, {}, {}, {}
for block in re.split(r"\n(?=S\d{3} 门店)", lease_text):
    sid = re.search(r"门店ID[:：]\s*(S\d{3})", block)
    if not sid:
        continue
    sid = sid.group(1)
    rent = re.search(r"月租金[:：]\s*([\d.]+)", block)
    prop = re.search(r"物业费[:：]\s*([\d.]+)", block)
    end = re.search(r"租期结束[:：]\s*(\S+)", block)
    note = re.search(r"备注[:：]\s*(\S+)", block)
    region = re.search(r"区域[:：]\s*(\S+)", block)
    lease_rent[sid] = float(rent.group(1)) if rent else np.nan
    lease_property[sid] = float(prop.group(1)) if prop else np.nan
    lease_end[sid] = end.group(1) if end else ''
    lease_note[sid] = note.group(1) if note else ''
    lease_region[sid] = region.group(1) if region else ''

# ============================================================
# 1. 渠道销售明细
# ============================================================
section("1. 2026H1渠道销售明细.xlsx")
expected_cols = ['订单日期', '订单号', '渠道ID', '渠道名称', '渠道类型', '门店ID', '会员ID',
                 'SKU', '品类', '数量', '标价', '成交价', '优惠金额', '收入', '商品成本',
                 '退款金额', '退款数量', '订单状态', '发货方式', '月份', '净收入']
miss = [c for c in expected_cols if c not in orders.columns]
check(not miss, f"  ✓ 字段完整({len(orders.columns)} 列)", f"  ✗ 缺少字段:{miss}")
check('渠道归属' not in orders.columns,
      "  ✓ 无冗余字段「渠道归属」", "  ✗ 仍存在与渠道ID 重复的「渠道归属」")

check(orders['订单号'].is_unique, f"  ✓ 订单号唯一({len(orders)} 行)", "  ✗ 订单号有重复")
check((abs(orders['收入'] - orders['成交价'] * orders['数量']) < 0.01).all(),
      "  ✓ 收入 = 成交价×数量", "  ✗ 收入计算错误")
check((abs(orders['优惠金额'] - (orders['标价'] - orders['成交价']) * orders['数量']) < 0.01).all(),
      "  ✓ 优惠金额 = (标价−成交价)×数量", "  ✗ 优惠金额错误")
check((abs(orders['净收入'] - (orders['收入'] - orders['退款金额'])) < 0.01).all(),
      "  ✓ 净收入 = 收入−退款", "  ✗ 净收入错误")
check((abs(orders['退款金额'] - orders['成交价'] * orders['退款数量']) < 0.01).all(),
      "  ✓ 退款金额 = 成交价×退款数量", "  ✗ 退款金额错误")
check((orders['退款数量'] <= orders['数量']).all(), "  ✓ 退款数量 ≤ 数量", "  ✗ 退款数量超过数量")
check((orders['成交价'] <= orders['标价']).all(), "  ✓ 成交价 ≤ 标价", "  ✗ 成交价超过标价")

od = pd.to_datetime(orders['订单日期'])
check(((od >= '2026-01-01') & (od <= '2026-06-30')).all(),
      f"  ✓ 订单日期在 H1 内({od.min().date()} ~ {od.max().date()})", "  ✗ 订单日期越界")
check((orders['月份'] == od.dt.strftime('%Y-%m')).all(), "  ✓ 月份与订单日期一致", "  ✗ 月份与订单日期不一致")

sku_cost_map = dict(zip(sku_master['SKU编码'], sku_master['单位成本']))
sku_cat_map = dict(zip(sku_master['SKU编码'], sku_master['品类']))
check((abs(orders['商品成本'] - orders['SKU'].map(sku_cost_map) * orders['数量']) < 0.01).all(),
      "  ✓ 商品成本 = SKU单位成本×数量", "  ✗ 商品成本与 SKU 主数据不符")
check((orders['品类'] == orders['SKU'].map(sku_cat_map)).all(),
      "  ✓ 品类与 SKU 主数据一致", "  ✗ 品类与 SKU 主数据不符")

rj = orders[['订单号', '订单日期']].merge(refund[['订单号', '退款日期']], on='订单号')
check(pd.to_datetime(rj['退款日期']).le('2026-06-30').all(),
      "  ✓ 退款日期不跨期(全部落在 H1)", "  ✗ 存在 H1 之后的退款日期(跨期)")
check((pd.to_datetime(rj['退款日期']) >= rj['订单日期']).all(),
      "  ✓ 退款日期 ≥ 订单日期", "  ✗ 退款日期早于订单日期")
check(len(refund) == (orders['退款金额'] > 0).sum(),
      f"  ✓ 退款明细与订单退款一一对应({len(refund)} 笔)", "  ✗ 退款明细与订单退款数不符")

check(orders[orders['发货方式'] == '快递']['门店ID'].isna().all(),
      "  ✓ 快递订单(中心仓)不关联门店ID", "  ✗ 快递订单存在门店ID")
check(orders[orders['发货方式'].isin(['门店发货', '到店自提', '即时配送'])]['门店ID'].notna().all(),
      "  ✓ 门店履约订单均关联门店ID", "  ✗ 门店履约订单存在空门店ID")
check(set(orders['SKU']) <= set(sku_master[sku_master['状态'] != '停售']['SKU编码']),
      "  ✓ 停售 SKU 未产生订单", "  ✗ 停售 SKU 仍出现在订单中")
check(set(orders['会员ID'].dropna()) <= set(member_base['会员ID']),
      "  ✓ 会员ID 与《会员基础》同一编号空间", "  ✗ 会员ID 无法与《会员基础》关联")

agg = orders.groupby(['渠道ID', '月份']).agg(
    收入=('收入', 'sum'), 净收入=('净收入', 'sum'), 成本=('商品成本', 'sum')
).reset_index()
mm = ch_month.merge(agg, on=['渠道ID', '月份'], suffixes=('_表', '_算'))
check((abs(mm['收入_表'] - mm['收入_算']) < 0.01).all()
      and (abs(mm['净收入_表'] - mm['净收入_算']) < 0.01).all(),
      "  ✓ 渠道月度汇总可由订单明细复算", "  ✗ 渠道月度汇总与订单明细不符")
check((abs(ch_month['毛利'] - (ch_month['净收入'] - ch_month['商品成本'])) < 0.01).all(),
      "  ✓ 渠道月度汇总·毛利 = 净收入−商品成本", "  ✗ 渠道月度汇总毛利口径错误")

# ============================================================
# 2. 门店经营台账
# ============================================================
section("2. 2026H1门店经营台账.xlsx")
check(store_base['门店ID'].is_unique and store_base['租约ID'].is_unique,
      f"  ✓ 门店ID/租约ID 唯一({len(store_base)} 家)", "  ✗ 门店ID 或租约ID 有重复")
check((abs(store_month['客单价'] - store_month['到店销售额'] / store_month['订单数']) < 0.01).all(),
      "  ✓ 客单价 = 到店销售额 ÷ 订单数", "  ✗ 客单价错误")
check((abs(store_month['毛利率'] - store_month['毛利额'] / store_month['到店销售额']) < 0.0005).all(),
      "  ✓ 毛利率 = 毛利额 ÷ 到店销售额", "  ✗ 毛利率与毛利额不符")
check((store_month['线上订单履约单量'] == store_month['自提单量'] + store_month['门店发货单量']).all(),
      "  ✓ 线上订单履约单量 = 自提 + 门店发货", "  ✗ 线上履约单量构成错误")

o2o = orders[orders['发货方式'].isin(['门店发货', '到店自提'])]
check(store_month['线上订单履约单量'].sum() == len(o2o),
      f"  ✓ 台账履约单量 = 渠道明细 O2O 订单数({len(o2o)} 单)",
      f"  ✗ 台账履约单量({store_month['线上订单履约单量'].sum()})与渠道明细({len(o2o)})不符")
check(store_month['自提单量'].sum() == (orders['发货方式'] == '到店自提').sum(),
      "  ✓ 自提单量 = 渠道明细到店自提单数", "  ✗ 自提单量与渠道明细不符")
check(store_month['门店发货单量'].sum() == (orders['发货方式'] == '门店发货').sum(),
      "  ✓ 门店发货单量 = 渠道明细门店发货单数", "  ✗ 门店发货单量与渠道明细不符")

o2o_amt = o2o.groupby(['门店ID', '月份'])['净收入'].sum().reset_index()
o2o_amt.columns = ['门店ID', '月份', '_O2O']
sm = store_month.merge(o2o_amt, on=['门店ID', '月份'], how='left')
sm['_O2O'] = sm['_O2O'].fillna(0)
check((abs(sm['门店O2O履约贡献'] - sm['_O2O']) < 0.01).all(),
      "  ✓ 门店O2O履约贡献 = 渠道 O2O 聚合", "  ✗ 门店O2O履约贡献与渠道不符")

check('去年同期到店销售额' not in store_month.columns,
      "  ✓ 《门店月度经营》不含去年同期字段(门店口径不含增长类指标)",
      "  ✗ 仍存在去年同期到店销售额字段")
check('面积是否调整' in store_base.columns,
      '  ✓ 《门店基础信息》以「面积是否调整」标识口径失真门店',
      '  ✗ 缺少「面积是否调整」字段')

exp_rent = store_exp.groupby('门店ID')['租金及物业费'].first()
rent_bad = [s for s in store_base['门店ID']
            if abs(exp_rent[s] - (lease_rent.get(s, np.nan) + lease_property.get(s, np.nan))) > 0.01]
check(not rent_bad, "  ✓ 《门店费用·租金及物业费》= 租约月租金 + 物业费(逐店一致)",
      f"  ✗ 租金及物业费与租约不一致:{rent_bad[:5]}")
check(all(exp_rent[s] > lease_rent.get(s, np.nan) for s in store_base['门店ID']),
      "  ✓ 门店费用已完整计入物业费(不再少计)",
      "  ✗ 存在未计入物业费的门店")

j = store_month.merge(store_exp[['门店ID', '月份', '人力', '营销', '履约包装']], on=['门店ID', '月份'])
check(j['人力'].corr(j['到店销售额']) > 0.5,
      f"  ✓ 人力费与到店销售额相关({j['人力'].corr(j['到店销售额']):.2f})",
      f"  ✗ 人力费与到店销售额不相关({j['人力'].corr(j['到店销售额']):.2f})")
check(j['营销'].corr(j['到店销售额']) > 0.5,
      f"  ✓ 营销费与到店销售额相关({j['营销'].corr(j['到店销售额']):.2f})",
      f"  ✗ 营销费与到店销售额不相关({j['营销'].corr(j['到店销售额']):.2f})")
check(j['履约包装'].corr(j['线上订单履约单量']) > 0.5,
      f"  ✓ 履约包装与履约单量相关({j['履约包装'].corr(j['线上订单履约单量']):.2f})",
      f"  ✗ 履约包装与履约单量不相关({j['履约包装'].corr(j['线上订单履约单量']):.2f})")

neg = [c for c in ['租金及物业费', '人力', '营销', '水电', '履约包装', '其他', '折旧摊销']
       if (store_exp[c] < 0).any()]
check(not neg, "  ✓ 门店费用无负值", f"  ✗ 门店费用存在负值:{neg}")
check(((store_inv['库存数量'] > 0) | (store_inv['库存金额'] == 0)).all(),
      "  ✓ 库存数量与金额逻辑一致", "  ✗ 库存数量与金额矛盾")

inv_sum = store_inv.groupby(['门店ID', '月份'])['库存金额'].sum().reset_index()
inv_sum.columns = ['门店ID', '月份', '明细合计']
inv_cmp = store_month[['门店ID', '月份', '库存金额', '库存周转天数', '到店销售额', '毛利率']].merge(
    inv_sum, on=['门店ID', '月份'])
check((abs(inv_cmp['库存金额'] - inv_cmp['明细合计']) < 0.01).all(),
      "  ✓ 门店库存明细合计 = 门店月度经营·库存金额",
      "  ✗ 库存明细与月度经营库存金额不符")
inv_cmp = inv_cmp.assign(日均成本=inv_cmp['到店销售额'] * (1 - inv_cmp['毛利率']) / 30)
calc_days = inv_cmp['库存金额'] / inv_cmp['日均成本']
check((abs(calc_days - inv_cmp['库存周转天数']) < 0.2).all(),
      f"  ✓ 库存周转天数 = 库存金额 ÷ 日均销售成本(中位 {calc_days.median():.0f} 天)",
      "  ✗ 库存周转天数与库存金额不自洽")

# ============================================================
# 3. 线上投放与平台费用
# ============================================================
section("3. 2026H1线上投放与平台费用.xlsx")
check((abs(ch_fee['净收入'] - (ch_fee['收入'] - ch_fee['退款'])) < 0.01).all(),
      "  ✓ 净收入 = 收入 − 退款", "  ✗ 净收入口径错误")
check((abs(ch_fee['毛利'] - (ch_fee['净收入'] - ch_fee['商品成本'])) < 0.01).all(),
      "  ✓ 毛利 = 净收入 − 商品成本", "  ✗ 毛利口径错误")

grp = orders.groupby(['渠道ID', '月份', '品类']).agg(收入=('收入', 'sum')).reset_index()
grp = grp.merge(commission_rule[['渠道ID', '品类', '佣金率', '技术服务费率']], on=['渠道ID', '品类'])
grp['应佣金'] = grp['收入'] * grp['佣金率']
grp['应技服'] = grp['收入'] * grp['技术服务费率']
expect = grp.groupby(['渠道ID', '月份'])[['应佣金', '应技服']].sum().reset_index()
fee_cmp = ch_fee.merge(expect, on=['渠道ID', '月份'])
check((abs(fee_cmp['平台佣金'] - fee_cmp['应佣金']) < 0.01).all(),
      "  ✓ 平台佣金可由《平台佣金规则》复算", "  ✗ 平台佣金与佣金规则不符")
check((abs(fee_cmp['技术服务费'] - fee_cmp['应技服']) < 0.01).all(),
      "  ✓ 技术服务费可由《平台佣金规则》复算", "  ✗ 技术服务费与佣金规则不符")

ads = ad_detail.copy()
ads['月份'] = pd.to_datetime(ads['日期']).dt.strftime('%Y-%m')
ads_sum = ads.groupby(['渠道ID', '月份'])['投放金额'].sum().reset_index()
ads_sum.columns = ['渠道ID', '月份', '应广告费']
fee_ad = ch_fee.merge(ads_sum, on=['渠道ID', '月份'])
check((abs(fee_ad['广告费'] - fee_ad['应广告费']) < 0.01).all(),
      "  ✓ 广告费 = 投放明细金额合计", "  ✗ 广告费与投放明细不符")

check(len(ad_detail) == 180, f"  ✓ 投放明细 180 条({len(ad_detail)})", f"  ✗ 投放明细 {len(ad_detail)} 条")
per_ch_month = ads.groupby(['渠道ID', '月份']).size()
check((per_ch_month == 5).all(), "  ✓ 每渠道每月 5 条投放", "  ✗ 投放明细分布不均")

check((ad_detail['曝光'] >= ad_detail['点击']).all() and (ad_detail['点击'] >= ad_detail['加购']).all()
      and (ad_detail['加购'] >= ad_detail['支付订单']).all(),
      "  ✓ 投放漏斗单调(曝光 ≥ 点击 ≥ 加购 ≥ 支付订单)", "  ✗ 投放漏斗不单调")
cpm = (ad_detail['投放金额'] / ad_detail['曝光'].replace(0, np.nan) * 1000)
check(cpm.between(5, 100).all(),
      f"  ✓ CPM 合理(中位 {cpm.median():.1f} 元/千次曝光)", f"  ✗ CPM 异常(中位 {cpm.median():.1f})")
attr = ads.groupby(['渠道ID', '月份'])['支付金额'].sum().reset_index()
attr = attr.merge(ch_fee[['渠道ID', '月份', '收入']], on=['渠道ID', '月份'])
check((attr['支付金额'] < attr['收入']).all(),
      "  ✓ 归因支付金额 < 该渠道当月收入", "  ✗ 归因支付金额超过渠道收入")

linked = ad_detail[ad_detail['活动ID'].notna() & (ad_detail['活动ID'] != '')]
bad_link = 0
ch_orders_month = orders.groupby(['渠道ID', '月份']).size()
over_orders = 0
for _, r in linked.iterrows():
    a = activity[activity['活动ID'] == r['活动ID']]
    if len(a) == 0 or a.iloc[0]['渠道ID'] != r['渠道ID'] \
            or not (a.iloc[0]['开始日期'] <= r['日期'] <= a.iloc[0]['结束日期']):
        bad_link += 1
for _, r in ad_detail.iterrows():
    ms = pd.Timestamp(r['日期']).strftime('%Y-%m')
    if r['支付订单'] > ch_orders_month.get((r['渠道ID'], ms), 0):
        over_orders += 1
check(bad_link == 0, f"  ✓ 挂活动的投放渠道与日期均一致({len(linked)} 条)", f"  ✗ {bad_link} 条投放与活动不匹配")
check(over_orders == 0, "  ✓ 投放归因订单不超过该渠道当月订单量", f"  ✗ {over_orders} 条投放归因订单超量")
check((abs(ad_detail['投放ROI'] - ad_detail['支付金额'] / ad_detail['投放金额']) < 0.01).all(),
      "  ✓ 投放ROI = 支付金额 ÷ 投放金额", "  ✗ 投放ROI 计算错误")
check(len(commission_rule) == 36, f"  ✓ 平台佣金规则 36 条(6 渠道×6 品类)", "  ✗ 佣金规则条数错误")

# ============================================================
# 4. 促销活动记录
# ============================================================
section("4. 2026H1促销活动记录.xlsx")
check(activity['活动ID'].is_unique, f"  ✓ 活动ID 唯一({len(activity)} 个)", "  ✗ 活动ID 有重复")
check((activity['结束日期'] >= activity['开始日期']).all(), "  ✓ 活动结束 ≥ 开始", "  ✗ 活动日期倒置")
check(pd.to_datetime(activity['结束日期']).le('2026-06-30').all(),
      "  ✓ 活动日期不跨 H1", "  ✗ 存在跨 H1 的活动")
det = activity_detail.merge(sku_master[['SKU编码', '单位成本', '品类']], left_on='SKU',
                            right_on='SKU编码', suffixes=('', '_sku'))
check((abs(activity_detail['折扣率'] - (1 - activity_detail['活动价'] / activity_detail['原价'])) < 0.0002).all(),
      "  ✓ 折扣率 = 1 − 活动价 ÷ 原价", "  ✗ 折扣率与价格不自洽")
check((det['活动价'] >= det['单位成本']).all(), "  ✓ 活动价不低于单位成本", "  ✗ 存在低于成本的售价")
check((abs(activity_detail['销售额'] - activity_detail['活动价'] * activity_detail['销量']) < 0.01).all(),
      "  ✓ 销售额 = 活动价×销量", "  ✗ 活动明细销售额错误")
check((abs(det['毛利'] - (det['活动价'] - det['单位成本']) * det['销量']) < 0.01).all(),
      "  ✓ 毛利 = (活动价−成本)×销量", "  ✗ 活动明细毛利错误")
det_sum = activity_detail.groupby('活动ID')['销售额'].sum().reset_index()
eff_cmp = det_sum.merge(activity_effect, on='活动ID')
check((abs(eff_cmp['销售额'] - eff_cmp['活动期间销售额']) < 0.01).all(),
      "  ✓ 活动效果·期间销售额 = 明细销售额合计", "  ✗ 活动效果与明细销售额不符")
eff_fee = activity_effect.merge(activity[['活动ID', '实际费用']], on='活动ID')
check((abs(eff_fee['增量销售额'] - (eff_fee['活动期间销售额'] - eff_fee['活动前销售额'])) < 0.01).all(),
      "  ✓ 增量销售额 = 期间 − 活动前", "  ✗ 增量销售额错误")
check((abs(eff_fee['ROI'] - eff_fee['增量毛利'] / eff_fee['实际费用']) < 0.02).all(),
      "  ✓ ROI = 增量毛利 ÷ 实际费用", "  ✗ ROI 计算错误")

# ============================================================
# 5. 会员与跨渠道订单
# ============================================================
section("5. 2026H1会员与跨渠道订单.xlsx")
check(member_base['会员ID'].is_unique, f"  ✓ 会员ID 唯一({len(member_base)} 人)", "  ✗ 会员ID 有重复")
check(member_base['手机号脱敏'].is_unique,
      f"  ✓ 手机号脱敏唯一({member_base['手机号脱敏'].nunique()}/{len(member_base)})",
      "  ✗ 手机号脱敏存在重复(脱敏应保持唯一性)")
check(member_order['订单号'].is_unique, f"  ✓ 会员订单号唯一({len(member_order)} 笔)", "  ✗ 会员订单号有重复")
check(len(set(member_order['订单号']) & set(orders['订单号'])) == 0,
      "  ✓ 会员订单号与渠道订单号不撞号", "  ✗ 存在撞号")
check(set(member_base['会员ID']) <= set(member_order['会员ID']),
      "  ✓ 每个会员至少 1 笔订单", "  ✗ 存在无订单会员")

first = member_order[member_order['是否首购'] == '是']
check((first.groupby('会员ID').size() == 1).all(), "  ✓ 首购唯一(每人恰好 1 笔)", "  ✗ 首购标记不唯一")
earliest = member_order.sort_values(['会员ID', '订单日期']).drop_duplicates('会员ID', keep='first')
check((earliest['是否首购'] == '是').all(), "  ✓ 首购 = 该会员最早订单", "  ✗ 首购不是最早订单")
check(((member_order['是否首购'] == '是') ^ (member_order['是否复购'] == '是')).all(),
      "  ✓ 首购与复购互斥且完备", "  ✗ 首购/复购标记矛盾")

cmp_churn = cross_behavior.merge(member_base[['会员ID', '状态']], on='会员ID')
check((cmp_churn.apply(lambda r: (r['流失标记'] == '是') == (r['状态'] == '流失'), axis=1)).all(),
      "  ✓ 流失标记与会员状态口径一致", "  ✗ 两表流失口径不一致")
check(member_order[member_order['订单类型'] == '线下']['渠道ID'].isna().all(),
      "  ✓ 线下订单不填线上渠道ID", "  ✗ 线下订单存在渠道ID")
check(member_order[member_order['订单类型'] == '线上']['门店ID'].isna().all(),
      "  ✓ 线上订单不填门店ID", "  ✗ 线上订单存在门店ID")
check(member_order[member_order['订单类型'] == 'O2O'][['渠道ID', '门店ID']].notna().all().all(),
      "  ✓ O2O 订单同时具备渠道ID 与门店ID", "  ✗ O2O 订单字段缺失")
first_date = member_order.groupby('会员ID')['订单日期'].min().rename('首单')
reg_cmp = first_date.to_frame().join(member_base.set_index('会员ID')[['注册日期']])
check(reg_cmp.apply(lambda r: r['注册日期'] <= r['首单'], axis=1).all(),
      "  ✓ 注册日期不晚于首单日期", "  ✗ 存在注册晚于首单的会员")
check(abs(cross_behavior['累计消费'].sum() - member_order['金额'].sum()) < 0.01,
      "  ✓ 累计消费合计 = 会员订单金额合计", "  ✗ 累计消费与会员订单不符")
rfm_corr = cross_behavior['RFM分'].corr(cross_behavior['累计消费'])
check(rfm_corr > 0.3, f"  ✓ RFM 分与累计消费正相关({rfm_corr:.2f})",
      f"  ✗ RFM 分与累计消费无关({rfm_corr:.2f})")
lv = cross_behavior.merge(member_base[['会员ID', '会员等级']], on='会员ID').groupby('会员等级')['累计消费'].mean()
order_ok = lv.get('钻石', 0) > lv.get('金卡', 0) > lv.get('银卡', 0) > lv.get('普通', 0)
check(order_ok, f"  ✓ 会员等级与累计消费单调({lv.round(0).to_dict()})", "  ✗ 会员等级与累计消费不单调")
h1_reg = (pd.to_datetime(member_base['注册日期']) >= '2026-01-01').sum()
check(h1_reg > 20, f"  ✓ 2026H1 有新增注册会员({h1_reg} 人)", f"  ✗ H1 新增会员过少({h1_reg} 人)")

recency = (pd.Timestamp('2026-06-30') - cross_behavior['最近购买日期']).dt.days
rule_status = np.where(recency <= 60, '活跃', np.where(recency <= 120, '沉睡', '流失'))
actual_status = member_base.set_index('会员ID').loc[cross_behavior['会员ID'], '状态'].values
check((rule_status == actual_status).all(),
      "  ✓ 会员状态与最近购买天数一致(≤60 活跃 / 61-120 沉睡 / >120 流失)",
      f"  ✗ 会员状态与最近购买天数不一致({(rule_status != actual_status).sum()} 人)")

# ============================================================
# 6. 租约与制度文档
# ============================================================
section("6. 华东区门店租约汇总.docx / 渠道与门店管理策略.docx")
check(len(lease_rent) == 42, f"  ✓ 租约覆盖 42 家门店({len(lease_rent)})", f"  ✗ 租约门店数 {len(lease_rent)}")
bad_end = [s for s in store_base['门店ID'] if lease_end.get(s, '') <= '2026-06-30']
check(not bad_end, "  ✓ 在营门店租期结束晚于基准日", f"  ✗ 存在已过期租约:{bad_end[:5]}")
lease_start_map = {}
for block in re.split(r"\n(?=S\d{3} 门店)", lease_text):
    sid = re.search(r"门店ID[:：]\s*(S\d{3})", block)
    st = re.search(r"租期开始[:：]\s*(\S+)", block)
    if sid:
        lease_start_map[sid.group(1)] = st.group(1) if st else ''
open_map = store_base.set_index('门店ID')['开业日期'].astype(str).to_dict()
lease_order = [s for s in store_base['门店ID'] if lease_start_map.get(s, '') >= open_map.get(s, '')]
check(not lease_order, "  ✓ 租期开始早于门店开业日期",
      f"  ✗ 租期开始晚于开业日期:{lease_order[:5]}")
adjusted = sorted([s for s in store_base['门店ID'] if '面积调整' in lease_note.get(s, '')])
flag_yes = sorted(store_base[store_base['面积是否调整'] == '是']['门店ID'].tolist())
check(adjusted == flag_yes, f'  ✓ 租约备注与台账「面积是否调整」一致({adjusted})',
      f'  ✗ 不一致:租约{adjusted} 台账{flag_yes}')
bad_region = [s for s in store_base['门店ID']
              if lease_region.get(s, '') != store_base.set_index('门店ID').loc[s, '区域']]
check(not bad_region, "  ✓ 租约区域与台账一致", f"  ✗ 区域不一致:{bad_region[:5]}")
policy_keys = ['门店分级标准', 'O2O 履约规则', '数据来源优先级', '缺失与异常处理',
               '闭店与改造条件', '促销与投放效果口径', '商品与品类等级', '会员与订单口径']
miss_policy = [k for k in policy_keys if k not in policy_text]
check(not miss_policy, "  ✓ 制度文档含全部规则章节(含新增 ROI/品类等级/会员口径)",
      f"  ✗ 制度文档缺少章节:{miss_policy}")
check('租金口径' in policy_text and '租金及物业费' in policy_text,
      "  ✓ 制度文档已明确租金口径(合同看租约 / 损益看台账·租金及物业费)",
      "  ✗ 制度文档缺少租金口径条款")

# ============================================================
# 7. SKU 主数据 / 旧版门店清单
# ============================================================
section("7. SKU主数据.xlsx / 旧版门店清单_2024.xlsx")
check(sku_master['SKU编码'].is_unique, f"  ✓ SKU 编码唯一({len(sku_master)} 个)", "  ✗ SKU 编码有重复")
check((sku_master['单位成本'] < sku_master['标准标价']).all(), "  ✓ 单位成本 < 标准标价", "  ✗ 成本 ≥ 标价")
check((abs(sku_master['建议最低售价'] - sku_master['单位成本'] * 1.15) < 0.01).all(),
      "  ✓ 建议最低售价 = 成本×1.15", "  ✗ 建议最低售价错误")
margin = (sku_master['标准标价'] - sku_master['单位成本']) / sku_master['标准标价']
grades = set(sku_master['品类等级'])
check({'A', 'B', 'C'} <= grades,
      f"  ✓ 品类等级 A/B/C 均存在({sku_master['品类等级'].value_counts().to_dict()})",
      f"  ✗ 品类等级分布异常:{sku_master['品类等级'].value_counts().to_dict()}")
check(margin.min() > 0.10 and margin.max() < 0.50,
      f"  ✓ SKU 毛利率区间合理({margin.min():.2%} ~ {margin.max():.2%})",
      f"  ✗ SKU 毛利率区间异常({margin.min():.2%} ~ {margin.max():.2%})")
price_by_cat = sku_master.groupby('品类')['标准标价'].mean().round(0).to_dict()
check(price_by_cat.get('食品', 0) < price_by_cat.get('家电', 0) / 5,
      f"  ✓ 价格体系按品类分层({price_by_cat})", f"  ✗ 价格未按品类分层:{price_by_cat}")

check(len(old_store) == 50, f"  ✓ 旧版清单 50 家门店", "  ✗ 旧版清单门店数错误")
merged_old = old_store.merge(store_base, on='门店ID', suffixes=('_旧', '_新'))
consistent = merged_old.apply(
    lambda r: r['区域_旧'] == r['区域_新'] and r['城市_旧'] == r['城市_新'] and r['门店名称_旧'] == r['门店名称_新'],
    axis=1,
).all()
check(consistent, "  ✓ 旧版清单在营门店的区域/城市/名称与当前台账一致",
      "  ✗ 旧版清单与当前台账在同一门店ID 上存在区域/城市冲突")
banned = ['测试', '模拟', '样例', '合成', 'synthetic', 'test', '噪音', '答案']
hits = []
for w in banned:
    for col in old_store.columns:
        if old_store[col].dtype == object:
            n = old_store[col].astype(str).str.contains(w, na=False).sum()
            if n > 0:
                hits.append(f"{w}@{col}×{n}")
check(not hits, "  ✓ 旧版清单无出题视角禁用词", f"  ✗ 发现禁用词:{hits}")

# 门店口径不含任何"增长/同店"类概念
all_columns = []
for _df in [orders, refund, ch_month, store_base, store_month, store_inv, store_exp, ch_fee,
            ad_detail, commission_rule, member_base, member_order, cross_behavior, activity,
            activity_detail, activity_effect, sku_master, old_store]:
    all_columns += list(_df.columns)
banned_terms = ['同店', '门店增长', '去年同期', '同比']
col_hits = sorted({c for c in all_columns if any(t in str(c) for t in banned_terms)})
doc_hits = sorted({t for t in banned_terms if t in lease_text or t in policy_text})
check(not col_hits and not doc_hits,
      '  ✓ 附件字段与文档均不含「同店 / 门店增长 / 去年同期 / 同比」类概念',
      f'  ✗ 仍存在增长类概念:字段{col_hits} 文档{doc_hits}')

# ============================================================
# 8. 业务合理性
# ============================================================
section("8. 业务合理性检查")
online_aov = orders['净收入'].sum() / len(orders)
offline_aov = store_month['客单价'].mean()
ratio = online_aov / offline_aov
check(ratio < 8, f"  ✓ 线上/到店客单价差距合理({online_aov:.0f} : {offline_aov:.0f} = {ratio:.1f}×)",
      f"  ✗ 线上与到店客单价差距过大({ratio:.1f}×)")
item_price = orders['成交价'].mean()
check(item_price < 1500, f"  ✓ 件均成交价合理({item_price:.0f} 元)", f"  ✗ 件均成交价异常({item_price:.0f} 元)")

# 门店分级可判定性:按制度文档规则应能分出 A/B/C/D,且无"无法归类"
store_perf = store_month.groupby('门店ID').agg(
    到店销售额=('到店销售额', 'sum'),
).join(store_base.set_index('门店ID')[['营业面积', '面积是否调整']])
store_perf['月均销售'] = store_perf['到店销售额'] / 6
store_perf['坪效'] = store_perf['月均销售'] / store_perf['营业面积']
store_perf['月均租金'] = store_exp.groupby('门店ID')['租金及物业费'].mean()
store_perf['租售比'] = store_perf['月均租金'] / store_perf['月均销售']


def grade(row):
    """制度文档 2.2/2.3:先按坪效定级,租售比 >35% 直接列 D。"""
    if row['租售比'] > 0.35:
        return 'D'
    if row['坪效'] >= 5000:
        return 'A'
    if row['坪效'] >= 2000:
        return 'B'
    if row['坪效'] >= 1000:
        return 'C'
    return 'D'


store_perf['分级'] = store_perf.apply(grade, axis=1)
dist = store_perf['分级'].value_counts().to_dict()
check(len(dist) >= 3 and all(k in dist for k in ['A', 'B', 'C', 'D']),
      f"  ✓ 42 家门店均可判定分级且四档齐全({dist})",
      f"  ✗ 分级判定异常(存在无法归类或档位缺失):{dist}")
high_rent = set(store_perf[store_perf['租售比'] > 0.35].index)
check(all(store_perf.loc[s, '分级'] == 'D' for s in high_rent),
      f"  ✓ 租售比 >35% 的门店已全部列为 D 级({len(high_rent)} 家)",
      "  ✗ 存在租售比 >35% 但未列为 D 级的门店")
expired_d = [s for s in store_perf[store_perf['分级'] == 'D'].index if lease_end.get(s, '') <= '2026-12-31']
print(f"  · D 级门店 {len(store_perf[store_perf['分级'] == 'D'])} 家,其中租约在 H2 前到期 {len(expired_d)} 家"
      f"(闭店评估触发条件可达成)")

# ============================================================
# 汇总
# ============================================================
section("汇总")
print(f"\n  问题数:{len(issues)}")
if issues:
    for i, msg in enumerate(issues, 1):
        print(f"    {i}. {msg}")
    print("\n  ❌ 存在数据问题,请修复后重新生成。")
    sys.exit(1)
print("\n  ✅ 全部校验通过:L1 完整性 + L2 自洽 + L3 跨文件勾稽 + L4 合理性。")
sys.exit(0)
