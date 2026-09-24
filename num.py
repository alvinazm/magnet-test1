"""
Magnet 出题 - 线下/线上零售渠道与门店策略附件生成器
v3.1 - 2026-09-23

v3.2 - 2026-09-24(按 docs/skill/出题工作流.md 的 L5 探针与培训文档标准修复):
 - 修复 D1:活动明细改为由《渠道销售明细》订单按"渠道 + 活动日期区间 + SKU"聚合,
   活动期间销售额与订单台账可逐笔勾稽;活动前销售额改用"非活动期日均 × 活动天数"基线
 - 修复 D11:活动明细按 SKU 聚合后不再出现重复行;投放明细的投放项目同渠道当月消重
 - 修复 D12:渠道月度费用"费用合计"改为先对各分项四舍五入再求和,消除 ±0.01 舍入差
 - 新增活动期销售增量:22% 的订单日期定向落入本渠道活动窗口(独立随机源,不影响主随机流)
 - 《渠道与门店管理策略》补充 4.5–4.12、6.1、7.1、7.2、7.5、7.6、9.4、9.7、9.8 条款,
   声明无明细来源的汇总值、数据提取时点、比对容差、客单价口径与各字段语义

v3.1 变更:
 - 删除《门店月度经营·去年同期到店销售额》字段;门店口径不再包含任何"门店增长/同店"概念
 - 《门店基础信息》的"是否同店口径"改为"面积是否调整"(是/否),租约备注同步改为"计划期内发生面积调整"
 - 《渠道与门店管理策略》第 2 章简化为"坪效定级 + 租售比 >35% 直接 D",并删除全部同店表述;
   面积调整门店改为按"坪效口径可能失真,须披露"处理(2.5)

相对 v2.0 的修复(逐条对应 input 数据缺陷清单):
 1. 移除冗余字段「渠道归属」(原与渠道ID 逐行完全重复)
 2. 退款日期全部落在 2026H1 内,不再把 H2 退款计入 H1 净收入
 3. 商品价格改为按品类定价(食品/日用低、家电高),件数 1-3 件为主,线上/线下客单价回归合理差距
 4. 发货方式按渠道分布;门店ID 规则统一:快递=中心仓(无门店ID),门店发货/到店自提/即时配送=门店履约(必填)
 5. 门店「线上订单履约单量/自提单量/门店发货单量」由渠道订单明细聚合,不再与明细脱节(原差 505 倍)
 6. 门店费用与业务驱动因素挂钩(人力/营销/其他 ∝ 到店销售,水电/折旧 ∝ 面积,履约包装 ∝ 履约单量)
 7. 《门店费用·租金》与《门店租约汇总·月租金》单一来源,逐店完全一致
 8. 门店基础信息标注"面积是否调整",用于面积调整导致坪效口径失真的披露
 9. 在营门店租期结束日晚于基准日;部分低效门店租期落在 2026H2,支撑闭店评估
10. 平台佣金/技术服务费由《平台佣金规则》精确复算(原差 -41.7%~+58.7%)
11. 《渠道月度费用·毛利》= 净收入 − 商品成本,并补齐退款/净收入/费用合计/净利等口径列
12. 广告费 = 该渠道当月《投放明细》投放金额合计(原两表差 7% 且无口径)
13. 投放明细与同渠道活动关联且日期落在活动期内;新增「支付金额」,投放 ROI = 支付金额 ÷ 投放金额
14. 折扣率 = 1 − 活动价 ÷ 原价;活动价不低于成本;活动日期不跨 H1
15. 活动效果·活动期间销售额 = 《活动明细》销售额合计,增量与 ROI 由明细复算
16. 会员手机号脱敏唯一;RFM 分、会员等级、流失状态由真实购买行为派生且两表口径统一
17. 会员订单类型与渠道ID/门店ID 一致;每个会员至少 1 笔订单;首购唯一;会员ID 与《渠道销售明细》同一编号空间
18. 会员注册日期覆盖 2023-01 至 2026H1,含 H1 新增会员
19. 旧版门店清单的区域/城市/名称与当前台账一致(仅状态与存续不同)
20. SKU 按品类定价,成本比例使毛利率覆盖 15%-45%,A/B/C 三档等级均存在
21. 停售 SKU 不再出现在 2026H1 订单中
22. 《渠道与门店管理策略》补齐缺失规则:门店分级(无空档)、O2O 单量口径、
    促销与投放 ROI、SKU 品类等级、会员与订单类型口径

试跑(不覆盖 input/):
    MAGNET_INPUT_DIR=/tmp/magnet_dryrun python num_v3.py
正常运行:
    python num.py
"""
import os
import random
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
from docx import Document as DocxDocument

# ===================== 全局配置 =====================
# 输出目录:默认写入 input/(模拟真实业务数据导出);可用环境变量覆盖以便试跑
OUTPUT_DIR = os.environ.get("MAGNET_INPUT_DIR", "input")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 固定随机种子,保证可复现
SEED = 42
random.seed(SEED)
np.random.seed(SEED)

# 渠道定义
channels = [
    ('CH01', '自营APP', '自营'),
    ('CH02', '微信小程序', '自营'),
    ('CH03', '天猫旗舰店', '平台'),
    ('CH04', '京东自营', '平台'),
    ('CH05', '抖音小店', '平台'),
    ('CH06', '美团闪购', '即时零售'),
]
channel_ids = [c[0] for c in channels]
channel_names = {c[0]: c[1] for c in channels}
channel_types = {c[0]: c[2] for c in channels}

# 渠道订单权重:抖音 + 天猫合计约 60%
channel_order_weights = [0.10, 0.08, 0.30, 0.12, 0.30, 0.10]

# 门店ID 占位:线下注册会员写入《会员基础》
STORE_CHANNEL_LABEL = '门店'

# 区域与城市
region_cities = {
    '华东': ['上海', '杭州', '南京', '苏州', '宁波', '合肥', '无锡', '常州'],
    '华南': ['广州', '深圳', '厦门', '福州', '南宁'],
    '华北': ['北京', '天津', '济南', '青岛', '石家庄'],
}
regions = list(region_cities.keys())
all_cities = sum(region_cities.values(), [])
business_districts = ['核心商圈', '社区商圈', '交通枢纽', '办公区', '郊区']

# 品类与价格体系(修复 3):按品类设定真实价格区间
categories = ['食品', '日用', '美妆', '家电', '服饰', '母婴']
category_price_range = {
    '食品': (8, 120),
    '日用': (12, 200),
    '美妆': (59, 700),
    '家电': (199, 6000),
    '服饰': (99, 1200),
    '母婴': (39, 800),
}
# 渠道品类结构:线上家电/服饰占比更高,即时零售以食品日用为主
channel_category_weights = {
    'CH01': [0.15, 0.15, 0.15, 0.25, 0.15, 0.15],
    'CH02': [0.15, 0.15, 0.20, 0.15, 0.20, 0.15],
    'CH03': [0.15, 0.10, 0.20, 0.20, 0.25, 0.10],
    'CH04': [0.10, 0.10, 0.15, 0.30, 0.25, 0.10],
    'CH05': [0.15, 0.10, 0.25, 0.10, 0.30, 0.10],
    'CH06': [0.30, 0.30, 0.15, 0.05, 0.05, 0.15],
}
# 门店到店销售的品类结构:以便利型消费品为主,不含大家电
store_category_weights = [0.30, 0.28, 0.14, 0.02, 0.14, 0.12]

# 发货方式分布(修复 4):按渠道类型区分履约主体
channel_ship_weights = {
    'CH01': {'快递': 0.45, '门店发货': 0.25, '到店自提': 0.20, '即时配送': 0.10},
    'CH02': {'快递': 0.15, '门店发货': 0.25, '到店自提': 0.50, '即时配送': 0.10},
    'CH03': {'快递': 0.50, '门店发货': 0.25, '到店自提': 0.18, '即时配送': 0.07},
    'CH04': {'快递': 0.55, '门店发货': 0.25, '到店自提': 0.15, '即时配送': 0.05},
    'CH05': {'快递': 0.45, '门店发货': 0.30, '到店自提': 0.18, '即时配送': 0.07},
    'CH06': {'快递': 0.00, '门店发货': 0.05, '到店自提': 0.20, '即时配送': 0.75},
}

# 日期范围
start_date = datetime(2026, 1, 1)
end_date = datetime(2026, 6, 30)
date_range = pd.date_range(start_date, end_date, freq='D')
months = [(2026, m) for m in range(1, 7)]
month_strs = [f'{y}-{m:02d}' for (y, m) in months]


def rand_datetime(a: datetime, b: datetime) -> datetime:
    """在 [a, b] 之间随机取一个时间点(含端点,按天粒度)。"""
    delta = max(0, (b - a).days)
    return a + timedelta(days=random.randint(0, delta))


def sample_category_price(cat: str) -> float:
    """按品类的价格区间取价,分布偏低价段(长尾)。"""
    lo, hi = category_price_range[cat]
    return round(lo + (hi - lo) * float(np.random.beta(1.6, 4.0)), 2)


# ===================== 0. 基础主数据(供各附件共同使用)=====================
print("[0/9] 生成基础主数据(SKU / 门店 / 租约 / 佣金规则 / 促销 / 投放)...")

# ---- 门店主数据 ----
store_region_assign = ['华东'] * 18 + ['华南'] * 12 + ['华北'] * 12
store_ids = [f'S{i:03d}' for i in range(1, 43)]
store_names = [f'门店{i:02d}' for i in range(1, 43)]
store_region = list(store_region_assign)
store_city = [random.choice(region_cities[r]) for r in store_region]
store_district = [random.choice(business_districts) for _ in range(42)]
# 营业面积:社区型门店以 60-140 ㎡ 为主,少量 140-260 ㎡
store_area = [
    int(random.uniform(140, 260)) if random.random() < 0.10 else int(random.uniform(60, 140))
    for _ in range(42)
]
store_open_date = [datetime(2019, 1, 1) + timedelta(days=random.randint(0, 2300)) for _ in range(42)]
store_type = [random.choice(['标准店', '社区店', '旗舰店', '前置仓']) for _ in range(42)]
# 3 家计划期内发生面积调整的门店(与租约"备注"字段一致)
adjusted_store_idx = [5, 12, 30]  # S006 / S013 / S031
area_adjusted_flag = ['是' if i in adjusted_store_idx else '否' for i in range(42)]

# ---- 到店销售能力模型(决定坪效与门店分级)----
store_ability = {}
store_monthly_sales = {}   # (门店ID, 月份) -> dict
for i, sid in enumerate(store_ids):
    area = store_area[i]
    ability = random.choices(
        ['问题店', '低效店', '正常店', '明星店'],
        weights=[12, 18, 55, 15],
    )[0]
    store_ability[sid] = ability
    base_range = {
        '问题店': (60000, 140000),
        '低效店': (150000, 220000),
        '正常店': (220000, 350000),
        '明星店': (350000, 600000),
    }[ability]
    area_factor = 0.85 + area / 140 * 0.30         # 60㎡→0.98, 140㎡→1.15

    for (y, m) in months:
        ms = f'{y}-{m:02d}'
        sales_target = random.uniform(*base_range) * area_factor * random.uniform(0.85, 1.15)
        aov = round(random.uniform(150, 320), 2)
        orders = max(1, int(sales_target / aov))
        sales = round(orders * aov, 2)
        conv_rate = random.uniform(0.10, 0.30)
        traffic = max(orders, int(orders / conv_rate))
        gross_rate = random.uniform(0.25, 0.45)
        store_monthly_sales[(sid, ms)] = {
            'ability': ability,
            '到店销售额': sales,
            '订单数': orders,
            '客单价': aov,
            '进店客流': traffic,
            '成交率': round(orders / traffic, 4),
            '连带率': round(random.uniform(1.2, 2.5), 2),
            '毛利额': round(sales * gross_rate, 2),
            '毛利率': round(gross_rate, 4),
        }

# ---- SKU 主数据(修复 3 / 20 / 21)----
skus = [f'SKU{str(i).zfill(4)}' for i in range(1, 201)]
sku_category = {sku: random.choice(categories) for sku in skus}
sku_price = {sku: sample_category_price(sku_category[sku]) for sku in skus}
# 成本占标价 55%-85% → 毛利率 15%-45%,使 A/B/C 三档等级都存在
sku_cost = {sku: round(sku_price[sku] * random.uniform(0.55, 0.85), 2) for sku in skus}
sku_min_price = {sku: round(sku_cost[sku] * 1.15, 2) for sku in skus}
sku_margin = {
    sku: round((sku_price[sku] - sku_cost[sku]) / sku_price[sku], 4) for sku in skus
}
# 品类等级阈值:A ≥40%、B 25%-40%、C <25%
sku_grade = {
    sku: ('A' if sku_margin[sku] >= 0.40 else ('B' if sku_margin[sku] >= 0.25 else 'C'))
    for sku in skus
}
sku_lead_time = {sku: random.choice([7, 15, 30, 45, 60]) for sku in skus}
sku_launch = {
    sku: rand_datetime(datetime(2023, 1, 1), datetime(2025, 6, 30)).strftime('%Y-%m-%d')
    for sku in skus
}
sku_status = {
    sku: random.choices(['在售', '季节性', '停售'], weights=[0.80, 0.15, 0.05])[0]
    for sku in skus
}
subcategories = {
    '食品': ['零食', '饮料', '生鲜', '调味'],
    '日用': ['洗护', '清洁', '纸品', '家居'],
    '美妆': ['护肤', '彩妆', '香水', '工具'],
    '家电': ['小家电', '厨房', '数码', '个护'],
    '服饰': ['男装', '女装', '童装', '配件'],
    '母婴': ['奶粉', '辅食', '用品', '玩具'],
}
sku_subcategory = {sku: random.choice(subcategories[sku_category[sku]]) for sku in skus}

df_sku_master = pd.DataFrame({
    'SKU编码': skus,
    '品类': [sku_category[s] for s in skus],
    '子品类': [sku_subcategory[s] for s in skus],
    '单位成本': [sku_cost[s] for s in skus],
    '标准标价': [sku_price[s] for s in skus],
    '建议最低售价': [sku_min_price[s] for s in skus],
    '品类等级': [sku_grade[s] for s in skus],
    '采购周期': [sku_lead_time[s] for s in skus],
    '上市日期': [sku_launch[s] for s in skus],
    '状态': [sku_status[s] for s in skus],
})
# 停售 SKU 不得出现在 2026H1 订单中(修复 21)
sellable_skus_by_cat = {
    cat: [s for s in skus if sku_category[s] == cat and sku_status[s] != '停售']
    for cat in categories
}

# ---- 租约主数据(修复 7 / 9):租金与租期单一来源 ----
lease = {}
d_grade_candidates = [sid for sid in store_ids if store_ability[sid] == '问题店']
for i, sid in enumerate(store_ids):
    area = store_area[i]
    unit_rate = random.uniform(150, 700)                 # 元/㎡/月
    rent = round(area * unit_rate, 2)
    # 租约先签后开业:租期开始早于开业日期 30-180 天
    lease_start = store_open_date[i] - timedelta(days=random.randint(30, 180))
    if sid in d_grade_candidates and random.random() < 0.6:
        # 部分低效门店租期落在 2026H2,使"租约到期 + 连续 D 级"具备触发条件
        lease_end = rand_datetime(datetime(2026, 7, 1), datetime(2026, 12, 31))
    else:
        lease_end = rand_datetime(datetime(2027, 1, 1), datetime(2032, 12, 31))
    lease[sid] = {
        '租约ID': f'L{str(i + 1).zfill(3)}',
        '租期开始': lease_start.strftime('%Y-%m-%d'),
        '租期结束': lease_end.strftime('%Y-%m-%d'),
        '免租期': f'{random.randint(0, 90)}天',
        '月租金': rent,
        '物业费': round(rent * 0.12, 2),
        '递增条款': random.choice(['无', '每年递增3%', '每年递增5%', '每两年递增8%']),
        '违约金': round(rent * random.uniform(2, 6), 2),
        '备注': '计划期内发生面积调整' if i in adjusted_store_idx else '正常',
    }

# ---- 平台佣金规则(修复 10):先生成规则,费用表按规则精确计算 ----
# 自营渠道(自营APP / 微信小程序)不承担平台佣金,仅支付通道费率
commission_range = {
    'CH01': (0.005, 0.020),
    'CH02': (0.004, 0.018),
    'CH03': (0.040, 0.090),
    'CH04': (0.035, 0.080),
    'CH05': (0.050, 0.100),
    'CH06': (0.030, 0.060),
}
tech_range = {
    'CH01': (0.001, 0.005),
    'CH02': (0.001, 0.004),
    'CH03': (0.010, 0.030),
    'CH04': (0.008, 0.025),
    'CH05': (0.012, 0.032),
    'CH06': (0.005, 0.015),
}
commission_rules = {}
rules = []
for ch in channel_ids:
    for cat in categories:
        rate = round(random.uniform(*commission_range[ch]), 4)
        tech_rate = round(random.uniform(*tech_range[ch]), 4)
        commission_rules[(ch, cat)] = (rate, tech_rate)
        rules.append({
            '渠道ID': ch,
            '品类': cat,
            '佣金率': rate,
            '技术服务费率': tech_rate,
            '结算周期': random.choice(['T+1', 'T+7', 'T+15', '月结']),
            '备注': '',
        })
df_rules = pd.DataFrame(rules)

# ---- 促销活动(修复 14 / 15)----
activities = []
activity_channels = random.choices(channel_ids, weights=channel_order_weights, k=50)
for i in range(1, 51):
    aid = f'A{i:03d}'
    ch = activity_channels[i - 1]
    # 活动日期不跨 H1
    act_start = rand_datetime(start_date, datetime(2026, 6, 15))
    act_end = min(act_start + timedelta(days=random.randint(2, 14)), end_date)
    budget = round(random.uniform(2, 15), 2) * 1000
    actual = round(budget * random.uniform(0.9, 1.1), 2)
    activities.append({
        '活动ID': aid,
        '活动名称': f'活动{i}',
        '活动类型': random.choice(['满减', '折扣', '直播', '会员日', '节日促销']),
        '开始日期': act_start,
        '结束日期': act_end,
        '渠道ID': ch,
        '门店ID': random.choice(store_ids) if random.random() > 0.5 else None,
        '预算': budget,
        '实际费用': actual,
    })
df_activities = pd.DataFrame(activities)

# 活动期的销售增量:把一部分订单日期定向落到该渠道的活动窗口内,使"活动前 vs 活动期间"
# 的比较具有真实增量(而不是纯噪声)。使用独立随机源,不消耗主随机流,避免影响其他附件。
activity_by_channel = {}
for _, _act in df_activities.iterrows():
    activity_by_channel.setdefault(_act['渠道ID'], []).append(
        (_act['开始日期'], _act['结束日期'])
    )
rng_promo = random.Random(20260625)
promo_order_share = 0.22          # 22% 的订单落在活动窗口内,形成可观测的活动增量

# ===================== 1. 渠道销售明细 =====================
print("生成 2026H1渠道销售明细.xlsx ...")
member_ids = [f'M{str(i).zfill(6)}' for i in range(1, 2501)]
orders = []
order_id_counter = 100000
order_channels = random.choices(channel_ids, weights=channel_order_weights, k=7500)

for idx in range(7500):
    channel = order_channels[idx]
    # 订单日期:7.5% 的订单会退款,退款需在 H1 内完成,故这批订单日期提前
    has_refund = random.random() < 0.075
    if has_refund:
        order_date = rand_datetime(start_date, datetime(2026, 6, 5))
    else:
        order_date = rand_datetime(start_date, end_date)      # 照旧消耗 1 次主随机数
        windows = activity_by_channel.get(channel, [])
        if windows and rng_promo.random() < promo_order_share:
            ws, we = rng_promo.choice(windows)
            order_date = ws + timedelta(days=rng_promo.randint(0, (we - ws).days))

    cat = random.choices(categories, weights=channel_category_weights[channel])[0]
    sku = random.choice(sellable_skus_by_cat[cat])
    qty = random.choices([1, 2, 3, 4, 5, 6], weights=[58, 24, 10, 4, 2, 2])[0]
    list_price = sku_price[sku]
    deal_price = round(list_price * random.uniform(0.85, 1.0), 2) if random.random() < 0.7 else list_price
    revenue = round(deal_price * qty, 2)
    cost = round(sku_cost[sku] * qty, 2)

    refund = 0.0
    refund_qty = 0
    status = '已完成'
    if has_refund:
        refund_qty = random.randint(1, qty)
        refund = round(deal_price * refund_qty, 2)
        status = '部分退款' if refund_qty < qty else '已退款'

    ship_method = random.choices(
        list(channel_ship_weights[channel].keys()),
        weights=list(channel_ship_weights[channel].values()),
    )[0]
    # 修复 4:快递=中心仓发货,不关联门店;其余三种由门店履约,必须关联门店
    store_id = None if ship_method == '快递' else random.choice(store_ids)
    member_id = random.choice(member_ids) if random.random() < 0.45 else None

    orders.append({
        '订单日期': order_date,
        '订单号': f'O{order_id_counter}',
        '渠道ID': channel,
        '渠道名称': channel_names[channel],
        '渠道类型': channel_types[channel],
        '门店ID': store_id,
        '会员ID': member_id,
        'SKU': sku,
        '品类': cat,
        '数量': qty,
        '标价': list_price,
        '成交价': deal_price,
        '优惠金额': round((list_price - deal_price) * qty, 2),
        '收入': revenue,
        '商品成本': cost,
        '退款金额': refund,
        '退款数量': refund_qty,
        '订单状态': status,
        '发货方式': ship_method,
        '月份': f'{order_date.year}-{order_date.month:02d}',
    })
    order_id_counter += 1

df_orders = pd.DataFrame(orders)
df_orders['净收入'] = df_orders['收入'] - df_orders['退款金额']

# ---- 促销活动明细与效果(修复 D1/D11):活动明细直接由订单台账聚合,可与订单逐笔勾稽 ----
# 口径(《渠道与门店管理策略》7.2):
#   · 活动期间销售额 = 该渠道(活动绑定门店时为该门店)在活动起止日期内订单的收入合计;
#   · 活动前销售额   = 活动开始前「等长窗口」的同口径收入合计;
#   · 活动价 = 期间该 SKU 收入 ÷ 销量(加权成交价),原价取 SKU 标准标价。
rev_by_ch_month = df_orders.groupby(['渠道ID', '月份'])['收入'].sum()
cnt_by_ch_month = df_orders.groupby(['渠道ID', '月份']).size()
aov_by_ch_month = (rev_by_ch_month / cnt_by_ch_month).round(2)

# 为保持随机流与上一版一致(其余 8 个附件数据不变),按旧逻辑的「两段式」顺序消耗同等数量的随机数:
# 第一段:每个活动的 target_sales 与明细行抽样;第二段:每个活动的 uplift
for _, act in df_activities.iterrows():
    random.uniform(0.20, 0.60)
    for _ in range(random.randint(5, 20)):
        cat = random.choices(categories, weights=channel_category_weights[act['渠道ID']])[0]
        random.choice(sellable_skus_by_cat[cat])
        random.uniform(0.80, 0.97)
for _ in df_activities.iterrows():
    random.uniform(0.40, 2.00)

df_orders['订单日期'] = pd.to_datetime(df_orders['订单日期'])


def _activity_orders(act, start, end):
    """取活动窗口内该渠道的订单(活动清单的 `门店ID` 为发起门店,不改变聚合口径)。"""
    return df_orders[(df_orders['渠道ID'] == act['渠道ID'])
                     & (df_orders['订单日期'] >= start)
                     & (df_orders['订单日期'] <= end)]


# 活动前基线 = 该渠道「非活动期日均收入」× 活动天数(H1 内不属于该渠道任何活动窗口的日期)
h1_days = pd.date_range(start_date, end_date, freq='D')
baseline_daily = {}
for ch in channel_ids:
    act_days = set()
    for ws, we in activity_by_channel.get(ch, []):
        act_days |= set(pd.date_range(ws, we, freq='D'))
    non_act_days = [d for d in h1_days if d not in act_days]
    daily_rev = df_orders[df_orders['渠道ID'] == ch].groupby('订单日期')['收入'].sum()
    baseline_daily[ch] = float(daily_rev.reindex(non_act_days).fillna(0).mean()) if non_act_days \
        else float(daily_rev.mean())


details, effects = [], []
for _, act in df_activities.iterrows():
    aid = act['活动ID']
    start = pd.Timestamp(act['开始日期'])
    end = pd.Timestamp(act['结束日期'])
    span = (end - start).days + 1

    during_sel = _activity_orders(act, start, end)
    grp = during_sel.groupby('SKU').agg(
        销量=('数量', 'sum'), 销售额=('收入', 'sum'), 商品成本=('商品成本', 'sum'),
    ).reset_index()
    for _, row in grp.iterrows():
        sku = row['SKU']
        orig = sku_price[sku]
        sales = round(float(row['销售额']), 2)
        qty = int(row['销量'])
        avg_price = round(sales / qty, 2) if qty else 0.0
        details.append({
            '活动ID': aid,
            'SKU': sku,
            '品类': sku_category[sku],
            '原价': orig,
            '活动价': avg_price,
            '折扣率': round(1 - avg_price / orig, 4),
            '销量': qty,
            '销售额': sales,
            '毛利': round(sales - float(row['商品成本']), 2),
        })

    during = round(float(grp['销售额'].sum()), 2)              # 与活动明细严格一致
    before = round(baseline_daily[act['渠道ID']] * span, 2)     # 非活动期日均 × 活动天数
    inc = round(during - before, 2)
    during_cost = float(grp['商品成本'].sum())
    act_margin = ((during - during_cost) / during) if during > 0 else 0.3
    inc_gross = round(inc * act_margin, 2)
    roi = round(inc_gross / act['实际费用'], 2) if act['实际费用'] > 0 else 0
    effects.append({
        '活动ID': aid,
        '渠道ID': act['渠道ID'],
        '门店ID': act['门店ID'],
        '活动期间销售额': during,
        '活动前销售额': before,
        '增量销售额': inc,
        '增量毛利': inc_gross,
        'ROI': roi,
    })
df_details = pd.DataFrame(details)
df_effects = pd.DataFrame(effects)

# ---- 投放明细(修复 13):每月每渠道 5 条,合计 180 条 ----
# 广告预算与该渠道当月收入挂钩(按渠道设定费率),再由逆向漏斗反推曝光/点击/加购/支付订单,
# 使「广告费 = 投放金额合计」且「支付订单不超过该渠道实际订单量」。
ad_rate_range = {
    'CH01': (0.14, 0.24),   # 自营APP 投放依赖度高
    'CH02': (0.04, 0.08),   # 微信小程序 私域为主,投放低
    'CH03': (0.12, 0.20),
    'CH04': (0.08, 0.15),
    'CH05': (0.12, 0.22),
    'CH06': (0.06, 0.12),
}
投放 = []
for ch in channel_ids:
    ch_acts = df_activities[df_activities['渠道ID'] == ch]
    for (y, m) in months:
        ms = f'{y}-{m:02d}'
        m_start = datetime(y, m, 1)
        m_end = datetime(y, m + 1, 1) - timedelta(days=1) if m < 12 else end_date
        revenue = float(rev_by_ch_month.get((ch, ms), 0.0))
        ch_orders = int(cnt_by_ch_month.get((ch, ms), 0))
        ch_aov = float(aov_by_ch_month.get((ch, ms), 0.0))
        ad_budget = revenue * random.uniform(*ad_rate_range[ch])
        weights = [random.uniform(0.7, 1.3) for _ in range(5)]
        amounts = [round(ad_budget * w / sum(weights), 2) for w in weights]
        amounts[-1] = round(ad_budget - sum(amounts[:-1]), 2)
        # 该渠道当月可归因订单上限(20%-45%),按 5 条投放均分,避免归因订单超过真实订单量
        cap_orders = max(1, int(ch_orders * random.uniform(0.20, 0.45) / 5))
        month_rows = []
        for k in range(5):
            amount = amounts[k]
            ad_date = rand_datetime(m_start, m_end)
            activity_id = ''
            overlapping = ch_acts[
                (ch_acts['开始日期'] <= m_end) & (ch_acts['结束日期'] >= m_start)
            ]
            if len(overlapping) > 0 and random.random() < 0.7:
                act_row = overlapping.sample(n=1).iloc[0]
                activity_id = act_row['活动ID']
                ad_date = rand_datetime(
                    max(act_row['开始日期'], m_start), min(act_row['结束日期'], m_end)
                )
            # 曝光由投放金额与 CPM 决定(CPM 15-40 元/千次曝光),
            # 再按真实漏斗率(CTR / 加购率 / 支付率)向前推算点击与加购
            cpm = random.uniform(15, 40)
            exposure = int(amount / cpm * 1000)
            click = int(exposure * random.uniform(0.008, 0.035))
            add_cart = int(click * random.uniform(0.06, 0.18))
            funnel_pay = int(add_cart * random.uniform(0.08, 0.22))
            pay_orders = max(0, min(funnel_pay, cap_orders))
            pay_amount = round(pay_orders * ch_aov * random.uniform(0.95, 1.05), 2)
            proj_no = random.randint(1, 20)          # 保持与原实现相同的抽样位置与顺序
            month_rows.append({
                '日期': ad_date,
                '渠道ID': ch,
                '投放项目': proj_no,
                '活动ID': activity_id,
                '投放金额': amount,
                '曝光': exposure,
                '点击': click,
                '加购': add_cart,
                '支付订单': pay_orders,
                '支付金额': pay_amount,
                '投放ROI': round(pay_amount / amount, 2) if amount > 0 else 0,
            })
        # 同一渠道当月 5 个投放项目消重,避免出现重复的(日期,渠道,项目)组合
        used_projects, projects = set(), []
        for row in month_rows:
            value = row['投放项目']
            while value in used_projects:
                value = value % 20 + 1
            used_projects.add(value)
            projects.append(value)
        for row, value in zip(month_rows, projects):
            row['投放项目'] = f'项目{value}'
            投放.append(row)
df_ad = pd.DataFrame(投放)

# 渠道退款明细(修复 2):退款日期一律落在 H1 内
refund_records = df_orders[df_orders['退款金额'] > 0].copy()
refund_records['退款日期'] = refund_records['订单日期'] + pd.to_timedelta(
    np.random.randint(1, 22, len(refund_records)), unit='D'
)
refund_records['退款日期'] = refund_records['退款日期'].clip(upper=pd.Timestamp(end_date))
refund_records['退款原因'] = np.random.choice(
    ['七天无理由', '质量问题', '错拍', '物流损坏', '其他'], len(refund_records)
)
refund_records['是否影响毛利'] = '是'
df_refund = refund_records[['订单号', '渠道ID', '退款日期', '退款原因', '退款金额', '是否影响毛利']]

# 渠道月度汇总
monthly = df_orders.groupby(['月份', '渠道ID', '渠道名称']).agg(
    收入=('收入', 'sum'),
    退款=('退款金额', 'sum'),
    净收入=('净收入', 'sum'),
    商品成本=('商品成本', 'sum'),
    订单量=('订单号', 'count'),
).reset_index()
monthly['毛利'] = (monthly['净收入'] - monthly['商品成本']).round(2)
monthly['毛利率'] = (monthly['毛利'] / monthly['净收入']).round(4)
monthly['客单价'] = (monthly['净收入'] / monthly['订单量']).round(2)

with pd.ExcelWriter(f'{OUTPUT_DIR}/2026H1渠道销售明细.xlsx', engine='openpyxl') as writer:
    df_orders.to_excel(writer, sheet_name='渠道销售明细', index=False)
    df_refund.to_excel(writer, sheet_name='渠道退款明细', index=False)
    monthly.to_excel(writer, sheet_name='渠道月度汇总', index=False)
print("  -> 完成")

# ===================== 2. 门店经营台账 =====================
print("生成 2026H1门店经营台账.xlsx ...")
store_base = pd.DataFrame({
    '门店ID': store_ids,
    '门店名称': store_names,
    '区域': store_region,
    '城市': store_city,
    '商圈类型': store_district,
    '开业日期': store_open_date,
    '营业面积': store_area,
    '租约ID': [lease[sid]['租约ID'] for sid in store_ids],
    '店型': store_type,
    '状态': ['营业'] * 42,
    '面积是否调整': area_adjusted_flag,
})

# 修复 5:门店 O2O 履约单量与贡献均由渠道订单明细聚合
o2o_orders = df_orders[df_orders['发货方式'].isin(['门店发货', '到店自提'])]
o2o_amount = o2o_orders.groupby(['门店ID', '月份'])['净收入'].sum()
o2o_ship_cnt = o2o_orders[o2o_orders['发货方式'] == '门店发货'].groupby(['门店ID', '月份']).size()
o2o_pickup_cnt = o2o_orders[o2o_orders['发货方式'] == '到店自提'].groupby(['门店ID', '月份']).size()

store_monthly_rows = []
store_month_inventory = {}          # (门店ID, 月份) -> 库存金额(与《门店库存》明细合计一致)
for sid in store_ids:
    area = store_base.loc[store_base['门店ID'] == sid, '营业面积'].values[0]
    # 门店到店销售的品类结构(用于反推动销 SKU 数与库存)
    for ms in month_strs:
        row = store_monthly_sales[(sid, ms)]
        ship_cnt = int(o2o_ship_cnt.get((sid, ms), 0))
        pickup_cnt = int(o2o_pickup_cnt.get((sid, ms), 0))
        o2o_contribution = float(o2o_amount.get((sid, ms), 0.0))
        # 库存金额按库存周转天数反推:库存金额 = 日均销售成本 × 周转天数(20-90 天)
        daily_cost = row['到店销售额'] * (1 - row['毛利率']) / 30
        turnover_days = random.uniform(20, 90)
        inventory_amt = round(daily_cost * turnover_days, 2)
        store_month_inventory[(sid, ms)] = inventory_amt
        store_monthly_rows.append({
            '月份': ms,
            '门店ID': sid,
            '到店销售额': row['到店销售额'],
            '订单数': row['订单数'],
            '客单价': row['客单价'],
            '进店客流': row['进店客流'],
            '成交率': row['成交率'],
            '连带率': row['连带率'],
            '毛利额': row['毛利额'],
            '毛利率': row['毛利率'],
            '线上订单履约单量': ship_cnt + pickup_cnt,
            '自提单量': pickup_cnt,
            '门店发货单量': ship_cnt,
            '退货额': round(row['到店销售额'] * random.uniform(0.01, 0.08), 2),
            '库存金额': inventory_amt,
            '库存周转天数': round(inventory_amt / daily_cost, 1) if daily_cost else 0.0,
            '动销SKU数': random.randint(80, 180),
            '缺货率': round(random.uniform(0.01, 0.15), 4),
            '损耗率': round(random.uniform(0.005, 0.05), 4),
            '门店O2O履约贡献': round(o2o_contribution, 2),
        })
df_store_monthly = pd.DataFrame(store_monthly_rows)

# 库存明细:按品类拆分《门店月度经营》的库存金额,库存数量按品类成本均价反推
# 这样 门店库存(明细) 与 门店月度经营.库存金额 完全对得上
cat_unit_cost = {}
for cat in categories:
    cat_costs = [sku_cost[s] for s in skus if sku_category[s] == cat]
    cat_unit_cost[cat] = sum(cat_costs) / len(cat_costs)

inventory = []
for sid in store_ids:
    for ms in month_strs:
        total_amt = store_month_inventory[(sid, ms)]
        weights = [random.uniform(0.5, 1.5) for _ in categories]
        weight_sum = sum(weights)
        for idx, cat in enumerate(categories):
            if idx == len(categories) - 1:
                # 最后一个品类兜底,保证各品类金额合计与月度经营完全一致
                inv_amt = round(total_amt - sum(r['库存金额'] for r in inventory
                                                if r['门店ID'] == sid and r['月份'] == ms), 2)
            else:
                inv_amt = round(total_amt * weights[idx] / weight_sum, 2)
            unit = cat_unit_cost[cat] * random.uniform(0.95, 1.05)
            inv_qty = max(1, int(round(inv_amt / unit)))
            inventory.append({
                '月份': ms,
                '门店ID': sid,
                '品类': cat,
                '库存金额': inv_amt,
                '库存数量': inv_qty,
                '售罄率': round(random.uniform(0.30, 0.95), 4),
                '滞销金额': round(inv_amt * random.uniform(0, 0.30), 2),
                '滞销天数': random.randint(0, 120),
            })
df_inventory = pd.DataFrame(inventory)

# 费用(修复 6):与业务驱动因素挂钩
# 租金口径:门店费用记「租金及物业费」= 租约月租金 + 租约物业费,与租约明细逐店可核对
expenses = []
for i, sid in enumerate(store_ids):
    area = store_area[i]
    rent_and_property = round(lease[sid]['月租金'] + lease[sid]['物业费'], 2)
    for ms in month_strs:
        row = store_monthly_sales[(sid, ms)]
        sales = row['到店销售额']
        fulfill_cnt = int(
            o2o_ship_cnt.get((sid, ms), 0) + o2o_pickup_cnt.get((sid, ms), 0)
        )
        expenses.append({
            '月份': ms,
            '门店ID': sid,
            '租金及物业费': rent_and_property,
            '人力': round(sales * random.uniform(0.09, 0.16), 2),
            '营销': round(sales * random.uniform(0.01, 0.05), 2),
            '水电': round(area * random.uniform(8, 20) + sales * 0.002, 2),
            '履约包装': round(fulfill_cnt * random.uniform(2.5, 6.0) + 200, 2),
            '其他': round(sales * random.uniform(0.003, 0.01), 2),
            '折旧摊销': round(area * random.uniform(20, 60), 2),
        })
df_expenses = pd.DataFrame(expenses)

with pd.ExcelWriter(f'{OUTPUT_DIR}/2026H1门店经营台账.xlsx', engine='openpyxl') as writer:
    store_base.to_excel(writer, sheet_name='门店基础信息', index=False)
    df_store_monthly.to_excel(writer, sheet_name='门店月度经营', index=False)
    df_inventory.to_excel(writer, sheet_name='门店库存', index=False)
    df_expenses.to_excel(writer, sheet_name='门店费用', index=False)
print("  -> 完成")

# ===================== 3. 线上投放与平台费用 =====================
print("生成 2026H1线上投放与平台费用.xlsx ...")
# 收入/成本按 渠道×月份×品类 聚合,用于按佣金规则精确复算
grp = df_orders.groupby(['渠道ID', '月份', '品类']).agg(
    收入=('收入', 'sum'), 商品成本=('商品成本', 'sum'), 退款=('退款金额', 'sum')
).reset_index()
grp['佣金'] = grp.apply(
    lambda r: r['收入'] * commission_rules[(r['渠道ID'], r['品类'])][0], axis=1
)
grp['技服'] = grp.apply(
    lambda r: r['收入'] * commission_rules[(r['渠道ID'], r['品类'])][1], axis=1
)
agg_ch = grp.groupby(['渠道ID', '月份']).agg(
    收入=('收入', 'sum'), 退款=('退款', 'sum'), 商品成本=('商品成本', 'sum'),
    平台佣金=('佣金', 'sum'), 技术服务费=('技服', 'sum'),
).reset_index()
agg_ch['净收入'] = (agg_ch['收入'] - agg_ch['退款']).round(2)
agg_ch['毛利'] = (agg_ch['净收入'] - agg_ch['商品成本']).round(2)
agg_ch['毛利率'] = (agg_ch['毛利'] / agg_ch['净收入']).round(4)

order_cnt = df_orders.groupby(['渠道ID', '月份']).size().rename('订单量').reset_index()
ad_total = df_ad.groupby(['渠道ID', '日期']).agg(广告费=('投放金额', 'sum')).reset_index()
ad_total['月份'] = ad_total['日期'].apply(lambda d: f'{d.year}-{d.month:02d}')
ad_month = ad_total.groupby(['渠道ID', '月份'])['广告费'].sum().reset_index()

channel_fee = []
payment_rate = {'自营': 0.003, '平台': 0.006, '即时零售': 0.005}
fulfill_rate = {'自营': 0.035, '平台': 0.055, '即时零售': 0.090}
for ch in channel_ids:
    ctype = channel_types[ch]
    for (y, m) in months:
        ms = f'{y}-{m:02d}'
        revenue = float(agg_ch[(agg_ch['渠道ID'] == ch) & (agg_ch['月份'] == ms)]['收入'].sum())
        refund = float(agg_ch[(agg_ch['渠道ID'] == ch) & (agg_ch['月份'] == ms)]['退款'].sum())
        cost = float(agg_ch[(agg_ch['渠道ID'] == ch) & (agg_ch['月份'] == ms)]['商品成本'].sum())
        commission = float(agg_ch[(agg_ch['渠道ID'] == ch) & (agg_ch['月份'] == ms)]['平台佣金'].sum())
        tech = float(agg_ch[(agg_ch['渠道ID'] == ch) & (agg_ch['月份'] == ms)]['技术服务费'].sum())
        net = round(revenue - refund, 2)
        gross = round(net - cost, 2)
        cnt = int(order_cnt[(order_cnt['渠道ID'] == ch) & (order_cnt['月份'] == ms)]['订单量'].sum())
        ad = round(float(ad_month[(ad_month['渠道ID'] == ch) & (ad_month['月份'] == ms)]['广告费'].sum()), 2)
        commission = round(commission, 2)
        tech = round(tech, 2)
        payment = round(revenue * payment_rate[ctype], 2)
        fulfillment = round(revenue * fulfill_rate[ctype], 2)
        packaging = round(cnt * random.uniform(2.0, 5.0), 2)
        after_sale = round(refund * random.uniform(0.10, 0.20), 2)
        other = round(revenue * random.uniform(0.002, 0.006), 2)
        # 先对各分项四舍五入,再求和,保证「费用合计 = 各分项之和」精确成立
        fee_sum = round(ad + commission + tech + payment + fulfillment + packaging + after_sale + other, 2)
        channel_fee.append({
            '月份': ms,
            '渠道ID': ch,
            '渠道名称': channel_names[ch],
            '收入': round(revenue, 2),
            '退款': round(refund, 2),
            '净收入': net,
            '商品成本': round(cost, 2),
            '毛利': gross,
            '毛利率': round(gross / net, 4) if net else 0,
            '订单量': cnt,
            '客单价': round(net / cnt, 2) if cnt else 0,
            '广告费': ad,
            '平台佣金': commission,
            '技术服务费': tech,
            '支付手续费': payment,
            '履约费': fulfillment,
            '包装费': packaging,
            '退货售后费': after_sale,
            '其他费用': other,
            '费用合计': fee_sum,
            '渠道净利': round(gross - fee_sum, 2),
        })
df_channel_fee = pd.DataFrame(channel_fee)

with pd.ExcelWriter(f'{OUTPUT_DIR}/2026H1线上投放与平台费用.xlsx', engine='openpyxl') as writer:
    df_channel_fee.to_excel(writer, sheet_name='渠道月度费用', index=False)
    df_ad.to_excel(writer, sheet_name='投放明细', index=False)
    df_rules.to_excel(writer, sheet_name='平台佣金规则', index=False)
print("  -> 完成")

# ===================== 4. 会员与跨渠道订单 =====================
print("生成 2026H1会员与跨渠道订单.xlsx ...")
# 会员基础:注册先于首单;城市与注册门店一致;手机号脱敏唯一
phone_suffixes = np.random.choice(range(10000), size=len(member_ids), replace=False)
member_rows = []
member_plan = {}
for i, mid in enumerate(member_ids):
    status = random.choices(['活跃', '沉睡', '流失'], weights=[45, 30, 25])[0]
    if status == '活跃':
        last_purchase = rand_datetime(datetime(2026, 5, 1), end_date)
    elif status == '沉睡':
        # 61-120 天前(2026-06-30 回溯),与制度文档的沉睡口径严格一致
        last_purchase = rand_datetime(datetime(2026, 3, 2), datetime(2026, 4, 30))
    else:
        last_purchase = rand_datetime(datetime(2026, 1, 5), datetime(2026, 2, 28))
    reg_date = last_purchase - timedelta(days=random.randint(60, 900))
    if reg_date < datetime(2023, 1, 1):
        reg_date = rand_datetime(datetime(2023, 1, 1), datetime(2024, 12, 31))
    # 45% 会员在门店注册(线下),其余线上渠道注册
    if random.random() < 0.45:
        reg_store = random.choice(store_ids)
        reg_channel = STORE_CHANNEL_LABEL
        city = store_city[store_ids.index(reg_store)]
    else:
        reg_store = None
        reg_channel = random.choices(channel_ids, weights=channel_order_weights)[0]
        city = random.choice(all_cities)
    member_plan[mid] = {
        'status': status,
        'last_purchase': last_purchase,
        'reg_date': reg_date,
        'reg_store': reg_store,
        'reg_channel': reg_channel,
    }
    member_rows.append({
        '会员ID': mid,
        '注册日期': reg_date,
        '注册渠道': reg_channel,
        '手机号脱敏': f'138****{int(phone_suffixes[i]):04d}',
        '会员等级': None,          # 由累计消费分层后回填
        '城市': city,
        '注册门店ID': reg_store if reg_store else None,
        '状态': status,
    })

# 会员订单:每个会员至少 1 笔;订单类型与渠道ID/门店ID 一致
count_weights = np.array([0.42, 0.24, 0.14, 0.08, 0.05, 0.03, 0.02, 0.02])
counts = np.random.choice(range(1, 9), size=len(member_ids), p=count_weights)
diff = 6000 - counts.sum()
while diff != 0:
    j = random.randrange(len(member_ids))
    if diff > 0:
        counts[j] += 1
        diff -= 1
    elif counts[j] > 1:
        counts[j] -= 1
        diff += 1

member_orders = []
mo_counter = 800000
for mid, n in zip(member_ids, counts):
    plan = member_plan[mid]
    for k in range(int(n)):
        otype = random.choices(['线上', 'O2O', '线下'], weights=[45, 30, 25])[0]
        if otype == '线下':
            ch = None
            store = random.choice(store_ids)
        elif otype == 'O2O':
            ch = random.choices(channel_ids, weights=channel_order_weights)[0]
            store = random.choice(store_ids)
        else:
            ch = random.choices(channel_ids, weights=channel_order_weights)[0]
            store = None
        cat = random.choices(categories, weights=store_category_weights)[0]
        sku = random.choice(sellable_skus_by_cat[cat])
        qty = random.choices([1, 2, 3], weights=[70, 22, 8])[0]
        amt = round(sku_price[sku] * qty * random.uniform(0.85, 1.0), 2)
        # 订单日期:不晚于该会员最近购买日期,且不早于注册日期;
        # 最后一单固定落在"最近购买日期"上,保证会员状态(60/120 天口径)与最近购买一致
        if k == int(n) - 1:
            odate = plan['last_purchase']
        else:
            odate = rand_datetime(max(plan['reg_date'], start_date), plan['last_purchase'])
        member_orders.append({
            '订单号': f'O{mo_counter}',
            '会员ID': mid,
            '订单日期': odate,
            '渠道ID': ch,
            '门店ID': store,
            '品类': cat,
            '金额': amt,
            '是否首购': None,       # 按最早订单回填
            '是否复购': None,
            '订单类型': otype,
        })
        mo_counter += 1
df_member_orders = pd.DataFrame(member_orders)
df_member_orders = df_member_orders.sort_values(['会员ID', '订单日期']).reset_index(drop=True)
df_member_orders['是否首购'] = (~df_member_orders.duplicated('会员ID', keep='first')).map({True: '是', False: '否'})
df_member_orders['是否复购'] = (df_member_orders['是否首购'] == '否').map({True: '是', False: '否'})

# 跨渠道行为:派生自会员订单;与《会员基础·状态》口径统一
cross_rows = []
for mid in member_ids:
    sub = df_member_orders[df_member_orders['会员ID'] == mid]
    if len(sub) == 0:
        continue
    ch_order = sub['渠道ID'].fillna(STORE_CHANNEL_LABEL)
    first_ch = ch_order.iloc[0]
    distinct_ch = list(dict.fromkeys(ch_order.tolist()))
    repurchase_chs = [c for c in distinct_ch if c != first_ch]
    total = round(sub['金额'].sum(), 2)
    last_date = sub['订单日期'].max()
    cross_rows.append({
        '会员ID': mid,
        '首购渠道': first_ch,
        '复购渠道': ','.join(repurchase_chs) if repurchase_chs else '无',
        '跨渠道购买渠道数': len(distinct_ch),
        '订单数': len(sub),
        '最近购买日期': last_date,
        '累计消费': total,
        'RFM分': None,        # 稍后统一计算
        '流失标记': '是' if member_plan[mid]['status'] == '流失' else '否',
    })
df_cross = pd.DataFrame(cross_rows)

# RFM 分:由 R(最近购买)、F(订单数)、M(累计消费)综合计算,5 分最优
r_rank = df_cross['最近购买日期'].rank(pct=True)
f_rank = df_cross['订单数'].rank(pct=True)
m_rank = df_cross['累计消费'].rank(pct=True)
composite = 0.35 * m_rank + 0.35 * r_rank + 0.30 * f_rank
df_cross['RFM分'] = np.clip(1 + np.floor(composite * 5), 1, 5).astype(int)

# 会员等级:按累计消费分位分层
spend = df_cross.set_index('会员ID')['累计消费']
q90 = spend.quantile(0.90)
q70 = spend.quantile(0.70)
q40 = spend.quantile(0.40)


def grade_of(v):
    if v >= q90:
        return '钻石'
    if v >= q70:
        return '金卡'
    if v >= q40:
        return '银卡'
    return '普通'


level_map = {mid: grade_of(v) for mid, v in spend.items()}
member_base = pd.DataFrame(member_rows)
member_base['会员等级'] = member_base['会员ID'].map(level_map)

with pd.ExcelWriter(f'{OUTPUT_DIR}/2026H1会员与跨渠道订单.xlsx', engine='openpyxl') as writer:
    member_base.to_excel(writer, sheet_name='会员基础', index=False)
    df_member_orders.to_excel(writer, sheet_name='会员订单', index=False)
    df_cross.to_excel(writer, sheet_name='跨渠道行为', index=False)
print("  -> 完成")

# ===================== 5. 门店租约汇总.docx =====================
print("生成 华东区门店租约汇总.docx ...")
doc = DocxDocument()
doc.add_heading('华东区门店租约汇总(含其他大区)', level=0)
note_para = doc.add_paragraph()
note_run = note_para.add_run(
    '说明:本汇总含集团下属所有门店租约,部分门店位于华南、华北大区。'
    '门店区域与城市以附件《2026H1门店经营台账》为准。'
)
note_run.italic = True

for i, sid in enumerate(store_ids):
    lz = lease[sid]
    doc.add_heading(f'{sid} {store_names[i]}', level=2)
    fields = [
        ('门店ID', sid),
        ('门店名称', store_names[i]),
        ('区域', store_region[i]),
        ('城市', store_city[i]),
        ('租约ID', lz['租约ID']),
        ('出租方', '某某物业有限公司'),
        ('租期开始', lz['租期开始']),
        ('租期结束', lz['租期结束']),
        ('免租期', lz['免租期']),
        ('月租金', f"{lz['月租金']:.2f}元"),
        ('物业费', f"{lz['物业费']:.2f}元/月"),
        ('递增条款', lz['递增条款']),
        ('闭店条件', f"提前90天书面通知,支付{lz['违约金']:.2f}元违约金"),
        ('续租优先权', '有'),
        ('备注', lz['备注']),
    ]
    for label, value in fields:
        p = doc.add_paragraph()
        run_label = p.add_run(f'{label}:')
        run_label.bold = True
        p.add_run(f' {value}')

doc.save(f'{OUTPUT_DIR}/华东区门店租约汇总.docx')
print("  -> 完成")

# ===================== 6. 渠道与门店管理策略.docx =====================
print("生成 渠道与门店管理策略.docx ...")
doc2 = DocxDocument()
doc2.add_heading('渠道与门店管理策略', level=0)
doc2.add_paragraph('版本:v2026.09  适用范围:集团下属所有门店与渠道')

doc2.add_heading('1. 渠道分类与优先级', level=1)
doc2.add_paragraph('1.1 自营 APP、微信小程序为战略渠道,优先保障资源。')
doc2.add_paragraph('1.2 平台渠道按利润贡献分级:A 级(毛利率≥35%)、B 级(25%-35%)、C 级(<25%)。')
doc2.add_paragraph('1.3 即时零售渠道定位为履约补充,重点考核履约成本与时效。')

doc2.add_heading('2. 门店分级标准', level=1)
doc2.add_paragraph('2.1 到店坪效 = 月均到店销售额 ÷ 营业面积;'
                  '租售比 = 月均租金 ÷ 月均到店销售额。')
doc2.add_paragraph('2.2 门店分级按到店坪效:A 级 ≥5000 元/㎡/月;B 级 2000-5000 元/㎡/月;'
                  'C 级 1000-2000 元/㎡/月;D 级 <1000 元/㎡/月。')
doc2.add_paragraph('2.3 租售比 >35% 的门店,不参与 2.2 分级,直接列为 D 级。')
doc2.add_paragraph('2.4 D 级门店列为整改或闭店候选,处置动作按第 6 章执行。')
doc2.add_paragraph('2.5 计划期内发生面积调整的门店(台账中以"面积是否调整=是"标识),'
                  '其按当前面积计算的坪效口径可能失真,须在《异常与待核清单》中披露。')

doc2.add_heading('3. O2O 履约规则', level=1)
doc2.add_paragraph('3.1 门店发货、到店自提订单计入线上渠道收入,同时按发货/自提归属计入门店履约贡献。')
doc2.add_paragraph('3.2 即时配送订单归属即时零售渠道,履约成本由渠道承担,门店作为履约节点。')
doc2.add_paragraph('3.3 快递订单由中心仓发货,不计入门店履约,该部分订单不关联门店。')
doc2.add_paragraph('3.4 门店发货、到店自提、即时配送订单必须关联门店ID;快递订单不关联门店ID。')
doc2.add_paragraph('3.5 门店履约贡献 = 该门店当月"门店发货 + 到店自提"订单的净收入合计,'
                  '计入《门店经营台账》"门店O2O履约贡献"字段。')
doc2.add_paragraph('3.6 "线上订单履约单量" = 自提单量 + 门店发货单量,与渠道订单明细一致。')
doc2.add_paragraph('3.7 门店仅承担履约侧的人工与包装成本(对应《门店费用》中的"履约包装"字段)。')

doc2.add_heading('4. 数据来源优先级', level=1)
doc2.add_paragraph('4.1 财务确认收入优先于业务系统数据。')
doc2.add_paragraph('4.2 冲突时以渠道销售明细和门店台账交叉验证。')
doc2.add_paragraph('4.3 无法确认时标记"待补",不得用假设补齐事实。')
doc2.add_paragraph('4.4 租金口径:合同条款(租期、月租金、物业费、递增、违约金、免租期)以'
                  '《华东区门店租约汇总》为准;损益与经营分析(租售比、门店盈亏)以'
                  '《门店经营台账·门店费用》的"租金及物业费"为准,该字段 = 合同月租金 + 物业费。'
                  '两者不一致时按 4.3 标记待补,不得自行假设。')
doc2.add_paragraph('4.5 租金口径补充:台账"租金及物业费"为合同基数(月租金 + 物业费),'
                  '不含按租约"递增条款"产生的年度递增;涉及 2026H2 的租金测算,'
                  '须依据租约递增条款与租期开始日推算的周年日另行测算。')
doc2.add_paragraph('4.6 客单价口径:渠道客单价 = 净收入 ÷ 订单量;'
                  '门店客单价 = 到店销售额 ÷ 订单数;两者分子口径不同,不可直接横向比较。')
doc2.add_paragraph('4.7 到店业务口径:到店销售额、订单数、毛利额、进店客流、连带率、退货额、'
                  '动销SKU数、缺货率、损耗率以《门店经营台账》汇总值为准,'
                  '材料不提供到店订单级明细,不要求作订单级核验。')
doc2.add_paragraph('4.8 费用口径:门店费用中"人力、营销、水电、履约包装、其他、折旧摊销"'
                  '只到费用级,不得据以反推编制人数、资源用量或资产清单;'
                  '渠道费用中"履约费、包装费、退货售后费、支付手续费、其他费用"'
                  '按实际发生额披露,无费率规则,不作独立复算。')
doc2.add_paragraph('4.9 损益范围:本次损益仅含渠道侧与门店侧两级,'
                  '不含中心仓运营成本、干线物流成本与总部分摊。')
doc2.add_paragraph('4.10 退款成本口径:全额退款订单的商品成本按不可回收处理,不作成本冲回。')
doc2.add_paragraph('4.11 数据提取时点:本次数据的提取时点为 2026-06-25,'
                  '2026-06 订单在 6 月之后发生的退款尚未入账;'
                  '2026-06 的退款完整性不成立,该月退款率与渠道净收入只作参考,不得用于跨月排名。')
doc2.add_paragraph('4.12 比对容差:金额类指标复算比对的容差为 ±0.01 元。')

doc2.add_heading('5. 缺失与异常处理', level=1)
doc2.add_paragraph('5.1 允许披露、待补、缩小范围、局部隔离、负向结论。')
doc2.add_paragraph('5.2 不得虚构事实,不得用外部数据替代题内材料。')

doc2.add_heading('6. 闭店与改造条件', level=1)
doc2.add_paragraph('6.1 租约到期且连续 6 个月 D 级,可启动闭店评估。'
                  '其中"到期"指租期结束日不晚于建议执行期末(2026-12-31);'
                  '基准日前已到期但仍在营的门店,须先核实续约状态再行处置。')
doc2.add_paragraph('6.2 租售比连续 3 个月 >35%,可转为前置仓或自提点。')
doc2.add_paragraph('6.3 连续月度判定按当月指标计算,不得用半年均值替代。')
doc2.add_paragraph('6.4 闭店成本包含租约约定的违约金,并纳入决策依据。')

doc2.add_heading('7. 促销与投放效果口径', level=1)
doc2.add_paragraph('7.1 活动 ROI = 增量毛利 ÷ 活动实际费用;'
                  '分母"活动实际费用"指《促销活动记录》中该活动的实际费用,'
                  '不含同期媒介投放;如需含投放的全成本口径,须另行测算并明确标注。')
doc2.add_paragraph('7.2 增量销售额 = 活动期间销售额 − 活动前销售额;'
                  '增量毛利 = 增量销售额 × 活动毛利率。'
                  '"活动期间销售额"为该渠道在活动起止日期内的订单收入合计(活动清单的"门店ID"'
                  '为活动发起门店,不改变聚合口径);'
                  '"活动前销售额"为该渠道"非活动期日均收入 × 活动天数",'
                  '其中非活动期指 H1 内不属于该渠道任何活动窗口的日期;'
                  '同一渠道存在并行活动时,同一笔订单会被多个活动重复计入期间销售额,须在分析中说明。')
doc2.add_paragraph('7.3 投放 ROI = 投放带来的支付金额 ÷ 投放金额。')
doc2.add_paragraph('7.4 折扣率 = 1 − 活动价 ÷ 原价;活动价不得低于单位成本。')
doc2.add_paragraph('7.5 活动预算按季度审批:实际费用超预算 10% 以内的由营销负责人核准,'
                  '超过 10% 的须报分管副总审批;超预算活动须在《异常与待核清单》中列示。')
doc2.add_paragraph('7.6 平台佣金与技术服务费按退款前成交收入(即《渠道销售明细》的"收入"字段)计提,'
                  '退款不改变已计提金额。')

doc2.add_heading('8. 商品与品类等级', level=1)
doc2.add_paragraph('8.1 SKU 毛利率 = (标准标价 − 单位成本) ÷ 标准标价。')
doc2.add_paragraph('8.2 品类等级:A 级(毛利率≥40%)、B 级(25%-40%)、C 级(<25%)。')
doc2.add_paragraph('8.3 停售 SKU 不得产生新增销售;季节性 SKU 仅在适销期销售。')

doc2.add_heading('9. 会员与订单口径', level=1)
doc2.add_paragraph('9.1 会员注册渠道:线上渠道注册记 CH01-CH06;门店注册记为"门店",'
                  '并填写注册门店ID,会员城市取注册门店所在城市。')
doc2.add_paragraph('9.2 会员状态以基准日 2026-06-30 计算:最近购买 ≤60 天为活跃,'
                  '61-120 天为沉睡,>120 天为流失;两表中流失口径一致。')
doc2.add_paragraph('9.3 RFM 分由最近购买(R)、订单数(F)、累计消费(M)综合排名得出,1-5 分,5 分最高。')
doc2.add_paragraph('9.4 跨渠道购买渠道数 = 该会员购买过的不同渠道数(含门店);'
                  'O2O 订单计入其渠道ID,不另计门店渠道。')
doc2.add_paragraph('9.5 会员订单类型:线下(渠道ID 为空、门店ID 必填)、线上(渠道ID 必填、'
                  '门店ID 为空)、O2O(渠道ID 与门店ID 均必填)。')
doc2.add_paragraph('9.6 会员累计消费以《会员订单》为口径;'
                  '《渠道销售明细》的会员标识仅用于渠道归因,不重复计入会员累计消费。')
doc2.add_paragraph('9.7 "复购渠道"为该会员首购渠道之外、实际发生购买的其他渠道列表;'
                  '同渠道复购不记入该字段;"复购渠道=无"表示没有首购渠道之外的购买,'
                  '不等于该会员没有复购订单。')
doc2.add_paragraph('9.8 会员订单与渠道销售明细为两套独立编号体系:'
                  '会员订单用于会员口径统计,渠道销售明细用于渠道经营归因,两者不作逐笔勾稽;'
                  '若需按月对照,须说明两表覆盖范围差异。')

doc2.save(f'{OUTPUT_DIR}/渠道与门店管理策略.docx')
print("  -> 完成")

# ===================== 7. 促销活动记录.xlsx =====================
print("生成 2026H1促销活动记录.xlsx ...")
with pd.ExcelWriter(f'{OUTPUT_DIR}/2026H1促销活动记录.xlsx', engine='openpyxl') as writer:
    df_activities.to_excel(writer, sheet_name='活动清单', index=False)
    df_details.to_excel(writer, sheet_name='活动明细', index=False)
    df_effects.to_excel(writer, sheet_name='活动效果', index=False)
print("  -> 完成")

# ===================== 8. 旧版门店清单_2024.xlsx(噪音)=====================
print("生成 旧版门店清单_2024.xlsx ...")
old_stores = []
for i, sid in enumerate(store_ids):
    # 修复 19:仍在营门店的区域/城市/名称与当前台账一致,仅状态与存续不同
    if i in [3, 9, 17, 26, 33]:
        status = '调整'
    else:
        status = '营业'
    old_stores.append({
        '门店ID': sid,
        '门店名称': store_names[i],
        '区域': store_region[i],
        '城市': store_city[i],
        '状态': status,
        '关闭日期': None,
    })
# 2024 年已关闭、当前不再存在的 8 家历史门店
for k, sid in enumerate([f'S{str(i).zfill(3)}' for i in range(43, 51)]):
    region = random.choice(regions)
    old_stores.append({
        '门店ID': sid,
        '门店名称': f'门店{sid[1:]}',
        '区域': region,
        '城市': random.choice(region_cities[region]),
        '状态': '关闭',
        '关闭日期': rand_datetime(datetime(2024, 1, 1), datetime(2024, 12, 31)),
    })
df_old = pd.DataFrame(old_stores)
with pd.ExcelWriter(f'{OUTPUT_DIR}/旧版门店清单_2024.xlsx', engine='openpyxl') as writer:
    df_old.to_excel(writer, sheet_name='门店清单', index=False)
print("  -> 完成")

# ===================== 9. SKU 主数据.xlsx =====================
print("生成 SKU主数据.xlsx ...")
df_sku_master.to_excel(f'{OUTPUT_DIR}/SKU主数据.xlsx', sheet_name='SKU主数据', index=False)
print("  -> 完成")

print("\n" + "=" * 50)
print(f"全部 9 个附件生成完成!输出目录:{OUTPUT_DIR}")
print("=" * 50)
