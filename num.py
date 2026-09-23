"""
Magnet 出题 - 线下/线上零售渠道与门店策略附件生成器
v2.0 修复版 - 2026-09-21

修复内容(相对 v1.0):
1. 门店到店销售量级修正(442 元/月 → 5-80 万/月)
2. 门店费用量级修正(人力 5-30 元 → 3-8 万元)
3. 修复销售 vs 订单数 不同步异常(19 行 → 5 行可控)
4. 修复库存数量 vs 库存金额 逻辑矛盾(65% → 5%)
5. 投放明细量级修正(1857 元 → 接近月度广告费)
6. 区域从 1 个扩展到 3 个
7. 新增「门店 O2O 履约贡献」字段,反映真实门店收入结构
8. 会员订单号范围调整,避免与渠道订单号撞号
"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from docx import Document as DocxDocument
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.units import cm

# ===================== 全局配置 =====================
# 输出目录:所有附件统一输出到 input/ 子目录(模拟真实业务数据导出)
import os
OUTPUT_DIR = "input"
os.makedirs(OUTPUT_DIR, exist_ok=True)

random.seed(42)
np.random.seed(42)

# 中文支持
# 中文支持:使用 ReportLab 内置 CID 字体,无需系统字体,跨平台稳定
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
FONT_NAME = 'STSong-Light'

# 渠道定义(不变)
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

# 渠道订单权重:抖音 + 天猫合计 60%,符合当前零售实际流量分布
# CH01 自营APP 10%, CH02 微信小程序 8%, CH03 天猫旗舰店 30%,
# CH04 京东自营 12%, CH05 抖音小店 30%, CH06 美团闪购 10%
channel_order_weights = [0.10, 0.08, 0.30, 0.12, 0.30, 0.10]

# 【修复 7】三个区域,不再只有华东
region_cities = {
    '华东': ['上海', '杭州', '南京', '苏州', '宁波', '合肥', '无锡', '常州'],
    '华南': ['广州', '深圳', '厦门', '福州', '南宁'],
    '华北': ['北京', '天津', '济南', '青岛', '石家庄'],
}
regions = list(region_cities.keys())
cities = sum(region_cities.values(), [])
business_districts = ['核心商圈', '社区商圈', '交通枢纽', '办公区', '郊区']

# 42 家门店(分配到三个区域:华东 18、华南 12、华北 12)
store_region_assign = (
    ['华东'] * 18 + ['华南'] * 12 + ['华北'] * 12
)
assert len(store_region_assign) == 42

store_ids = [f'S{str(i).zfill(3)}' for i in range(1, 43)]
store_names = [f'门店{str(i).zfill(2)}' for i in range(1, 43)]
store_region = store_region_assign
store_city = [random.choice(region_cities[r]) for r in store_region]
store_district = [random.choice(business_districts) for _ in range(42)]
store_area = [random.randint(80, 500) for _ in range(42)]
# 开业日期 2018-2025 之间
store_open_date = [datetime(2018,1,1) + timedelta(days=random.randint(0, 2500)) for _ in range(42)]
store_type = [random.choice(['标准店', '社区店', '旗舰店', '前置仓']) for _ in range(42)]
store_status = ['营业'] * 42

# 品类
categories = ['食品', '日用', '美妆', '家电', '服饰', '母婴']
skus = [f'SKU{str(i).zfill(4)}' for i in range(1, 201)]
sku_category = {sku: random.choice(categories) for sku in skus}
# 【修复 1 配套】调整 SKU 价格区间,让客单价更合理
sku_price = {sku: round(random.uniform(50, 2000), 2) for sku in skus}
sku_cost = {sku: round(sku_price[sku] * random.uniform(0.4, 0.75), 2) for sku in skus}

# 日期范围
start_date = datetime(2026, 1, 1)
end_date = datetime(2026, 6, 30)
date_range = pd.date_range(start_date, end_date, freq='D')
months = [(2026, m) for m in range(1, 7)]

# ===================== 1. 渠道销售明细 =====================
print("生成 2026H1渠道销售明细.xlsx ...")
orders = []
order_id_counter = 100000
# 按权重预生成 7500 个订单的渠道
order_channels = random.choices(channel_ids, weights=channel_order_weights, k=7500)
for idx in range(7500):
    order_date = random.choice(date_range)
    channel = order_channels[idx]
    store_id = random.choice(store_ids) if random.random() > 0.15 else None
    member_id = f'M{random.randint(10000, 99999)}' if random.random() > 0.2 else None
    sku = random.choice(skus)
    cat = sku_category[sku]
    qty = random.randint(1, 5)
    list_price = sku_price[sku]
    discount = random.uniform(0.7, 1.0) if random.random() > 0.3 else 1.0
    deal_price = round(list_price * discount, 2)
    revenue = round(deal_price * qty, 2)
    cost = round(sku_cost[sku] * qty, 2)
    refund = 0.0
    refund_qty = 0
    status = '已完成'
    if random.random() < 0.08:
        refund_qty = random.randint(1, qty)
        refund = round(deal_price * refund_qty, 2)
        status = '部分退款' if refund_qty < qty else '已退款'
    ship_method = random.choice(['快递', '门店发货', '到店自提', '即时配送'])
    if ship_method in ['门店发货', '到店自提'] and store_id is None:
        store_id = random.choice(store_ids)
    channel_attr = channel
    order_id = f'O{order_id_counter}'
    order_id_counter += 1
    orders.append({
        '订单日期': order_date,
        '订单号': order_id,
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
        '渠道归属': channel_attr,
        '月份': f'{order_date.year}-{order_date.month:02d}'
    })

df_orders = pd.DataFrame(orders)

# 渠道退款明细
refund_records = df_orders[df_orders['退款金额'] > 0].copy()
refund_records['退款日期'] = refund_records['订单日期'] + pd.to_timedelta(
    np.random.randint(1, 30, len(refund_records)), unit='D'
)
refund_records['退款原因'] = np.random.choice(
    ['七天无理由', '质量问题', '错拍', '物流损坏', '其他'], len(refund_records)
)
refund_records['是否影响毛利'] = '是'
df_refund = refund_records[['订单号','渠道ID','退款日期','退款原因','退款金额','是否影响毛利']]

# 渠道月度汇总
df_orders['净收入'] = df_orders['收入'] - df_orders['退款金额']
monthly = df_orders.groupby(['月份','渠道ID','渠道名称']).agg(
    收入=('收入','sum'),
    退款=('退款金额','sum'),
    净收入=('净收入','sum'),
    商品成本=('商品成本','sum'),
    订单量=('订单号','count'),
).reset_index()
monthly['毛利'] = monthly['净收入'] - monthly['商品成本']
monthly['客单价'] = (monthly['净收入'] / monthly['订单量']).round(2)

with pd.ExcelWriter(f'{OUTPUT_DIR}/2026H1渠道销售明细.xlsx', engine='openpyxl') as writer:
    df_orders.to_excel(writer, sheet_name='渠道销售明细', index=False)
    df_refund.to_excel(writer, sheet_name='渠道退款明细', index=False)
    monthly.to_excel(writer, sheet_name='渠道月度汇总', index=False)
print("  -> 完成")

# ===================== 2. 门店经营台账 =====================
print("生成 2026H1门店经营台账.xlsx ...")

# 基础信息
store_base = pd.DataFrame({
    '门店ID': store_ids,
    '门店名称': store_names,
    '区域': store_region,
    '城市': store_city,
    '商圈类型': store_district,
    '开业日期': store_open_date,
    '营业面积': store_area,
    '租约ID': [f'L{str(i).zfill(3)}' for i in range(1, 43)],
    '店型': store_type,
    '状态': store_status,
    '是否同店口径': ['是'] * 42
})
# 3 家"非同店"(因面积调整)
for idx in [5, 12, 30]:
    store_base.loc[idx, '是否同店口径'] = '否'

# 月度经营 + 【修复 1, 3, 8】三个修复一起处理
monthly_store = []
# 先聚合每个门店每月的 O2O 履约贡献(从渠道销售明细中"门店发货 + 自提"订单的金额聚合)
# 这是门店的真实收入贡献之二
o2o_orders = df_orders[df_orders['发货方式'].isin(['门店发货', '到店自提'])].copy()
o2o_by_store_month = o2o_orders.groupby(['门店ID', '月份'])['净收入'].sum().reset_index()
o2o_by_store_month.columns = ['门店ID', '月份', '门店O2O履约贡献']

for sid in store_ids:
    area = store_base.loc[store_base['门店ID']==sid, '营业面积'].values[0]
    # 给每家门店一个"基础销售能力"等级,模拟不同门店的真实表现差异
    # 大部分门店正常(5-80 万),少数低(2-5 万,问题店),少数高(80-150 万,明星店)
    ability = np.random.choice(
        ['问题店', '正常店', '明星店'],
        p=[0.15, 0.70, 0.15]
    )
    if ability == '问题店':
        sales_base = random.uniform(5000, 30000)  # 5-30 万/月
    elif ability == '明星店':
        sales_base = random.uniform(80000, 150000)  # 80-150 万/月
    else:
        sales_base = random.uniform(30000, 80000)  # 3-8 万... 不,改成 30-80 万

    # 修正:统一改成 5-150 万区间,避免数学计算混乱
    if ability == '问题店':
        sales_base = random.uniform(5, 30)  # 万
    elif ability == '明星店':
        sales_base = random.uniform(80, 150)
    else:
        sales_base = random.uniform(30, 80)

    for (y, m) in months:
        month_str = f'{y}-{m:02d}'
        # 到店销售:在该店能力基础上 × 月度波动 × 面积系数
        monthly_factor = random.uniform(0.85, 1.15)
        area_factor = 0.6 + (area / 500) * 0.8  # 80㎡→0.73, 500㎡→1.40
        sales = round(sales_base * monthly_factor * area_factor * 10000, 2)  # 转为元

        # 【修复 3】销售和订单数严格挂钩
        aov_min, aov_max = 80, 200
        aov = random.uniform(aov_min, aov_max)
        orders = max(1, int(sales / aov))
        # 重新校准 sales 保证与 orders 一致
        sales = round(orders * aov, 2)
        aov = round(aov, 2)

        # 客流 = 订单 / 成交率(成交率 10%-30%)
        conv_rate = random.uniform(0.10, 0.30)
        traffic = max(orders, int(orders / conv_rate))
        conv = round(orders / traffic, 4) if traffic > 0 else 0
        attach = round(random.uniform(1.2, 2.5), 2)

        # 毛利:25%-45%
        gross_rate = random.uniform(0.25, 0.45)
        gross = round(sales * gross_rate, 2)
        gross_margin = round(gross / sales, 4) if sales > 0 else 0

        # O2O 履约单量与销售额
        online_orders = int(orders * random.uniform(0.5, 2.5))  # O2O 单量通常远大于到店
        pickup = int(online_orders * random.uniform(0.3, 0.6))
        ship_from_store = online_orders - pickup

        # 从 o2o_by_store_month 取该门店该月的 O2O 履约贡献金额
        o2o_row = o2o_by_store_month[
            (o2o_by_store_month['门店ID'] == sid) &
            (o2o_by_store_month['月份'] == month_str)
        ]
        o2o_contribution = float(o2o_row['门店O2O履约贡献'].values[0]) if len(o2o_row) > 0 else 0.0

        return_amt = round(sales * random.uniform(0.01, 0.08), 2)
        inventory = round(sales * random.uniform(0.5, 1.5), 2)
        turnover_days = round(random.uniform(20, 90), 1)
        active_sku = random.randint(80, 180)
        stockout_rate = round(random.uniform(0.01, 0.15), 4)
        loss_rate = round(random.uniform(0.005, 0.05), 4)

        monthly_store.append({
            '月份': month_str,
            '门店ID': sid,
            '到店销售额': sales,  # 【修复 1】改名,与"门店O2O履约贡献"区分
            '订单数': orders,
            '客单价': aov,
            '进店客流': traffic,
            '成交率': conv,
            '连带率': attach,
            '毛利额': gross,
            '毛利率': gross_margin,
            '线上订单履约单量': online_orders,
            '自提单量': pickup,
            '门店发货单量': ship_from_store,
            '退货额': return_amt,
            '库存金额': inventory,
            '库存周转天数': turnover_days,
            '动销SKU数': active_sku,
            '缺货率': stockout_rate,
            '损耗率': loss_rate,
            '门店O2O履约贡献': round(o2o_contribution, 2),  # 【修复 8】新增字段
        })

df_store_monthly = pd.DataFrame(monthly_store)

# 库存 - 【修复 4】让库存数量与金额逻辑一致
inventory = []
# 各类目预设平均单价(用于反推库存数量)
cat_unit_price = {
    '食品': 25, '日用': 35, '美妆': 120,
    '家电': 800, '服饰': 200, '母婴': 80,
}
for sid in store_ids:
    for (y, m) in months:
        for cat in categories:
            inv_amt = round(random.uniform(500, 8000), 2)  # 库存金额 500-8000 元
            unit = cat_unit_price[cat] * random.uniform(0.7, 1.3)
            inv_qty = max(0, int(inv_amt / unit))  # 【修复 4】按真实单价反推
            # 保证 inv_qty 与 inv_amt 逻辑一致
            if inv_qty == 0:
                # 如果按金额推不出整数件数,补一个最小件数
                inv_qty = max(1, int(inv_amt / (unit * 2)))
            sell_through = round(random.uniform(0.3, 0.95), 4)
            slow_amt = round(inv_amt * random.uniform(0, 0.3), 2)
            slow_days = random.randint(0, 180)
            inventory.append({
                '月份': f'{y}-{m:02d}',
                '门店ID': sid,
                '品类': cat,
                '库存金额': inv_amt,
                '库存数量': inv_qty,
                '售罄率': sell_through,
                '滞销金额': slow_amt,
                '滞销天数': slow_days
            })
df_inventory = pd.DataFrame(inventory)

# 费用 - 【修复 2】真实费用量级
expenses = []
for sid in store_ids:
    area = store_base.loc[store_base['门店ID']==sid, '营业面积'].values[0]
    base_rent = area * random.uniform(80, 200)
    for (y, m) in months:
        rent = round(base_rent * random.uniform(0.95, 1.05), 2)
        # 人力:3-8 万/月(按面积缩放)
        labor = round(random.uniform(30000, 80000) * (0.7 + area/500*0.6), 2)
        marketing = round(random.uniform(2000, 15000), 2)
        utility = round(random.uniform(2000, 8000), 2)
        fulfillment = round(random.uniform(1000, 5000), 2)
        other = round(random.uniform(500, 3000), 2)
        depreciation = round(random.uniform(3000, 10000), 2)
        expenses.append({
            '月份': f'{y}-{m:02d}',
            '门店ID': sid,
            '租金': rent,
            '人力': labor,
            '营销': marketing,
            '水电': utility,
            '履约包装': fulfillment,
            '其他': other,
            '折旧摊销': depreciation
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
channel_fee = []
for ch in channel_ids:
    for (y, m) in months:
        revenue = df_orders[
            (df_orders['渠道ID']==ch) &
            (df_orders['月份']==f'{y}-{m:02d}')
        ]['收入'].sum()
        gross = revenue * random.uniform(0.3, 0.5)
        ad = round(revenue * random.uniform(0.05, 0.25), 2)
        commission = round(revenue * random.uniform(0.02, 0.1), 2)
        tech = round(revenue * random.uniform(0.005, 0.03), 2)
        payment = round(revenue * random.uniform(0.003, 0.01), 2)
        fulfillment = round(revenue * random.uniform(0.02, 0.12), 2)
        packaging = round(revenue * random.uniform(0.005, 0.02), 2)
        after_sale = round(revenue * random.uniform(0.005, 0.04), 2)
        other = round(revenue * random.uniform(0.001, 0.01), 2)
        channel_fee.append({
            '月份': f'{y}-{m:02d}',
            '渠道ID': ch,
            '渠道名称': channel_names[ch],
            '广告费': ad,
            '平台佣金': commission,
            '技术服务费': tech,
            '支付手续费': payment,
            '履约费': fulfillment,
            '包装费': packaging,
            '退货售后费': after_sale,
            '其他费用': other,
            '收入': revenue,
            '毛利': gross
        })
df_channel_fee = pd.DataFrame(channel_fee)

# 【修复 5】投放明细量级修正
# 渠道月度广告费半年合计约 352 万,180 行投放明细应接近这个量级
# 即平均每条 2 万元左右
投放 = []
# 按权重预生成 180 条投放的渠道
ad_channels = random.choices(channel_ids, weights=channel_order_weights, k=180)
for idx in range(180):
    ch = ad_channels[idx]
    date = random.choice(date_range)
    activity_id = f'A{random.randint(1, 50):03d}'
    amount = round(random.uniform(5000, 30000), 2)  # 5000-30000 元/行
    exposure = int(amount * random.uniform(500, 2000))
    click = int(exposure * random.uniform(0.01, 0.08))
    add_cart = int(click * random.uniform(0.1, 0.4))
    pay_orders = int(add_cart * random.uniform(0.2, 0.6))
    roi = round(random.uniform(0.5, 5), 2)
    投放.append({
        '日期': date,
        '渠道ID': ch,
        '投放项目': f'项目{random.randint(1,20)}',
        '活动ID': activity_id,
        '投放金额': amount,
        '曝光': exposure,
        '点击': click,
        '加购': add_cart,
        '支付订单': pay_orders,
        'ROI': roi
    })
df_ad = pd.DataFrame(投放)

# 平台佣金规则
rules = []
for ch in channel_ids:
    for cat in categories:
        rules.append({
            '渠道ID': ch,
            '品类': cat,
            '佣金率': round(random.uniform(0.02, 0.12), 4),
            '技术服务费率': round(random.uniform(0.005, 0.03), 4),
            '结算周期': random.choice(['T+1', 'T+7', 'T+15', '月结']),
            '备注': ''
        })
df_rules = pd.DataFrame(rules)

with pd.ExcelWriter(f'{OUTPUT_DIR}/2026H1线上投放与平台费用.xlsx', engine='openpyxl') as writer:
    df_channel_fee.to_excel(writer, sheet_name='渠道月度费用', index=False)
    df_ad.to_excel(writer, sheet_name='投放明细', index=False)
    df_rules.to_excel(writer, sheet_name='平台佣金规则', index=False)
print("  -> 完成")

# ===================== 4. 会员与跨渠道订单 =====================
print("生成 2026H1会员与跨渠道订单.xlsx ...")
member_ids = [f'M{str(i).zfill(6)}' for i in range(1, 2501)]
member_base = pd.DataFrame({
    '会员ID': member_ids,
    '注册日期': [start_date - timedelta(days=random.randint(0, 1000)) for _ in member_ids],
    '注册渠道': random.choices(channel_ids, weights=channel_order_weights, k=len(member_ids)),
    '手机号脱敏': [f'138****{random.randint(1000,9999)}' for _ in member_ids],
    '会员等级': [random.choice(['普通','银卡','金卡','钻石']) for _ in member_ids],
    '城市': [random.choice(cities) for _ in member_ids],
    '注册门店ID': [random.choice(store_ids) if random.random()>0.3 else None for _ in member_ids],
    '状态': [random.choice(['活跃','沉睡','流失']) for _ in member_ids]
})

# 【修复 6】会员订单号范围调整到 O800000+,避免与渠道订单号 O100000-107500 撞号
# 改为自增,避免随机冲突
mo_counter = 800000
# 按权重预生成 6000 个会员订单的渠道
member_order_channels = random.choices(channel_ids, weights=channel_order_weights, k=6000)
member_orders = []
for idx in range(6000):
    mid = random.choice(member_ids)
    oid = f'O{mo_counter}'
    mo_counter += 1
    ch = member_order_channels[idx]
    store = random.choice(store_ids) if random.random()>0.2 else None
    cat = random.choice(categories)
    # 调整金额范围,与渠道销售客单价匹配(渠道客单价 ~2700)
    amt = round(random.uniform(100, 3000), 2)
    first = random.random() < 0.15
    repurchase = not first
    member_orders.append({
        '订单号': oid,
        '会员ID': mid,
        '订单日期': random.choice(date_range),
        '渠道ID': ch,
        '门店ID': store,
        '品类': cat,
        '金额': amt,
        '是否首购': '是' if first else '否',
        '是否复购': '是' if repurchase else '否',
        '订单类型': random.choice(['线上', '线下', 'O2O'])
    })
df_member_orders = pd.DataFrame(member_orders)

cross = []
for mid in member_ids:
    orders = df_member_orders[df_member_orders['会员ID']==mid]
    if len(orders) == 0:
        continue
    first_ch = orders.sort_values('订单日期').iloc[0]['渠道ID']
    repurchase_chs = orders['渠道ID'].unique()
    cross_count = len(repurchase_chs)
    last_date = orders['订单日期'].max()
    total = orders['金额'].sum()
    rfm = random.randint(1,5)
    churn = '是' if random.random() < 0.2 else '否'
    cross.append({
        '会员ID': mid,
        '首购渠道': first_ch,
        '复购渠道': ','.join(repurchase_chs),
        '跨渠道购买次数': cross_count,
        '最近购买日期': last_date,
        '累计消费': round(total, 2),
        'RFM分': rfm,
        '流失标记': churn
    })
df_cross = pd.DataFrame(cross)

with pd.ExcelWriter(f'{OUTPUT_DIR}/2026H1会员与跨渠道订单.xlsx', engine='openpyxl') as writer:
    member_base.to_excel(writer, sheet_name='会员基础', index=False)
    df_member_orders.to_excel(writer, sheet_name='会员订单', index=False)
    df_cross.to_excel(writer, sheet_name='跨渠道行为', index=False)
print("  -> 完成")

# ===================== 5. 华东区门店租约汇总.docx =====================
print("生成 华东区门店租约汇总.docx ...")
# 注:虽然文件名是"华东区",但实际包含所有三个区域的门店(为了与门店台账对齐)
# 这是有意的"现实业务常见命名偏差",专家可识别并归类
doc = DocxDocument()

# 标题
title = doc.add_heading('华东区门店租约汇总(含其他大区)', level=0)

# 说明
note_para = doc.add_paragraph()
run = note_para.add_run('说明:本汇总含集团下属所有门店租约,部分门店位于华南、华北大区。区域分布详见附件 1《2026H1 门店经营台账》。')
run.italic = True

for i, sid in enumerate(store_ids):
    lease_id = f'L{str(i+1).zfill(3)}'
    area = store_area[i]
    rent_base = area * random.uniform(80, 200)
    start = datetime(2020, 1, 1) + timedelta(days=random.randint(0, 1500))
    end = start + timedelta(days=365 * random.randint(3, 8))
    escalation = random.choice(['无', '每年递增3%', '每年递增5%', '每两年递增8%'])
    free_rent = random.randint(0, 90)
    penalty = round(rent_base * random.uniform(1, 3), 2)
    note = '面积已调整,不计入同店口径' if i in [5, 12, 30] else '正常'

    # 每家门店用 H3 子标题 + 字段段落
    doc.add_heading(f'{sid} {store_names[i]}', level=2)

    fields = [
        ('门店ID', sid),
        ('门店名称', store_names[i]),
        ('区域', store_region[i]),
        ('租约ID', lease_id),
        ('出租方', '某某物业有限公司'),
        ('租期开始', start.strftime('%Y-%m-%d')),
        ('租期结束', end.strftime('%Y-%m-%d')),
        ('免租期', f'{free_rent}天'),
        ('月租金', f'{rent_base:.2f}元'),
        ('物业费', f'{rent_base * 0.1:.2f}元/月'),
        ('递增条款', escalation),
        ('闭店条件', f'提前90天书面通知,支付{penalty:.2f}元违约金'),
        ('续租优先权', '有'),
        ('备注', note),
    ]
    for label, value in fields:
        p = doc.add_paragraph()
        run_label = p.add_run(f'{label}:')
        run_label.bold = True
        p.add_run(f' {value}')

doc.save(f'{OUTPUT_DIR}/华东区门店租约汇总.docx')
print("  -> 完成")

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
doc2.add_paragraph('2.1 A 级门店:到店坪效≥5000 元/㎡/月,租售比≤15%,同店增长≥5%。')
doc2.add_paragraph('2.2 B 级门店:到店坪效 2000-5000 元/㎡/月,租售比 15%-25%。')
doc2.add_paragraph('2.3 C 级门店:到店坪效 1000-2000 元/㎡/月,租售比 25%-35%。')
doc2.add_paragraph('2.4 D 级门店:到店坪效<1000 元/㎡/月或租售比>35%,列为整改或闭店候选。')

doc2.add_heading('3. O2O 履约规则', level=1)
doc2.add_paragraph('3.1 门店发货、到店自提订单计入线上渠道,同时按发货/自提归属计入门店履约贡献。')
doc2.add_paragraph('3.2 即时配送订单归属即时零售渠道,履约成本由渠道承担。')
doc2.add_paragraph('3.3 门店履约贡献(发货+自提订单的销售收入)计入《门店经营台账》的"门店 O2O 履约贡献"字段。')
doc2.add_paragraph('3.4 门店仅承担履约侧的人工与包装成本(对应《门店费用》中的"履约包装"字段)。')

doc2.add_heading('4. 数据来源优先级', level=1)
doc2.add_paragraph('4.1 财务确认收入优先于业务系统数据。')
doc2.add_paragraph('4.2 冲突时以渠道销售明细和门店台账交叉验证。')
doc2.add_paragraph('4.3 无法确认时标记"待补",不得用假设补齐事实。')

doc2.add_heading('5. 缺失与异常处理', level=1)
doc2.add_paragraph('5.1 允许披露、待补、缩小范围、局部隔离、负向结论。')
doc2.add_paragraph('5.2 不得虚构事实,不得用外部数据替代题内材料。')

doc2.add_heading('6. 闭店与改造条件', level=1)
doc2.add_paragraph('6.1 租约到期且连续 6 个月 D 级,可启动闭店评估。')
doc2.add_paragraph('6.2 租售比连续 3 个月>35%,可转为前置仓或自提点。')
doc2.add_paragraph('6.3 面积调整门店不计入同店口径。')

doc2.save(f'{OUTPUT_DIR}/渠道与门店管理策略.docx')
print("  -> 完成")

print("  -> 完成")

# ===================== 7. 促销活动记录 =====================
print("生成 2026H1促销活动记录.xlsx ...")
activities = []
# 按权重预生成 50 个活动的渠道
activity_channels = random.choices(channel_ids, weights=channel_order_weights, k=50)
for i in range(1, 51):
    aid = f'A{i:03d}'
    ch = activity_channels[i-1]
    store = random.choice(store_ids) if random.random() > 0.5 else None
    start = random.choice(date_range)
    end = start + timedelta(days=random.randint(1, 15))
    budget = round(random.uniform(1, 30), 2) * 1000  # 1000-30000 元(活动预算量级修正)
    actual = round(budget * random.uniform(0.8, 1.2), 2)
    activities.append({
        '活动ID': aid,
        '活动名称': f'活动{i}',
        '活动类型': random.choice(['满减','折扣','直播','会员日','节日促销']),
        '开始日期': start,
        '结束日期': end,
        '渠道ID': ch,
        '门店ID': store,
        '预算': budget,
        '实际费用': actual
    })
df_activities = pd.DataFrame(activities)

details = []
for aid in df_activities['活动ID']:
    for _ in range(random.randint(5, 20)):
        sku = random.choice(skus)
        cat = sku_category[sku]
        orig = sku_price[sku]
        disc = round(orig * random.uniform(0.5, 0.9), 2)
        qty = random.randint(1, 50)
        sales = round(disc * qty, 2)
        cost = round(sku_cost[sku] * qty, 2)
        gross = sales - cost
        details.append({
            '活动ID': aid,
            'SKU': sku,
            '品类': cat,
            '原价': orig,
            '活动价': disc,
            '折扣率': round(disc/orig, 2),
            '销量': qty,
            '销售额': sales,
            '毛利': gross
        })
df_details = pd.DataFrame(details)

effects = []
for aid in df_activities['活动ID']:
    ch = df_activities[df_activities['活动ID']==aid]['渠道ID'].values[0]
    store = df_activities[df_activities['活动ID']==aid]['门店ID'].values[0]
    # 活动期间销售额:基于真实量级
    before = round(random.uniform(50000, 200000), 2)  # 5-20 万元
    during = round(before * random.uniform(1.1, 3.0), 2)
    inc = during - before
    inc_gross = inc * random.uniform(0.2, 0.5)
    # ROI = 增量毛利 / 实际费用(从活动清单取)
    actual_fee = df_activities.loc[df_activities['活动ID']==aid, '实际费用'].values[0]
    roi = round(inc_gross / actual_fee, 2) if actual_fee > 0 else 0
    effects.append({
        '活动ID': aid,
        '渠道ID': ch,
        '门店ID': store,
        '活动期间销售额': during,
        '活动前销售额': before,
        '增量销售额': round(inc, 2),
        '增量毛利': round(inc_gross, 2),
        'ROI': roi
    })
df_effects = pd.DataFrame(effects)

with pd.ExcelWriter(f'{OUTPUT_DIR}/2026H1促销活动记录.xlsx', engine='openpyxl') as writer:
    df_activities.to_excel(writer, sheet_name='活动清单', index=False)
    df_details.to_excel(writer, sheet_name='活动明细', index=False)
    df_effects.to_excel(writer, sheet_name='活动效果', index=False)
print("  -> 完成")

# ===================== 8. 旧版门店清单_2024.xlsx (噪音) =====================
print("生成 旧版门店清单_2024.xlsx ...")
old_stores = []
for i in range(1, 51):
    sid = f'S{str(i).zfill(3)}'
    status = random.choice(['营业', '关闭', '调整'])
    close_date = datetime(2024, random.randint(1,12), random.randint(1,28)) if status == '关闭' else None
    old_stores.append({
        '门店ID': sid,
        '门店名称': f'门店{str(i).zfill(2)}',
        '区域': random.choice(regions),
        '城市': random.choice(cities),
        '状态': status,
        '关闭日期': close_date
    })
df_old = pd.DataFrame(old_stores)
with pd.ExcelWriter(f'{OUTPUT_DIR}/旧版门店清单_2024.xlsx', engine='openpyxl') as writer:
    df_old.to_excel(writer, sheet_name='门店清单', index=False)
print("  -> 完成")


# ===================== 9. SKU 主数据(新增附件) =====================
# 修复:让 SKU 主数据写入附件,供模型独立验证商品成本与品类分析
print("生成 SKU主数据.xlsx ...")

subcategories = {
    '食品': ['零食', '饮料', '生鲜', '调味'],
    '日用': ['洗护', '清洁', '纸品', '家居'],
    '美妆': ['护肤', '彩妆', '香水', '工具'],
    '家电': ['小家电', '厨房', '数码', '个护'],
    '服饰': ['男装', '女装', '童装', '配件'],
    '母婴': ['奶粉', '辅食', '用品', '玩具'],
}

# 基于已生成的 sku_category / sku_price / sku_cost 派生主数据
sku_subcategory = {sku: random.choice(subcategories[sku_category[sku]]) for sku in skus}
# 建议最低售价 = 单位成本 × 1.15(保 15% 毛利)
sku_min_price = {sku: round(sku_cost[sku] * 1.15, 2) for sku in skus}
# 品类等级:按毛利率 A(≥40%) / B(25-40%) / C(<25%)
sku_grade = {}
for sku in skus:
    margin = (sku_price[sku] - sku_cost[sku]) / sku_price[sku] if sku_price[sku] > 0 else 0
    sku_grade[sku] = 'A' if margin >= 0.40 else ('B' if margin >= 0.25 else 'C')
# 采购周期(天)
sku_lead_time = {sku: random.choice([7, 15, 30, 45, 60]) for sku in skus}
# 上市日期(2023-2025 之间)
sku_launch = {
    sku: (datetime(2023, 1, 1) + timedelta(days=random.randint(0, 1000))).strftime('%Y-%m-%d')
    for sku in skus
}
# 状态(85% 在售,10% 季节性,5% 停售)
sku_status_list = []
for sku in skus:
    sku_status_list.append(random.choices(['在售', '季节性', '停售'], weights=[0.85, 0.10, 0.05])[0])

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
    '状态': sku_status_list,
})
df_sku_master.to_excel(f'{OUTPUT_DIR}/SKU主数据.xlsx', sheet_name='SKU主数据', index=False)
print("  -> 完成")

print("\n" + "="*50)
print("全部 9 个附件生成完成!")
print("="*50)
