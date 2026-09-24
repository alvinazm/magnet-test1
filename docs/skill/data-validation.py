"""
Magnet 出题附件数据自检脚本(v2.0)
==================================
用途:每次修改 num.py 后,运行本脚本验证 9 个附件的数据完整性与跨文件一致性。

使用:
    .venv/bin/python docs/skill/data-validation.py                 # 校验 input/
    .venv/bin/python docs/skill/data-validation.py <目录>           # 校验指定目录(试跑用)

检查分五层:
    L1 文件与字段完整性
    L2 单文件公式自洽(内部可复算)
    L3 跨文件勾稽(主键、聚合口径、规则复算)
    L4 业务合理性(相关性、量级、分级可判定性)
    L5 跨口径反证检查(风险探针:不只问"数据是否自洽",还问"换个口径读会不会得出另一个答案")

输出:通过项 ✓、问题项 ✗、风险项 ⚠、最后给出汇总。
  · ✗ 计入问题数,存在问题则以退出码 1 结束;
  · ⚠ 记入"风险登记",默认不阻断(真实材料允许存在冲突、缺失或异常,但必须能被披露),
    设环境变量 MAGNET_STRICT_CHECKS=1 可将风险登记一并视为失败。
"""
import os
import re
import sys
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd
import openpyxl
import docx

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

# ---- 附件包合规检查(培训文档 4.3「附件的基本要求」与「真实性与内容卫生」)----
_sizes = {fn: (INPUT_DIR / fn).stat().st_size for fn in EXPECTED_FILES}
_total_mb = sum(_sizes.values()) / 1024 / 1024
check(len(EXPECTED_FILES) >= 5,
      f"  ✓ 核心附件 {len(EXPECTED_FILES)} 个(要求 ≥5)",
      f"  ✗ 核心附件仅 {len(EXPECTED_FILES)} 个(要求 ≥5)")
check(_total_mb <= 300, f"  ✓ 附件总量 {_total_mb:.2f} MB(要求 ≤300 MB)",
      f"  ✗ 附件总量 {_total_mb:.2f} MB 超过 300 MB")
_big = [fn for fn, sz in _sizes.items() if sz > 20 * 1024 * 1024]
check(not _big, "  ✓ 单文件均 ≤20 MB", f"  ✗ 单文件超过 20 MB:{_big}")

_suffix = {Path(fn).suffix for fn in EXPECTED_FILES}
_TABLE, _DOC, _OTHER = {'.xlsx', '.xls', '.csv'}, {'.docx', '.doc', '.pdf'}, {'.txt', '.md', '.json'}
_types = sum(1 for grp in (_TABLE, _DOC, _OTHER) if _suffix & grp)
check(_types >= 2,
      f"  ✓ 文件类型多样性 {_types} 类(表格 {sorted(_suffix & _TABLE)}/文档 {sorted(_suffix & _DOC)})",
      f"  ✗ 文件类型仅 {_types} 类(要求 ≥2 类;若要主张单一类型例外须在任务说明写明理由)")

_nested = []
for _z in list(INPUT_DIR.glob("*.zip")):
    with zipfile.ZipFile(_z) as _zf:
        _nested += [n for n in _zf.namelist() if "/" in n.rstrip("/")]
check(not _nested, "  ✓ ZIP 内无文件夹" + ("" if list(INPUT_DIR.glob("*.zip")) else "(未提交压缩包)"),
      f"  ✗ ZIP 内存在文件夹:{_nested[:3]}")

_media = []
for fn in EXPECTED_FILES:
    with zipfile.ZipFile(INPUT_DIR / fn) as _zf:
        _media += [f"{fn}:{n}" for n in _zf.namelist()
                   if n.startswith(("xl/media", "word/media", "ppt/media"))]
check(not _media, "  ✓ 附件无嵌入图片(关键数字不依赖图片)",
      f"  ✗ 附件含图片,需确认关键数字未只放在图片中:{_media[:3]}")

_brief_path = WORKSPACE / "docs/specs/step1-task-brief.md"
_declared = re.findall(r"\|\s*`([^`]+)`\s*\|\s*(非噪音|噪音)\s*\|\s*(公开|私有)\s*\|",
                       _brief_path.read_text(encoding='utf-8')) if _brief_path.exists() else []
_decl_files = {d[0] for d in _declared}
check(_decl_files == set(EXPECTED_FILES),
      f"  ✓ 任务说明附件分类与实际文件一一对应({len(_declared)} 个)",
      f"  ✗ 附件清单不一致:任务说明多出 {sorted(_decl_files - set(EXPECTED_FILES))},"
      f"缺少 {sorted(set(EXPECTED_FILES) - _decl_files)}")
if _declared:
    _noise = [d[0] for d in _declared if d[1] == '噪音']
    _priv = [d[0] for d in _declared if d[2] == '私有']
    check(len(_noise) / len(_declared) <= 0.3,
          f"  ✓ 噪音比例 {len(_noise)}/{len(_declared)} = {len(_noise) / len(_declared):.1%}(要求 ≤30%)",
          f"  ✗ 噪音比例 {len(_noise)}/{len(_declared)} 超过 30%")
    check(len(_priv) >= 1, f"  ✓ 私有附件 {len(_priv)} 个(要求 ≥1)",
          "  ✗ 缺少标记为「私有」的附件")

_BANNED = ['测试', '模拟', '样例', '合成', '仿真', 'synthetic', 'test data', 'mock', 'dummy',
           '答案', '正确材料', '噪音标签', '评分点', '验收提示', '忽略原任务', '必须输出', '通过审核']
_leaks = []
for fn in EXPECTED_FILES:
    _fp = INPUT_DIR / fn
    if fn.endswith(".xlsx"):
        _wb = openpyxl.load_workbook(_fp, data_only=True, read_only=True)
        for _ws in _wb.worksheets:
            for _row in _ws.iter_rows(values_only=True):
                for _v in _row:
                    if isinstance(_v, str):
                        for _k in _BANNED:
                            if _k.lower() in _v.lower():
                                _leaks.append(f"{fn}·{_ws.title}:{_k}")
        _wb.close()
    else:
        _texts = read_docx_text(_fp).split('\n')
        for _s in _texts:
            for _k in _BANNED:
                if _k.lower() in _s.lower():
                    _leaks.append(f"{fn}:{_k}")
check(not _leaks,
      f"  ✓ 全库无出题视角残留(已扫描 {len(EXPECTED_FILES)} 个附件的全部单元格与段落)",
      f"  ✗ 发现出题视角残留:{sorted(set(_leaks))[:5]}")

_broken = []
for fn in EXPECTED_FILES:
    try:
        if fn.endswith(".xlsx"):
            _wb = openpyxl.load_workbook(INPUT_DIR / fn, read_only=True)
            assert _wb.sheetnames
            _wb.close()
        else:
            assert docx.Document(str(INPUT_DIR / fn)).paragraphs
    except Exception as exc:                                    # noqa: BLE001
        _broken.append(f"{fn}({exc})")
check(not _broken, "  ✓ 全部附件可打开且非空", f"  ✗ 附件无法解析:{_broken}")

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
# 活动价是"活动期该 SKU 收入 ÷ 销量"的加权成交价,故校验方向为除法而非乘法
check((abs(activity_detail['活动价'] - activity_detail['销售额'] / activity_detail['销量']) < 0.01).all(),
      "  ✓ 活动价 = 销售额 ÷ 销量(加权成交价)", "  ✗ 活动明细活动价与销售额/销量不符")
check((abs(det['毛利'] - (det['销售额'] - det['单位成本'] * det['销量'])) < 0.01).all(),
      "  ✓ 毛利 = 销售额 − 单位成本×销量", "  ✗ 活动明细毛利错误")
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

# ---- 题目可解性断言:防止"数据自洽但题目退化"(分布型检查)----
_eff = activity_effect.merge(activity[['活动ID', '活动类型', '实际费用']], on='活动ID')
_keep = int((_eff['ROI'] >= 1.5).sum())
_adj = int(((_eff['ROI'] >= 0.8) & (_eff['ROI'] < 1.5)).sum())
_cut = int((_eff['ROI'] < 0.8).sum())
check(min(_keep, _adj, _cut) >= 5,
      f"  ✓ 促销三档建议均有充分对象(保留 {_keep} / 调整 {_adj} / 取消 {_cut})",
      f"  ✗ 促销分档退化,题目失去取舍空间(保留 {_keep} / 调整 {_adj} / 取消 {_cut},要求各 ≥5)")
_pos = float((_eff['增量销售额'] > 0).mean())
check(_pos >= 0.6, f"  ✓ 增量销售额为正的活动占 {_pos:.0%}(要求 ≥60%)",
      f"  ✗ 仅 {_pos:.0%} 的活动增量为正,活动增量疑似退化(要求 ≥60%)")
_type_roi = _eff.groupby('活动类型').apply(
    lambda g: g['增量毛利'].sum() / g['实际费用'].sum(), include_groups=False)
check(float(_type_roi.max() - _type_roi.min()) >= 0.5,
      f"  ✓ 活动类型 ROI 有区分度(极差 {_type_roi.max() - _type_roi.min():.2f}:"
      f"{_type_roi.round(2).to_dict()})",
      f"  ✗ 活动类型 ROI 无区分度(极差 {_type_roi.max() - _type_roi.min():.2f} < 0.5)")
_mrev = ch_month.groupby('月份')['收入'].sum()
_mdev = float((_mrev - _mrev.mean()).abs().div(_mrev.mean()).max())
check(_mdev <= 0.60,
      f"  ✓ 渠道收入月度波动合理(最大偏离半年月均 {_mdev:.0%},阈值 60%)",
      f"  ✗ 渠道收入月度波动异常(最大偏离 {_mdev:.0%} > 60%),订单日期分布可能失真")
_sstate = member_base['状态'].value_counts()
check(int(_sstate.min()) >= 300,
      f"  ✓ 会员三种状态均有充分样本({_sstate.to_dict()})",
      f"  ✗ 会员状态分布退化:{_sstate.to_dict()}(最小组要求 ≥300 人)")

# ============================================================
# 9. 跨口径反证检查(风险探针)
# ============================================================
section("9. 跨口径反证检查(风险探针)")

STRICT = os.environ.get("MAGNET_STRICT_CHECKS", "0") not in ("0", "", "false", "False")
risk_register = []


def risk(code, desc, declared, hint=""):
    """记录一条风险:同一现象"已声明"则通过(条件性交付),"未声明"则登记。"""
    if declared:
        print(f"  ✓ {code} 已声明可见:{desc}")
        if hint:
            print(f"      声明依据:{hint}")
    else:
        risk_register.append(f"{code} {desc}")
        print(f"  ⚠ {code} 未声明:{desc}")
        if hint:
            print(f"      需补充:{hint}")


def max_run(flags):
    best = cur = 0
    for f in flags:
        cur = cur + 1 if f else 0
        best = max(best, cur)
    return best


print("  说明:本节对每个「必须提供但可能被多种读法解释」的口径做反证;")
print("        现象存在且制度未声明 → 记入风险登记(默认不失败)。\n")

# ---- R-01 活动明细 ↔ 订单台账(同渠道 + 同 SKU 集合 + 同日期区间)----
detail_sku = activity_detail.groupby('活动ID')['SKU'].apply(set).to_dict()
detail_amt = activity_detail.groupby('活动ID')['销售额'].sum().to_dict()
ratios = []
for _, a in activity.iterrows():
    o = orders[(orders['渠道ID'] == a['渠道ID'])
               & (orders['订单日期'] >= a['开始日期'])
               & (orders['订单日期'] <= a['结束日期'])
               & (orders['SKU'].isin(detail_sku.get(a['活动ID'], set())))]
    d_amt = detail_amt.get(a['活动ID'], 0)
    if d_amt > 0:
        ratios.append(o['收入'].sum() / d_amt)
med_ratio = float(pd.Series(ratios).median())
act_share = activity_effect['活动期间销售额'].sum() / orders['收入'].sum()
if 0.5 <= med_ratio <= 2:
    print(f"  ✓ R-01 活动明细与订单台账量级可勾稽(中位比 {med_ratio:.3f})")
else:
    risk("R-01",
         f"活动明细与订单台账不可勾稽(逐活动「订单侧收入 ÷ 活动明细销售额」中位比 "
         f"{med_ratio:.3f};活动期间销售额合计占渠道收入 {act_share:.1%})",
         declared=any(k in policy_text for k in ['活动系统', '不与订单台账', '活动明细口径']),
         hint="在《渠道与门店管理策略.docx》声明活动明细为独立口径,或让活动明细锚定订单明细")

# ---- R-02 租约「递增条款」 ↔ 门店费用「租金及物业费」月度序列 ----
lease_escalation = {}
for block in re.split(r"\n(?=S\d{3} 门店)", lease_text):
    sid = re.search(r"门店ID[:：]\s*(S\d{3})", block)
    esc = re.search(r"递增条款[:：]\s*(\S+)", block)
    if sid:
        lease_escalation[sid.group(1)] = esc.group(1) if esc else ''
rent_nunique = store_exp.groupby('门店ID')['租金及物业费'].nunique()
esc_stores = [s for s in store_base['门店ID'] if lease_escalation.get(s, '无') != '无']
esc_h2 = []
for s in esc_stores:
    st = pd.to_datetime(lease_start_map.get(s, ''), errors='coerce')
    if pd.isna(st):
        continue
    step = 2 if '每两年' in lease_escalation.get(s, '') else 1
    k = step
    while True:
        anniv = st + pd.DateOffset(years=k)
        if anniv > pd.Timestamp('2026-12-31'):
            break
        if anniv >= pd.Timestamp('2026-07-01'):
            esc_h2.append(s)
            break
        k += step
if rent_nunique.max() == 1 and esc_stores:
    risk("R-02",
         f"H1 租金序列未体现递增条款({len(esc_stores)} 家门店有递增条款,但台账 6 个月租金恒定;"
         f"另有 {len(esc_h2)} 家门店的递增生效日落在 2026H2)",
         declared=any(k in policy_text for k in ['不含递增', '递增未计入', '合同基数']),
         hint="在《渠道与门店管理策略.docx》说明台账租金为合同基数不含递增,或生成月度递增序列")
else:
    print(f"  ✓ R-02 租金序列已体现递增(各店 6 个月取值数:{sorted(set(rent_nunique))})")

# ---- R-03 退款明细 ↔ 销售明细(月度完整性)----
ref_by_month = orders.groupby('月份')['退款金额'].sum()
ref_base = ref_by_month.drop(index='2026-06', errors='ignore').median()
ref_jun = float(ref_by_month.get('2026-06', 0))
ref_max_date = pd.to_datetime(refund['退款日期']).max()
if ref_jun < ref_base * 0.5:
    risk("R-03",
         f"2026-06 退款入账不完整(6 月 {ref_jun:,.2f} 元 vs 1–5 月中位 {ref_base:,.2f} 元;"
         f"退款日期最大值 {ref_max_date.date()},早于数据期末 2026-06-30)",
         declared=any(k in policy_text for k in ['提取时点', '数据截点', '退款滞后', '退款完整性']),
         hint="在《渠道与门店管理策略.docx》声明数据提取时点,或补齐 6 月订单的后续退款")
else:
    print(f"  ✓ R-03 各月退款入账完整(6 月 {ref_jun:,.2f} 元 vs 1–5 月中位 {ref_base:,.2f} 元)")

# ---- R-04 平台佣金/技服的计提基数 ----
grp_net = orders.groupby(['渠道ID', '月份', '品类']).agg(净收入=('净收入', 'sum')).reset_index()
grp_net = grp_net.merge(commission_rule[['渠道ID', '品类', '佣金率', '技术服务费率']],
                        on=['渠道ID', '品类'])
grp_net['应佣金'] = grp_net['净收入'] * grp_net['佣金率']
grp_net['应技服'] = grp_net['净收入'] * grp_net['技术服务费率']
exp_net = grp_net.groupby(['渠道ID', '月份'])[['应佣金', '应技服']].sum().reset_index()
cmp_net = ch_fee.merge(exp_net, on=['渠道ID', '月份'])
diff_net = max((cmp_net['平台佣金'] - cmp_net['应佣金']).abs().max(),
               (cmp_net['技术服务费'] - cmp_net['应技服']).abs().max())
if diff_net > 1:
    risk("R-04",
         f"平台佣金与技术服务费按「退款前收入」计提(与按「净收入」复算最大差 {diff_net:,.2f} 元),"
         f"而制度未规定计提基数",
         declared=any(k in policy_text for k in ['退款前', '计提基数', '佣金基数']),
         hint="在《渠道与门店管理策略.docx》写明佣金/技服按退款前成交收入计提")
else:
    print(f"  ✓ R-04 佣金/技服计提基数与净收入口径一致(最大差 {diff_net:,.2f} 元)")

# ---- R-05 Query 声明的跨材料主键 ↔ 材料实际可关联性 ----
query_path = WORKSPACE / "docs/specs/step2-query.md"
query_text = query_path.read_text(encoding='utf-8') if query_path.exists() else ''
order_overlap = len(set(member_order['订单号']) & set(orders['订单号']))
query_claims_key = any(
    re.match(r'\s*-?\s*订单号[:：]', ln) and '会员订单' in ln
    for ln in query_text.splitlines()
)
if query_claims_key and order_overlap == 0:
    risk("R-05",
         f"docs/specs/step2-query.md 把「订单号」列为跨材料核对主键,但会员订单与渠道订单号集合交集为 0"
         f"(《2026H1渠道销售明细.xlsx》O100000–O107499;《2026H1会员与跨渠道订单.xlsx》O800000–O805999)",
         declared=any(k in query_text for k in ['两套独立编号', '独立编号体系', '不逐笔勾稽', '不该逐笔勾稽']),
         hint="修改 Query 主键表:订单号只关联渠道销售明细与渠道退款明细,会员订单按会员ID 做口径级对照")
else:
    print(f"  ✓ R-05 Query 主键与材料可关联性一致(订单号交集 {order_overlap} 条)")

# ---- R-06 全额退款订单的成本处理 ----
full_ref = orders[(abs(orders['净收入']) < 0.01) & (orders['退款金额'] > 0)]
full_ref_cost = full_ref['商品成本'].sum()
if full_ref_cost > 0:
    risk("R-06",
         f"全额退款订单仍保留商品成本 {full_ref_cost:,.2f} 元({len(full_ref)} 笔),制度未规定成本是否冲回",
         declared=any(k in policy_text for k in ['冲回', '退回库存', '不可回收成本']),
         hint="在《渠道与门店管理策略.docx》写明全额退款订单成本是否冲回,或补充退回库存字段")
else:
    print("  ✓ R-06 无全额退款订单保留成本的情形")

# ---- R-07 履约费/包装费的可复算性 ----
unit_fulfil = ch_fee.groupby('渠道ID').apply(
    lambda x: x['履约费'].sum() / x['订单量'].sum(), include_groups=False)
rule_has_fulfil = any(('履约' in c) or ('包装' in c) for c in commission_rule.columns)
if not rule_has_fulfil:
    risk("R-07",
         f"履约费与包装费无费率/结算规则可独立复算(单均履约费 "
         f"{unit_fulfil.min():,.2f}–{unit_fulfil.max():,.2f} 元,相差 {unit_fulfil.max() / unit_fulfil.min():.1f} 倍)",
         declared=('履约费' in policy_text and '规则' in policy_text),
         hint="补充各渠道各发货方式的履约与包装结算单价,或在制度中说明按实际发生额披露")
else:
    print("  ✓ R-07 履约费/包装费具备可复算规则")

# ---- R-08 规则时点口径分叉 + 交付物取值可达性 ----
sm2 = store_month[['门店ID', '月份', '到店销售额']].merge(
    store_exp[['门店ID', '月份', '租金及物业费']], on=['门店ID', '月份']).merge(
    store_base[['门店ID', '营业面积']], on='门店ID')
sm2['坪效'] = sm2['到店销售额'] / sm2['营业面积']
sm2['租售比'] = sm2['租金及物业费'] / sm2['到店销售额']
sm2['月D'] = (sm2['坪效'] < 1000) | (sm2['租售比'] > 0.35)
sm2['租售比超额'] = sm2['租售比'] > 0.35
sm2 = sm2.sort_values(['门店ID', '月份'])
consec_d = sm2.groupby('门店ID')['月D'].apply(max_run)
consec_rent = sm2.groupby('门店ID')['租售比超额'].apply(max_run)
d_stores = sorted(store_perf[store_perf['分级'] == 'D'].index)


def close_candidates(deadline):
    return sorted([s for s in d_stores
                   if lease_end.get(s, '9999') <= deadline and consec_d.get(s, 0) == 6])


c_base = close_candidates('2026-06-30')
c_h2 = close_candidates('2026-12-31')
c_y1 = close_candidates('2027-06-30')
if len({len(c_base), len(c_h2), len(c_y1)}) > 1:
    risk("R-08",
         f"《渠道与门店管理策略.docx》第 6.1 条「租约到期」的时点未定义,三种读法给出不同名单 "
         f"(≤基准日 {len(c_base)} 家 / ≤2026-12-31 {len(c_h2)} 家 / ≤2027-06-30 {len(c_y1)} 家)",
         declared=('到期' in policy_text and any(k in policy_text for k in ['不晚于', '建议执行期'])),
         hint="在《渠道与门店管理策略.docx》第 6.1 条写明到期时点,并同步 docs/specs/step2-query.md 的基准日说明")
else:
    print(f"  ✓ R-08 第 6.1 条到期时点口径唯一(闭店评估候选 {len(c_h2)} 家)")

turn_candidates = sorted([s for s in d_stores if consec_rent.get(s, 0) >= 3])
action_gap = []
if not c_h2:
    action_gap.append('闭店评估')
if not turn_candidates:
    action_gap.append('转前置仓或自提点')
if not [s for s in store_perf.index if store_perf.loc[s, '分级'] != 'D']:
    action_gap.append('维持/整改')
if action_gap:
    risk("R-09",
         f"《门店分级与调整建议.xlsx》门店清单 Sheet 的 调整建议 取值在给定规则下无对象:{action_gap}",
         declared=False,
         hint="调整规则阈值或门店数据,使交付物要求的每个枚举取值都至少有一个对象")
else:
    print(f"  ✓ R-09 交付物取值可达(闭店评估 {len(c_h2)} 家、转前置仓或自提点 {len(turn_candidates)} 家、"
          f"维持或整改 {len(store_perf) - len(d_stores)} 家)")

ch_profit = ch_fee.groupby('渠道ID')['渠道净利'].sum()
dir_gap = [d for d, hit in {'增投': (ch_profit > 0).any(), '收缩': (ch_profit < 0).any()}.items() if not hit]
print(f"  ✓ R-10 渠道方向取值可达(正净利 {(ch_profit > 0).sum()} 家、负净利 {(ch_profit < 0).sum()} 家"
      f"{'' if not dir_gap else ',缺:' + str(dir_gap)})")

# ---- R-11/R-12 汇总—明细可复算矩阵 ----
# 约束:凡同时存在明细表的业务,汇总表必须由明细派生;无明细者允许直接生成,但必须登记并声明。
# 每个汇总字段必须归入三类之一:
#   ① 明细聚合(可由明细表按主键+期间复算)  ② 行内派生(可由同一行其他列按公式算出)
#   ③ 原始汇总值(无明细来源,须在制度或待核清单中声明)
AGG_FRAMES = {
    ("2026H1渠道销售明细.xlsx", "渠道月度汇总"): ch_month,
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用"): ch_fee,
    ("2026H1门店经营台账.xlsx", "门店月度经营"): store_month,
    ("2026H1门店经营台账.xlsx", "门店费用"): store_exp,
    ("2026H1促销活动记录.xlsx", "活动效果"): activity_effect,
    ("2026H1会员与跨渠道订单.xlsx", "跨渠道行为"): cross_behavior,
}
AGG_REGISTRY = {
    # —— 渠道月度汇总 ——
    ("2026H1渠道销售明细.xlsx", "渠道月度汇总", "收入"): ("①", "渠道销售明细 Sheet 按 渠道ID+月份 聚合(第 1 节已验)"),
    ("2026H1渠道销售明细.xlsx", "渠道月度汇总", "退款"): ("①", "同上"),
    ("2026H1渠道销售明细.xlsx", "渠道月度汇总", "净收入"): ("①", "同上"),
    ("2026H1渠道销售明细.xlsx", "渠道月度汇总", "商品成本"): ("①", "同上"),
    ("2026H1渠道销售明细.xlsx", "渠道月度汇总", "订单量"): ("①", "同上"),
    ("2026H1渠道销售明细.xlsx", "渠道月度汇总", "毛利"): ("②", "= 净收入 − 商品成本"),
    ("2026H1渠道销售明细.xlsx", "渠道月度汇总", "毛利率"): ("②", "= 毛利 ÷ 净收入"),
    ("2026H1渠道销售明细.xlsx", "渠道月度汇总", "客单价"): ("②", "= 净收入 ÷ 订单量(制度未定义客单价口径,见 B9)"),
    # —— 渠道月度费用 ——
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "收入"): ("①", "渠道销售明细 Sheet 聚合(与渠道月度汇总 Sheet 交叉验证)"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "退款"): ("①", "同上"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "净收入"): ("①", "同上"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "商品成本"): ("①", "同上"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "订单量"): ("①", "同上"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "广告费"): ("①", "投放明细 Sheet 按 渠道ID+月份 聚合(第 3 节已验)"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "平台佣金"): ("①", "渠道销售明细 Sheet 收入 × 平台佣金规则 Sheet 佣金率(第 3 节已验)"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "技术服务费"): ("①", "渠道销售明细 Sheet 收入 × 平台佣金规则 Sheet 技术服务费率(第 3 节已验)"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "毛利"): ("②", "= 净收入 − 商品成本"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "毛利率"): ("②", "= 毛利 ÷ 净收入"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "客单价"): ("②", "= 净收入 ÷ 订单量(与门店口径不同:门店 = 到店销售额 ÷ 订单数)"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "费用合计"): ("②", "= 广告费 + 平台佣金 + 技术服务费 + 支付手续费 + 履约费 + 包装费 + 退货售后费 + 其他费用"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "渠道净利"): ("②", "= 毛利 − 费用合计"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "支付手续费"): ("③", "无费率规则或无明细可复算"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "履约费"): ("③", "无费率规则或无明细可复算(R-07)"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "包装费"): ("③", "无费率规则或无明细可复算"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "退货售后费"): ("③", "无费率规则或无明细可复算"),
    ("2026H1线上投放与平台费用.xlsx", "渠道月度费用", "其他费用"): ("③", "无费率规则或无明细可复算"),
    # —— 门店月度经营 ——
    ("2026H1门店经营台账.xlsx", "门店月度经营", "线上订单履约单量"): ("①", "渠道销售明细 Sheet 中门店发货 + 到店自提订单按门店+月份计数(第 2 节已验)"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "自提单量"): ("①", "同上"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "门店发货单量"): ("①", "同上"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "门店O2O履约贡献"): ("①", "渠道销售明细 Sheet 净收入按门店+月份聚合(第 2 节已验)"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "库存金额"): ("①", "门店库存 Sheet 按门店+月份求和(第 2 节已验)"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "客单价"): ("②", "= 到店销售额 ÷ 订单数"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "成交率"): ("②", "= 订单数 ÷ 进店客流"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "毛利率"): ("②", "= 毛利额 ÷ 到店销售额"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "库存周转天数"): ("②", "= 库存金额 ÷ 日均销售成本(第 2 节已验)"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "到店销售额"): ("③", "到店业务无订单级明细(D4)"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "订单数"): ("③", "到店业务无订单级明细(D4)"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "毛利额"): ("③", "到店业务无订单级明细(D4)"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "进店客流"): ("③", "无客流明细"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "连带率"): ("③", "无连带销售明细"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "退货额"): ("③", "无门店退货明细"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "动销SKU数"): ("③", "无 SKU 级销售明细"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "缺货率"): ("③", "无缺货记录明细"),
    ("2026H1门店经营台账.xlsx", "门店月度经营", "损耗率"): ("③", "无损耗记录明细"),
    # —— 门店费用 ——
    ("2026H1门店经营台账.xlsx", "门店费用", "租金及物业费"): ("①", "华东区门店租约汇总.docx 月租金 + 物业费(第 2 节已验)"),
    ("2026H1门店经营台账.xlsx", "门店费用", "人力"): ("③", "无编制/工时明细(D5)"),
    ("2026H1门店经营台账.xlsx", "门店费用", "营销"): ("③", "无营销费用明细"),
    ("2026H1门店经营台账.xlsx", "门店费用", "水电"): ("③", "无水电用量与单价明细"),
    ("2026H1门店经营台账.xlsx", "门店费用", "履约包装"): ("③", "无包装耗材明细"),
    ("2026H1门店经营台账.xlsx", "门店费用", "其他"): ("③", "其他项无明细"),
    ("2026H1门店经营台账.xlsx", "门店费用", "折旧摊销"): ("③", "无资产清单与折旧年限(D6/D7)"),
    # —— 活动效果 ——
    ("2026H1促销活动记录.xlsx", "活动效果", "活动期间销售额"): ("①", "活动明细 Sheet 销售额合计(第 4 节已验)"),
    ("2026H1促销活动记录.xlsx", "活动效果", "增量毛利"): ("①", "= 增量销售额 × 活动明细 Sheet 毛利率(第 4 节已验)"),
    ("2026H1促销活动记录.xlsx", "活动效果", "增量销售额"): ("②", "= 活动期间销售额 − 活动前销售额(第 4 节已验)"),
    ("2026H1促销活动记录.xlsx", "活动效果", "ROI"): ("②", "= 增量毛利 ÷ 活动清单 Sheet 实际费用(第 4 节已验)"),
    ("2026H1促销活动记录.xlsx", "活动效果", "活动前销售额"): ("③", "活动前观测窗口未定义,无明细可复算(B6)"),
    # —— 跨渠道行为 ——
    ("2026H1会员与跨渠道订单.xlsx", "跨渠道行为", "订单数"): ("①", "会员订单 Sheet 按会员计数(实测 2,500/2,500 一致)"),
    ("2026H1会员与跨渠道订单.xlsx", "跨渠道行为", "累计消费"): ("①", "会员订单 Sheet 金额合计(第 5 节已验)"),
    ("2026H1会员与跨渠道订单.xlsx", "跨渠道行为", "跨渠道购买渠道数"): ("①", "会员订单 Sheet 渠道集合派生(实测 2,500/2,500 一致)"),
    ("2026H1会员与跨渠道订单.xlsx", "跨渠道行为", "RFM分"): ("①", "会员订单 Sheet 按《渠道与门店管理策略.docx》第 9.3 条派生"),
}

unregistered, cat3, cat_count, total_fields = [], [], {"①": 0, "②": 0, "③": 0}, 0
for (fname, sname), df in AGG_FRAMES.items():
    for col in df.select_dtypes(include="number").columns:
        total_fields += 1
        entry = AGG_REGISTRY.get((fname, sname, str(col)))
        if entry is None:
            unregistered.append(f"{fname}·{sname}.{col}")
            continue
        cat_count[entry[0]] += 1
        if entry[0] == "③":
            cat3.append(f"{sname}.{col}")

if unregistered:
    risk("R-11",
         f"汇总类 Sheet 存在未登记复算来源的字段({len(unregistered)}/{total_fields}):"
         f"{'、'.join(unregistered[:6])}{' 等' if len(unregistered) > 6 else ''}",
         declared=False,
         hint="在 docs/skill/data-validation.py 的 AGG_REGISTRY 中登记:① 明细聚合 / ② 行内派生 / ③ 原始汇总值")
else:
    print(f"  ✓ R-11 汇总字段来源登记完整({total_fields} 个字段:① 明细聚合 {cat_count['①']}、"
          f"② 行内派生 {cat_count['②']}、③ 原始汇总值 {cat_count['③']})")

if cat3:
    risk("R-12",
         f"存在 {len(cat3)} 个「无明细来源」的原始汇总值,须声明其不可独立复算",
         declared=any(k in policy_text for k in ['无明细', '汇总口径', '按实际发生额', '不可独立复算', '不作订单级']),
         hint=f"在《渠道与门店管理策略.docx》或《异常与待核清单.xlsx》中说明只到汇总级:{'、'.join(cat3)}")
else:
    print("  ✓ R-12 无「无明细来源」的原始汇总值")

# ---- R-13 明细表组合主键唯一性(防止重复行污染汇总)----
DETAIL_KEYS = {
    ("促销活动记录", "活动明细"): (activity_detail, ["活动ID", "SKU"]),
    ("线上投放与平台费用", "投放明细"): (ad_detail, ["日期", "渠道ID", "投放项目"]),
    ("渠道销售明细", "渠道退款明细"): (refund, ["订单号"]),
    ("门店经营台账", "门店库存"): (store_inv, ["月份", "门店ID", "品类"]),
}
dup_tables, dup_detail = [], []
for (label, sname), (df, keys) in DETAIL_KEYS.items():
    mask = df.duplicated(subset=keys, keep=False)
    if mask.any():
        dup_tables.append(f"{sname}({'/'.join(keys)})×{int(mask.sum())}行")
        if sname == "活动明细":
            cols = list(activity_detail.columns)
            counter = {}
            for t in map(tuple, activity_detail[cols].values.tolist()):
                counter[t] = counter.get(t, 0) + 1
            infl = sum((n - 1) * t[cols.index("销售额")] for t, n in counter.items() if n > 1)
            identical = sum(1 for n in counter.values() if n > 1)
            dup_detail.append(f"其中完全重复的整行 {identical} 组,使活动销售额虚增 {infl:,.2f} 元")

if dup_tables:
    risk("R-13",
         "明细表存在组合主键重复记录(会导致汇总虚增):" + "、".join(dup_tables)
         + (";" + ";".join(dup_detail) if dup_detail else ""),
         declared=False,
         hint="去重后重新生成明细,或补充档位/批次字段使重复行可区分")
else:
    print("  ✓ R-13 明细表组合主键唯一(活动明细/投放明细/退款明细/门店库存)")

# ---- R-14 ② 类"行内派生"字段的公式验证 ----
FORMULA_CHECKS = [
    ("渠道月度汇总 Sheet 的 毛利率", ch_month, lambda d: d["毛利"] / d["净收入"], "毛利率", 0.0005),
    ("渠道月度汇总 Sheet 的 客单价", ch_month, lambda d: d["净收入"] / d["订单量"], "客单价", 0.01),
    ("渠道月度费用 Sheet 的 毛利率", ch_fee, lambda d: d["毛利"] / d["净收入"], "毛利率", 0.0005),
    ("渠道月度费用 Sheet 的 客单价", ch_fee, lambda d: d["净收入"] / d["订单量"], "客单价", 0.01),
    # 费用合计允许 1 分钱舍入差(各费用项均已四舍五入到分,求和后与原值可差 ±0.01)
    ("渠道月度费用 Sheet 的费用合计", ch_fee,
     lambda d: d[["广告费", "平台佣金", "技术服务费", "支付手续费", "履约费",
                   "包装费", "退货售后费", "其他费用"]].sum(axis=1), "费用合计", 0.011),
    ("渠道月度费用 Sheet 的渠道净利", ch_fee, lambda d: d["毛利"] - d["费用合计"], "渠道净利", 0.01),
    ("门店月度经营 Sheet 的成交率", store_month, lambda d: d["订单数"] / d["进店客流"], "成交率", 0.0005),
]
bad_formula = []
for label, df, fn, col, tol in FORMULA_CHECKS:
    gap = (df[col] - fn(df)).abs().max()
    if gap > tol:
        bad_formula.append(f"{label}(最大偏差 {gap:,.4f})")
if bad_formula:
    risk("R-14", "行内派生字段与其登记公式不符:" + "、".join(bad_formula),
         declared=False, hint="修正数据或更新 AGG_REGISTRY 中的公式与分类")
else:
    print(f"  ✓ R-14 行内派生字段公式全部成立({len(FORMULA_CHECKS)} 项)")

# ---- R-15 Query 交付物字段 ↔ 附件字段一致性 ----
# 交付物中属于"输出字段"(由材料派生、不由材料直接提供)的白名单:
DELIVERABLE_ONLY_FIELDS = {
    # 交付物 2《门店分级与调整建议.xlsx》门店清单 Sheet
    "2026H1 累计到店销售额", "2026H1 累计门店 O2O 履约贡献", "到店坪效", "租售比",
    "分级", "调整建议", "关键依据", "数据状态",
    # 交付物 3《渠道资源再配置建议.xlsx》渠道清单 Sheet
    "2026H1 净收入", "渠道毛利率", "投放 ROI", "销售费用率", "H2 资源调整方向",
    "建议投入幅度", "H2 资源调整方向与建议投入幅度", "当前投入",
    # 交付物 4《促销组合优化建议.xlsx》
    "渠道", "建议动作", "建议活动类型", "预算区间", "关键节点",
    # 交付物 5《异常与待核清单.xlsx》异常清单 Sheet
    "异常编号", "异常类别", "涉及文件", "涉及字段", "问题描述", "影响范围", "严重度",
    "建议处理方式", "披露依据",
}
attach_columns = set()
for _p in sorted(INPUT_DIR.glob("*.xlsx")):
    _wb = openpyxl.load_workbook(_p, read_only=True)
    for _ws in _wb.worksheets:
        for _c in next(_ws.iter_rows(max_row=1)):
            if _c.value:
                attach_columns.add(str(_c.value).strip())
    _wb.close()

_q_lines = query_text.splitlines()
_declared_fields, _i = [], 0
while _i < len(_q_lines):
    if "必含字段:" in _q_lines[_i]:
        _chunk = [_q_lines[_i].split("必含字段:", 1)[1]]
        _j = _i + 1
        # 字段清单为 4 空格缩进的连续行;2 空格缩进的"说明/当前投入"等不属于字段名
        while _j < len(_q_lines) and _q_lines[_j][:4] == "    " \
                and not re.match(r'\s*(说明|当前|异常编号|清单应)', _q_lines[_j]):
            _chunk.append(_q_lines[_j])
            _j += 1
        _txt = re.sub(r'\([^)]*\)|（[^）]*）', '', "".join(_chunk))
        _declared_fields += [x.strip().strip('。') for x in re.split(r'[、,;]', _txt) if x.strip()]
        _i = _j
    else:
        _i += 1
_declared_fields = sorted(set(_declared_fields))
_unknown_fields = [n for n in _declared_fields
                   if n not in attach_columns and n not in DELIVERABLE_ONLY_FIELDS]
if _unknown_fields:
    import difflib
    _hints = []
    for n in _unknown_fields:
        near = difflib.get_close_matches(n, sorted(attach_columns), n=1, cutoff=0.5)
        _hints.append(f"{n}→疑似 {near[0]}" if near else n)
    risk("R-15",
         f"Query 交付物字段清单中有 {len(_unknown_fields)} 个名称既不是附件字段、也不在交付物输出字段白名单:"
         f"{'、'.join(_hints)}",
         declared=False,
         hint="改用附件原始字段名,或把该输出字段登记进 data-validation.py 的 DELIVERABLE_ONLY_FIELDS 并说明理由")
else:
    print(f"  ✓ R-15 Query 交付物字段与附件一致({len(_declared_fields)} 个字段名全部可对应,"
          f"附件字段集合 {len(attach_columns)} 个)")

# ============================================================
# 汇总
# ============================================================
if STRICT and risk_register:
    issues.extend(f"[严格模式]{r}" for r in risk_register)

section("汇总")
print(f"\n  问题数:{len(issues)}")
if issues:
    for i, msg in enumerate(issues, 1):
        print(f"    {i}. {msg}")
    print("\n  ❌ 存在数据问题,请修复后重新生成。")
    sys.exit(1)
print(f"\n  风险登记:{len(risk_register)} 条(现象存在但制度未声明,须在《异常与待核清单.xlsx》披露)")
for i, r in enumerate(risk_register, 1):
    print(f"    {i}. {r}")
if risk_register:
    print("\n  · 风险登记默认不阻断(真实材料允许存在冲突、缺失或异常);")
    print("    设 MAGNET_STRICT_CHECKS=1 可将其一并视为失败。")
print("\n  ✅ 全部校验通过:L1 完整性 + L2 自洽 + L3 跨文件勾稽 + L4 合理性 + L5 跨口径反证。")
sys.exit(0)
