# 出题工具箱(docs/skill)

> **这是什么**:Magnet 出题过程中**可跨题复用**的文档与脚本,与"某一道题的过程文档"严格分开存放。
> **怎么用**:项目开始前先读本目录的 `出题工作流.md` 与 `rules1.md`,再按其中的命令跑脚本。
> **迁移记录**:2026-09-24 由 `docs/`(培训文档)与 `docs/specs/`(SOP、写作规则、脚本)统一迁入本目录。

---

## 一、目录约定

| 目录 | 放什么 | 是否跨题复用 |
| --- | --- | --- |
| **`docs/skill/`** | 培训文档、工作流 SOP、写作规则、机审与自检脚本 | **可复用**,下道题直接沿用 |
| `docs/specs/` | 本题产物:任务说明、Query、评估表 | 每题重写 |
| `docs/reviews/` | 本题复盘:材料问题清单、M1/M2/M3 结果评估 | 每题重写 |
| `input/`、`output/`、`m1~m3-deliverables/` | 本题材料与各产物 | 每题重写 |
| 项目根目录 | `num.py`(材料生成器)、`build_solution.py`(Solution 生成器) | 骨架复用、业务逻辑重写,见下文第四节 |

---

## 二、文件清单

| 文件 | 类型 | 作用 | 复用性 |
| --- | --- | --- | --- |
| `rules1.md` | 文档 | 「Magnet」出题专家培训文档(v3.0),全部判据的最终依据 | 平台提供,**逐字引用,不修改** |
| `rules2.md` | 文档 | 培训文档要点摘要(六对象、四标准、A/B/C 口径、附件要求、流程与退回) | 同上 |
| `出题工作流.md` | 文档 | 通用 SOP:环节 0 + 8 个环节 × 6 道闸门、通用探针库、跨领域适配、提交前清单 | **原样复用** |
| `step3-writing-guide.md` | 文档 | 评估表模块 A 的写作规则(依据 `rules1.md` 6.2–6.4 节) | **原样复用** |
| `data-validation.py` | 脚本 | 材料自检 L1–L5(含 10 条跨口径反证探针、风险登记与严格模式) | 骨架复用,L1–L5 业务项按题重写 |
| `validate_a_module_levels.py` | 脚本 | 评估表模块 A 的档位机审(字段完整、档位显式对应) | **原样复用** |
| `validate_lineage.py` | 脚本 | 数据血缘说明的可复核性机审 | 骨架复用,字段清单按题改 |
| `add_workbook_checks.py` | 脚本 | **交付物自检**:按 Query/制度口径复算 Excel 交付物的派生值、枚举取值与对象行数(默认只读,结果写 `docs/reviews/交付物自检.md`);`--mode validations` 刷新枚举下拉,`--mode inject` 注入 Check 工作表(仅工作版) | 骨架复用,口径按题改 |
| `validate_all.py` | 脚本 | 统一机审入口:数据自检 + A 模块档位 + 附件/Sheet 名一致性 + 引用完整性 | **原样复用**,只改文件清单常量 |
| `validate_baseline.py` | 脚本 | **材料基线比对**:把当前 `input/` 与基线快照逐 Sheet 比对,列出变化并判断是否超出预期 | **原样复用** |
| `baseline_snapshot.json` | 快照 | 材料基线指纹(各 Sheet 行数/列数/内容哈希/数值列合计),由 `--update` 生成 | 每题一份 |
| `口径登记.md` | 文档 | **口径单点事实源**:登记每个企业口径的唯一定义位置与在 Query / 评估表中的锚点 | **原样复用**,换题时整表重填 |
| `validate_calibers.py` | 脚本 | **口径登记一致性机审**:核对该口径在制度 / Query / 评估表三处是否都有锚点,并打印贯通矩阵 | **原样复用** |

---

## 三、使用顺序

```bash
# 0 先读文档(环节 0:项目开始前通读,每环节开工前复读对应章节)
#   docs/skill/rules1.md  +  docs/skill/出题工作流.md

# 1 生成材料
.venv/bin/python num.py

# 2 材料自检(L1–L5)
.venv/bin/python docs/skill/data-validation.py
MAGNET_STRICT_CHECKS=1 .venv/bin/python docs/skill/data-validation.py   # 严格模式:风险登记转失败

# 3 材料基线比对(改过生成器之后必做)
.venv/bin/python docs/skill/validate_baseline.py                          # 报告变化
.venv/bin/python docs/skill/validate_baseline.py \\
    --expect "2026H1促销活动记录.xlsx,渠道与门店管理策略.docx"             # 只允许这些变化
.venv/bin/python docs/skill/validate_baseline.py --update                 # 确认后更新基线

# 4 生成专家 Solution
.venv/bin/python build_solution.py

# 5 提交前统一机审(五项阻断 + 一项报告)
.venv/bin/python docs/skill/validate_all.py

# 其中可单独运行的两项:
.venv/bin/python docs/skill/validate_calibers.py    # 口径登记一致性(制度/Query/评估表贯通)
.venv/bin/python docs/skill/validate_baseline.py    # 材料基线比对
```

---

## 四、哪些能直接复用、哪些必须重写

| 组件 | 换题后 | 说明 |
| --- | --- | --- |
| 8 个环节流程、6 道闸门、通用探针库(6 类)、提交前清单 | **原样复用** | 与题材无关,见 `出题工作流.md` 第一篇 |
| `validate_all.py`、`validate_a_module_levels.py`、`step3-writing-guide.md` | **原样复用** | 只改附件清单等常量 |
| `validate_baseline.py` | **原样复用** | 首次使用先 `--update` 建立本题基线 |
| `口径登记.md` + `validate_calibers.py` | **原样复用** | 换题时重填口径表(通常 15–25 条),脚本不用改 |
| `data-validation.py` 的 `check()` / `risk()` 骨架、风险登记、严格模式开关 | **原样复用** | 已是通用写法 |
| `data-validation.py` 的 L1–L2(字段与公式)、L3(主键与聚合)、L4(量级与相关性) | **重写** | 换字段、公式、主键、判据 |
| `data-validation.py` 的 L5 探针 | **按 6 类模板重写** | 不变量 / 派生 / 勾稽 / 时点 / 可达 / 歧义 |
| `num.py`(材料生成器) | 骨架复用,业务逻辑重写 | 保留固定种子、单一事实源、结果表由明细派生等写法 |
| `build_solution.py`(Solution 生成器) | 骨架复用,业务逻辑重写 | 注意:**Solution 不能用来证明材料没问题**(同源复算必然通过) |

---

## 五、维护约定

1. **培训文档只读**:`rules1.md`、`rules2.md` 是平台原文,任何改动都会使判据失去依据;发现版本更新时,
   替换整个文件并在 `出题工作流.md` 中更新记录的版本号。
2. **引用检查范围**:统一机审的"引用完整性"只扫描本题产物与工具文档(`docs/specs/*.md`、
   `docs/skill/step3-writing-guide.md`、`docs/skill/出题工作流.md`、`docs/reviews/*.md`、
   `output/数据血缘说明.md`);培训文档原文按"逐字引用"处理,不纳入该检查。
3. **脚本路径变更**:脚本一律通过绝对路径常量定位工作区,移动目录后只需确认
   `validate_all.py` 中的 `SPECS` / `SKILL` 两个常量指向正确。
4. **本目录不存放题目答案**:任何含本题具体数据、结论或交付物的文件都不应放在这里。
