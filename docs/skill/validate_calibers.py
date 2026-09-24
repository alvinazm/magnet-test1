"""
口径登记一致性机审
==================

用途:读取 `docs/skill/口径登记.md` 的口径表,逐行核对每个口径是否在
  · 附件《渠道与门店管理策略.docx》(条款号 + 关键词)
  · `docs/specs/step2-query.md`(关键词,可填"—"表示由制度唯一承载)
  · `docs/specs/step3-evaluation-rubric.md`(关键词,可填"—"表示尚无对应标准)
三处留下锚点,并打印"口径贯通矩阵"。

背景:本项目曾出现同一口径在制度与 Query 各写一半且互相矛盾(第 6.1 条"租约到期"时点),
     导致同一判断出现 0/2/3 家三种答案。本脚本把这种"半截口径"变成可机审的缺失项。

使用:
    .venv/bin/python docs/skill/validate_calibers.py

退出码:
    0 = 所有已登记的锚点均存在
    1 = 存在缺失锚点(或口径登记表格式异常)
"""
import re
import sys
from pathlib import Path

import docx

WORKSPACE = Path(__file__).resolve().parent.parent.parent
REGISTRY = Path(__file__).resolve().parent / "口径登记.md"
POLICY = WORKSPACE / "input/渠道与门店管理策略.docx"
QUERY = WORKSPACE / "docs/specs/step2-query.md"
RUBRIC = WORKSPACE / "docs/specs/step3-evaluation-rubric.md"


def read_docx(path: Path) -> str:
    doc = docx.Document(str(path))
    parts = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for r in t.rows:
            parts.append(" | ".join(c.text for c in r.cells))
    return "\n".join(parts)


def parse_registry() -> list:
    rows = []
    for line in REGISTRY.read_text(encoding="utf-8").splitlines():
        if not line.strip().startswith("| K-"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) != 7:
            raise ValueError(f"口径登记表列数异常:{line[:60]}")
        rows.append(dict(zip(
            ["编号", "口径", "制度条款", "制度关键词", "Query 关键词", "评估表关键词", "备注"], cells)))
    return rows


def main() -> int:
    if not REGISTRY.exists():
        print(f"✗ 口径登记表不存在:{REGISTRY}")
        return 1
    policy_text = read_docx(POLICY)
    query_text = QUERY.read_text(encoding="utf-8")
    rubric_text = RUBRIC.read_text(encoding="utf-8")

    try:
        rows = parse_registry()
    except ValueError as exc:
        print(f"✗ {exc}")
        return 1

    print("=" * 78)
    print("【口径登记一致性检查】")
    print("=" * 78)
    print(f"  登记口径 {len(rows)} 条;核对对象:制度条款 / Query / 评估表\n")
    print(f"  {'编号':6s}{'口径':34s}{'制度条款':10s}{'制度':6s}{'Query':7s}{'评估表':7s}")

    issues = []
    n_query, n_rubric, n_na_q, n_na_r = 0, 0, 0, 0
    for r in rows:
        clause_ok = r["制度条款"] in policy_text
        policy_ok = r["制度关键词"] in policy_text
        q_kw, rb_kw = r["Query 关键词"], r["评估表关键词"]
        query_ok = (q_kw == "—") or (q_kw in query_text)
        rubric_ok = (rb_kw == "—") or (rb_kw in rubric_text)
        n_query += 1 if query_ok else 0
        n_rubric += 1 if rubric_ok else 0
        n_na_q += 1 if q_kw == "—" else 0
        n_na_r += 1 if rb_kw == "—" else 0
        print(f"  {r['编号']:6s}{r['口径']:34s}{r['制度条款']:10s}"
              f"{'✓' if clause_ok and policy_ok else '✗':6s}"
              f"{('—' if q_kw == '—' else ('✓' if query_ok else '✗')):7s}"
              f"{('—' if rb_kw == '—' else ('✓' if rubric_ok else '✗')):7s}")
        if not clause_ok:
            issues.append(f"{r['编号']} 制度条款 {r['制度条款']} 不存在")
        if not policy_ok:
            issues.append(f"{r['编号']} 制度缺少关键词:{r['制度关键词']}")
        if not query_ok:
            issues.append(f"{r['编号']} Query 缺少关键词:{q_kw}")
        if not rubric_ok:
            issues.append(f"{r['编号']} 评估表缺少关键词:{rb_kw}")

    print(f"\n  锚点覆盖:制度 {len(rows)}/{len(rows)};"
          f"Query {n_query}/{len(rows)}(其中 {n_na_q} 条由制度唯一承载);"
          f"评估表 {n_rubric}/{len(rows)}(其中 {n_na_r} 条标注无对应标准)")

    if issues:
        print(f"\n  ✗ 发现 {len(issues)} 处缺失锚点:")
        for i, msg in enumerate(issues, 1):
            print(f"    {i}. {msg}")
        print("\n  处理方式:先改《渠道与门店管理策略.docx》,再按本表同步 Query 与评估表;")
        print("            若该口径确实不需在 Query/评估表出现,把对应单元格改为「—」。")
        return 1

    print("\n  ✅ 所有已登记口径的锚点均一致。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
