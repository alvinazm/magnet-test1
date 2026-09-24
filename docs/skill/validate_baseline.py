"""
材料基线比对(回归检查)
========================

用途:把当前 `input/` 与一份「基线快照」逐 Sheet 比对,回答两个问题:
  1. 这次改动**实际**动了哪些附件、哪些 Sheet、哪些列;
  2. 实际变化是否**超出预期**(防止"改一处、动全库"的事故)。

背景:本项目曾出现一次事故——修改促销生成逻辑时无意改变了随机数调用顺序,导致 9 个附件
全量漂移,当时只能靠人工逐格比对才发现。本脚本把那次手工动作固化为可重复的检查。

使用:
    .venv/bin/python docs/skill/validate_baseline.py --update
        # 用当前 input/ 生成/覆盖基线快照(docs/skill/baseline_snapshot.json)
        # 仅在"这次变化是预期的"时才执行,并在文档中记录变更内容
    .venv/bin/python docs/skill/validate_baseline.py
        # 与基线比对,打印变化清单(不判定失败)
    .venv/bin/python docs/skill/validate_baseline.py --expect "2026H1促销活动记录.xlsx,渠道与门店管理策略.docx"
        # 只允许列出的文件(或 文件:Sheet)发生变化;出现其他变化即退出码 1

退出码:
    0 = 无变化,或变化均在 --expect 允许范围内,或未指定 --expect
    1 = 存在超出 --expect 允许范围的变化,或基线快照不存在
"""
import argparse
import hashlib
import json
import sys
from pathlib import Path

import docx
import pandas as pd

WORKSPACE = Path(__file__).resolve().parent.parent.parent
INPUT_DIR = WORKSPACE / "input"
SNAPSHOT_PATH = Path(__file__).resolve().parent / "baseline_snapshot.json"


def _hash_rows(rows) -> str:
    h = hashlib.md5()
    for row in rows:
        h.update(("\t".join("" if v is None else str(v) for v in row) + "\n").encode("utf-8"))
    return h.hexdigest()[:12]


def scan_file(path: Path) -> dict:
    """扫描一个附件,返回可比较的指纹。"""
    info = {"size_kb": round(path.stat().st_size / 1024, 1), "sheets": {}}
    if path.suffix == ".xlsx":
        sheets = pd.read_excel(path, sheet_name=None)
        for name, df in sheets.items():
            numeric = df.select_dtypes(include="number")
            info["sheets"][name] = {
                "rows": int(len(df)),
                "cols": int(len(df.columns)),
                "hash": _hash_rows(df.itertuples(index=False, name=None)),
                "sums": {c: round(float(numeric[c].sum()), 2) for c in numeric.columns},
            }
    else:
        doc = docx.Document(str(path))
        rows = [(p.text,) for p in doc.paragraphs]
        for t in doc.tables:
            for r in t.rows:
                rows.append(tuple(c.text for c in r.cells))
        info["sheets"]["docx"] = {
            "rows": len(rows), "cols": 1, "hash": _hash_rows(rows), "sums": {},
        }
    return info


def scan_all(input_dir: Path = None) -> dict:
    root = input_dir or INPUT_DIR
    return {p.name: scan_file(p) for p in sorted(root.iterdir())
            if p.suffix in (".xlsx", ".docx")}


def diff_files(base: dict, cur: dict) -> list:
    """返回变化清单:[(文件, Sheet, 变化描述)]"""
    changes = []
    for fn in sorted(set(base) | set(cur)):
        if fn not in cur:
            changes.append((fn, "-", "文件缺失"))
            continue
        if fn not in base:
            changes.append((fn, "-", "新增文件"))
            continue
        b, c = base[fn], cur[fn]
        for s in sorted(set(b["sheets"]) | set(c["sheets"])):
            if s not in c["sheets"]:
                changes.append((fn, s, "Sheet 缺失"))
            elif s not in b["sheets"]:
                changes.append((fn, s, "新增 Sheet"))
            else:
                bs, cs = b["sheets"][s], c["sheets"][s]
                if bs["hash"] == cs["hash"]:
                    continue
                bits = []
                if bs["rows"] != cs["rows"]:
                    bits.append(f"行数 {bs['rows']}→{cs['rows']}")
                if bs["cols"] != cs["cols"]:
                    bits.append(f"列数 {bs['cols']}→{cs['cols']}")
                for col in sorted(set(bs["sums"]) | set(cs["sums"])):
                    ov, nv = bs["sums"].get(col), cs["sums"].get(col)
                    if ov is None or nv is None or abs(ov - nv) > 0.01:
                        bits.append(f"{col} 合计 {ov}→{nv}")
                changes.append((fn, s, "; ".join(bits) if bits else "内容变化(合计未变)"))
    return changes


def main() -> int:
    ap = argparse.ArgumentParser(description="材料基线比对")
    ap.add_argument("dir", nargs="?", default=None, help="要比对的目录(默认 input/,试跑用)")
    ap.add_argument("--update", action="store_true", help="用当前 input/ 覆盖基线快照")
    ap.add_argument("--expect", default="", help="允许变化的文件或 文件:Sheet,逗号分隔")
    args = ap.parse_args()
    input_dir = Path(args.dir) if args.dir else INPUT_DIR

    if args.update:
        SNAPSHOT_PATH.write_text(
            json.dumps(scan_all(input_dir), ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"✓ 已更新基线快照:{SNAPSHOT_PATH.relative_to(WORKSPACE)}")
        print("  提示:基线只应在「这次变化是预期的」时更新,并在文档中记录变更内容。")
        return 0

    if not SNAPSHOT_PATH.exists():
        print(f"✗ 基线快照不存在:{SNAPSHOT_PATH.relative_to(WORKSPACE)}")
        print("  请先运行:--update")
        return 1

    base = json.loads(SNAPSHOT_PATH.read_text(encoding="utf-8"))
    changes = diff_files(base, scan_all(input_dir))

    print("=" * 78)
    print("【材料基线比对】")
    print("=" * 78)
    if not changes:
        print(f"  ✓ 与基线完全一致({len(base)} 个附件逐 Sheet 指纹相同)")
        return 0

    print(f"  与基线相比共有 {len(changes)} 处变化:")
    for fn, sheet, desc in changes:
        print(f"    · {fn} / {sheet}:{desc}")

    if not args.expect:
        print("\n  · 未指定 --expect,仅作报告(退出码 0)。")
        print("    确认这些变化都是预期的之后,运行 --update 更新基线。")
        return 0

    allowed = {x.strip() for x in args.expect.split(",") if x.strip()}
    unexpected = [(f, s, d) for f, s, d in changes
                  if f not in allowed and f"{f}:{s}" not in allowed]
    if unexpected:
        print(f"\n  ✗ 存在 {len(unexpected)} 处**超出预期**的变化:")
        for fn, sheet, desc in unexpected:
            print(f"    · {fn} / {sheet}:{desc}")
        print(f"\n  预期范围:{sorted(allowed)}")
        return 1
    print(f"\n  ✓ 全部变化均在预期范围内({sorted(allowed)})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
