"""
Magnet 专家工作台导出(只读)
============================

用途:把工作台页面上的信息(本题记录 + 评估表填写内容)抓取下来,导出成一份 Markdown 文档,
便于离线核对与复盘。**全程只做 GET 请求,不提交、不修改任何平台数据。**

用法:
    .venv/bin/python docs/skill/export_workbench.py \\
        --url "https://<host>/?talent_id=<tid>&uuid=<uuid>" \\
        --record-id <record_id> \\
        --out docs/reviews/平台工作台导出_YYYY-MM-DD.md [--all-records]

说明:
    · 页面是单页应用,数据经 GET /api/login → /api/records → /api/record → /api/eval 取得;
    · uuid 属认领凭证,文档中默认脱敏成 <uuid 已脱敏>,原始 JSON 不写入仓库;
    · --all-records 会把该专家名下全部记录一并列出(默认只导出 --record-id 指定的记录)。
"""
from __future__ import annotations

import argparse
import json
import re
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/126 Safari/537.36"


def get(url: str, token: str | None = None) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    if token:
        req.add_header("Authorization", f"Bearer {token}")
    with urllib.request.urlopen(req, timeout=60) as r:
        raw = r.read().decode("utf-8", "replace")
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {"_raw": raw}


def redact(s: str) -> str:
    """脱敏:uuid、对象存储预签名 URL 的凭证/签名参数、访问密钥、平台 file_token。"""
    out = str(s)
    out = re.sub(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", "<uuid 已脱敏>", out)
    out = re.sub(r"(X-Tos-Credential=)[^&\s]+", r"\1<已脱敏>", out)
    out = re.sub(r"(X-Tos-Signature=)[^&\s]+", r"\1<已脱敏>", out)
    out = re.sub(r"AK[A-Za-z]{2}[A-Za-z0-9]{8,}", "<访问密钥已脱敏>", out)
    out = re.sub(r'("file_token"\s*:\s*")[^"]+(")', r"\1<已脱敏>\2", out)
    return out


def md_line(label: str, value) -> str:
    return f"| {label} | {redact(value)} |" if value not in (None, "") else f"| {label} | — |"


def render_eval_form(ef: dict) -> str:
    out: list[str] = []
    out.append(f"表单 schema:`{ef.get('schema')}`;导出时间:{ef.get('exported_at')};参与模型:{', '.join(ef.get('models') or [])}\n")

    out.append("### A · 优秀结果标准\n")
    for it in ef.get("module_a") or []:
        out.append(f"**{it.get('id')} · {it.get('type')} · 影响程度 {it.get('importance')}"
                   f"{' · 有中间状态' if it.get('has_mid_state') else ' · 无中间状态'}**\n")
        out.append("```\n" + str(it.get("standard", "")).strip() + "\n```\n")
        for g in it.get("grades") or []:
            out.append(f"- {g}")
        out.append("")

    mb = ef.get("module_b") or {}
    items = mb.get("items") if isinstance(mb, dict) else mb
    out.append("### B · 模型评估项\n")
    for it in items or []:
        out.append(f"**{it.get('id')} · {it.get('type')} · 关联 A 标准 {it.get('related_a_id') or '无'} · 影响程度 {it.get('importance')}**\n")
        out.append(f"- 评估项描述:{str(it.get('defect','')).strip()}")
        out.append(f"- 合格时应满足:{str(it.get('expected','')).strip()}")
        for m, a in (it.get("model_assessments") or {}).items():
            ev = str((a or {}).get("evidence") or "").strip() or "—"
            out.append(f"- **{m}**:{(a or {}).get('status')};原因及证据:{ev}")
        if it.get("grades"):
            out.append("- 中间状态档位:" + ";".join(str(g) for g in it["grades"]))
        if it.get("boundary"):
            out.append(f"- 例外边界:{it['boundary']}")
        out.append("")

    mh = ef.get("module_h") or {}
    out.append("### H · 敏感红线\n")
    out.append(f"- 是否命中:{'是' if mh.get('has_redline') else '否'}")
    out.append(f"- 命中记录:{mh.get('hits') or '无'}")
    out.append(f"- 建议设立的红线标准:{mh.get('suggested_standards') or '无'}\n")
    return "\n".join(out)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", required=True)
    ap.add_argument("--record-id", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--all-records", action="store_true")
    args = ap.parse_args()

    q = urllib.parse.parse_qs(urllib.parse.urlparse(args.url).query)
    base = f"{urllib.parse.urlparse(args.url).scheme}://{urllib.parse.urlparse(args.url).netloc}"
    uuid = (q.get("uuid") or [""])[0]
    talent = (q.get("talent_id") or [""])[0]
    if not uuid:
        raise SystemExit("URL 缺少 uuid 参数")

    login = get(f"{base}/api/login?uuid={urllib.parse.quote(uuid)}&talent_id={urllib.parse.quote(talent)}")
    token = login.get("token", "")
    if not token:
        raise SystemExit(f"登录失败:{login}")
    records = (get(f"{base}/api/records", token).get("records") or [])
    rec = get(f"{base}/api/record?record_id={urllib.parse.quote(args.record_id)}&fresh=1", token).get("record") or {}
    ev = get(f"{base}/api/eval?record_id={urllib.parse.quote(args.record_id)}", token)

    L: list[str] = []
    L.append(f"# Magnet 专家工作台导出({rec.get('domain','')})\n")
    L.append(f"> **导出时间**:{datetime.now():%Y-%m-%d %H:%M}(只读 GET 抓取,未对平台做任何写操作)")
    L.append(f"> **抓取方式**:`GET /api/login` → `/api/records` → `/api/record?record_id=…&fresh=1` → `/api/eval?record_id=…`")
    L.append(f"> **凭证**:uuid 已脱敏;原始 JSON 未入库(仅本地 `/tmp`)\n")
    L.append("---\n")

    L.append("## 一、工作台概览\n")
    L.append(f"- 页面标题:Magnet 专家工作台;专家:{rec.get('annotator_name')};期次:{rec.get('period')}")
    L.append(f"- 名下记录数:{len(records)}\n")
    L.append("| record_id | 题目领域 | 状态 | 最近更新 |")
    L.append("| --- | --- | --- | --- |")
    for r in records:
        same = r.get("record_id") == args.record_id
        updated = rec.get("topic_status_updated_at") if same else ""
        L.append(f"| `{r.get('record_id')}` | {r.get('domain')} | {r.get('task_status')} | {updated or '—'} |")
    L.append("")

    L.append("## 二、本题记录字段\n")
    L.append("| 字段 | 值 |")
    L.append("| --- | --- |")
    for k, v in rec.items():
        if isinstance(v, (dict, list)):
            continue
        if len(str(v)) <= 120:
            L.append(md_line(k, v))
    L.append("")
    L.append("### 2.1 Source 字段(任务说明)\n")
    for k in ["context", "purpose", "requirements", "scope", "deliverable"]:
        L.append(f"**{k}**\n\n```\n{rec.get(k,'')}\n```\n")
    L.append("### 2.2 附件清单(classification)\n")
    L.append("```json\n" + str(rec.get("classification", "")) + "\n```\n")

    L.append("## 三、审核结果与人审意见\n")
    L.append(f"- 附件机审:{rec.get('attach_result')};题包机审:{rec.get('pkg_result')};"
             f"评估表机审:{rec.get('eval_result')}({rec.get('eval_opinion')})")
    L.append(f"- 人审状态:{rec.get('first_status') or '—'};返回状态:{rec.get('second_status') or '—'}")
    L.append(f"- 评估表机审报告(TOS):`{rec.get('eval_overall')}`\n")
    for title, key in [("附件机审报告", "attach_report"), ("题包机审报告", "pkg_report"),
                       ("题包人审结论", "pkg_opinion"), ("人审意见(评估表)", "first_advice")]:
        body = str(rec.get(key) or "").strip()
        L.append(f"### {title}\n")
        L.append(body or "(空)")
        L.append("")

    L.append("## 四、候选 Query 与最终 Query\n")
    for k in ["q1", "q2", "q3"]:
        L.append(f"### {k}(系统生成候选)\n\n```\n{rec.get(k,'')}\n```\n")
    L.append("### 最终 Query\n\n```\n" + str(rec.get("query", "")) + "\n```\n")
    L.append("### Query 引用材料(query_ref)\n\n```json\n" + json.dumps(rec.get("query_ref"), ensure_ascii=False, indent=1) + "\n```\n")

    L.append("## 五、评估表填写内容(平台保存版)\n")
    if ev.get("success"):
        L.append(f"- 记录状态:{ev.get('task_status')};可编辑:{ev.get('editable')};"
                 f"模型:{', '.join(ev.get('available_models') or [])};返修:{ev.get('rework')}\n")
        L.append(render_eval_form(ev.get("eval_form") or {}))
    else:
        L.append(f"(未能读取评估表:{ev})\n")

    L.append("## 六、模型产物与 Solution\n")
    L.append(f"- M1:{rec.get('m1')}")
    L.append(f"- M2:{rec.get('m2')}")
    L.append(f"- M3:{rec.get('m3')}")
    L.append(f"- 已提交 Solution:{json.dumps(rec.get('expert_solution_files'), ensure_ascii=False)}")
    L.append(f"- Solution 人审意见:{str(rec.get('solution_human_advice') or '—')[:200]}")
    L.append("")

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(redact("\n".join(L)), encoding="utf-8")
    print(f"已导出:{out}({out.stat().st_size/1024:.1f} KB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
