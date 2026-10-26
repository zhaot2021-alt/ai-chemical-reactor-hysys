# -*- coding: utf-8 -*-
"""
export_sessions.py — 导出 Claude Code 开发会话为 Markdown（已脱敏）

生成 doc/AI_DEV_RECORD.md：
  - 用户/助手文字消息（超长截断）
  - 工具调用摘要（不嵌入超长工具输出）
  - 完整原始会话 JSONL 与 output 探针报告路径
  - 自动替换 API key / 密码 / 远程主机信息
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

PROJECT = Path(r"E:/desktop/chem_ai_model")
OUT = PROJECT / "doc" / "AI_DEV_RECORD.md"
SESSIONS = [
    ("Session A — 架构/COM/阶段1-2B 探索",
     Path(r"C:/Users/HONOR/.claude/projects/E--desktop-hysys/"
          r"7f3af75e-e758-4b16-a267-327b23c1121f.jsonl")),
    ("Session B — Kinetic/报告收尾",
     Path(r"C:/Users/HONOR/.claude/projects/E--desktop-chem-ai-model/"
          r"fc0355d6-2b85-4ffa-8723-9093b379b77b.jsonl")),
]
MAX_TEXT = 8000
MAX_TOOL = 800


def redact(text: str) -> str:
    text = re.sub(r"sk-[A-Za-z0-9_\-]{10,}", "<REDACTED_API_KEY>", text)
    text = re.sub(
        r"(?i)(密码|password|passwd|pwd)(\s*[:：=]\s*)\S+",
        r"\1\2<REDACTED_PASSWORD>",
        text,
    )
    text = re.sub(
        r"(?i)(Authorization:\s*Bearer\s+)\S+", r"\1<REDACTED>", text
    )
    text = re.sub(
        r"(?i)((?:DEEPSEEK|OPENAI|DASHSCOPE|CHEM_AI_LLM)_API_KEY\s*=\s*)"
        r"['\"]?[^\s'\"]+",
        r"\1<REDACTED>",
        text,
    )
    text = text.replace("52.237.119.161", "<REMOTE_HOST>")
    text = text.replace("Azureuser", "<REMOTE_USER>")
    return text


def clip(text: str, limit: int) -> str:
    text = text.strip()
    if len(text) <= limit:
        return text
    return text[:limit] + f"\n... [截断，原文 {len(text)} 字符]"


def skip_user(text: str, meta: bool) -> bool:
    if meta:
        return True
    s = text.lstrip()
    return s.startswith(
        ("<local-command", "<command-name>", "<command-message>",
         "<command-args>", "<system-reminder>")
    )


def export_session(path: Path, title: str) -> list[str]:
    lines = [f"\n\n# {title}\n", f"> 原始会话：`{path}`\n"]
    if not path.exists():
        return lines + ["\n（原始会话文件不存在）\n"]
    start = end = None
    with path.open(encoding="utf-8", errors="replace") as fh:
        for raw in fh:
            try:
                data = json.loads(raw)
            except Exception:
                continue
            ts = data.get("timestamp")
            if ts:
                start = start or ts
                end = ts
            typ = data.get("type")
            msg = data.get("message")
            if typ == "user" and isinstance(msg, dict):
                content = msg.get("content")
                if isinstance(content, str):
                    if not skip_user(content, bool(data.get("isMeta"))):
                        lines.append(
                            f"\n### 用户 · {ts or ''}\n\n{redact(clip(content, MAX_TEXT))}\n"
                        )
                elif isinstance(content, list):
                    for item in content:
                        if isinstance(item, dict) and item.get("type") == "text":
                            txt = item.get("text", "")
                            if not skip_user(txt, bool(data.get("isMeta"))):
                                lines.append(
                                    f"\n### 用户 · {ts or ''}\n\n{redact(clip(txt, MAX_TEXT))}\n"
                                )
            elif typ == "assistant" and isinstance(msg, dict):
                content = msg.get("content")
                if not isinstance(content, list):
                    continue
                for item in content:
                    if not isinstance(item, dict):
                        continue
                    if item.get("type") == "text":
                        txt = item.get("text", "").strip()
                        if txt:
                            lines.append(
                                f"\n### 助手 · {ts or ''}\n\n{redact(clip(txt, MAX_TEXT))}\n"
                            )
                    elif item.get("type") == "tool_use":
                        name = item.get("name", "tool")
                        inp = json.dumps(item.get("input", {}), ensure_ascii=False)
                        lines.append(
                            f"\n> 工具调用 `{name}`：\n\n```text\n"
                            + redact(clip(inp, MAX_TOOL))
                            + "\n```\n"
                        )
    lines.append(f"\n> 会话时间：{start or '?'} → {end or '?'}\n")
    return lines


def evidence() -> list[str]:
    lines = ["\n\n# 完整会话与验证证据路径\n"]
    lines.append("\n## 原始 Claude 会话（JSONL，完整未截断）\n")
    for title, path in SESSIONS:
        lines.append(f"- **{title}**：`{path}`\n")
    lines.append("\n## output 下的探针/验证记录（txt）\n")
    for path in sorted((PROJECT / "output").glob("*.txt")):
        lines.append(f"- `{path}`（{path.stat().st_size} 字节）\n")
    lines.append(
        "\n> 原始 JSONL 与探针报告可能包含完整命令输出；公开前请再次脱敏，"
        "不要提交 API key、密码或远程主机信息。\n"
    )
    return lines


def main() -> int:
    parts = [
        "# AI 开发记录（Claude Code 会话导出 / 摘要版）\n\n",
        "- 项目：AI 驱动的 HYSYS 反应器建模（Aspen HYSYS V15）\n",
        f"- 导出时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n",
        "- 内容：用户需求、助手决策与工具调用摘要；完整工具日志未嵌入\n",
        "- 安全说明：API key、密码、远程主机/用户等已自动替换\n",
        "- 完整原始会话与 output 探针报告路径见文末\n\n",
        "## 摘要\n\n",
        "- 分层架构：config / nlp_parser / hysys_driver / main / agent，"
        "COM 只出现在驱动层。\n",
        "- HYSYS V15 COM 实测：Case、Peng-Robinson（token `pengrob`）、"
        "组分、FEED、组成、Solver、保存与释放。\n",
        "- 反应器 TypeName：`gibbsreactorop` / `conversionreactorop` / "
        "`equilibriumreactorop`。\n",
        "- 通用反应配置：`conversionrxn`、`equilibriumrxn`，支持化学计量、"
        "转化率、数值 Ka；Gibbs 支持能量流补热。\n",
        "- 两段式 AI：本地 LLM 生成 plan.json，远程只执行 plan。\n",
        "- 两个必测场景与三类核心反应器路径均有真实 HYSYS 结果；"
        "完整边界见 README/REPORT。\n",
    ]
    for title, path in SESSIONS:
        parts.extend(export_session(path, title))
    parts.extend(evidence())
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("".join(parts), encoding="utf-8")
    text = OUT.read_text(encoding="utf-8")
    patterns = [
        r"sk-[A-Za-z0-9_\-]{10,}",
        r"(?i)(密码|password|passwd|pwd)\s*[:：=]\s*(?!<REDACTED)\S+",
        r"52\.237\.119\.161",
        r"Azureuser",
    ]
    hits = sum(len(re.findall(p, text)) for p in patterns)
    print(f"wrote {OUT} ({OUT.stat().st_size} bytes); redaction_hits={hits}")
    return 0 if hits == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
