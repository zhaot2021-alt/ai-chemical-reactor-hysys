# -*- coding: utf-8 -*-
"""
agent.py — 两段式 AI 反应器建模入口
================================================================
本地（有 LLM key）：
    python src/agent.py plan --text "自然语言需求" --out output/plan.json
    # 调用 LLM，只生成结构化 plan，不连接 HYSYS

远程 HYSYS 工作站（不需要 LLM key）：
    python src/agent.py run --plan output/plan.json --visible
    # 读取 plan，执行 Gibbs / Conversion，保存 .hsc 并输出结果

这样 key 只在本地环境变量中，不写入仓库、不复制到远程环境。
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path

import config


def _setup_logging(level: str) -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format=config.LOG_FORMAT,
        datefmt=config.LOG_DATEFMT,
    )


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="两段式 AI 反应器建模：本地 plan，远程 run",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    plan = sub.add_parser("plan", help="本地调用 LLM 生成 plan.json（不连接 HYSYS）")
    plan.add_argument("--text", default=None, help="自然语言需求")
    plan.add_argument(
        "--out",
        default=str(config.OUTPUT_DIR / "plan.json"),
        help="输出 plan 路径",
    )
    plan.add_argument(
        "--log-level",
        choices=("DEBUG", "INFO", "WARNING", "ERROR"),
        default=config.LOG_LEVEL,
    )

    run = sub.add_parser("run", help="远程执行 plan（HYSYS，无需 LLM key）")
    run.add_argument("--plan", required=True, help="plan.json 路径")
    run.add_argument("--template", default=None, help="Conversion 模板路径")
    run.add_argument(
        "--visible",
        action=argparse.BooleanOptionalAction,
        default=None,
    )
    run.add_argument(
        "--log-level",
        choices=("DEBUG", "INFO", "WARNING", "ERROR"),
        default=config.LOG_LEVEL,
    )
    return parser


def _cmd_plan(args: argparse.Namespace) -> int:
    from llm_planner import LLMPlanError, plan_request

    text = args.text
    if not text:
        print("请输入自然语言需求，回车结束：")
        try:
            text = input("> ").strip()
        except EOFError:
            text = ""
    if not text:
        logging.error("输入为空")
        return 2

    try:
        plan = plan_request(text)
    except LLMPlanError as exc:
        logging.error("LLM 规划失败：%s", exc)
        return 1

    out_path = Path(args.out)
    if not out_path.is_absolute():
        out_path = (Path.cwd() / out_path).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"source_text": text, "plan": plan}
    out_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    print(f"\nplan 已写入：{out_path}")
    print("远程执行：python src\\agent.py run --plan "
          + str(out_path))
    return 0


def _cmd_run(args: argparse.Namespace) -> int:
    plan_path = Path(args.plan)
    if not plan_path.is_absolute():
        plan_path = (Path.cwd() / plan_path).resolve()
    if not plan_path.exists():
        logging.error("plan 文件不存在：%s", plan_path)
        return 1

    payload = json.loads(plan_path.read_text(encoding="utf-8"))
    plan = payload.get("plan") if isinstance(payload, dict) and "plan" in payload else payload
    if not isinstance(plan, dict):
        logging.error("plan 文件结构错误")
        return 1

    print("=" * 72)
    print(f"反应器类型：{plan.get('reactor_type')}")
    print(f"选择理由：{plan.get('reason')}")
    print("=" * 72)

    try:
        from hysys_driver import HYSYSDriver
        from scenario_runner import run_conversion, run_gibbs
    except Exception as exc:  # noqa: BLE001
        logging.error("导入 HYSYS 驱动/执行器失败：%s", exc)
        return 1

    reactor_type = str(plan.get("reactor_type", "")).lower()
    driver = HYSYSDriver(visible=args.visible)
    try:
        driver.connect()
        if reactor_type == "gibbs":
            summary = run_gibbs(plan, driver)
        elif reactor_type == "conversion":
            summary = run_conversion(plan, driver, args.template)
        else:
            logging.error(
                "当前执行器只支持 gibbs / conversion，plan 为 %r", reactor_type
            )
            return 1
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:  # noqa: BLE001
        logging.exception("执行 plan 失败：%s", exc)
        return 1
    finally:
        driver.shutdown()


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    _setup_logging(args.log_level)
    if args.command == "plan":
        return _cmd_plan(args)
    if args.command == "run":
        return _cmd_run(args)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
