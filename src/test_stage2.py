# -*- coding: utf-8 -*-
"""
test_stage2.py — 阶段2A 联调入口：FEED 组成 + 求解 + 读回
================================================================
演示：
  固定测试文本（10 个标准组分 + 等摩尔组成）
  → nlp_parser 输出 composition
  → HYSYSDriver.set_stream_composition()
  → HYSYSDriver.solve_case()
  → HYSYSDriver.read_stream_results()
  → 保存 .hsc

用法：
  本地 dry-run（无需 HYSYS）：
      python src/test_stage2.py --dry-run
  远程 HYSYS 工作站：
      python src/test_stage2.py --no-visible
      python src/test_stage2.py --visible --log-level DEBUG

【本地可调试】--dry-run 不导入驱动层、不连接 HYSYS。
【仅远程可运行】不带 --dry-run 的路径需要 HYSYS V15 + pywin32。
"""

from __future__ import annotations

import argparse
import json
import logging
import math
import sys

import config
import nlp_parser
import main as app_main

logger = logging.getLogger("test_stage2")

# 覆盖：10 个标准组分、显式“等摩尔”组成、温度/压力/摩尔流量
TEST_TEXT = (
    "测试用例：新建一个仿真案例。进料包含甲烷、水蒸气、一氧化碳、二氧化碳、"
    "氢气、甲苯、苯、邻二甲苯、间二甲苯、对二甲苯；各组分等摩尔；"
    "进料温度 25℃，压力 1 bar，摩尔流量 100 kmol/h。"
)

EXPECTED_TEMPERATURE_C = 25.0
EXPECTED_PRESSURE_BAR = 1.0
EXPECTED_MOLAR_FLOW_KMOLH = 100.0
EXPECTED_FRACTION = 0.1


def _verify_fixed_parse(parsed: nlp_parser.ParsedRequest) -> None:
    problems: list[str] = []

    expected_components = set(config.STANDARD_COMPONENT_IDS)
    actual_components = set(parsed["components"])
    if actual_components != expected_components:
        problems.append(
            f"组分集合不匹配：缺少={sorted(expected_components - actual_components)}，"
            f"多出={sorted(actual_components - expected_components)}"
        )

    composition = parsed["composition"]
    if composition is None:
        problems.append("未解析出 composition")
    else:
        if set(composition) != expected_components:
            problems.append(
                f"组成键与组分不一致：{sorted(composition)}"
            )
        for component_id, value in composition.items():
            if abs(value - EXPECTED_FRACTION) > 1e-9:
                problems.append(
                    f"等摩尔组成错误：{component_id}={value}，"
                    f"期望 {EXPECTED_FRACTION}"
                )
        total = sum(composition.values())
        if abs(total - 1.0) > 1e-9:
            problems.append(f"组成总和不是 1.0：{total}")

    if (
        parsed["temperature_c"] is None
        or abs(parsed["temperature_c"] - EXPECTED_TEMPERATURE_C) > 1e-9
    ):
        problems.append(
            f"温度解析错误：期望 {EXPECTED_TEMPERATURE_C}℃，"
            f"实际 {parsed['temperature_c']}"
        )
    if (
        parsed["pressure_bar"] is None
        or abs(parsed["pressure_bar"] - EXPECTED_PRESSURE_BAR) > 1e-9
    ):
        problems.append(
            f"压力解析错误：期望 {EXPECTED_PRESSURE_BAR} bar，"
            f"实际 {parsed['pressure_bar']}"
        )
    if (
        parsed["molar_flow_kmolh"] is None
        or abs(parsed["molar_flow_kmolh"] - EXPECTED_MOLAR_FLOW_KMOLH) > 1e-9
    ):
        problems.append(
            f"摩尔流量解析错误：期望 {EXPECTED_MOLAR_FLOW_KMOLH} kmol/h，"
            f"实际 {parsed['molar_flow_kmolh']}"
        )

    if parsed["missing_info"]:
        problems.append(f"阶段2A固定文本不应有缺失字段：{parsed['missing_info']}")

    if problems:
        raise app_main.Stage1Error(
            "阶段2A固定文本解析校验失败：\n- " + "\n- ".join(problems)
        )

    logger.info(
        "阶段2A解析校验通过：10 个标准组分，等摩尔组成，"
        "%.2f℃ / %.2f bar / %.2f kmol/h",
        parsed["temperature_c"],
        parsed["pressure_bar"],
        parsed["molar_flow_kmolh"],
    )


def _verify_live_summary(summary: dict) -> None:
    problems: list[str] = []
    if summary.get("stage") != "stage2a":
        problems.append(f"stage 期望 stage2a，实际 {summary.get('stage')!r}")
    if not summary.get("composition_set"):
        problems.append("composition_set 不是 True")
    if not summary.get("solved"):
        problems.append("求解标志 solved 不是 True")

    feed = summary.get("feed_stream") or {}
    composition = feed.get("mole_fractions")
    if not composition:
        problems.append("feed_stream.mole_fractions 为空")
    else:
        total = sum(float(value) for value in composition.values())
        if abs(total - 1.0) > 1e-6:
            problems.append(f"读回组成总和不是 1.0：{total}")

    vapour_fraction = feed.get("vapour_fraction")
    if vapour_fraction is None:
        problems.append("feed_stream.vapour_fraction 为空")
    elif not math.isfinite(float(vapour_fraction)):
        problems.append(f"汽化率不是有限值：{vapour_fraction!r}")

    for key, label in (
        ("temperature_k", "温度"),
        ("pressure_kpa", "压力"),
        ("molar_flow_kgmoleh", "摩尔流量"),
    ):
        if feed.get(key) is None:
            problems.append(f"feed_stream.{key}（{label}）为空")

    if problems:
        raise app_main.Stage1Error(
            "阶段2A联调结果校验失败：\n- " + "\n- ".join(problems)
        )


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="阶段2A HYSYS 联调：FEED 组成 + 求解 + 读回",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只解析 + 打印驱动调用计划，不连接 HYSYS（本地可运行）",
    )
    parser.add_argument(
        "--visible",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="是否显示 HYSYS 窗口（--visible / --no-visible）",
    )
    parser.add_argument(
        "--save",
        type=str,
        default=None,
        help=".hsc 保存路径（默认 output/<CaseName>.hsc）",
    )
    parser.add_argument(
        "--log-level",
        choices=("DEBUG", "INFO", "WARNING", "ERROR"),
        default=config.LOG_LEVEL,
        help="日志级别",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    args = _build_arg_parser().parse_args(argv)
    app_main.setup_logging(args.log_level)

    print("=" * 72)
    print("阶段2A 测试文本：")
    print(TEST_TEXT)
    print("=" * 72)

    try:
        parsed = nlp_parser.parse(TEST_TEXT)
        print("\n[解析结果]")
        print(json.dumps(parsed, ensure_ascii=False, indent=2))
        _verify_fixed_parse(parsed)

        summary = app_main.run_stage1(
            TEST_TEXT,
            save_path=args.save,
            visible=args.visible,
            strict=True,
            dry_run=args.dry_run,
        )
    except app_main.Stage1Error as exc:
        logger.error("阶段2A测试失败：%s", exc)
        return 1
    except Exception as exc:  # noqa: BLE001
        logger.exception("阶段2A测试出现未预期错误：%s", exc)
        return 1

    print("\n[阶段2A摘要]")
    print(json.dumps(summary, ensure_ascii=False, indent=2))

    if args.dry_run:
        plan = summary.get("driver_plan") or {}
        if not plan.get("solve"):
            logger.error("dry-run 计划中没有 solve=True")
            return 1
        print("\n✓ 本地 dry-run 完成：组成已解析，未连接 HYSYS。")
        print("  下一步：在 HYSYS 工作站执行 `python src/test_stage2.py --no-visible`。")
    else:
        try:
            _verify_live_summary(summary)
        except app_main.Stage1Error as exc:
            logger.error("%s", exc)
            return 1
        print("\n✓ 阶段2A真实联调完成：FEED 组成已设置、求解并读回。")
        print("  下一步：阶段2B 反应器创建与产物分布。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
