# -*- coding: utf-8 -*-
"""
test_stage1.py — 阶段1 联调 / 测试入口示例（不是 pytest 单元测试）
================================================================
演示完整链路：
  固定测试文本 → nlp_parser.parse() 结构化参数 → main.run_stage1()
  → HYSYSDriver 新建 Case / 流体包 / 10 个组分 / FEED 物流 → 保存 .hsc

用法：
  1) 本地（无需 HYSYS / pywin32）只验证解析，并做 dry-run：
        python src/test_stage1.py --dry-run
  2) 远程 HYSYS 工作站真实联调：
        python src/test_stage1.py --no-visible
        python src/test_stage1.py --visible --save output/stage1_test.hsc

【本地可调试】--dry-run 分支不会导入 hysys_driver、不会连接 HYSYS。
【仅远程可运行】不带 --dry-run 的分支需要 HYSYS V15 + pywin32 + 可用许可证。
"""

from __future__ import annotations

import argparse
import json
import logging
import sys

import config
import nlp_parser

# 注意：main 模块在顶层只导入 config / nlp_parser，驱动层是延迟导入的，
# 因此本地导入 main 不会加载 win32com。
import main as app_main


logger = logging.getLogger("test_stage1")

# ==================== 固定测试文本 ====================
# 覆盖：10 个标准组分、温度℃、压力 bar、摩尔流量 kmol/h
TEST_TEXT = (
    "测试用例：新建一个仿真案例。进料包含甲烷、水蒸气、一氧化碳、二氧化碳、"
    "氢气、甲苯、苯、邻二甲苯、间二甲苯、对二甲苯；"
    "进料温度 25℃，压力 1 bar，摩尔流量 100 kmol/h。"
)

EXPECTED_TEMPERATURE_C = 25.0
EXPECTED_PRESSURE_BAR = 1.0
EXPECTED_MOLAR_FLOW_KMOLH = 100.0


def _verify_fixed_parse(parsed: nlp_parser.ParsedRequest) -> None:
    """对固定测试文本的解析结果做严格校验，任何偏差都明确报错。"""
    problems: list[str] = []

    expected_components = set(config.STANDARD_COMPONENT_IDS)
    actual_components = set(parsed["components"])
    if actual_components != expected_components:
        missing = sorted(expected_components - actual_components)
        extra = sorted(actual_components - expected_components)
        problems.append(f"组分集合不匹配：缺少={missing}，多出={extra}")

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

    if parsed["composition"] is not None:
        problems.append(
            "阶段1固定文本不应推断出组成；请检查解析规则"
        )

    unexpected_missing = [
        field for field in parsed["missing_info"] if field != "composition"
    ]
    if unexpected_missing:
        problems.append(
            f"阶段1固定文本存在意外缺失字段：{unexpected_missing}"
        )

    if problems:
        raise app_main.Stage1Error(
            "固定测试文本解析校验失败：\n- " + "\n- ".join(problems)
        )

    logger.info(
        "固定测试文本解析校验通过：10 个标准组分；%.2f℃；%.2f bar；%.2f kmol/h；"
        "组成缺失（阶段1预期，main 会跳过组成/求解）",
        parsed["temperature_c"],
        parsed["pressure_bar"],
        parsed["molar_flow_kmolh"],
    )


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="阶段1 HYSYS 联调测试入口",
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
    print("阶段1 测试文本：")
    print(TEST_TEXT)
    print("=" * 72)

    try:
        # 第 1 步：本地解析（不涉及 COM），并严格校验
        parsed = nlp_parser.parse(TEST_TEXT)
        print("\n[解析结果]")
        print(json.dumps(parsed, ensure_ascii=False, indent=2))
        _verify_fixed_parse(parsed)

        # 第 2 步：调度层 → 驱动层（dry-run 时到计划为止）
        summary = app_main.run_stage1(
            TEST_TEXT,
            save_path=args.save,
            visible=args.visible,
            strict=False,  # 阶段1固定文本不含组成，非 strict 下跳过求解
            dry_run=args.dry_run,
        )
    except app_main.Stage1Error as exc:
        logger.error("阶段1测试失败：%s", exc)
        return 1
    except Exception as exc:  # noqa: BLE001 - 测试入口保留完整 traceback
        logger.exception("阶段1测试出现未预期错误：%s", exc)
        return 1

    print("\n[阶段1摘要]")
    print(json.dumps(summary, ensure_ascii=False, indent=2))

    if args.dry_run:
        print("\n✓ 本地 dry-run 完成：解析与参数校验通过，未连接 HYSYS。")
        print("  下一步：在 HYSYS 工作站执行 `python src/test_stage1.py --no-visible` 做真实联调。")
    else:
        print("\n✓ 阶段1真实联调完成：Case、流体包、组分、FEED 物流已创建并保存。")
        print("  注意：物流组成未设置、未求解，也未创建任何反应器（阶段1范围）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
