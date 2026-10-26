# -*- coding: utf-8 -*-
"""
test_stage2b_gibbs.py — 阶段2B：Gibbs 反应器端到端联调
================================================================
链路：
  创建 Case/PR 流体包/5 组分
  → FEED 组成/T/P/F
  → 创建空产物流 VAPOUR/LIQUID
  → 创建 gibbsreactorop
  → Feeds.Add(FEED)、VapourProduct/LiquidProduct、PressureDropValue
  → solve
  → 读回产物组成/温度/压力/汽化率
  → 保存 .hsc

未设置任何未验证的反应器温度/压力 setpoint；工况仅用 FEED 条件。
运行（远程 HYSYS 工作站）：
    python src\test_stage2b_gibbs.py --visible --log-level DEBUG
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
import time
from pathlib import Path

import config

logger = logging.getLogger("test_stage2b")

COMPONENTS = ("Methane", "H2O", "CO", "CO2", "Hydrogen")
FEED_T_K = 793.15          # 520℃
FEED_P_KPA = 1350.0        # 13.5 bar
FEED_FLOW_KGMOLEH = 1000.0
FEED_COMPOSITION = {
    "Methane": 1.0 / 3.7,
    "H2O": 2.7 / 3.7,
    "CO": 0.0,
    "CO2": 0.0,
    "Hydrogen": 0.0,
}

FEED_STREAM = "FEED"
VAPOUR_STREAM = "VAPOUR"
LIQUID_STREAM = "LIQUID"
REACTOR_NAME = "R-GIBBS"


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    parser = argparse.ArgumentParser(description="阶段2B Gibbs 反应器联调")
    parser.add_argument(
        "--visible",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="是否显示 HYSYS 窗口",
    )
    parser.add_argument(
        "--log-level",
        choices=("DEBUG", "INFO", "WARNING", "ERROR"),
        default=config.LOG_LEVEL,
    )
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=getattr(logging, args.log_level, logging.INFO),
        format=config.LOG_FORMAT,
        datefmt=config.LOG_DATEFMT,
    )

    try:
        from hysys_driver import HYSYSDriver
    except Exception as exc:  # noqa: BLE001
        logger.error("无法导入 hysys_driver：%s", exc)
        return 1

    driver = HYSYSDriver(visible=args.visible)
    try:
        driver.connect()
        case_name = driver.new_case(
            f"ChemAI_Gibbs_{time.strftime('%Y%m%d_%H%M%S')}"
        )
        driver.create_fluid_package(
            config.FLUID_PACKAGE_NAME,
            property_package=config.PROPERTY_PACKAGE_NAME,
            components=COMPONENTS,
        )

        driver.create_material_stream(
            FEED_STREAM, FEED_T_K, FEED_P_KPA, FEED_FLOW_KGMOLEH
        )
        driver.set_stream_composition(FEED_STREAM, FEED_COMPOSITION)
        driver.solve_case()

        driver.create_empty_material_stream(VAPOUR_STREAM)
        driver.create_empty_material_stream(LIQUID_STREAM)
        driver.create_reactor("gibbs", REACTOR_NAME)
        connection = driver.connect_reactor(
            REACTOR_NAME,
            [FEED_STREAM],
            VAPOUR_STREAM,
            LIQUID_STREAM,
            pressure_drop_kpa=0.0,
        )
        driver.solve_case()

        vapour = driver.read_stream_results(VAPOUR_STREAM)
        liquid = driver.read_stream_results(LIQUID_STREAM)
        saved_path = driver.save_case(
            Path(config.OUTPUT_DIR) / f"{case_name}.hsc"
        )

        summary = {
            "case_name": case_name,
            "reactor": connection,
            "feed_temperature_k": FEED_T_K,
            "feed_pressure_kpa": FEED_P_KPA,
            "vapour_product": vapour,
            "liquid_product": liquid,
            "hsc_path": saved_path,
        }
        print(json.dumps(summary, ensure_ascii=False, indent=2))

        if not vapour.get("solved") or not vapour.get("mole_fractions"):
            logger.error("气相产物未求解或无组成")
            return 1
        if vapour.get("temperature_k") is None or vapour.get("pressure_kpa") is None:
            logger.error("气相产物温度/压力读回为空")
            return 1
        print("\n✓ 阶段2B Gibbs 反应器联调完成。")
        return 0
    except Exception as exc:  # noqa: BLE001
        logger.exception("阶段2B Gibbs 联调失败：%s", exc)
        return 1
    finally:
        driver.shutdown()


if __name__ == "__main__":
    raise SystemExit(main())
