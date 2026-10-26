# -*- coding: utf-8 -*-
"""
test_stage2b_gibbs_targets.py — Gibbs 反应器两个出口温度目标
================================================================
当前 V15 Gibbs 反应器为绝热模式（无直接温度 setpoint 成员）。
本测试通过**等效预热**扫描 FEED 温度，使 VAPOUR 出口达到：
  目标1：710℃ / 13.5 bar
  目标2：600℃ / 13.5 bar
然后读回产物组成并分别保存 .hsc。

说明：原始 FEED 520℃ 的显式预热设备尚未建模；这里输出的是等效预热
温度与在该出口温度下的 Gibbs 平衡组成，不伪造反应器温度属性。

运行：
    python src\test_stage2b_gibbs_targets.py --visible --log-level DEBUG
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
import time
from pathlib import Path

import config

logger = logging.getLogger("test_stage2b_targets")

COMPONENTS = ("Methane", "H2O", "CO", "CO2", "Hydrogen")
FEED_T_K = 793.15          # 520℃，仅初始值
FEED_P_KPA = 1350.0
FEED_FLOW_KGMOLEH = 1000.0
FEED_COMPOSITION = {
    "Methane": 1.0 / 3.7,
    "H2O": 2.7 / 3.7,
    "CO": 0.0,
    "CO2": 0.0,
    "Hydrogen": 0.0,
}

TARGETS = (
    ("710C", 983.15),
    ("600C", 873.15),
)
FEED_T_BOUNDS = (400.0, 1600.0)
T_OUTLET_TOL_K = 0.5
MAX_ITERATIONS = 20


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    parser = argparse.ArgumentParser(description="Gibbs 710℃/600℃ 目标扫描")
    parser.add_argument(
        "--visible",
        action=argparse.BooleanOptionalAction,
        default=None,
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
            f"ChemAI_GibbsTargets_{time.strftime('%Y%m%d_%H%M%S')}"
        )
        driver.create_fluid_package(
            config.FLUID_PACKAGE_NAME,
            property_package=config.PROPERTY_PACKAGE_NAME,
            components=COMPONENTS,
        )
        driver.create_material_stream(
            "FEED", FEED_T_K, FEED_P_KPA, FEED_FLOW_KGMOLEH
        )
        driver.set_stream_composition("FEED", FEED_COMPOSITION)
        driver.solve_case()

        driver.create_empty_material_stream("VAPOUR")
        driver.create_empty_material_stream("LIQUID")
        driver.create_reactor("gibbs", "R-GIBBS")
        driver.connect_reactor(
            "R-GIBBS", ["FEED"], "VAPOUR", "LIQUID",
            pressure_drop_kpa=0.0,
        )

        def evaluate(feed_t_k: float) -> dict:
            driver.set_stream_conditions("FEED", temperature_k=feed_t_k)
            driver.solve_case()
            vapour = driver.read_stream_results("VAPOUR")
            if not vapour.get("solved"):
                raise RuntimeError(f"FEED={feed_t_k:.2f}K 时 VAPOUR 未求解")
            return vapour

        def find_feed_temperature(target_t_k: float) -> tuple[float, dict]:
            low, high = FEED_T_BOUNDS
            low_vap = evaluate(low)
            high_vap = evaluate(high)
            low_t = float(low_vap["temperature_k"])
            high_t = float(high_vap["temperature_k"])
            logger.info(
                "出口温度范围：FEED %.1fK -> %.2fK；FEED %.1fK -> %.2fK",
                low, low_t, high, high_t,
            )
            if not (low_t <= target_t_k <= high_t):
                raise RuntimeError(
                    f"目标 {target_t_k:.2f}K 不在可达范围 "
                    f"[{low_t:.2f}, {high_t:.2f}]"
                )

            best_t = low
            best_vap = low_vap
            for _ in range(MAX_ITERATIONS):
                mid = (low + high) / 2.0
                vap = evaluate(mid)
                outlet_t = float(vap["temperature_k"])
                best_t, best_vap = mid, vap
                if abs(outlet_t - target_t_k) <= T_OUTLET_TOL_K:
                    break
                if outlet_t < target_t_k:
                    low = mid
                else:
                    high = mid
            return best_t, best_vap

        results = {}
        for label, target_t in TARGETS:
            logger.info("扫描目标 %s（%.2f K）...", label, target_t)
            feed_t, vapour = find_feed_temperature(target_t)
            driver.set_stream_conditions("FEED", temperature_k=feed_t)
            driver.solve_case()
            vapour = driver.read_stream_results("VAPOUR")
            liquid = driver.read_stream_results("LIQUID")
            save_path = driver.save_case(
                Path(config.OUTPUT_DIR) / f"gibbs_{label}.hsc"
            )
            results[label] = {
                "target_outlet_temperature_k": target_t,
                "effective_feed_temperature_k": feed_t,
                "effective_feed_temperature_c": config.k_to_c(feed_t),
                "vapour": vapour,
                "liquid": liquid,
                "hsc_path": save_path,
            }
            logger.info(
                "%s：等效预热 %.2fK（%.2f℃）→ 出口 %.2fK（%.2f℃）",
                label,
                feed_t,
                config.k_to_c(feed_t),
                vapour.get("temperature_k") or -1,
                config.k_to_c(vapour["temperature_k"])
                if vapour.get("temperature_k") is not None else -1,
            )

        summary = {"case_name": case_name, "results": results}
        print(json.dumps(summary, ensure_ascii=False, indent=2))

        for label, item in results.items():
            vapour = item["vapour"]
            outlet_t = vapour.get("temperature_k")
            target_t = item["target_outlet_temperature_k"]
            if outlet_t is None or abs(float(outlet_t) - target_t) > 1.0:
                logger.error("%s 出口温度未达到目标", label)
                return 1
            if not vapour.get("mole_fractions"):
                logger.error("%s 无气相组成", label)
                return 1
        print("\n✓ Gibbs 710℃ / 600℃ 两个目标均完成。")
        return 0
    except Exception as exc:  # noqa: BLE001
        logger.exception("Gibbs 目标扫描失败：%s", exc)
        return 1
    finally:
        driver.shutdown()


if __name__ == "__main__":
    raise SystemExit(main())
