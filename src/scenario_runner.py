# -*- coding: utf-8 -*-
"""
scenario_runner.py — 执行 LLM plan（远程 HYSYS 侧，不需要 LLM key）
================================================================
支持：
  - gibbs：动态创建 Gibbs 反应器，用等效预热扫描达到 plan 中的出口温度目标
  - conversion：打开手动结构模板 ref_conversion.hsc，覆写进料后求解、读产物

所有参数来自 plan，不在代码里写死工艺数值。
"""

from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Any

import config

logger = logging.getLogger("scenario_runner")

TOLUENE_MW_KG_KMOL = 92.1402
DEFAULT_GIBBS_FLOW_KGMOLEH = 1000.0
FEED_T_BOUNDS_K = (400.0, 1600.0)
OUTLET_TOL_K = 0.5
MAX_ITERATIONS = 20
CONVERSION_TOL = 0.02


def _require_components(plan: dict[str, Any]) -> list[str]:
    components = plan.get("components") or []
    if not components:
        raise ValueError("plan.components 为空")
    unknown = [item for item in components if item not in config.STANDARD_COMPONENT_IDS]
    if unknown:
        raise ValueError(f"plan 含非标准组分：{unknown}")
    return list(components)


def _require_feed_composition(plan: dict[str, Any]) -> dict[str, float]:
    composition = plan.get("feed_composition")
    if not composition:
        raise ValueError("plan.feed_composition 为空，拒绝猜测组成")
    total = sum(float(value) for value in composition.values())
    if total <= 0:
        raise ValueError("plan.feed_composition 总和必须大于 0")
    return {str(key): float(value) / total for key, value in composition.items()}


def run_gibbs(plan: dict[str, Any], driver: Any) -> dict[str, Any]:
    components = _require_components(plan)
    composition = _require_feed_composition(plan)

    feed_t_c = plan.get("feed_temperature_c")
    pressure_bar = plan.get("pressure_bar")
    targets_c = plan.get("outlet_temperatures_c") or []
    if feed_t_c is None or pressure_bar is None:
        raise ValueError("Gibbs plan 缺少 feed_temperature_c 或 pressure_bar")
    if not targets_c:
        raise ValueError("Gibbs plan 缺少 outlet_temperatures_c")

    flow = plan.get("molar_flow_kmolh")
    if flow is None:
        flow = DEFAULT_GIBBS_FLOW_KGMOLEH
        logger.warning(
            "plan 未给出摩尔流量，使用默认 %.1f kgmole/h（仅用于可运行测试）",
            flow,
        )
    flow = float(flow)

    case_name = driver.new_case(
        f"ChemAI_Gibbs_{time.strftime('%Y%m%d_%H%M%S')}"
    )
    driver.create_fluid_package(
        config.FLUID_PACKAGE_NAME,
        property_package=config.PROPERTY_PACKAGE_NAME,
        components=components,
    )
    driver.create_material_stream(
        "FEED",
        config.c_to_k(float(feed_t_c)),
        config.bar_to_kpa(float(pressure_bar)),
        flow,
    )
    driver.set_stream_composition("FEED", composition)
    driver.solve_case()

    driver.create_empty_material_stream("VAPOUR")
    driver.create_empty_material_stream("LIQUID")
    driver.create_reactor("gibbs", "R-GIBBS")
    driver.connect_reactor(
        "R-GIBBS", ["FEED"], "VAPOUR", "LIQUID",
        pressure_drop_kpa=0.0,
    )

    def evaluate(feed_t_k: float) -> dict[str, Any]:
        driver.set_stream_conditions("FEED", temperature_k=feed_t_k)
        driver.solve_case()
        vapour = driver.read_stream_results("VAPOUR")
        if not vapour.get("solved"):
            raise RuntimeError(f"FEED={feed_t_k:.2f} K 时 VAPOUR 未求解")
        return vapour

    def find_feed_temperature(target_k: float) -> tuple[float, dict[str, Any]]:
        low, high = FEED_T_BOUNDS_K
        low_vap = evaluate(low)
        high_vap = evaluate(high)
        low_t = float(low_vap["temperature_k"])
        high_t = float(high_vap["temperature_k"])
        if not (low_t <= target_k <= high_t):
            raise RuntimeError(
                f"出口目标 {target_k:.2f} K 不在可达范围 [{low_t:.2f}, {high_t:.2f}]"
            )
        best_t, best_vap = low, low_vap
        for _ in range(MAX_ITERATIONS):
            mid = (low + high) / 2.0
            vap = evaluate(mid)
            best_t, best_vap = mid, vap
            outlet_t = float(vap["temperature_k"])
            if abs(outlet_t - target_k) <= OUTLET_TOL_K:
                break
            if outlet_t < target_k:
                low = mid
            else:
                high = mid
        return best_t, best_vap

    results: dict[str, Any] = {}
    for target_c in targets_c:
        target_k = config.c_to_k(float(target_c))
        logger.info("Gibbs 目标 %.2f℃ → %.2f K", float(target_c), target_k)
        feed_t_k, _ = find_feed_temperature(target_k)
        driver.set_stream_conditions("FEED", temperature_k=feed_t_k)
        driver.solve_case()
        vapour = driver.read_stream_results("VAPOUR")
        liquid = driver.read_stream_results("LIQUID")
        label = f"{float(target_c):.0f}C"
        save_path = driver.save_case(
            Path(config.OUTPUT_DIR) / f"gibbs_{label}.hsc"
        )
        results[label] = {
            "target_outlet_temperature_c": float(target_c),
            "effective_feed_temperature_c": config.k_to_c(feed_t_k),
            "vapour": vapour,
            "liquid": liquid,
            "hsc_path": save_path,
        }
        outlet_t = vapour.get("temperature_k")
        if outlet_t is None or abs(float(outlet_t) - target_k) > 1.0:
            raise RuntimeError(f"Gibbs {label} 出口温度未达到目标")
        if not vapour.get("mole_fractions"):
            raise RuntimeError(f"Gibbs {label} 无气相组成")

    return {
        "scenario": "gibbs",
        "case_name": case_name,
        "reactor_type": plan.get("reactor_type"),
        "reason": plan.get("reason"),
        "components": components,
        "feed_composition": composition,
        "feed_temperature_c": float(feed_t_c),
        "pressure_bar": float(pressure_bar),
        "results": results,
        "note": "Gibbs 反应器为绝热模式；出口温度通过等效预热达成，未伪造温度属性",
    }


def run_conversion(
    plan: dict[str, Any],
    driver: Any,
    template_path: str | Path | None = None,
) -> dict[str, Any]:
    components = _require_components(plan)
    if "Toluene" not in components:
        raise ValueError("当前 Conversion 模板只支持甲苯歧化，plan 必须含 Toluene")

    template = Path(
        template_path or config.CONVERSION_TEMPLATE_PATH
    )
    if not template.exists():
        raise FileNotFoundError(
            f"Conversion 结构模板不存在：{template}；"
            "请在 HYSYS 中保存 ref_conversion.hsc"
        )

    feed_t_c = plan.get("feed_temperature_c")
    pressure_bar = plan.get("pressure_bar")
    mass_flow = plan.get("mass_flow_kg_h")
    conversion_percent = plan.get("conversion_percent")
    if feed_t_c is None or pressure_bar is None or mass_flow is None:
        raise ValueError("Conversion plan 缺少 feed_temperature_c / pressure_bar / mass_flow_kg_h")
    if conversion_percent is None:
        raise ValueError("Conversion plan 缺少 conversion_percent")
    if abs(float(conversion_percent) - 50.0) > CONVERSION_TOL * 100:
        raise ValueError(
            f"模板固定 50% 转化率，plan 给出 {conversion_percent}%，拒绝错误映射"
        )

    composition = _require_feed_composition(plan)
    non_toluene = {
        key: value for key, value in composition.items()
        if key != "Toluene" and value > 1e-9
    }
    if non_toluene:
        raise ValueError(
            f"当前模板假设纯甲苯进料，plan 组成含其他组分：{non_toluene}"
        )

    driver.open_case(template)
    logger.info("模板物流：%s", driver.list_streams())
    driver.set_stream_conditions(
        "Feed1",
        temperature_k=config.c_to_k(float(feed_t_c)),
        pressure_kpa=config.bar_to_kpa(float(pressure_bar)),
        molar_flow_kgmoleh=float(mass_flow) / TOLUENE_MW_KG_KMOL,
    )
    driver.set_stream_composition("Feed1", {"Toluene": 1.0})
    driver.solve_case()

    feed = driver.read_stream_results("Feed1")
    vapour = driver.read_stream_results("Prod1")
    liquid = driver.read_stream_results("Liq1")
    save_path = driver.save_case(
        Path(config.OUTPUT_DIR) / "conversion_380C_25bar.hsc"
    )

    feed_tol = float(feed.get("molar_flow_kgmoleh") or 0.0) * float(
        (feed.get("mole_fractions") or {}).get("Toluene", 0.0)
    )
    product_tol = 0.0
    for product in (vapour, liquid):
        flow = float(product.get("molar_flow_kgmoleh") or 0.0)
        fractions = product.get("mole_fractions") or {}
        product_tol += flow * float(fractions.get("Toluene", 0.0))
    actual_conversion = (
        (feed_tol - product_tol) / feed_tol * 100.0 if feed_tol > 0 else 0.0
    )
    if abs(actual_conversion - 50.0) > 2.0:
        raise RuntimeError(f"实际甲苯转化率偏离 50%：{actual_conversion:.2f}%")

    return {
        "scenario": "conversion",
        "template": str(template),
        "reactor_type": plan.get("reactor_type"),
        "reason": plan.get("reason"),
        "feed": feed,
        "vapour_product": vapour,
        "liquid_product": liquid,
        "actual_conversion_percent": actual_conversion,
        "hsc_path": save_path,
    }
