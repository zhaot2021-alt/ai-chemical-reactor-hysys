# -*- coding: utf-8 -*-
"""
scenario_runner.py — 执行 LLM plan（远程 HYSYS 侧，不需要 LLM key）
================================================================
支持：
  - gibbs：动态创建 Gibbs 反应器，用等效预热扫描达到 plan 中的出口温度目标
  - conversion：从 plan.reactions 动态创建反应集/反应/反应器（不再依赖模板）

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
COMPONENT_MW_KG_KMOL = {
    "Hydrogen": 2.016, "Nitrogen": 28.014, "Oxygen": 31.999,
    "CO": 28.010, "CO2": 44.009, "Methane": 16.043, "H2O": 18.015,
    "H2S": 34.081, "Ethylene": 28.054, "Ethane": 30.070,
    "Propene": 42.081, "Propane": 44.097,
    "i-Butane": 58.124, "n-Butane": 58.124,
    "i-Butene": 56.108, "1-Butene": 56.108,
    "13-Butadiene": 54.092, "tr2-Butene": 56.108, "cis2-Butene": 56.108,
    "i-Pentane": 72.146, "n-Pentane": 72.146, "1-Pentene": 70.132,
    "Ethanol": 46.069, "MTBE": 88.150, "ETBE": 102.177,
    "Ammonia": 17.031,
    "Toluene": 92.1402, "Benzene": 78.114,
    "o-Xylene": 106.167, "m-Xylene": 106.167, "p-Xylene": 106.167,
}
DEFAULT_GIBBS_FLOW_KGMOLEH = 1000.0
ENERGY_DUTY_BOUNDS_KJ_H = (0.0, 1.0e8)
OUTLET_TOL_K = 0.5
MAX_ITERATIONS = 20
CONVERSION_TOL = 0.02


def _require_components(plan: dict[str, Any]) -> list[str]:
    components = plan.get("components") or []
    if not components:
        raise ValueError("plan.components 为空")
    unknown = [item for item in components if item not in config.KNOWN_COMPONENT_IDS]
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
    driver.create_energy_stream("Q-GIBBS")
    driver.create_reactor("gibbs", "R-GIBBS")
    driver.connect_reactor(
        "R-GIBBS", ["FEED"], "VAPOUR", "LIQUID",
        pressure_drop_kpa=0.0,
        energy_stream_name="Q-GIBBS",
    )

    def evaluate(duty_kj_h: float) -> dict[str, Any]:
        driver.set_energy_stream_heat_flow("Q-GIBBS", duty_kj_h)
        driver.solve_case()
        vapour = driver.read_stream_results("VAPOUR")
        if not vapour.get("solved"):
            raise RuntimeError(f"duty={duty_kj_h:.6g} kJ/h 时 VAPOUR 未求解")
        return vapour

    def find_duty(target_k: float) -> tuple[float, dict[str, Any]]:
        low, high = ENERGY_DUTY_BOUNDS_KJ_H
        low_vap = evaluate(low)
        high_vap = evaluate(high)
        low_t = float(low_vap["temperature_k"])
        high_t = float(high_vap["temperature_k"])
        if not (low_t <= target_k <= high_t):
            raise RuntimeError(
                f"出口目标 {target_k:.2f} K 不在可达范围 "
                f"[{low_t:.2f}, {high_t:.2f}]（duty {low:g}~{high:g} kJ/h）"
            )
        best_duty, best_vap = low, low_vap
        for _ in range(MAX_ITERATIONS):
            mid = (low + high) / 2.0
            vap = evaluate(mid)
            best_duty, best_vap = mid, vap
            outlet_t = float(vap["temperature_k"])
            if abs(outlet_t - target_k) <= OUTLET_TOL_K:
                break
            if outlet_t < target_k:
                low = mid
            else:
                high = mid
        return best_duty, best_vap

    results: dict[str, Any] = {}
    for target_c in targets_c:
        target_k = config.c_to_k(float(target_c))
        logger.info("Gibbs 目标 %.2f℃ → %.2f K", float(target_c), target_k)
        duty_kj_h, _ = find_duty(target_k)
        driver.set_energy_stream_heat_flow("Q-GIBBS", duty_kj_h)
        driver.solve_case()
        vapour = driver.read_stream_results("VAPOUR")
        liquid = driver.read_stream_results("LIQUID")
        label = f"{float(target_c):.0f}C"
        save_path = driver.save_case(
            Path(config.OUTPUT_DIR) / f"gibbs_{label}.hsc"
        )
        results[label] = {
            "target_outlet_temperature_c": float(target_c),
            "energy_stream_duty_kj_h": duty_kj_h,
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
        "note": "FEED 保持 plan 温度；Gibbs 通过能量流提供重整热，duty 由目标出口温度反算",
    }


def run_conversion(
    plan: dict[str, Any],
    driver: Any,
    template_path: str | Path | None = None,
) -> dict[str, Any]:
    """
    通用 Conversion：从 plan.reactions 动态创建反应集/反应/反应器。
    不再依赖 ref_conversion.hsc 模板。
    """
    components = _require_components(plan)
    composition = _require_feed_composition(plan)
    reactions = plan.get("reactions") or []
    if not reactions:
        raise ValueError(
            "Conversion plan 缺少 reactions（化学计量）；"
            "请让 LLM/用户提供已配平反应式，拒绝猜测"
        )

    feed_t_c = plan.get("feed_temperature_c")
    pressure_bar = plan.get("pressure_bar")
    if feed_t_c is None or pressure_bar is None:
        raise ValueError("Conversion plan 缺少 feed_temperature_c / pressure_bar")

    molar_flow = plan.get("molar_flow_kmolh")
    if molar_flow is None:
        mass_flow = plan.get("mass_flow_kg_h")
        if mass_flow is None:
            raise ValueError("Conversion plan 缺少摩尔流量或质量流量")
        mixture_mw = 0.0
        for component_id, fraction in composition.items():
            if fraction <= 1e-12:
                continue
            if component_id not in COMPONENT_MW_KG_KMOL:
                raise ValueError(
                    f"质量流量换摩尔流量缺少 {component_id} 的分子量；"
                    "请提供 molar_flow_kmolh"
                )
            mixture_mw += fraction * COMPONENT_MW_KG_KMOL[component_id]
        molar_flow = float(mass_flow) / mixture_mw
    molar_flow = float(molar_flow)
    if molar_flow <= 0:
        raise ValueError(f"摩尔流量必须为正，收到 {molar_flow!r}")

    case_name = driver.new_case(
        f"ChemAI_Conversion_{time.strftime('%Y%m%d_%H%M%S')}"
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
        molar_flow,
    )
    driver.set_stream_composition("FEED", composition)

    set_name = "Set-1"
    driver.create_reaction_set(set_name)
    driver.associate_reaction_set(set_name)
    for index, reaction in enumerate(reactions, start=1):
        reaction_name = str(reaction.get("name") or f"Rxn-{index}")
        stoich = reaction.get("stoichiometry") or {}
        base_component = str(reaction.get("base_component") or "")
        conversion = reaction.get("conversion_percent")
        if conversion is None:
            conversion = plan.get("conversion_percent")
        if conversion is None:
            raise ValueError(f"反应 {reaction_name!r} 缺少 conversion_percent")
        if not base_component:
            raise ValueError(f"反应 {reaction_name!r} 缺少 base_component")
        driver.add_conversion_reaction(
            set_name,
            reaction_name,
            stoich,
            base_component,
            float(conversion),
        )

    driver.create_empty_material_stream("VAPOUR")
    driver.create_empty_material_stream("LIQUID")
    driver.create_reactor("conversion", "R-CONV")
    driver.connect_reactor(
        "R-CONV", ["FEED"], "VAPOUR", "LIQUID",
        pressure_drop_kpa=0.0,
        reaction_set_name=set_name,
    )
    driver.solve_case()

    feed = driver.read_stream_results("FEED")
    vapour = driver.read_stream_results("VAPOUR")
    liquid = driver.read_stream_results("LIQUID")
    label_t = f"{float(feed_t_c):.0f}C"
    label_p = f"{float(pressure_bar):g}bar"
    save_path = driver.save_case(
        Path(config.OUTPUT_DIR) / f"conversion_{label_t}_{label_p}.hsc"
    )

    base_component = str(reactions[0].get("base_component") or "")
    expected = reactions[0].get("conversion_percent")
    if expected is None:
        expected = plan.get("conversion_percent")
    feed_base = float(feed.get("molar_flow_kgmoleh") or 0.0) * float(
        (feed.get("mole_fractions") or {}).get(base_component, 0.0)
    )
    product_base = 0.0
    for product in (vapour, liquid):
        flow = float(product.get("molar_flow_kgmoleh") or 0.0)
        fractions = product.get("mole_fractions") or {}
        product_base += flow * float(fractions.get(base_component, 0.0))
    actual_conversion = (
        (feed_base - product_base) / feed_base * 100.0 if feed_base > 0 else 0.0
    )
    if len(reactions) == 1 and expected is not None:
        if abs(actual_conversion - float(expected)) > 2.0:
            raise RuntimeError(
                f"实际 {base_component} 转化率 {actual_conversion:.2f}% "
                f"偏离 plan {expected}%"
            )

    return {
        "scenario": "conversion",
        "case_name": case_name,
        "reactor_type": plan.get("reactor_type"),
        "reason": plan.get("reason"),
        "components": components,
        "reactions": reactions,
        "feed": feed,
        "vapour_product": vapour,
        "liquid_product": liquid,
        "actual_conversion_percent": actual_conversion,
        "hsc_path": save_path,
    }


def _resolve_molar_flow(plan: dict[str, Any], composition: dict[str, float]) -> float:
    molar_flow = plan.get("molar_flow_kmolh")
    if molar_flow is not None:
        return float(molar_flow)
    mass_flow = plan.get("mass_flow_kg_h")
    if mass_flow is None:
        raise ValueError("kinetic plan 缺少摩尔流量或质量流量")
    mixture_mw = 0.0
    for component_id, fraction in composition.items():
        if fraction <= 1e-12:
            continue
        if component_id not in COMPONENT_MW_KG_KMOL:
            raise ValueError(
                f"质量流量换摩尔流量缺少 {component_id} 的分子量；"
                "请提供 molar_flow_kmolh"
            )
        mixture_mw += fraction * COMPONENT_MW_KG_KMOL[component_id]
    molar_flow = float(mass_flow) / mixture_mw
    if molar_flow <= 0:
        raise ValueError(f"摩尔流量必须为正，收到 {molar_flow!r}")
    return molar_flow


def run_kinetic(plan: dict[str, Any], driver: Any) -> dict[str, Any]:
    """
    通用 kinetic：按 plan 创建 kineticrxn + CSTR/PFR，求解并读产物。
    plan.reactor_type 为 "pfr" 或 "cstr"（LLM 按相态选择）。
    """
    reactor_kind = str(plan.get("reactor_type", "")).strip().lower()
    if reactor_kind not in ("pfr", "cstr"):
        raise ValueError(f"run_kinetic 只支持 pfr/cstr，收到 {reactor_kind!r}")
    components = _require_components(plan)
    composition = _require_feed_composition(plan)
    reactions = plan.get("reactions") or []
    if not reactions:
        raise ValueError("kinetic plan 缺少 reactions（动力学反应）")
    for reaction in reactions:
        kinetics = reaction.get("kinetics") or {}
        if kinetics.get("activation_energy_kj_kgmole") is None:
            raise ValueError(
                f"反应 {reaction.get('name')!r} 缺少 kinetics.activation_energy_kj_kgmole"
            )
        if kinetics.get("frequency_factor_m3_kgmole_s") is None:
            raise ValueError(
                f"反应 {reaction.get('name')!r} 缺少 kinetics.frequency_factor_m3_kgmole_s"
            )
        if not reaction.get("stoichiometry"):
            raise ValueError(f"反应 {reaction.get('name')!r} 缺少 stoichiometry")

    feed_t_c = plan.get("feed_temperature_c")
    pressure_bar = plan.get("pressure_bar")
    if feed_t_c is None or pressure_bar is None:
        raise ValueError("kinetic plan 缺少 feed_temperature_c / pressure_bar")
    molar_flow = _resolve_molar_flow(plan, composition)

    spec = plan.get("reactor_spec") or {}
    if reactor_kind == "pfr":
        if spec.get("total_volume_m3") is None and (
            spec.get("tube_length_m") is None or spec.get("tube_diameter_m") is None
        ):
            raise ValueError("PFR 需要 total_volume_m3 或 tube_length_m+tube_diameter_m")
    else:
        if spec.get("volume_m3") is None:
            raise ValueError("CSTR 需要 volume_m3")

    case_name = driver.new_case(
        f"ChemAI_Kinetic_{time.strftime('%Y%m%d_%H%M%S')}"
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
        molar_flow,
    )
    driver.set_stream_composition("FEED", composition)
    driver.solve_case()

    set_name = "KinSet"
    driver.create_reaction_set(set_name, set_type="Kinetic")
    driver.associate_reaction_set(set_name)
    for index, reaction in enumerate(reactions, start=1):
        reaction_name = str(reaction.get("name") or f"KinRxn-{index}")
        kinetics = reaction.get("kinetics") or {}
        driver.add_kinetic_reaction(
            set_name,
            reaction_name,
            reaction.get("stoichiometry") or {},
            kinetics.get("forward_orders") or {},
            float(kinetics["activation_energy_kj_kgmole"]),
            float(kinetics["frequency_factor_m3_kgmole_s"]),
            reverse_orders=kinetics.get("reverse_orders") or {},
            reverse_activation_energy=float(
                kinetics.get("reverse_activation_energy_kj_kgmole", 0.0)
            ),
            reverse_frequency_factor=float(
                kinetics.get("reverse_frequency_factor_m3_kgmole_s", 0.0)
            ),
            reaction_phase=int(kinetics.get("reaction_phase", 0)),
        )

    if reactor_kind == "pfr":
        driver.create_empty_material_stream("PRODUCT")
        driver.create_energy_stream("Q-KIN")
        driver.create_reactor("pfr", "R-PFR")
        driver.connect_kinetic_reactor(
            "R-PFR", ["FEED"],
            product_name="PRODUCT",
            reaction_set_name=set_name,
            energy_stream_name="Q-KIN",
            pressure_drop_kpa=float(spec.get("pressure_drop_kpa", 0.0)),
            total_volume_m3=spec.get("total_volume_m3"),
            tube_length_m=spec.get("tube_length_m"),
            tube_diameter_m=spec.get("tube_diameter_m"),
            heat_flow_kj_h=spec.get("heat_flow_kj_h"),
            phase_type=int(spec.get("phase_type", 5)),
            holdup_init_type=int(spec.get("holdup_init_type", 1)),
        )
    else:
        driver.create_empty_material_stream("VAPOUR")
        driver.create_empty_material_stream("LIQUID")
        driver.create_energy_stream("Q-KIN")
        driver.create_reactor("cstr", "R-CSTR")
        driver.connect_kinetic_reactor(
            "R-CSTR", ["FEED"],
            vapour_product_name="VAPOUR",
            liquid_product_name="LIQUID",
            reaction_set_name=set_name,
            energy_stream_name="Q-KIN",
            pressure_drop_kpa=float(spec.get("pressure_drop_kpa", 0.0)),
            volume_m3=spec.get("volume_m3"),
            liquid_level_pct=spec.get("liquid_level_pct", 50.0),
            heat_flow_kj_h=spec.get("heat_flow_kj_h"),
        )

    driver.solve_case()
    feed = driver.read_stream_results("FEED")
    results: dict[str, Any] = {}
    if reactor_kind == "pfr":
        product = driver.read_stream_results("PRODUCT")
        results["product"] = product
        solved = product.get("solved")
    else:
        vapour = driver.read_stream_results("VAPOUR")
        liquid = driver.read_stream_results("LIQUID")
        results["vapour_product"] = vapour
        results["liquid_product"] = liquid
        solved = vapour.get("solved") or liquid.get("solved")
    if not solved:
        raise RuntimeError(f"{reactor_kind.upper()} 产物未求解")
    save_path = driver.save_case(
        Path(config.OUTPUT_DIR) / f"kinetic_{reactor_kind}_case.hsc"
    )
    return {
        "scenario": reactor_kind,
        "case_name": case_name,
        "reactor_type": reactor_kind,
        "reason": plan.get("reason"),
        "components": components,
        "reactions": reactions,
        "reactor_spec": spec,
        "feed": feed,
        "results": results,
        "hsc_path": save_path,
    }


def run_equilibrium(plan: dict[str, Any], driver: Any) -> dict[str, Any]:
    """
    通用 Equilibrium：按 plan.reactions 的化学计量和 EquilibriumConstant(Ka)
    动态创建反应集/平衡反应/平衡反应器，求解并读产物。
    """
    components = _require_components(plan)
    composition = _require_feed_composition(plan)
    reactions = plan.get("reactions") or []
    if not reactions:
        raise ValueError("Equilibrium plan 缺少 reactions（化学计量和 Ka）")

    feed_t_c = plan.get("feed_temperature_c")
    pressure_bar = plan.get("pressure_bar")
    molar_flow = plan.get("molar_flow_kmolh")
    if feed_t_c is None or pressure_bar is None:
        raise ValueError("Equilibrium plan 缺少 feed_temperature_c / pressure_bar")
    if molar_flow is None:
        raise ValueError(
            "Equilibrium plan 需要 molar_flow_kmolh（当前不猜质量流量换算）"
        )
    molar_flow = float(molar_flow)
    if molar_flow <= 0:
        raise ValueError(f"摩尔流量必须为正，收到 {molar_flow!r}")

    case_name = driver.new_case(
        f"ChemAI_Equilibrium_{time.strftime('%Y%m%d_%H%M%S')}"
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
        molar_flow,
    )
    driver.set_stream_composition("FEED", composition)

    set_name = "EqSet"
    driver.create_reaction_set(set_name, set_type="Equilibrium")
    driver.associate_reaction_set(set_name)
    for index, reaction in enumerate(reactions, start=1):
        reaction_name = str(reaction.get("name") or f"EqRxn-{index}")
        stoich = reaction.get("stoichiometry") or {}
        ka = reaction.get("equilibrium_constant")
        if ka is None:
            raise ValueError(
                f"反应 {reaction_name!r} 缺少 equilibrium_constant（Ka）；"
                "Equilibrium 不接受猜测的热力学参数"
            )
        driver.add_equilibrium_reaction(
            set_name, reaction_name, stoich, float(ka)
        )

    driver.create_empty_material_stream("VAPOUR")
    driver.create_empty_material_stream("LIQUID")
    driver.create_reactor("equilibrium", "R-EQ")
    driver.connect_reactor(
        "R-EQ", ["FEED"], "VAPOUR", "LIQUID",
        pressure_drop_kpa=0.0,
        reaction_set_name=set_name,
    )
    driver.solve_case()
    feed = driver.read_stream_results("FEED")
    vapour = driver.read_stream_results("VAPOUR")
    liquid = driver.read_stream_results("LIQUID")
    save_path = driver.save_case(
        Path(config.OUTPUT_DIR) / "equilibrium_case.hsc"
    )
    if not vapour.get("solved") and not liquid.get("solved"):
        raise RuntimeError("Equilibrium 产物未求解")
    return {
        "scenario": "equilibrium",
        "case_name": case_name,
        "reactor_type": plan.get("reactor_type"),
        "reason": plan.get("reason"),
        "components": components,
        "reactions": reactions,
        "feed": feed,
        "vapour_product": vapour,
        "liquid_product": liquid,
        "hsc_path": save_path,
    }
