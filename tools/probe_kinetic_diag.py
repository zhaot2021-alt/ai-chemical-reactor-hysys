# -*- coding: utf-8 -*-
r"""
probe_kinetic_diag.py — kinetic 参数有效性 + PFR/CSTR 求解诊断（一次跑完）
=========================================================================
在 HYSYS 工作站运行：

    python tools\probe_kinetic_diag.py

按参考 Case 的顺序配置（先关联 FP，再激活反应，后写参数），
每一步都读回参数确认写入是否生效，并做两个 CSTR 对照：
  CSTR-A: FF=1000（参考值）
  CSTR-B: FF=1e6（放大 1000 倍，确认反应引擎是否生效）
输出 output\probe_kinetic_diag_report.txt。
"""

from __future__ import annotations

import gc
import sys
import time
import traceback
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tools"))
import config  # noqa: E402
import probe_hysys_kinetic_api as base  # noqa: E402

PROGID = "HYSYS.Application"
REPORT_PATH = ROOT / "output" / "probe_kinetic_diag_report.txt"
REPORT: list[str] = []


def log(message: str = "") -> None:
    print(message, flush=True)
    REPORT.append(str(message))


def err(exc: BaseException) -> str:
    if isinstance(exc, pywintypes.com_error):
        hr = getattr(exc, "hresult", None)
        return f"HRESULT={hr and f'0x{hr & 0xFFFFFFFF:08X}'} {getattr(exc, 'excepinfo', None)!r}"
    return f"{type(exc).__name__}: {exc}"


def try_set(obj: Any, attr: str, value: Any, units: tuple[str, ...] = ()) -> None:
    try:
        current = getattr(obj, attr)
    except Exception:
        current = None
    if current is not None and hasattr(current, "SetValue") and units:
        for unit in units:
            try:
                current.SetValue(value, unit)
                log(f"    SET {attr} = {value!r} [{unit}] OK (SetValue)")
                return
            except Exception as exc:  # noqa: BLE001
                log(f"    SET {attr} [{unit}] failed: {err(exc)}")
    try:
        setattr(obj, attr, value)
        log(f"    SET {attr} = {value!r} OK (direct)")
    except Exception as exc:  # noqa: BLE001
        log(f"    SET {attr} = {value!r} failed: {err(exc)}")


def try_get(obj: Any, attr: str, units: tuple[str, ...] = ()) -> None:
    try:
        value = getattr(obj, attr)
    except Exception as exc:  # noqa: BLE001
        log(f"    GET {attr} failed: {err(exc)}")
        return
    if hasattr(value, "GetValue"):
        for unit in units:
            try:
                log(f"    GET {attr} = {value.GetValue(unit)} [{unit}]")
                return
            except Exception:
                continue
        try:
            log(f"    GET {attr} = {float(value)} [default]")
            return
        except Exception:
            pass
    log(f"    GET {attr} = {value!r}")


def read_stream(stream: Any, components: list[str], label: str) -> None:
    log(f"    --- {label} ---")
    try:
        t = getattr(stream, "Temperature").GetValue("K")
        f = getattr(stream, "MolarFlow").GetValue("kgmole/h")
        log(f"      T={float(t):.4g} K, F={float(f):.6g} kgmole/h")
    except Exception as exc:  # noqa: BLE001
        log(f"      T/F read failed: {err(exc)}")
    try:
        raw = [float(v) for v in stream.ComponentMolarFractionValue]
        log("      mole frac: " +
            str({components[i]: round(v, 6) for i, v in enumerate(raw[: len(components)])}))
    except Exception as exc:  # noqa: BLE001
        log(f"      mole frac read failed: {err(exc)}")


def make_feed(case: Any, name: str, components: list[str]) -> Any:
    stream = case.Flowsheet.MaterialStreams.Add(name)
    stream.Temperature.SetValue(config.c_to_k(300.0), "K")
    stream.Pressure.SetValue(config.bar_to_kpa(10.0), "kPa")
    stream.MolarFlow.SetValue(100.0, "kgmole/h")
    fractions = {"CO2": 0.5, "Hydrogen": 0.5, "Methane": 0.0, "CO": 0.0, "H2O": 0.0}
    stream.ComponentMolarFractionValue = tuple(
        float(fractions.get(c, 0.0)) for c in components
    )
    base._wait_solver(case)
    return stream


def make_reaction(
    manager: Any,
    rset: Any,
    fp: Any,
    name: str,
    ff: float,
    components: list[str],
) -> Any:
    rxn = manager.Reactions.Add(name, "kineticrxn")
    reactant_names = ["CO2", "Hydrogen", "CO", "H2O"]
    stoich = {"CO2": -1.0, "Hydrogen": -1.0, "CO": 1.0, "H2O": 1.0}
    orders = {"CO2": 1.0, "Hydrogen": 1.0}
    for c in reactant_names:
        rxn.Reactants.Add(c)
    rxn.ReactantStoichCoefValue = tuple(stoich[c] for c in reactant_names)
    rxn.ComponentForwardOrderValue = tuple(orders.get(c, 0.0) for c in reactant_names)
    rxn.ComponentReverseOrderValue = tuple(0.0 for _ in reactant_names)
    rset.ActiveReactions.Add(rxn)
    log(f"  [{name}] 激活后写参数：")
    try_set(rxn, "ReactionPhase", 1)
    try_set(rxn, "Basis", 3)
    try_set(rxn, "BasisUnits2", "kgmole/m3")
    try_set(rxn, "RateUnits2", "kgmole/m3-s")
    try_set(rxn, "ForwardActivationEnergy", 50000.0, ("kJ/kgmole", "J/mol", "kJ/kmol"))
    try_set(rxn, "ForwardFrequencyFactor", ff, ("m3/kgmole-s", "m3/kmol-s", "1/s"))
    try_set(rxn, "ReverseActivationEnergy", 0.0, ("kJ/kgmole", "J/mol", "kJ/kmol"))
    try_set(rxn, "ReverseFrequencyFactor", 0.0, ("m3/kgmole-s", "m3/kmol-s", "1/s"))
    log(f"  [{name}] 参数读回（写入后）：")
    for attr in ("Basis", "BasisUnits2", "RateUnits2", "ReactionPhase",
                 "ForwardActivationEnergy", "ForwardFrequencyFactor",
                 "ReactantStoichCoefValue", "ComponentForwardOrderValue"):
        try_get(rxn, attr)
    log(f"  [{name}] 参数读回（关联 FP 后）：")
    rset.AssociateFluidPackage(fp)
    fp.ReactionPackage.ReactionSets.Add(rset)
    for attr in ("ForwardActivationEnergy", "ForwardFrequencyFactor", "Basis"):
        try_get(rxn, attr)
    return rxn


def build_pfr(case: Any, feed: Any, rset: Any, components: list[str]) -> None:
    log("\n=== PFR (V=1, L=5.09, D=0.5, PhaseType=5, Holdup=1) ===")
    pfr = case.Flowsheet.Operations.Add("PFR-DIAG", "PFReactorop")
    product = case.Flowsheet.MaterialStreams.Add("PFR-PROD")
    energy = case.Flowsheet.EnergyStreams.Add("Q-PFR")
    pfr.Feeds.Add(feed)
    pfr.Product = product
    pfr.EnergyStream = energy
    pfr.ReactionSet = rset
    pfr.PressureDropValue = 0.0
    # 参考 Case：V=1、L=5.09，D 由 HYSYS 计算（直接 PUT D 会被拒 E_INVALIDARG）
    try_set(pfr, "TotalVolumeValue", 1.0, ("m3",))
    try_set(pfr, "TubeLengthValue", 5.0929, ("m",))
    try_set(pfr, "HeatFlowValue", 0.0, ("kJ/h",))
    try_set(pfr, "PhaseType", 5)
    try_set(pfr, "HoldupInitType", 1)
    try:
        energy.HeatFlow.SetValue(0.0, "kJ/h")
    except Exception as exc:  # noqa: BLE001
        log(f"    Q-PFR set failed: {err(exc)}")
    base._wait_solver(case)
    read_stream(product, components, "PFR-PROD")


def build_cstr(case: Any, feed: Any, rset: Any, components: list[str],
               tag: str, liquid_level: float | None, with_energy: bool = True) -> None:
    log(f"\n=== CSTR-{tag} (V=1, LL={liquid_level}) ===")
    cstr = case.Flowsheet.Operations.Add(f"CSTR-{tag}", "KineticReactorop")
    vapour = case.Flowsheet.MaterialStreams.Add(f"CSTR-{tag}-VAP")
    liquid = case.Flowsheet.MaterialStreams.Add(f"CSTR-{tag}-LIQ")
    # CSTR 不能复用已连到其他反应器的进料流，必须自建进料
    own_feed = make_feed(case, f"FEED-{tag}", components)
    cstr.Feeds.Add(own_feed)
    cstr.VapourProduct = vapour
    cstr.LiquidProduct = liquid
    cstr.ReactionSet = rset
    cstr.PressureDropValue = 0.0
    try_set(cstr, "VolumeValue", 1.0, ("m3",))
    if liquid_level is not None:
        try_set(cstr, "LiquidLevelValue", liquid_level, ("%",))
    if with_energy:
        energy = case.Flowsheet.EnergyStreams.Add(f"Q-{tag}")
        cstr.EnergyStream = energy
        try:
            energy.HeatFlow.SetValue(0.0, "kJ/h")
        except Exception as exc:  # noqa: BLE001
            log(f"    Q-{tag} set failed: {err(exc)}")
    try_set(cstr, "HeatFlowValue", 0.0, ("kJ/h",))
    base._wait_solver(case)
    read_stream(vapour, components, f"CSTR-{tag}-VAP")
    read_stream(liquid, components, f"CSTR-{tag}-LIQ")
    for attr in ("RxnPercentConversionValue", "RxnExtentValue"):
        try_get(cstr, attr)


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    app = None
    case = None
    owns = False
    try:
        pythoncom.CoInitialize()
        try:
            app = win32com.client.GetActiveObject(PROGID)
            log("attached existing HYSYS")
        except Exception:
            app = win32com.client.Dispatch(PROGID)
            owns = True
            log("started new HYSYS")
        try:
            app.Visible = True
        except Exception:
            pass

        case, fp, feed, components = base._setup_case(app)
        read_stream(feed, components, "FEED")
        manager = case.BasisManager.ReactionPackageManager

        log("\n=== kinetic set A (FF=1000) ===")
        rset_a = manager.ReactionSets.Add("KinSetA", "Kinetic")
        make_reaction(manager, rset_a, fp, "Rxn-A", 1000.0, components)

        log("\n=== kinetic set B (FF=1e6, 灵敏度对照) ===")
        rset_b = manager.ReactionSets.Add("KinSetB", "Kinetic")
        make_reaction(manager, rset_b, fp, "Rxn-B", 1.0e6, components)

        build_pfr(case, feed, rset_a, components)
        build_cstr(case, feed, rset_a, components, "A", 50.0)
        build_cstr(case, feed, rset_b, components, "B", 50.0)
        build_cstr(case, feed, rset_b, components, "C", None)

        log("\n=== summary ===")
        log("  看 CSTR-B/C 是否有 CO/H2O 生成：有 -> 反应引擎生效，量级问题；")
        log("  无 -> 反应未被应用，看参数读回定位。")
    except Exception as exc:  # noqa: BLE001
        log(f"EXCEPTION: {err(exc)}")
        log(traceback.format_exc())
    finally:
        gc.collect()
        if case is not None:
            try:
                case.Close(False)
            except Exception:
                pass
            case = None
        if app is not None and owns:
            try:
                app.Quit()
                time.sleep(2)
            except Exception:
                pass
        app = None
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        REPORT_PATH.write_text("\n".join(REPORT), encoding="utf-8")
        print(f"\nreport: {REPORT_PATH}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
