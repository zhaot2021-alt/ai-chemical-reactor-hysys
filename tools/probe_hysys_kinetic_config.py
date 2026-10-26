# -*- coding: utf-8 -*-
r"""
probe_hysys_kinetic_config.py — 验证 kinetic 反应参数 + CSTR/PFR 闭环
=====================================================================
在 HYSYS 工作站运行：

    python tools\probe_hysys_kinetic_config.py --visible

只验证已经探测到的真实成员：
  - kineticrxn: ForwardActivationEnergy / ForwardFrequencyFactor /
    ComponentForwardOrderValue / ReactionPhase / Basis / RateUnits2
  - PFReactorop: Product / EnergyStream / ReactionSet / TotalVolumeValue /
    TubeLengthValue / TubeDiameterValue
  - CSTR 候选 TypeName: KineticReactorop 等

反应：CO2 + H2 -> CO + H2O（探针用，不做工艺结论）。
生成 output\probe_hysys_kinetic_config_report.txt。
"""

from __future__ import annotations

import gc
import sys
import time
import traceback
from pathlib import Path
from typing import Any, Iterable

import pythoncom
import pywintypes
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tools"))
import config  # noqa: E402
import probe_hysys_kinetic_api as base  # noqa: E402

PROGID = "HYSYS.Application"
REPORT_PATH = ROOT / "output" / "probe_hysys_kinetic_config_report.txt"
COMPONENTS = ("Methane", "H2O", "CO", "CO2", "Hydrogen")
REPORT: list[str] = []


def log(message: str = "") -> None:
    print(message, flush=True)
    REPORT.append(str(message))


def com_detail(exc: BaseException) -> str:
    if isinstance(exc, pywintypes.com_error):
        hr = getattr(exc, "hresult", None)
        htxt = f"0x{hr & 0xFFFFFFFF:08X}" if isinstance(hr, int) else "N/A"
        return f"HRESULT={htxt}, excepinfo={getattr(exc, 'excepinfo', None)!r}"
    return f"{type(exc).__name__}: {exc}"


def describe_prop(obj: Any, attr: str) -> Any:
    try:
        value = getattr(obj, attr)
    except Exception as exc:  # noqa: BLE001
        log(f"  {attr} GET failed: {com_detail(exc)}")
        return None
    extra = ""
    if hasattr(value, "SetValue"):
        extra = " [ProcessVariable]"
    log(f"  {attr} = {value!r} (type={type(value).__name__}){extra}")
    return value


def try_set_prop(
    obj: Any,
    attr: str,
    value: Any,
    units: Iterable[str] = (),
    *,
    label: str = "",
) -> bool:
    label = label or attr
    try:
        current = getattr(obj, attr)
    except Exception as exc:  # noqa: BLE001
        log(f"  SET {label}: GET failed {com_detail(exc)}")
        return False

    if hasattr(current, "SetValue"):
        for unit in units:
            try:
                current.SetValue(value, unit)
                log(f"  SET {label} = {value!r} [{unit}] OK (ProcessVariable)")
                return True
            except Exception as exc:  # noqa: BLE001
                log(f"  SET {label} = {value!r} [{unit}] failed: {com_detail(exc)}")
    try:
        setattr(obj, attr, value)
        log(f"  SET {label} = {value!r} OK (direct)")
        return True
    except Exception as exc:  # noqa: BLE001
        log(f"  SET {label} = {value!r} failed (direct): {com_detail(exc)}")
    return False


def read_process_value(obj: Any, attr: str, units: Iterable[str] = ()) -> None:
    try:
        current = getattr(obj, attr)
    except Exception as exc:  # noqa: BLE001
        log(f"  {attr} readback GET failed: {com_detail(exc)}")
        return
    if hasattr(current, "GetValue"):
        for unit in units:
            try:
                log(f"  {attr} readback = {current.GetValue(unit)!r} [{unit}]")
                return
            except Exception:
                continue
    log(f"  {attr} readback = {current!r}")


def read_stream(stream: Any, components: list[str], label: str) -> None:
    log(f"  --- {label} ---")
    for attr, unit in (
        ("Temperature", "K"), ("Pressure", "kPa"),
        ("MolarFlow", "kgmole/h"), ("MassFlow", "kg/h"),
    ):
        try:
            value = getattr(stream, attr).GetValue(unit)
            log(f"    {attr:16s} = {float(value):.6g} {unit}")
        except Exception as exc:  # noqa: BLE001
            log(f"    {attr:16s} GET failed: {com_detail(exc)}")
    try:
        log(f"    VapourFraction   = {float(stream.VapourFractionValue):.6g}")
    except Exception as exc:  # noqa: BLE001
        log(f"    VapourFraction GET failed: {com_detail(exc)}")
    try:
        raw = [float(value) for value in stream.ComponentMolarFractionValue]
        shown = {
            components[index]: round(value, 6)
            for index, value in enumerate(raw[: len(components)])
        }
        log(f"    mole fractions   = {shown}")
    except Exception as exc:  # noqa: BLE001
        log(f"    mole fractions GET failed: {com_detail(exc)}")


def setup(app: Any) -> tuple[Any, Any, Any, list[str]]:
    log("\n=== 1. case / feed / kinetic reaction ===")
    case, fp, feed, components = base._setup_case(app)
    read_stream(feed, components, "FEED")
    manager = case.BasisManager.ReactionPackageManager

    rset = manager.ReactionSets.Add("KineticSet", "Kinetic")
    rset.AssociateFluidPackage(fp)
    fp.ReactionPackage.ReactionSets.Add(rset)
    log(f"  reaction set: {rset!r}")

    rxn = manager.Reactions.Add("RWGS-KIN", "kineticrxn")
    log(f"  kinetic reaction: {rxn!r}; TypeName={rxn.TypeName!r}")
    for component_name in components:
        rxn.Reactants.Add(component_name)
    coefficients = {
        "CO2": -1.0, "Hydrogen": -1.0, "CO": 1.0, "H2O": 1.0,
        "Methane": 0.0,
    }
    rxn.ReactantStoichCoefValue = tuple(
        coefficients.get(name, 0.0) for name in components
    )
    log("  Reactants.Add + ReactantStoichCoefValue OK")

    log("\n--- current kinetic members ---")
    for attr in (
        "ReactionPhase", "Basis", "BasisUnits2", "RateUnits2",
        "ForwardActivationEnergy", "ForwardFrequencyFactor",
        "ReverseActivationEnergy", "ReverseFrequencyFactor",
        "ComponentForwardOrderValue", "ComponentReverseOrderValue",
        "ForwardTemperatureExponent", "ReverseTemperatureExponent",
    ):
        describe_prop(rxn, attr)

    log("\n--- set kinetic parameters ---")
    orders = tuple(1.0 if name in ("CO2", "Hydrogen") else 0.0 for name in components)
    try_set_prop(rxn, "ComponentForwardOrderValue", orders)
    try_set_prop(rxn, "ComponentReverseOrderValue", tuple(0.0 for _ in components))
    try_set_prop(rxn, "ReactionPhase", 1)
    try_set_prop(
        rxn, "ForwardActivationEnergy", 50000.0,
        ("kJ/kgmole", "J/mol", "kJ/kmol", "Btu/lbmole"),
    )
    try_set_prop(
        rxn, "ForwardFrequencyFactor", 1.0e3,
        ("m3/kgmole-s", "m3/kmol-s", "kgmole/m3-s", "1/s"),
    )
    try_set_prop(rxn, "ReverseActivationEnergy", 0.0,
                 ("kJ/kgmole", "J/mol", "kJ/kmol", "Btu/lbmole"))
    try_set_prop(rxn, "ReverseFrequencyFactor", 0.0,
                 ("m3/kgmole-s", "m3/kmol-s", "kgmole/m3-s", "1/s"))
    try_set_prop(rxn, "ForwardTemperatureExponent", 0.0)
    try_set_prop(rxn, "ReverseTemperatureExponent", 0.0)

    log("\n--- readback after set ---")
    for attr in (
        "ReactionPhase", "Basis", "BasisUnits2", "RateUnits2",
        "ForwardActivationEnergy", "ForwardFrequencyFactor",
        "ComponentForwardOrderValue",
    ):
        read_process_value(rxn, attr, ("kJ/kgmole", "J/mol", "kJ/kmol"))

    try:
        rset.ActiveReactions.Add(rxn)
        log("  ActiveReactions.Add OK")
    except Exception as exc:  # noqa: BLE001
        log(f"  ActiveReactions.Add failed: {com_detail(exc)}")
    return case, fp, feed, components


def _new_feed(case: Any, name: str, components: list[str]) -> Any:
    stream = case.Flowsheet.MaterialStreams.Add(name)
    stream.Temperature.SetValue(config.c_to_k(500.0), "K")
    stream.Pressure.SetValue(config.bar_to_kpa(10.0), "kPa")
    stream.MolarFlow.SetValue(100.0, "kgmole/h")
    fractions = {
        "CO2": 0.5, "Hydrogen": 0.5,
        "Methane": 0.0, "H2O": 0.0, "CO": 0.0,
    }
    stream.ComponentMolarFractionValue = tuple(
        fractions.get(name, 0.0) for name in components
    )
    base._wait_solver(case)
    return stream


def set_design_value(obj: Any, attr: str, value: float, units: str) -> None:
    try:
        current = getattr(obj, attr)
    except Exception as exc:  # noqa: BLE001
        log(f"    {attr} GET failed: {com_detail(exc)}")
        return
    if hasattr(current, "SetValue"):
        try:
            current.SetValue(value, units)
            log(f"    SET {attr} = {value} [{units}] OK")
            return
        except Exception as exc:  # noqa: BLE001
            log(f"    SET {attr} = {value} [{units}] failed: {com_detail(exc)}")
    try:
        setattr(obj, attr, value)
        log(f"    SET {attr} = {value} OK (direct)")
    except Exception as exc:  # noqa: BLE001
        log(f"    SET {attr} = {value} failed: {com_detail(exc)}")


def build_pfr(case: Any, feed: Any, rset: Any, components: list[str]) -> None:
    log("\n=== 2. PFR (PFReactorop) ===")
    try:
        pfr = case.Flowsheet.Operations.Add("PFR-KIN", "PFReactorop")
    except Exception as exc:  # noqa: BLE001
        log(f"  PFR add failed: {com_detail(exc)}")
        return
    log(f"  PFR: {pfr!r}; TypeName={pfr.TypeName!r}")
    product = case.Flowsheet.MaterialStreams.Add("PFR-PROD")
    energy = case.Flowsheet.EnergyStreams.Add("Q-PFR")
    try:
        pfr.Feeds.Add(feed)
        pfr.Product = product
        pfr.EnergyStream = energy
        pfr.ReactionSet = rset
        pfr.PressureDropValue = 0.0
        log("  Feeds/Product/EnergyStream/ReactionSet connected")
    except Exception as exc:  # noqa: BLE001
        log(f"  connection failed: {com_detail(exc)}")

    set_design_value(pfr, "TotalVolumeValue", 1.0, "m3")
    set_design_value(pfr, "TubeLengthValue", 10.0, "m")
    set_design_value(pfr, "TubeDiameterValue", 0.3568, "m")
    set_design_value(pfr, "HeatFlowValue", 0.0, "kJ/h")
    try:
        energy.HeatFlow.SetValue(0.0, "kJ/h")
        log("  Q-PFR heat flow = 0 kJ/h")
    except Exception as exc:  # noqa: BLE001
        log(f"  Q-PFR heat flow failed: {com_detail(exc)}")

    base._wait_solver(case)
    read_stream(product, components, "PFR-PROD")


def build_cstr(case: Any, rset: Any, components: list[str]) -> None:
    log("\n=== 3. CSTR (KineticReactorop candidates) ===")
    cstr = None
    used_type = ""
    for type_string in (
        "KineticReactorop", "kineticreactorop", "KineticReactorOp",
        "kineticreactorOp", "KineticReactor", "kineticreactor",
    ):
        try:
            cstr = case.Flowsheet.Operations.Add("CSTR-KIN", type_string)
        except Exception as exc:  # noqa: BLE001
            log(f"  add {type_string!r} failed: {com_detail(exc)}")
            continue
        used_type = type_string
        log(f"  CSTR add OK: {type_string!r} -> {cstr!r}; "
            f"TypeName={cstr.TypeName!r}")
        break
    if cstr is None:
        log("  no CSTR type string succeeded")
        return

    feed = _new_feed(case, "FEED-CSTR", components)
    vapour = case.Flowsheet.MaterialStreams.Add("CSTR-VAP")
    liquid = case.Flowsheet.MaterialStreams.Add("CSTR-LIQ")
    energy = case.Flowsheet.EnergyStreams.Add("Q-CSTR")
    try:
        cstr.Feeds.Add(feed)
        cstr.VapourProduct = vapour
        cstr.LiquidProduct = liquid
        cstr.EnergyStream = energy
        cstr.ReactionSet = rset
        cstr.PressureDropValue = 0.0
        log("  Feeds/VapourProduct/LiquidProduct/EnergyStream/ReactionSet connected")
    except Exception as exc:  # noqa: BLE001
        log(f"  connection failed: {com_detail(exc)}")

    set_design_value(cstr, "VolumeValue", 1.0, "m3")
    set_design_value(cstr, "LiquidLevelValue", 50.0, "%")
    set_design_value(cstr, "HeatFlowValue", 0.0, "kJ/h")
    try:
        energy.HeatFlow.SetValue(0.0, "kJ/h")
        log("  Q-CSTR heat flow = 0 kJ/h")
    except Exception as exc:  # noqa: BLE001
        log(f"  Q-CSTR heat flow failed: {com_detail(exc)}")

    base._wait_solver(case)
    read_stream(vapour, components, "CSTR-VAP")
    read_stream(liquid, components, "CSTR-LIQ")
    log(f"  CSTR type used: {used_type!r}")


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
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
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

        case, fp, feed, components = setup(app)
        build_pfr(case, feed, case.BasisManager.ReactionPackageManager
                  .ReactionSets.Item("KineticSet"), components)
        build_cstr(case, case.BasisManager.ReactionPackageManager
                   .ReactionSets.Item("KineticSet"), components)

        log("\n=== 4. summary ===")
        log("  if PFR/CSTR products show real T/P/F and compositions,")
        log("  the kinetic + reactor path is usable.")
    except Exception as exc:  # noqa: BLE001
        log(f"EXCEPTION: {com_detail(exc)}")
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
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        REPORT_PATH.write_text("\n".join(REPORT), encoding="utf-8")
        print(f"\nreport: {REPORT_PATH}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
