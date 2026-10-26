# -*- coding: utf-8 -*-
r"""
probe_kinetic_phase.py — ReactionPhase 枚举值探测 + 参考 Case 求解复验
======================================================================
在 HYSYS 工作站运行：

    python tools\probe_kinetic_phase.py

1) 打开 kinetic_cstr_ref.hsc，补进料组成、求解、读 Vap1/Liq1（本脚本自行记录）；
2) 对 kineticrxn 逐个尝试 ReactionPhase = 0..8 和常见枚举名，记录写后读回值，
   找到"读回=1（与 GUI 参考一致）"的写入值；
3) 用该值重建 CSTR 并求解，验证产物是否出现 CO/H2O。
输出 output\probe_kinetic_phase_report.txt。
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
import probe_kinetic_diag as diag  # noqa: E402

PROGID = "HYSYS.Application"
REPORT_PATH = ROOT / "output" / "probe_kinetic_phase_report.txt"
REF_CSTR = ROOT / "output" / "kinetic_cstr_ref.hsc"
REPORT: list[str] = []


def log(message: str = "") -> None:
    print(message, flush=True)
    REPORT.append(str(message))


def err(exc: BaseException) -> str:
    if isinstance(exc, pywintypes.com_error):
        hr = getattr(exc, "hresult", None)
        return f"HRESULT={hr and f'0x{hr & 0xFFFFFFFF:08X}'} {getattr(exc, 'excepinfo', None)!r}"
    return f"{type(exc).__name__}: {exc}"


def local_read_stream(stream: Any, components: list[str], label: str) -> None:
    log(f"    --- {label} ---")
    try:
        t = getattr(stream, "Temperature").GetValue("K")
        f = getattr(stream, "MolarFlow").GetValue("kgmole/h")
        log(f"      T={float(t):.4g} K, F={float(f):.6g} kgmole/h")
    except Exception as exc:  # noqa: BLE001
        log(f"      T/F read failed: {err(exc)}")
    try:
        raw = [float(v) for v in stream.ComponentMolarFractionValue]
        log("      mole frac: " + str(
            {components[i]: round(v, 6) for i, v in enumerate(raw[: len(components)])}))
    except Exception as exc:  # noqa: BLE001
        log(f"      mole frac read failed: {err(exc)}")


def solve_ref_case(app: Any) -> None:
    log("=" * 68)
    log("PART 1: GUI reference case solve")
    log("=" * 68)
    if not REF_CSTR.exists():
        log(f"  MISSING: {REF_CSTR}")
        return
    ref_case = app.SimulationCases.Open(str(REF_CSTR))
    try:
        fp = ref_case.BasisManager.FluidPackages.Item(0)
        components = base._component_names(fp)
        feed = ref_case.Flowsheet.MaterialStreams.Item("Feed1")
        fractions = {"CO2": 0.5, "Hydrogen": 0.5,
                     "Methane": 0.0, "CO": 0.0, "H2O": 0.0}
        feed.ComponentMolarFractionValue = tuple(
            float(fractions.get(name, 0.0)) for name in components
        )
        base._wait_solver(ref_case)
        for name in ("Vap1", "Liq1"):
            try:
                local_read_stream(
                    ref_case.Flowsheet.MaterialStreams.Item(name),
                    components, name,
                )
            except Exception as exc:  # noqa: BLE001
                log(f"    read {name} failed: {err(exc)}")
    finally:
        try:
            ref_case.Close(False)
        except Exception:
            pass


def probe_phase_values(case: Any, fp: Any, components: list[str]) -> int:
    log("=" * 68)
    log("PART 2: ReactionPhase write/readback mapping")
    log("=" * 68)
    manager = case.BasisManager.ReactionPackageManager
    mapping: dict[str, Any] = {}
    candidates: list[Any] = [0, 1, 2, 3, 4, 5, 6, 7, 8,
                             "Overall", "Liquid", "Vapour", "CombinedLiquid",
                             "CombinedVapour", "Aqueous", "VapourPhase",
                             "LiquidPhase", "OverallPhase"]
    for index, candidate in enumerate(candidates):
        rset = manager.ReactionSets.Add(f"PSet_{index}", "Kinetic")
        rxn = manager.Reactions.Add(f"PRxn_{index}", "kineticrxn")
        for name in ("CO2", "Hydrogen", "CO", "H2O"):
            rxn.Reactants.Add(name)
        rxn.ReactantStoichCoefValue = (-1.0, -1.0, 1.0, 1.0)
        rxn.ComponentForwardOrderValue = (1.0, 1.0, 0.0, 0.0)
        rset.ActiveReactions.Add(rxn)
        try:
            setattr(rxn, "ReactionPhase", candidate)
            write_status = "SET-OK"
        except Exception as exc:  # noqa: BLE001
            write_status = f"SET-FAIL({err(exc)})"
        try:
            readback = getattr(rxn, "ReactionPhase")
        except Exception as exc:  # noqa: BLE001
            readback = f"GET-FAIL({err(exc)})"
        log(f"  write {candidate!r:20s} -> {write_status:10s} | readback = {readback!r}")
        mapping[str(candidate)] = readback
        rset.AssociateFluidPackage(fp)
        fp.ReactionPackage.ReactionSets.Add(rset)
    # 找读回 == 1 的候选
    winner = None
    for candidate, readback in mapping.items():
        if readback == 1:
            winner = candidate
            break
    log(f"  candidate giving readback=1: {winner!r}")
    if winner is None:
        return 1
    try:
        return int(winner)
    except (TypeError, ValueError):
        return str(winner)


def build_cstr_winner(case: Any, fp: Any, components: list[str], phase_value: Any) -> None:
    log("=" * 68)
    log(f"PART 3: CSTR with ReactionPhase={phase_value!r}")
    log("=" * 68)
    manager = case.BasisManager.ReactionPackageManager
    rset = manager.ReactionSets.Add("KinSetW", "Kinetic")
    rxn = manager.Reactions.Add("Rxn-W", "kineticrxn")
    for name in ("CO2", "Hydrogen", "CO", "H2O"):
        rxn.Reactants.Add(name)
    rxn.ReactantStoichCoefValue = (-1.0, -1.0, 1.0, 1.0)
    rxn.ComponentForwardOrderValue = (1.0, 1.0, 0.0, 0.0)
    rxn.ComponentReverseOrderValue = (0.0, 0.0, 1.0, 1.0)
    rset.ActiveReactions.Add(rxn)
    setattr(rxn, "ReactionPhase", phase_value)
    for attr, value in (("Basis", 3), ("BasisUnits2", "kgmole/m3"),
                        ("RateUnits2", "kgmole/m3-s")):
        setattr(rxn, attr, value)
    setattr(rxn, "ForwardActivationEnergy", 50000.0)
    setattr(rxn, "ForwardFrequencyFactor", 1000.0)
    setattr(rxn, "ReverseActivationEnergy", 0.0)
    setattr(rxn, "ReverseFrequencyFactor", 0.0)
    rset.AssociateFluidPackage(fp)
    fp.ReactionPackage.ReactionSets.Add(rset)
    log(f"  phase readback after full config: {getattr(rxn, 'ReactionPhase')!r}")

    feed = diag.make_feed(case, "FEED-W", components)
    cstr = case.Flowsheet.Operations.Add("CSTR-W", "KineticReactorop")
    vapour = case.Flowsheet.MaterialStreams.Add("CSTR-W-VAP")
    liquid = case.Flowsheet.MaterialStreams.Add("CSTR-W-LIQ")
    cstr.Feeds.Add(feed)
    cstr.VapourProduct = vapour
    cstr.LiquidProduct = liquid
    cstr.ReactionSet = rset
    cstr.PressureDropValue = 0.0
    setattr(cstr, "VolumeValue", 1.0)
    setattr(cstr, "LiquidLevelValue", 50.0)
    base._wait_solver(case)
    local_read_stream(vapour, components, "CSTR-W-VAP")
    local_read_stream(liquid, components, "CSTR-W-LIQ")


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

        solve_ref_case(app)
        case, fp, feed, components = base._setup_case(app)
        phase_value = probe_phase_values(case, fp, components)
        build_cstr_winner(case, fp, components, phase_value)

        log("\n=== summary ===")
        log("  PART1 产物有 CO/H2O -> GUI 构建可反应（对照基准）；")
        log("  PART2 给出写后读回=1 的枚举值；PART3 验证该相态下 CSTR 是否转化。")
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
