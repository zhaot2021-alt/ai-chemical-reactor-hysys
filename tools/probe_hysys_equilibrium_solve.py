# -*- coding: utf-8 -*-
"""
probe_hysys_equilibrium_solve.py — Equilibrium Ka 设置与闭环探针
================================================================
用参考反应 CO2 + H2 <-> CO + H2O（Ka=1.0，仅 API 验证，不是工艺结论）
验证：equilibriumrxn 的 KEqValue/EquilibriumConstant 设置、
      equilibriumreactorop 连接与求解。
"""

from __future__ import annotations

import gc
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
import config  # noqa: E402

PROGID = "HYSYS.Application"
REPORT_PATH = ROOT / "output" / "probe_hysys_equilibrium_solve_report.txt"
COMPONENTS = ("Methane", "H2O", "CO", "CO2", "Hydrogen")
STOICH = {"Methane": 0.0, "H2O": 1.0, "CO": 1.0, "CO2": -1.0, "Hydrogen": -1.0}
FEED_COMP = {"Methane": 0.0, "H2O": 0.0, "CO": 0.0, "CO2": 0.5, "Hydrogen": 0.5}
FEED_T_K = 800.0
FEED_P_KPA = 1000.0
FEED_FLOW = 100.0
KA_TEST = 1.0

_REPORT: list[str] = []


def report(msg: str = "") -> None:
    print(msg, flush=True)
    _REPORT.append(str(msg))


def com_detail(exc: BaseException) -> str:
    if isinstance(exc, pywintypes.com_error):
        hr = getattr(exc, "hresult", None)
        info = getattr(exc, "excepinfo", None)
        htxt = f"0x{hr & 0xFFFFFFFF:08X}" if isinstance(hr, int) else "N/A"
        return f"HRESULT={htxt}, excepinfo={info!r}"
    return f"{type(exc).__name__}: {exc}"


def safe(obj, attr: str, label: str):
    try:
        value = getattr(obj, attr)
        report(f"  {label}.{attr} = {value!r}")
        return value
    except Exception as exc:  # noqa: BLE001
        report(f"  {label}.{attr} failed: {com_detail(exc)}")
        return None


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    app = None
    case = None
    owns = False
    com_init = False
    try:
        pythoncom.CoInitialize()
        com_init = True
        try:
            app = win32com.client.GetActiveObject(PROGID)
            report("attached existing HYSYS")
        except Exception:
            app = win32com.client.Dispatch(PROGID)
            owns = True
            report("started new HYSYS")
        try:
            app.Visible = True
        except Exception:
            pass

        try:
            case = app.SimulationCases.Add(
                f"EqSolve_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
        except Exception:
            case = app.SimulationCases.Add()
        basis = case.BasisManager
        basis.StartBasisChange()
        fp = basis.FluidPackages.Add("EqSolvePkg")
        fp.PropertyPackageName = config.PROPERTY_PACKAGE_NAME
        for cid in COMPONENTS:
            fp.Components.Add(cid)
        basis.EndBasisChange()

        manager = basis.ReactionPackageManager
        rset = manager.ReactionSets.Add("EqSet", "Equilibrium")
        rset.AssociateFluidPackage(fp)
        rxn = manager.Reactions.Add("EqRxn", "equilibriumrxn")
        for cid in COMPONENTS:
            rxn.Reactants.Add(cid)
        rxn.ReactantStoichCoefValue = tuple(STOICH[cid] for cid in COMPONENTS)
        report("stoich set")

        for candidate in (
            ("KEqValue", KA_TEST),
            ("EquilibriumConstant", KA_TEST),
            ("EquilibriumConstantParameterArrayValue", (KA_TEST,)),
        ):
            attr, value = candidate
            try:
                setattr(rxn, attr, value)
                report(f"{attr} = {value!r} OK")
                break
            except Exception as exc:  # noqa: BLE001
                report(f"{attr} failed: {com_detail(exc)}")
        for attr in ("KEqValue", "EquilibriumConstant", "EquilibriumConstantParameterArrayValue",
                     "TemperatureApproachValue", "ReactionPhase"):
            safe(rxn, attr, "EqRxn")

        rset.ActiveReactions.Add(rxn)
        fp.ReactionPackage.ReactionSets.Add(rset)

        feed = case.Flowsheet.MaterialStreams.Add("FEED")
        feed.Temperature.SetValue(FEED_T_K, "K")
        feed.Pressure.SetValue(FEED_P_KPA, "kPa")
        feed.MolarFlow.SetValue(FEED_FLOW, "kgmole/h")
        feed.ComponentMolarFractionValue = tuple(FEED_COMP[cid] for cid in COMPONENTS)
        vapour = case.Flowsheet.MaterialStreams.Add("VAPOUR")
        liquid = case.Flowsheet.MaterialStreams.Add("LIQUID")
        reactor = case.Flowsheet.Operations.Add("R-EQ", "equilibriumreactorop")
        reactor.Feeds.Add(feed)
        reactor.VapourProduct = vapour
        reactor.LiquidProduct = liquid
        reactor.PressureDropValue = 0.0
        reactor.ReactionSet = rset
        case.Solver.CanSolve = True
        report(f"IsSolving={case.Solver.IsSolving!r}")
        for name, stream in (("FEED", feed), ("VAPOUR", vapour), ("LIQUID", liquid)):
            report(f"{name}: T={stream.Temperature.GetValue('K')!r} "
                   f"P={stream.Pressure.GetValue('kPa')!r} "
                   f"F={stream.MolarFlow.GetValue('kgmole/h')!r} "
                   f"z={stream.ComponentMolarFractionValue!r} "
                   f"VF={stream.VapourFractionValue!r}")

        try:
            case.Close(False)
        except Exception:
            pass
        case = None
    except Exception as exc:  # noqa: BLE001
        report(f"EXCEPTION: {com_detail(exc)}")
        report(traceback.format_exc())
    finally:
        gc.collect()
        if case is not None:
            try:
                case.Close(False)
            except Exception:
                pass
        if app is not None and owns:
            try:
                app.Quit()
                time.sleep(2)
            except Exception:
                pass
        app = None
        gc.collect()
        if com_init:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        REPORT_PATH.write_text("\n".join(_REPORT), encoding="utf-8")
        print(f"\nreport: {REPORT_PATH}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
