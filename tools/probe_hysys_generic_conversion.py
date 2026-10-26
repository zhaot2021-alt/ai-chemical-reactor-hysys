# -*- coding: utf-8 -*-
"""
probe_hysys_generic_conversion.py — 从零创建 Conversion 反应并跑通
================================================================
1. 只读参考模板，读取 Rxn-1 的 stoich/base/conversion/phase
2. 新 Case 中创建：ReactionSet(Conversion) → conversionrxn → Reactants.Add(5)
   → 设置化学计量、BaseComponent、50% 转化率 → AssociateFluidPackage
   → 创建 conversionreactorop、连接 Feed1/Prod1/Liq1 → 求解
目标：验证能否不再依赖 ref_conversion.hsc 模板。
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
TEMPLATE = ROOT / "output" / "ref_conversion.hsc"
REPORT_PATH = ROOT / "output" / "probe_hysys_generic_conversion_report.txt"

COMPONENTS = ("Toluene", "Benzene", "o-Xylene", "m-Xylene", "p-Xylene")
STOICH = (-2.0, 1.0, 1.0 / 3.0, 1.0 / 3.0, 1.0 / 3.0)
FEED_T_K = 653.15
FEED_P_KPA = 2500.0
FEED_FLOW = 10000.0 / 92.1402  # 纯甲苯，10000 kg/h

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


def read_template(app) -> None:
    report("\n" + "=" * 78)
    report("PART 1: reference reaction values")
    report("=" * 78)
    if not TEMPLATE.exists():
        report(f"template missing: {TEMPLATE}")
        return
    case = None
    try:
        case = app.SimulationCases.Open(str(TEMPLATE))
        fp = case.BasisManager.FluidPackages.Item(0)
        rset = fp.ReactionPackage.ReactionSets.Item("Set-1")
        active = rset.ActiveReactions
        count = int(active.Count)
        report(f"ActiveReactions.Count={count}")
        for idx in range(count):
            rxn = active.Item(idx)
            report(f"\nRxn[{idx}] {rxn!r}")
            for attr in (
                "Name", "TypeName", "ReactantName", "ReactantStoichCoefValue",
                "BaseComponent", "ConversionCoefficientsValue", "ReactionPhase",
            ):
                safe(rxn, attr, "Rxn")
            reactions = safe(rxn, "Reactants", "Rxn")
            if reactions is not None:
                try:
                    rcount = int(reactions.Count)
                    report(f"  Reactants.Count={rcount}, Names={reactions.Names!r}")
                    for j in range(rcount):
                        item = reactions.Item(j)
                        report(f"  Reactants.Item({j}) = {item!r} type={type(item).__name__}")
                        for attr in ("Name", "ReactantName", "ReactantStoichCoefValue",
                                     "StoichCoefValue", "Value"):
                            safe(item, attr, f"Reactant[{j}]")
                except Exception as exc:  # noqa: BLE001
                    report(f"  Reactants inspect failed: {com_detail(exc)}")
        try:
            case.Close(False)
        except Exception:
            pass
        case = None
    except Exception as exc:  # noqa: BLE001
        report(f"read template failed: {com_detail(exc)}")
        report(traceback.format_exc())
    finally:
        if case is not None:
            try:
                case.Close(False)
            except Exception:
                pass
        gc.collect()


def generic_test(app) -> None:
    report("\n" + "=" * 78)
    report("PART 2: create/configure/solve a generic Conversion reactor")
    report("=" * 78)
    case = None
    try:
        try:
            case = app.SimulationCases.Add(
                f"GenericConv_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
        except Exception:
            case = app.SimulationCases.Add()
        basis = case.BasisManager
        basis.StartBasisChange()
        fp = basis.FluidPackages.Add("GenericConvPkg")
        fp.PropertyPackageName = config.PROPERTY_PACKAGE_NAME
        for cid in COMPONENTS:
            fp.Components.Add(cid)
        basis.EndBasisChange()

        manager = basis.ReactionPackageManager
        rset = manager.ReactionSets.Add("GenericSet", "Conversion")
        report(f"ReactionSet created: {rset!r}")
        try:
            rset.AssociateFluidPackage(fp)
            report("AssociateFluidPackage OK")
        except Exception as exc:  # noqa: BLE001
            report(f"AssociateFluidPackage failed: {com_detail(exc)}")

        rxn = manager.Reactions.Add("GenericRxn", "conversionrxn")
        report(f"Reaction created: {rxn!r}")
        reactants = rxn.Reactants
        for cid in COMPONENTS:
            try:
                item = reactants.Add(cid)
                report(f"Reactants.Add({cid!r}) OK -> {item!r}")
            except Exception as exc:  # noqa: BLE001
                report(f"Reactants.Add({cid!r}) failed: {com_detail(exc)}")

        try:
            report(f"before: ReactantStoichCoefValue = {rxn.ReactantStoichCoefValue!r}")
        except Exception as exc:  # noqa: BLE001
            report(f"read stoich failed: {com_detail(exc)}")
        for value in (tuple(STOICH), list(STOICH)):
            try:
                rxn.ReactantStoichCoefValue = value
                report(f"ReactantStoichCoefValue = {value!r} OK")
                break
            except Exception as exc:  # noqa: BLE001
                report(f"set stoich {type(value).__name__} failed: {com_detail(exc)}")

        base = None
        try:
            base = fp.Components.Item("Toluene")
        except Exception as exc:  # noqa: BLE001
            report(f"get Toluene component failed: {com_detail(exc)}")
        for candidate in (base, "Toluene"):
            if candidate is None:
                continue
            try:
                rxn.BaseComponent = candidate
                report(f"BaseComponent = {candidate!r} OK")
                break
            except Exception as exc:  # noqa: BLE001
                report(f"BaseComponent set failed: {com_detail(exc)}")

        try:
            rxn.ConversionCoefficientsValue = (50.0, 0.0, 0.0)
            report("ConversionCoefficientsValue = (50,0,0) OK")
        except Exception as exc:  # noqa: BLE001
            report(f"ConversionCoefficientsValue failed: {com_detail(exc)}")
        try:
            rxn.ReactionPhase = 5
            report("ReactionPhase = 5 OK")
        except Exception as exc:  # noqa: BLE001
            report(f"ReactionPhase failed: {com_detail(exc)}")

        try:
            rset.ActiveReactions.Add(rxn)
            report("ActiveReactions.Add(rxn) OK")
        except Exception as exc:  # noqa: BLE001
            report(f"ActiveReactions.Add failed: {com_detail(exc)}")
        try:
            fp.ReactionPackage.ReactionSets.Add(rset)
            report("FP.ReactionPackage.ReactionSets.Add(set) OK")
        except Exception as exc:  # noqa: BLE001
            report(f"FP.ReactionSets.Add failed: {com_detail(exc)}")

        feed = case.Flowsheet.MaterialStreams.Add("Feed1")
        feed.Temperature.SetValue(FEED_T_K, "K")
        feed.Pressure.SetValue(FEED_P_KPA, "kPa")
        feed.MolarFlow.SetValue(FEED_FLOW, "kgmole/h")
        feed.ComponentMolarFractionValue = (1.0, 0.0, 0.0, 0.0, 0.0)

        prod = case.Flowsheet.MaterialStreams.Add("Prod1")
        liq = case.Flowsheet.MaterialStreams.Add("Liq1")
        reactor = case.Flowsheet.Operations.Add("CRV-GEN", "conversionreactorop")
        reactor.Feeds.Add(feed)
        reactor.VapourProduct = prod
        reactor.LiquidProduct = liq
        reactor.PressureDropValue = 0.0
        try:
            reactor.ReactionSet = rset
            report("reactor.ReactionSet = set OK")
        except Exception as exc:  # noqa: BLE001
            report(f"reactor.ReactionSet failed: {com_detail(exc)}")

        try:
            case.Solver.CanSolve = True
            report(f"solved state IsSolving={case.Solver.IsSolving!r}")
            report(f"Prod1 T = {prod.Temperature.GetValue('K')!r}")
            report(f"Prod1 MolarFlow = {prod.MolarFlow.GetValue('kgmole/h')!r}")
            report(f"Prod1 composition = {prod.ComponentMolarFractionValue!r}")
            report(f"Prod1 VF = {prod.VapourFractionValue!r}")
        except Exception as exc:  # noqa: BLE001
            report(f"solve/read failed: {com_detail(exc)}")

        try:
            case.Close(False)
        except Exception:
            pass
        case = None
    except Exception as exc:  # noqa: BLE001
        report(f"generic test failed: {com_detail(exc)}")
        report(traceback.format_exc())
    finally:
        if case is not None:
            try:
                case.Close(False)
            except Exception:
                pass
        gc.collect()


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    app = None
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
        read_template(app)
        generic_test(app)
    except Exception as exc:  # noqa: BLE001
        report(f"EXCEPTION: {com_detail(exc)}")
        report(traceback.format_exc())
    finally:
        gc.collect()
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
