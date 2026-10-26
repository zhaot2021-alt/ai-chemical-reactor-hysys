# -*- coding: utf-8 -*-
"""
probe_hysys_kinetics.py — CSTR/PFR 与 kinetic 反应 API 探针
================================================================
只验证：
  1. Operations.Add 的 CSTR/PFR 真实 TypeName
  2. ReactionSets.Add(..., Type) 的 Kinetic 类型
  3. Reactions.Add(..., Type) 的 kinetic 反应类型与成员
不配置速率参数、不求解；测试完关闭临时 Case。
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
REPORT_PATH = ROOT / "output" / "probe_hysys_kinetics_report.txt"
COMPONENTS = ("Methane", "H2O", "CO", "CO2", "Hydrogen")
INVKIND = {1: "FUNC", 2: "GET", 4: "PUT", 8: "PUTREF"}
IGNORE = {"QueryInterface", "AddRef", "Release", "GetTypeInfoCount",
          "GetTypeInfo", "GetIDsOfNames", "Invoke"}
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


def dump_members(obj, label: str, filters: tuple[str, ...] | None = None) -> None:
    report(f"\n--- {label} ---")
    try:
        ti = obj._oleobj_.GetTypeInfo()
        ta = ti.GetTypeAttr()
    except Exception as exc:  # noqa: BLE001
        report(f"  type info failed: {com_detail(exc)}")
        return
    report(f"  cFuncs={ta.cFuncs} cVars={ta.cVars} iid={ta.iid}")
    for i in range(ta.cFuncs):
        try:
            fd = ti.GetFuncDesc(i)
            names = ti.GetNames(fd.memid)
        except Exception:
            continue
        name = names[0] if names else "?"
        if name in IGNORE:
            continue
        if filters and not any(f.lower() in name.lower() for f in filters):
            continue
        inv = INVKIND.get(fd.invkind, str(fd.invkind))
        report(f"  {name:44s} | {inv:6s} | params={list(names[1:])}")


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
                f"KineticsProbe_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
        except Exception:
            case = app.SimulationCases.Add()
        basis = case.BasisManager
        basis.StartBasisChange()
        fp = basis.FluidPackages.Add("KineticsProbePkg")
        fp.PropertyPackageName = config.PROPERTY_PACKAGE_NAME
        for cid in COMPONENTS:
            fp.Components.Add(cid)
        basis.EndBasisChange()

        manager = basis.ReactionPackageManager
        dump_members(manager, "ReactionPackageManager", filters=None)

        rset = None
        for set_type in ("Kinetic", "kinetic", "Kinetics", "kinetics"):
            try:
                rset = manager.ReactionSets.Add(f"KinSet_{set_type}", set_type)
                report(f"ReactionSets.Add({set_type!r}) OK -> {rset!r}")
                break
            except Exception as exc:  # noqa: BLE001
                report(f"ReactionSets.Add({set_type!r}) failed: {com_detail(exc)}")

        reactions = manager.Reactions
        rxn = None
        for rxn_type in ("kineticrxn", "kineticsrxn", "Kinetic", "kinetic",
                         "KineticReaction", "kineticreaction"):
            try:
                rxn = reactions.Add(f"KinRxn_{rxn_type}", rxn_type)
                report(f"Reactions.Add({rxn_type!r}) OK -> {rxn!r}")
                break
            except Exception as exc:  # noqa: BLE001
                report(f"Reactions.Add({rxn_type!r}) failed: {com_detail(exc)}")

        if rxn is not None:
            report("\n=== NEW kinetic reaction ===")
            safe(rxn, "Name", "KinRxn")
            safe(rxn, "TypeName", "KinRxn")
            safe(rxn, "VisibleTypeName", "KinRxn")
            dump_members(rxn, "Kinetic Reaction", filters=None)
            for attr in (
                "Reactants", "ReactantStoichCoefValue", "ReactionPhase",
                "RateConstant", "RateConstantValue", "ActivationEnergyValue",
                "PreExponentialFactorValue", "RateUnits", "Basis",
            ):
                safe(rxn, attr, "KinRxn")
            reactants = safe(rxn, "Reactants", "KinRxn")
            if reactants is not None:
                dump_members(reactants, "KinRxn.Reactants", filters=None)
                for cid in COMPONENTS:
                    try:
                        reactants.Add(cid)
                        report(f"Reactants.Add({cid!r}) OK")
                    except Exception as exc:  # noqa: BLE001
                        report(f"Reactants.Add({cid!r}) failed: {com_detail(exc)}")
                try:
                    rxn.ReactantStoichCoefValue = (-1.0, 1.0, 1.0, -1.0, 3.0)
                    report("ReactantStoichCoefValue set OK (probe values)")
                except Exception as exc:  # noqa: BLE001
                    report(f"ReactantStoichCoefValue failed: {com_detail(exc)}")
            if rset is not None:
                try:
                    rset.AssociateFluidPackage(fp)
                    report("AssociateFluidPackage OK")
                except Exception as exc:  # noqa: BLE001
                    report(f"AssociateFluidPackage failed: {com_detail(exc)}")
                try:
                    rset.ActiveReactions.Add(rxn)
                    report("ActiveReactions.Add OK")
                except Exception as exc:  # noqa: BLE001
                    report(f"ActiveReactions.Add failed: {com_detail(exc)}")
                try:
                    fp.ReactionPackage.ReactionSets.Add(rset)
                    report("FP ReactionSets.Add OK")
                except Exception as exc:  # noqa: BLE001
                    report(f"FP ReactionSets.Add failed: {com_detail(exc)}")

        for reactor_type in (
            "cstrop", "cstr", "CSTROp", "CSTR",
            "pfrreactorop", "pfrreactor", "pfrop", "pfr", "PFROp", "PFR",
            "plugflowreactorop",
        ):
            try:
                op = case.Flowsheet.Operations.Add(f"Probe_{reactor_type}", reactor_type)
                report(f"\nOperations.Add({reactor_type!r}) OK -> {op!r}")
                dump_members(
                    op, f"{reactor_type}",
                    filters=("feed", "product", "reaction", "temperature", "pressure",
                             "energy", "duty", "volume", "length", "diameter",
                             "conversion", "kinetic"),
                )
                break
            except Exception as exc:  # noqa: BLE001
                report(f"Operations.Add({reactor_type!r}) failed: {com_detail(exc)}")

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
