# -*- coding: utf-8 -*-
"""
probe_hysys_equilibrium.py — Equilibrium 反应/反应器配置探针
================================================================
只验证：
  1. ReactionSets.Add(name, "Equilibrium") 是否可用
  2. Reactions.Add(name, "equilibriumrxn"/候选类型) 是否可用
  3. 新反应对象的成员（平衡常数/吉布斯自由能/相态/基准组分）
  4. Reactants.Add / ReactantStoichCoefValue / AssociateFluidPackage / ActiveReactions.Add
  5. 顺带尝试 equilibriumreactorop 创建与连接（不配置热力学）
不保存；测试完关闭临时 Case。
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
REPORT_PATH = ROOT / "output" / "probe_hysys_equilibrium_report.txt"
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


def items(coll) -> list:
    try:
        count = int(coll.Count)
    except Exception:
        return []
    for start in (0, 1):
        result = []
        for i in range(start, start + count):
            try:
                result.append(coll.Item(i))
            except Exception:
                result = []
                break
        if result:
            return result
    return []


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
                f"EqProbe_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
        except Exception:
            case = app.SimulationCases.Add()
        basis = case.BasisManager
        basis.StartBasisChange()
        fp = basis.FluidPackages.Add("EqProbePkg")
        fp.PropertyPackageName = config.PROPERTY_PACKAGE_NAME
        for cid in COMPONENTS:
            fp.Components.Add(cid)
        basis.EndBasisChange()

        manager = basis.ReactionPackageManager
        dump_members(manager, "ReactionPackageManager", filters=None)
        reactions = manager.Reactions
        dump_members(reactions, "Reactions", filters=None)

        rset = None
        for set_type in ("Equilibrium", "equilibrium", "EquilibriumReaction"):
            try:
                rset = manager.ReactionSets.Add(f"EqSet_{set_type}", set_type)
                report(f"ReactionSets.Add({set_type!r}) OK -> {rset!r}")
                break
            except Exception as exc:  # noqa: BLE001
                report(f"ReactionSets.Add({set_type!r}) failed: {com_detail(exc)}")

        rxn = None
        for rxn_type in (
            "equilibriumrxn", "Equilibrium", "equilibrium",
            "EquilibriumReaction", "equilibriumreaction",
        ):
            try:
                rxn = reactions.Add(f"EqRxn_{rxn_type}", rxn_type)
                report(f"Reactions.Add({rxn_type!r}) OK -> {rxn!r}")
                break
            except Exception as exc:  # noqa: BLE001
                report(f"Reactions.Add({rxn_type!r}) failed: {com_detail(exc)}")

        if rxn is None and rset is not None:
            for candidate in ("equilibriumrxn", "Equilibrium"):
                try:
                    rxn = rset.ActiveReactions.Add("EqRxnActive", candidate)
                    report(f"ActiveReactions.Add({candidate!r}) OK -> {rxn!r}")
                    break
                except Exception as exc:  # noqa: BLE001
                    report(f"ActiveReactions.Add({candidate!r}) failed: {com_detail(exc)}")

        if rxn is not None:
            report("\n=== NEW equilibrium reaction ===")
            safe(rxn, "Name", "EqRxn")
            safe(rxn, "TypeName", "EqRxn")
            safe(rxn, "VisibleTypeName", "EqRxn")
            dump_members(
                rxn, "Equilibrium Reaction",
                filters=(
                    "equilib", "constant", "keq", "gibbs", "free", "reactant",
                    "stoich", "base", "phase", "temperature", "composition",
                    "component",
                ),
            )
            for attr in (
                "BaseComponent", "Reactants", "ReactionPhase",
                "ReactantStoichCoefValue", "EquilibriumConstantValue",
                "EquilibriumConstants", "GibbsFreeEnergy",
            ):
                safe(rxn, attr, "EqRxn")
            reactants = safe(rxn, "Reactants", "EqRxn")
            if reactants is not None:
                dump_members(reactants, "EqRxn.Reactants", filters=None)
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
                    report(f"ReactantStoichCoefValue set failed: {com_detail(exc)}")

            if rset is not None:
                try:
                    rset.AssociateFluidPackage(fp)
                    report("AssociateFluidPackage OK")
                except Exception as exc:  # noqa: BLE001
                    report(f"AssociateFluidPackage failed: {com_detail(exc)}")
                try:
                    rset.ActiveReactions.Add(rxn)
                    report("ActiveReactions.Add(rxn) OK")
                except Exception as exc:  # noqa: BLE001
                    report(f"ActiveReactions.Add(rxn) failed: {com_detail(exc)}")
                try:
                    fp.ReactionPackage.ReactionSets.Add(rset)
                    report("FP ReactionSets.Add(set) OK")
                except Exception as exc:  # noqa: BLE001
                    report(f"FP ReactionSets.Add(set) failed: {com_detail(exc)}")

        for reactor_type in ("equilibriumreactorop", "cstrop", "pfrreactorop", "cstr", "pfr"):
            try:
                op = case.Flowsheet.Operations.Add(f"Probe_{reactor_type}", reactor_type)
                report(f"Operations.Add({reactor_type!r}) OK -> {op!r}")
                dump_members(
                    op, f"{reactor_type}",
                    filters=("feed", "product", "reaction", "temperature", "pressure",
                             "energy", "duty", "conversion", "equilibrium"),
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
