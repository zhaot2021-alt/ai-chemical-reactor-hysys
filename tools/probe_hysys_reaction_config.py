# -*- coding: utf-8 -*-
"""
probe_hysys_reaction_config.py — Conversion 反应集/反应 COM 结构探针
================================================================
1. 只读打开 output/ref_conversion.hsc，导出 Set-1 / ActiveReactions / 反应对象
2. 在新临时 Case 中尝试 Reactions.Add 与 ReactionSet 挂载
不保存、不修改参考模板。
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
REPORT_PATH = ROOT / "output" / "probe_hysys_reaction_config_report.txt"

COMPONENTS = ("Toluene", "Benzene", "o-Xylene", "m-Xylene", "p-Xylene")
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


def safe(obj, attr: str, label: str):
    try:
        value = getattr(obj, attr)
        report(f"  {label}.{attr} = {value!r}")
        return value
    except Exception as exc:  # noqa: BLE001
        report(f"  {label}.{attr} failed: {com_detail(exc)}")
        return None


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
        report(f"  {name:42s} | {inv:6s} | params={list(names[1:])}")
    for i in range(ta.cVars):
        try:
            vd = ti.GetVarDesc(i)
            names = ti.GetNames(vd.memid)
        except Exception:
            continue
        name = names[0] if names else "?"
        if filters and not any(f.lower() in name.lower() for f in filters):
            continue
        report(f"  VAR {name}")


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


def describe_rxn(rxn, label: str) -> None:
    report(f"\n=== {label} ===")
    safe(rxn, "Name", label)
    safe(rxn, "TypeName", label)
    safe(rxn, "VisibleTypeName", label)
    dump_members(rxn, label, filters=None)
    for attr in (
        "BaseComponent", "Components", "Reactants", "Products",
        "ConversionCoefficients", "ConversionCoefficientsValue",
        "StoichiometricCoefficients", "ConversionValue",
    ):
        safe(rxn, attr, label)


def inspect_template(app) -> None:
    report("\n" + "=" * 78)
    report("PART 1: reference conversion template")
    report("=" * 78)
    if not TEMPLATE.exists():
        report(f"template missing: {TEMPLATE}")
        return
    case = None
    try:
        case = app.SimulationCases.Open(str(TEMPLATE))
        fp = case.BasisManager.FluidPackages.Item(0)
        rp = fp.ReactionPackage
        report(f"ReactionPackage = {rp!r}")
        dump_members(rp, "ReactionPackage", filters=None)
        sets = rp.ReactionSets
        dump_members(sets, "ReactionPackage.ReactionSets", filters=None)
        set_items = items(sets)
        report(f"ReactionSets count={len(set_items)} names={[safe(x,'Name','Set') for x in set_items]}")
        for s in set_items:
            name = safe(s, "Name", "Set")
            if name != "Set-1":
                continue
            dump_members(s, "Set-1", filters=None)
            for coll_attr in ("ActiveReactions", "InactiveReactions"):
                coll = safe(s, coll_attr, "Set-1")
                if coll is None:
                    continue
                dump_members(coll, f"Set-1.{coll_attr}", filters=None)
                rxns = items(coll)
                report(f"  {coll_attr} count={len(rxns)}")
                for idx, rxn in enumerate(rxns):
                    describe_rxn(rxn, f"{coll_attr}[{idx}]")
        try:
            case.Close(False)
        except Exception:
            pass
        del case
    except Exception as exc:  # noqa: BLE001
        report(f"inspect template failed: {com_detail(exc)}")
        report(traceback.format_exc())
    finally:
        gc.collect()


def create_test(app) -> None:
    report("\n" + "=" * 78)
    report("PART 2: attempt new reaction in temp case")
    report("=" * 78)
    case = None
    try:
        try:
            case = app.SimulationCases.Add(
                f"RxnConfigProbe_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
        except Exception:
            case = app.SimulationCases.Add()
        basis = case.BasisManager
        basis.StartBasisChange()
        fp = basis.FluidPackages.Add("Probe_RxnPkg")
        fp.PropertyPackageName = config.PROPERTY_PACKAGE_NAME
        for cid in COMPONENTS:
            fp.Components.Add(cid)
        basis.EndBasisChange()

        manager = basis.ReactionPackageManager
        dump_members(manager, "ReactionPackageManager", filters=None)
        reactions = manager.Reactions
        dump_members(reactions, "Reactions", filters=None)
        report(f"Reactions count={len(items(reactions))}")

        rset = None
        for set_type in ("Conversion", "conversion", "ConversionReaction"):
            try:
                rset = manager.ReactionSets.Add(f"Probe_Set_{set_type}", set_type)
                report(f"ReactionSets.Add({set_type!r}) OK -> {rset!r}")
                break
            except Exception as exc:  # noqa: BLE001
                report(f"ReactionSets.Add({set_type!r}) failed: {com_detail(exc)}")

        rxn = None
        if reactions is not None:
            for rxn_type in ("Conversion", "conversion", "ConversionReaction"):
                for args in (("Probe_Rxn", rxn_type), (rxn_type, "Probe_Rxn"), ("Probe_Rxn",)):
                    try:
                        rxn = reactions.Add(*args)
                        report(f"Reactions.Add{args} OK -> {rxn!r}")
                        break
                    except Exception as exc:  # noqa: BLE001
                        report(f"Reactions.Add{args} failed: {com_detail(exc)}")
                if rxn is not None:
                    break
        if rxn is not None:
            describe_rxn(rxn, "NEW Reaction")

        if rset is not None:
            for coll_attr in ("ActiveReactions", "InactiveReactions"):
                coll = safe(rset, coll_attr, "Probe_Set")
                if coll is not None:
                    dump_members(coll, f"Probe_Set.{coll_attr}", filters=None)
                    if rxn is not None:
                        try:
                            coll.Add(rxn)
                            report(f"{coll_attr}.Add(rxn) OK")
                        except Exception as exc:  # noqa: BLE001
                            report(f"{coll_attr}.Add(rxn) failed: {com_detail(exc)}")
        # attempts to attach reaction set to FP
        try:
            fp_rp = fp.ReactionPackage
            fp_sets = fp_rp.ReactionSets
            dump_members(fp_sets, "FP.ReactionPackage.ReactionSets", filters=None)
            if rset is not None:
                try:
                    fp_sets.Add(rset)
                    report("FP ReactionSets.Add(rset) OK")
                except Exception as exc:  # noqa: BLE001
                    report(f"FP ReactionSets.Add(rset) failed: {com_detail(exc)}")
        except Exception as exc:  # noqa: BLE001
            report(f"attach attempts failed: {com_detail(exc)}")

        try:
            case.Close(False)
        except Exception:
            pass
        case = None
    except Exception as exc:  # noqa: BLE001
        report(f"create test failed: {com_detail(exc)}")
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
        inspect_template(app)
        create_test(app)
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
