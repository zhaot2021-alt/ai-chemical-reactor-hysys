# -*- coding: utf-8 -*-
r"""
probe_kinetic_diff.py — GUI 参考 Case vs COM 构建的 kinetic 反应逐字段对比
==========================================================================
在 HYSYS 工作站运行：

    python tools\probe_kinetic_diff.py

1) 打开 output\kinetic_cstr_ref.hsc，补上进料组成并求解，看 GUI 建的
   反应是否真正发生（产物是否有 CO/H2O）；
2) 用 COM 重建同样的 CSTR 并求解；
3) 把两个 reaction 对象的所有可读成员 dump 出来对比。
输出 output\probe_kinetic_diff_report.txt。
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
REPORT_PATH = ROOT / "output" / "probe_kinetic_diff_report.txt"
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


def dump_values(obj: Any, label: str) -> None:
    """把对象所有 GET 成员名+值 dump 出来（只读，不调用任何 FUNC/PUT）。"""
    log(f"\n--- {label} all readable members ---")
    try:
        type_info = obj._oleobj_.GetTypeInfo()
        attr = type_info.GetTypeAttr()
    except Exception as exc:  # noqa: BLE001
        log(f"  type info failed: {err(exc)}")
        return
    ignore = {"QueryInterface", "AddRef", "Release", "GetTypeInfoCount",
              "GetTypeInfo", "GetIDsOfNames", "Invoke"}
    for index in range(attr.cFuncs):
        try:
            func_desc = type_info.GetFuncDesc(index)
            names = type_info.GetNames(func_desc.memid)
            member_name = str(names[0]) if names else "?"
            if member_name in ignore or func_desc.invkind != 2:  # 只读 GET
                continue
            value = getattr(obj, member_name)
            log(f"  {member_name:36s} = {value!r}")
        except Exception as exc:  # noqa: BLE001
            if member_name not in ignore:
                log(f"  {member_name:36s} GET failed: {err(exc)}")


def solve_and_read(case: Any, feed_name: str, product_names: list[str]) -> None:
    fp = case.BasisManager.FluidPackages.Item(0)
    components = base._component_names(fp)
    feed = case.Flowsheet.MaterialStreams.Item(feed_name)
    fractions = {"CO2": 0.5, "Hydrogen": 0.5, "Methane": 0.0, "CO": 0.0, "H2O": 0.0}
    feed.ComponentMolarFractionValue = tuple(
        float(fractions.get(name, 0.0)) for name in components
    )
    base._wait_solver(case)
    for name in product_names:
        try:
            diag.read_stream(
                case.Flowsheet.MaterialStreams.Item(name), components, name
            )
        except Exception as exc:  # noqa: BLE001
            log(f"    read {name} failed: {err(exc)}")


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

        log("=" * 68)
        log(f"PART 1: GUI reference case {REF_CSTR.name}")
        log("=" * 68)
        if not REF_CSTR.exists():
            log(f"  MISSING: {REF_CSTR}")
        else:
            ref_case = app.SimulationCases.Open(str(REF_CSTR))
            try:
                fp = ref_case.BasisManager.FluidPackages.Item(0)
                log(f"  FP: {fp.Name} components={base._component_names(fp)}")
                rset = ref_case.BasisManager.ReactionPackageManager.ReactionSets.Item(0)
                rxn = rset.ActiveReactions.Item(0)
                dump_values(rxn, "GUI reaction Rxn-1")
                log("\n--- solve ref case (set Feed1 composition) ---")
                solve_and_read(ref_case, "Feed1", ["Vap1", "Liq1"])
            finally:
                try:
                    ref_case.Close(False)
                except Exception:
                    pass

        log("=" * 68)
        log("PART 2: COM-built CSTR (same parameters)")
        log("=" * 68)
        case, fp, feed, components = base._setup_case(app)
        manager = case.BasisManager.ReactionPackageManager
        rset = manager.ReactionSets.Add("KinSet", "Kinetic")
        rxn = diag.make_reaction(manager, rset, fp, "Rxn-COM", 1000.0, components)
        dump_values(rxn, "COM reaction Rxn-COM")
        diag.build_cstr(case, feed, rset, components, "D", 50.0)

        log("\n=== summary ===")
        log("  GUI ref case 产物若有 CO/H2O -> GUI 构建参数正确，对比两 dump 找差异；")
        log("  若 GUI 也不转化 -> 参考 Case 本身缺关键设置。")
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
