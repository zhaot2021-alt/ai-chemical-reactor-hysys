# -*- coding: utf-8 -*-
"""
probe_hysys_property_package.py — HYSYS V15 物性包 API 诊断探针
================================================================
用途（仅诊断，不参与正式建模）：
  1. 导出 FluidPackage / BasisManager / FluidPackages 的 COM 成员表（含 GET/PUT 权限）
  2. 用临时 Case 测试几种 PropertyPackageName 设置顺序与内部名
  3. 把完整报告写到 output/probe_hysys_property_package_report.txt

运行（远程 HYSYS 工作站）：
    python tools/probe_hysys_property_package.py

安全设计：
  - 优先附着运行中的 HYSYS；若由本脚本启动，结束时 Quit
  - 每个测试用一个临时 Case，测试完立即 Close(False)，不保存
  - 若 HYSYS 中已有重要未保存工作，建议先保存并关闭 HYSYS 再运行
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

PROGID = "HYSYS.Application"
INVKIND = {1: "FUNC", 2: "GET", 4: "PUT", 8: "PUTREF"}
UNKNOWN_MEMBERS = {
    "QueryInterface", "AddRef", "Release",
    "GetTypeInfoCount", "GetTypeInfo", "GetIDsOfNames", "Invoke",
}

ROOT = Path(__file__).resolve().parents[1]
REPORT_PATH = ROOT / "output" / "probe_hysys_property_package_report.txt"

_REPORT_LINES: list[str] = []


def report(message: str = "") -> None:
    print(message, flush=True)
    _REPORT_LINES.append(str(message))


def com_detail(exc: BaseException) -> str:
    if isinstance(exc, pywintypes.com_error):
        hresult = getattr(exc, "hresult", None)
        excepinfo = getattr(exc, "excepinfo", None)
        hr_text = f"0x{hresult & 0xFFFFFFFF:08X}" if isinstance(hresult, int) else "N/A"
        return f"HRESULT={hr_text}, excepinfo={excepinfo!r}"
    return f"{type(exc).__name__}: {exc}"


def dump_type_info(obj, label: str, filters: tuple[str, ...] | None = None) -> None:
    """枚举 IDispatch 类型信息中的函数/属性；invkind 能区分 GET 与 PUT。"""
    report(f"\n--- [{label}] COM 类型信息 ---")
    try:
        type_info = obj._oleobj_.GetTypeInfo()
        attr = type_info.GetTypeAttr()
    except Exception as exc:  # noqa: BLE001
        report(f"  GetTypeInfo 失败：{com_detail(exc)}")
        return

    report(f"  cFuncs={attr.cFuncs}, cVars={attr.cVars}, iid={attr.iid}")
    for index in range(attr.cFuncs):
        try:
            func_desc = type_info.GetFuncDesc(index)
            names = type_info.GetNames(func_desc.memid)
        except Exception as exc:  # noqa: BLE001
            report(f"  <FuncDesc {index} 读取失败：{com_detail(exc)}>")
            continue

        name = names[0] if names else f"<memid 0x{func_desc.memid & 0xFFFFFFFF:08X}>"
        if name in UNKNOWN_MEMBERS:
            continue
        if filters and not any(token.lower() in name.lower() for token in filters):
            continue

        params = list(names[1:])
        inv = INVKIND.get(func_desc.invkind, str(func_desc.invkind))
        try:
            doc = type_info.GetDocumentation(func_desc.memid)
            doc_text = doc[0] if doc and doc[0] else ""
        except Exception:  # noqa: BLE001
            doc_text = ""
        report(
            f"  {name:38s} | {inv:6s} | params={params}"
            + (f" | {doc_text}" if doc_text else "")
        )

    for index in range(attr.cVars):
        try:
            var_desc = type_info.GetVarDesc(index)
            names = type_info.GetNames(var_desc.memid)
        except Exception as exc:  # noqa: BLE001
            report(f"  <VarDesc {index} 读取失败：{com_detail(exc)}>")
            continue
        name = names[0] if names else f"<memid 0x{var_desc.memid & 0xFFFFFFFF:08X}>"
        if filters and not any(token.lower() in name.lower() for token in filters):
            continue
        report(f"  VAR {name:34s}")


def safe_read(obj, attribute: str, label: str) -> None:
    try:
        value = getattr(obj, attribute)
        report(f"  {label}.{attribute} = {value!r}")
    except Exception as exc:  # noqa: BLE001
        report(f"  {label}.{attribute} 读取失败：{com_detail(exc)}")


def new_case(app, tag: str):
    name = f"APIProbe_{tag}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}"
    try:
        case = app.SimulationCases.Add(name)
    except Exception:
        case = app.SimulationCases.Add()
    try:
        case.Visible = False
    except Exception:
        pass
    return case


def close_case(case) -> None:
    if case is None:
        return
    try:
        try:
            case.Close(False)
        except Exception:
            case.Close()
    except Exception as exc:  # noqa: BLE001
        report(f"  关闭临时 Case 失败：{com_detail(exc)}")
    finally:
        try:
            del case
        except Exception:
            pass


def end_basis(basis) -> bool:
    try:
        basis.EndBasisChange()
        report("  EndBasisChange OK")
        return True
    except Exception as exc:  # noqa: BLE001
        report(f"  EndBasisChange 失败：{com_detail(exc)}")
        return False


def introspect(app) -> None:
    """在一个临时 Case 中导出 BasisManager / FluidPackages / FluidPackage 成员表。"""
    report("\n" + "=" * 78)
    report("第 1 部分：COM 成员表")
    report("=" * 78)
    case = None
    basis = None
    started = False
    try:
        case = new_case(app, "introspect")
        basis = case.BasisManager
        dump_type_info(basis, "BasisManager", filters=("prop", "pack", "basis", "fluid"))

        try:
            fluid_packages = basis.FluidPackages
            dump_type_info(
                fluid_packages, "FluidPackages collection",
                filters=("add", "item", "count", "prop", "pack", "name"),
            )
        except Exception as exc:  # noqa: BLE001
            report(f"  BasisManager.FluidPackages 读取失败：{com_detail(exc)}")

        try:
            basis.StartBasisChange()
            started = True
        except Exception as exc:  # noqa: BLE001
            report(f"  StartBasisChange 失败：{com_detail(exc)}")

        fluid_pkg = basis.FluidPackages.Add()
        report(f"  已创建临时 FluidPackage：{fluid_pkg!r}")
        dump_type_info(fluid_pkg, "FluidPackage", filters=None)

        safe_read(fluid_pkg, "Name", "FluidPackage")
        safe_read(fluid_pkg, "PropertyPackageName", "FluidPackage")
        safe_read(fluid_pkg, "PropPackage", "FluidPackage")
        try:
            prop_pkg = fluid_pkg.PropPackage
            safe_read(prop_pkg, "TypeName", "PropPackage")
            dump_type_info(prop_pkg, "PropPackage", filters=None)
        except Exception as exc:  # noqa: BLE001
            report(f"  FluidPackage.PropPackage 读取失败：{com_detail(exc)}")
    except Exception as exc:  # noqa: BLE001
        report(f"  成员表探针异常：{com_detail(exc)}")
        report(traceback.format_exc())
    finally:
        if started and basis is not None:
            end_basis(basis)
        close_case(case)
        gc.collect()


def run_variant(
    app,
    tag: str,
    pp_name: str,
    *,
    rename_before: bool = False,
    rename_after_pp: bool = False,
    set_after_end: bool = False,
    solver_hold: bool = False,
) -> None:
    """
    测试一种设置顺序：
      - 默认：Add -> 设 PP -> End
      - rename_after_pp：Add -> 设 PP -> 改名 -> End（当前驱动的修复顺序）
      - rename_before：Add -> 改名 -> 设 PP -> End（复现当前失败）
      - set_after_end：Add -> End -> 新 Start -> 设 PP -> End
      - solver_hold：在 StartBasisChange 之前 Solver.CanSolve=False
    """
    report("\n" + "-" * 78)
    report(
        f"测试 [{tag}] pp={pp_name!r}, rename_before={rename_before}, "
        f"rename_after_pp={rename_after_pp}, set_after_end={set_after_end}, "
        f"solver_hold={solver_hold}"
    )
    case = None
    basis = None
    started = False
    solver_was_held = False
    try:
        case = new_case(app, tag)
        basis = case.BasisManager

        if solver_hold:
            try:
                old_can_solve = case.Solver.CanSolve
                report(f"  Solver.CanSolve 初始 = {old_can_solve!r}")
                case.Solver.CanSolve = False
                solver_was_held = True
                report("  Solver.CanSolve = False OK")
            except Exception as exc:  # noqa: BLE001
                report(f"  设置 Solver.CanSolve=False 失败：{com_detail(exc)}")

        try:
            basis.StartBasisChange()
            started = True
        except Exception as exc:  # noqa: BLE001
            report(f"  StartBasisChange 失败：{com_detail(exc)}")

        fluid_pkg = basis.FluidPackages.Add()
        report(f"  FluidPackages.Add OK: {fluid_pkg!r}")

        def rename_fp(when: str) -> None:
            try:
                fluid_pkg.Name = f"Probe_{tag}"
                report(f"  改名（{when}）OK")
            except Exception as exc:  # noqa: BLE001
                report(f"  改名（{when}）失败：{com_detail(exc)}")

        def set_property_package() -> bool:
            ok = False
            try:
                fluid_pkg.PropertyPackageName = pp_name
                report(f"  PropertyPackageName = {pp_name!r} OK")
                ok = True
            except Exception as exc:  # noqa: BLE001
                report(f"  PropertyPackageName = {pp_name!r} 失败：{com_detail(exc)}")
            safe_read(fluid_pkg, "PropertyPackageName", "FluidPackage")
            return ok

        if rename_before:
            rename_fp("物性包设置之前")

        if not set_after_end:
            pp_ok = set_property_package()
            if pp_ok and rename_after_pp:
                rename_fp("物性包设置之后")

        if started:
            end_basis(basis)
            started = False

        if set_after_end:
            try:
                basis.StartBasisChange()
                started = True
            except Exception as exc:  # noqa: BLE001
                report(f"  第二次 StartBasisChange 失败：{com_detail(exc)}")
            set_property_package()
            if started:
                end_basis(basis)
                started = False

        safe_read(fluid_pkg, "PropertyPackageName", "FluidPackage")
        try:
            report(f"  PropPackage.TypeName = {fluid_pkg.PropPackage.TypeName!r}")
        except Exception as exc:  # noqa: BLE001
            report(f"  PropPackage.TypeName 读取失败：{com_detail(exc)}")
        try:
            report(f"  FluidPackage.Name = {fluid_pkg.Name!r}")
        except Exception as exc:  # noqa: BLE001
            report(f"  FluidPackage.Name 读取失败：{com_detail(exc)}")
    except Exception as exc:  # noqa: BLE001
        report(f"  变体异常：{com_detail(exc)}")
        report(traceback.format_exc())
    finally:
        if started and basis is not None:
            end_basis(basis)
        if solver_was_held and case is not None:
            try:
                case.Solver.CanSolve = True
                report("  Solver.CanSolve = True 已恢复")
            except Exception as exc:  # noqa: BLE001
                report(f"  恢复 Solver.CanSolve=True 失败：{com_detail(exc)}")
        close_case(case)
        gc.collect()


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    report(f"时间：{datetime.now().isoformat(timespec='seconds')}")
    report(f"HYSYS ProgID：{PROGID}")
    report(f"Python：{sys.version}")
    report(f"报告文件：{REPORT_PATH}")

    app = None
    owns = False
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        try:
            app = win32com.client.GetActiveObject(PROGID)
            report("已附着运行中的 HYSYS 实例（不会 Quit 它）")
        except Exception:
            for attempt in range(1, 4):
                try:
                    app = win32com.client.Dispatch(PROGID)
                    owns = True
                    report("已启动新的 HYSYS 实例（结束时 Quit）")
                    break
                except Exception as exc:  # noqa: BLE001
                    report(f"启动 HYSYS 失败（{attempt}/3）：{com_detail(exc)}")
                    if attempt < 3:
                        time.sleep(3)
        if app is None:
            raise RuntimeError("无法连接 HYSYS")

        try:
            app.Visible = True
        except Exception:
            pass

        introspect(app)
        run_variant(
            app, "A_pp_then_rename", "PengRobinson", rename_after_pp=True
        )
        run_variant(app, "B_rename_before", "PengRobinson", rename_before=True)
        run_variant(app, "C_after_end", "PengRobinson", set_after_end=True)
        run_variant(app, "D_pr", "PR")
        run_variant(app, "E_hysyspr", "HysysPR")
        run_variant(app, "F_srk", "SRK")
        run_variant(
            app, "G_solver_hold", "PengRobinson",
            rename_after_pp=True, solver_hold=True,
        )
    except Exception as exc:  # noqa: BLE001
        report(f"探针异常：{com_detail(exc)}")
        report(traceback.format_exc())
    finally:
        gc.collect()
        if app is not None and owns:
            try:
                app.Quit()
                report("已 Quit 本次启动的 HYSYS 实例")
                time.sleep(2)
            except Exception as exc:  # noqa: BLE001
                report(f"Quit 失败：{com_detail(exc)}")
        app = None
        gc.collect()
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        try:
            REPORT_PATH.write_text("\n".join(_REPORT_LINES), encoding="utf-8")
            print(f"\n完整报告已写入：{REPORT_PATH}", flush=True)
        except Exception as exc:  # noqa: BLE001
            print(f"报告写入失败：{exc}", flush=True)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
