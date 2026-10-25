# -*- coding: utf-8 -*-
"""
probe_hysys_property_package.py — HYSYS V15 物性包 API 诊断探针（v2）
================================================================
用途（仅诊断，不参与正式建模）：
  1. 导出 BasisManager / FluidPackages / FluidPackage 的 COM 成员表
  2. 验证 FluidPackages.Add(name) / Add(name, Type) 的真实签名效果
  3. 在**含至少一个组分**的临时 Case 中测试：
       - 先加组分再设 PropertyPackageName
       - 先设 PropertyPackageName 再加组分
       - 多个 Peng-Robinson 内部名 / Add 的 Type 参数
  4. 把完整报告写到 output/probe_hysys_property_package_report.txt

运行（远程 HYSYS 工作站）：
    python tools/probe_hysys_property_package.py

安全设计：
  - 优先附着运行中的 HYSYS；若由本脚本启动，结束时 Quit
  - 每个测试用一个临时 Case，测试完立即 Close(False)，不保存
  - 若 HYSYS 中已有重要未保存工作，建议先保存
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
    """枚举 IDispatch 类型信息；invkind 能区分 GET 与 PUT。"""
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


def add_fluid_package(basis, name: str, type_name: str | None = None):
    """优先使用文档示例的 Add(name)；可选测试 Add(name, Type)。"""
    if type_name is not None:
        try:
            fp = basis.FluidPackages.Add(name, type_name)
            report(f"  FluidPackages.Add({name!r}, {type_name!r}) OK")
            return fp
        except Exception as exc:  # noqa: BLE001
            report(
                f"  FluidPackages.Add({name!r}, {type_name!r}) 失败："
                f"{com_detail(exc)}；回退 Add(name)"
            )
    try:
        fp = basis.FluidPackages.Add(name)
        report(f"  FluidPackages.Add({name!r}) OK")
        return fp
    except Exception as exc:  # noqa: BLE001
        report(f"  FluidPackages.Add({name!r}) 失败：{com_detail(exc)}；回退 Add()")
        return basis.FluidPackages.Add()


def add_component(fluid_pkg, component_id: str = "Methane") -> bool:
    try:
        component = fluid_pkg.Components.Add(component_id)
        report(f"  Components.Add({component_id!r}) OK -> {component!r}")
        component = None
        return True
    except Exception as exc:  # noqa: BLE001
        report(f"  Components.Add({component_id!r}) 失败：{com_detail(exc)}")
        return False


def read_fluid_package(fluid_pkg, label: str = "FluidPackage") -> None:
    safe_read(fluid_pkg, "Name", label)
    safe_read(fluid_pkg, "PropertyPackageName", label)
    try:
        prop_pkg = fluid_pkg.PropertyPackage
    except Exception as exc:  # noqa: BLE001
        report(f"  {label}.PropertyPackage 读取失败：{com_detail(exc)}")
    else:
        report(f"  {label}.PropertyPackage = {prop_pkg!r}")
        safe_read(prop_pkg, "Name", "PropertyPackage")
        safe_read(prop_pkg, "TypeName", "PropertyPackage")
        safe_read(prop_pkg, "VisibleTypeName", "PropertyPackage")

    try:
        count = int(fluid_pkg.Components.Count)
        names = []
        for index in range(count):
            names.append(str(fluid_pkg.Components.Item(index).Name))
        report(f"  {label}.Components = {names}")
    except Exception as exc:  # noqa: BLE001
        report(f"  {label}.Components 读取失败：{com_detail(exc)}")


def introspect(app) -> None:
    """在一个临时 Case 中导出成员表，并测试 Add(name) + 组分 + End。"""
    report("\n" + "=" * 78)
    report("第 1 部分：COM 成员表 + Add(name)/组分/End 基础体验")
    report("=" * 78)
    case = None
    basis = None
    started = False
    try:
        case = new_case(app, "introspect")
        basis = case.BasisManager
        dump_type_info(
            basis, "BasisManager", filters=("prop", "pack", "basis", "fluid")
        )
        dump_type_info(
            basis.FluidPackages, "FluidPackages collection",
            filters=("add", "item", "count", "prop", "pack", "name"),
        )

        try:
            basis.StartBasisChange()
            started = True
        except Exception as exc:  # noqa: BLE001
            report(f"  StartBasisChange 失败：{com_detail(exc)}")

        fluid_pkg = add_fluid_package(basis, "Probe_Introspect")
        dump_type_info(fluid_pkg, "FluidPackage", filters=None)

        read_fluid_package(fluid_pkg, "FluidPackage")
        add_component(fluid_pkg, "Methane")
        read_fluid_package(fluid_pkg, "FluidPackage")

        if started:
            end_basis(basis)
            started = False
    except Exception as exc:  # noqa: BLE001
        report(f"  成员表/基础体验异常：{com_detail(exc)}")
        report(traceback.format_exc())
    finally:
        if started and basis is not None:
            end_basis(basis)
        close_case(case)
        gc.collect()


def run_component_test(
    app,
    tag: str,
    pp_name: str,
    *,
    order: str,
    add_with_type: bool = False,
) -> None:
    """
    在含 Methane 的临时 Case 中测试：
      order="comp_first"：Add(name) -> Components.Add -> PropertyPackageName -> End
      order="pp_first"  ：Add(name) -> PropertyPackageName -> Components.Add -> End
      add_with_type=True ：Add(name, pp_name) -> Components.Add -> End
    """
    report("\n" + "-" * 78)
    report(
        f"测试 [{tag}] pp={pp_name!r}, order={order}, "
        f"add_with_type={add_with_type}"
    )
    case = None
    basis = None
    started = False
    try:
        case = new_case(app, tag)
        basis = case.BasisManager
        try:
            basis.StartBasisChange()
            started = True
        except Exception as exc:  # noqa: BLE001
            report(f"  StartBasisChange 失败：{com_detail(exc)}")

        fluid_pkg = add_fluid_package(
            basis, f"Probe_{tag}", pp_name if add_with_type else None
        )

        def set_property_package() -> bool:
            try:
                fluid_pkg.PropertyPackageName = pp_name
                report(f"  PropertyPackageName = {pp_name!r} OK")
                return True
            except Exception as exc:  # noqa: BLE001
                report(
                    f"  PropertyPackageName = {pp_name!r} 失败："
                    f"{com_detail(exc)}"
                )
                return False

        if add_with_type:
            # Add(name, Type) 理论上可能已设置 PP；先读一次
            safe_read(fluid_pkg, "PropertyPackageName", "FluidPackage")
            add_component(fluid_pkg)
        elif order == "comp_first":
            add_component(fluid_pkg)
            set_property_package()
        else:
            set_property_package()
            add_component(fluid_pkg)

        if started:
            end_basis(basis)
            started = False

        read_fluid_package(fluid_pkg, "FluidPackage")
    except Exception as exc:  # noqa: BLE001
        report(f"  变体异常：{com_detail(exc)}")
        report(traceback.format_exc())
    finally:
        if started and basis is not None:
            end_basis(basis)
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

        report("\n" + "=" * 78)
        report("第 2 部分：带组分的设置顺序 / 内部名测试")
        report("=" * 78)
        run_component_test(app, "A_srk_comp_first", "SRK", order="comp_first")
        run_component_test(app, "B_srk_pp_first", "SRK", order="pp_first")
        run_component_test(
            app, "C_pr_comp_first", "PengRobinson", order="comp_first"
        )
        run_component_test(
            app, "D_pr_pp_first", "PengRobinson", order="pp_first"
        )
        run_component_test(
            app, "E_pr_dash_comp_first", "Peng-Robinson", order="comp_first"
        )
        run_component_test(
            app, "F_pr_add_type", "PengRobinson",
            order="comp_first", add_with_type=True,
        )
        run_component_test(
            app, "G_srk_add_type", "SRK",
            order="comp_first", add_with_type=True,
        )
        run_component_test(
            app, "H_pengrob_comp_first", "pengrob", order="comp_first"
        )
        run_component_test(
            app, "I_pengrob_pp_first", "pengrob", order="pp_first"
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
