# -*- coding: utf-8 -*-
"""
probe_hysys_reactors.py — HYSYS V15 反应器对象模型探针
================================================================
用途（只读诊断，不参与正式建模）：
  1. 导出 Flowsheet / Operations / ReactionPackageManager 的 COM 成员
  2. 尝试用候选类型名创建 Gibbs / Conversion / Equilibrium 反应器
  3. 导出成功创建的反应器对象成员（反应集、进出口、温度、压力等）
  4. 导出 FluidPackage.ReactionPackage 成员，为 reaction set 编程做准备
  5. 报告写入 output/probe_hysys_reactors_report.txt

运行（远程 HYSYS 工作站，建议先关闭已有 HYSYS）：
    python tools/probe_hysys_reactors.py

安全设计：
  - 只创建临时 Case；不保存、不配置正式模型
  - 测试后 Close(False)；若由本脚本启动 HYSYS，结束时 Quit
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
REPORT_PATH = ROOT / "output" / "probe_hysys_reactors_report.txt"

PROBE_COMPONENTS = ("Methane", "H2O")
FEED_T_K = 298.15
FEED_P_KPA = 100.0
FEED_FLOW_KGMOLEH = 10.0

INVKIND = {1: "FUNC", 2: "GET", 4: "PUT", 8: "PUTREF"}
UNKNOWN_MEMBERS = {
    "QueryInterface", "AddRef", "Release",
    "GetTypeInfoCount", "GetTypeInfo", "GetIDsOfNames", "Invoke",
}

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


def safe_read(obj, attribute: str, label: str) -> None:
    try:
        value = getattr(obj, attribute)
        report(f"  {label}.{attribute} = {value!r}")
    except Exception as exc:  # noqa: BLE001
        report(f"  {label}.{attribute} 读取失败：{com_detail(exc)}")


def dump_type_info(obj, label: str, filters: tuple[str, ...] | None = None) -> None:
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
        report(f"  {name:42s} | {inv:6s} | params={params}")

    for index in range(attr.cVars):
        try:
            var_desc = type_info.GetVarDesc(index)
            names = type_info.GetNames(var_desc.memid)
        except Exception:
            continue
        name = names[0] if names else f"<memid 0x{var_desc.memid & 0xFFFFFFFF:08X}>"
        if filters and not any(token.lower() in name.lower() for token in filters):
            continue
        report(f"  VAR {name:38s}")


def new_case(app, tag: str):
    name = f"ReactorProbe_{tag}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}"
    try:
        case = app.SimulationCases.Add(name)
    except Exception:
        case = app.SimulationCases.Add()
    try:
        case.Visible = True
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


def setup_case(app):
    case = new_case(app, "reactors")
    basis = case.BasisManager
    basis.StartBasisChange()
    fluid_pkg = basis.FluidPackages.Add("Probe_ReactorPkg")
    fluid_pkg.PropertyPackageName = config.PROPERTY_PACKAGE_NAME
    for component_id in PROBE_COMPONENTS:
        fluid_pkg.Components.Add(component_id)
    basis.EndBasisChange()

    feed = case.Flowsheet.MaterialStreams.Add("FEED")
    feed.Temperature.SetValue(FEED_T_K, config.HYSYS_TEMPERATURE_UNIT)
    feed.Pressure.SetValue(FEED_P_KPA, config.HYSYS_PRESSURE_UNIT)
    feed.MolarFlow.SetValue(FEED_FLOW_KGMOLEH, config.HYSYS_MOLAR_FLOW_UNIT)

    report(
        f"  临时 Case 已建立：FP=Probe_ReactorPkg，"
        f"PP={config.PROPERTY_PACKAGE_NAME}，组分={list(PROBE_COMPONENTS)}"
    )
    return case, fluid_pkg, feed


def dump_operation(operation, label: str) -> None:
    report(f"\n--- [{label}] 反应器对象 ---")
    report(f"  对象：{operation!r}")
    for attribute in ("Name", "TypeName", "VisibleTypeName", "TaggedName"):
        safe_read(operation, attribute, label)
    dump_type_info(
        operation, label,
        filters=(
            "reaction", "feed", "product", "stream", "inlet", "outlet",
            "conversion", "gibbs", "equilibrium", "duty", "pressure",
            "temperature", "vapour", "vapor", "component",
        ),
    )


def operations_for_type(flowsheet, type_name: str):
    """Flowsheet.Operations 是带 OperClassOrType 参数的属性。"""
    attempts = (
        (f"Operations({type_name!r})", lambda: flowsheet.Operations(type_name)),
        (
            f"Operations(OperClassOrType={type_name!r})",
            lambda: flowsheet.Operations(OperClassOrType=type_name),
        ),
    )
    for description, call in attempts:
        try:
            collection = call()
            report(f"  {description} 成功 -> {collection!r}")
            return collection
        except Exception as exc:  # noqa: BLE001
            report(f"  {description} 失败：{com_detail(exc)}")
    return None


def try_add_operation(collection, type_name: str, operation_name: str):
    """按类型集合尝试 Add()/Add(name)/Add(name,Type)/Add(Type,name)。"""
    attempts = (
        ("Add()", lambda: collection.Add()),
        (f"Add({operation_name!r})", lambda: collection.Add(operation_name)),
        (
            f"Add({operation_name!r}, {type_name!r})",
            lambda: collection.Add(operation_name, type_name),
        ),
        (
            f"Add({type_name!r}, {operation_name!r})",
            lambda: collection.Add(type_name, operation_name),
        ),
    )
    for description, call in attempts:
        try:
            operation = call()
            report(f"  {description} 成功 -> {operation!r}")
            return operation
        except Exception as exc:  # noqa: BLE001
            report(f"  {description} 失败：{com_detail(exc)}")
    return None


def probe_operations(case) -> None:
    report("\n" + "=" * 78)
    report("第 1 部分：Flowsheet / 反应器对象（参数化 Operations）")
    report("=" * 78)

    flowsheet = case.Flowsheet
    dump_type_info(
        flowsheet, "Flowsheet",
        filters=("operat", "add", "reactor", "unit", "stream"),
    )

    groups = (
        (
            "Gibbs",
            ("GibbsReactor", "Gibbs Reactor", "HYSYS.GibbsReactor", "Gibbs"),
        ),
        (
            "Conversion",
            (
                "ConversionReactor", "Conversion Reactor",
                "HYSYS.ConversionReactor", "Conversion",
            ),
        ),
        (
            "Equilibrium",
            (
                "EquilibriumReactor", "Equilibrium Reactor",
                "HYSYS.EquilibriumReactor", "Equilibrium",
            ),
        ),
    )

    for label, candidates in groups:
        report("\n" + "-" * 78)
        report(f"尝试创建 {label} 反应器")
        created = None
        for index, type_name in enumerate(candidates, start=1):
            collection = operations_for_type(flowsheet, type_name)
            if collection is None:
                continue
            dump_type_info(
                collection, f"Operations({type_name!r})",
                filters=("add", "item", "count", "names", "name"),
            )
            operation = try_add_operation(
                collection, type_name, f"Probe_{label}_{index}"
            )
            if operation is not None:
                created = operation
                break
        if created is not None:
            dump_operation(created, f"{label}Reactor")
        else:
            report(f"  {label} 反应器：所有候选类型名均未创建成功")


def try_add_reaction_set(reaction_sets) -> None:
    """测试 ReactionSets.Add(name, Type) 的 Type 候选。"""
    report("\n  测试创建 Reaction Set")
    type_candidates = (
        "Conversion", "Equilibrium", "Kinetic",
        "ConversionReaction", "EquilibriumReaction",
    )
    for set_type in type_candidates:
        set_name = f"Probe_Set_{set_type}"
        attempts = (
            (
                f"Add({set_name!r}, {set_type!r})",
                lambda: reaction_sets.Add(set_name, set_type),
            ),
            (f"Add({set_name!r})", lambda: reaction_sets.Add(set_name)),
            ("Add()", lambda: reaction_sets.Add()),
        )
        for description, call in attempts:
            try:
                new_set = call()
                report(f"  ReactionSets.{description} 成功 -> {new_set!r}")
                safe_read(new_set, "Name", "ReactionSet")
                safe_read(new_set, "TypeName", "ReactionSet")
                dump_type_info(
                    new_set, "ReactionSet",
                    filters=("reaction", "add", "conversion", "name", "type"),
                )
                return
            except Exception as exc:  # noqa: BLE001
                report(f"  ReactionSets.{description} 失败：{com_detail(exc)}")


def probe_reaction_packages(case, fluid_pkg) -> None:
    report("\n" + "=" * 78)
    report("第 2 部分：ReactionPackageManager / FluidPackage.ReactionPackage")
    report("=" * 78)

    try:
        reaction_package = fluid_pkg.ReactionPackage
    except Exception as exc:  # noqa: BLE001
        report(f"  FluidPackage.ReactionPackage 读取失败：{com_detail(exc)}")
    else:
        report(f"\n  FluidPackage.ReactionPackage = {reaction_package!r}")
        dump_type_info(
            reaction_package, "FluidPackage.ReactionPackage",
            filters=("reaction", "set", "add", "component", "conversion", "name"),
        )
        for attribute in ("ReactionSets", "Name", "Count"):
            safe_read(reaction_package, attribute, "ReactionPackage")
        try:
            package_sets = reaction_package.ReactionSets
        except Exception as exc:  # noqa: BLE001
            report(f"  ReactionPackage.ReactionSets 读取失败：{com_detail(exc)}")
        else:
            dump_type_info(
                package_sets, "ReactionPackage.ReactionSets", filters=None
            )

    try:
        manager = case.BasisManager.ReactionPackageManager
    except Exception as exc:  # noqa: BLE001
        report(f"\n  BasisManager.ReactionPackageManager 读取失败：{com_detail(exc)}")
        return

    report(f"\n  BasisManager.ReactionPackageManager = {manager!r}")
    dump_type_info(
        manager, "ReactionPackageManager",
        filters=("reaction", "set", "package", "add", "item", "count", "name"),
    )
    for attribute in (
        "ReactionPackages", "ReactionSets", "Reactions", "Count", "Names",
    ):
        try:
            value = getattr(manager, attribute)
        except Exception as exc:  # noqa: BLE001
            report(f"  manager.{attribute} 读取失败：{com_detail(exc)}")
            continue
        report(f"  manager.{attribute} = {value!r}")
        if attribute in ("ReactionPackages", "ReactionSets", "Reactions"):
            dump_type_info(value, f"ReactionPackageManager.{attribute}", filters=None)
        if attribute == "ReactionSets":
            try_add_reaction_set(value)


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
    case = None
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

        case, fluid_pkg, _feed = setup_case(app)
        probe_operations(case)
        probe_reaction_packages(case, fluid_pkg)
    except Exception as exc:  # noqa: BLE001
        report(f"探针异常：{com_detail(exc)}")
        report(traceback.format_exc())
    finally:
        gc.collect()
        close_case(case)
        case = None
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
