# -*- coding: utf-8 -*-
"""
probe_hysys_stream_composition.py — HYSYS V15 物流组成/求解 COM 探针
================================================================
用途（只读诊断，不参与正式建模）：
  1. 创建一个含 10 个标准组分、Peng-Robinson 的临时 Case
  2. 导出 ProcessStream 与 Solver 的相关 COM 成员（含 GET/PUT）
  3. 测试多种物流组成写入方式并读回验证：
       - ComponentMolarFraction.Calculate(array)
       - ComponentMolarFraction.Values = array
       - MolarFractionsValue = array + TPFlash()
       - ComponentMolarFractionValue = array（预期不可直接写）
  4. 测试 Solver.CanSolve / IsSolving / Run 和求解后的流股状态
  5. 报告写入 output/probe_hysys_stream_composition_report.txt

运行（远程 HYSYS 工作站，建议先关闭已有 HYSYS）：
    python tools/probe_hysys_stream_composition.py

安全设计：
  - 优先附着已有 HYSYS；若由本脚本启动，结束时 Quit
  - 只创建临时 Case，不保存；测试后 Close(False)
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

import config  # noqa: E402  (需要先加入 src 路径)

PROGID = "HYSYS.Application"
REPORT_PATH = ROOT / "output" / "probe_hysys_stream_composition_report.txt"

FEED_T_K = 298.15
FEED_P_KPA = 100.0
FEED_FLOW_KGMOLEH = 100.0

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
    name = f"StreamProbe_{tag}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}"
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


def item_list(collection, label: str) -> list:
    """读取 HYSYS 集合，兼容 0-based / 1-based 索引。"""
    try:
        count = int(collection.Count)
    except Exception as exc:  # noqa: BLE001
        report(f"  {label}.Count 读取失败：{com_detail(exc)}")
        return []

    for start in (0, 1):
        items = []
        for index in range(start, start + count):
            try:
                items.append(collection.Item(index))
            except Exception:
                items = []
                break
        if items:
            return items
    return []


def create_basis(app):
    """创建含 10 标准组分 + PR 的临时 Case；返回 (case, fluid_pkg)。"""
    case = new_case(app, "basis")
    basis = case.BasisManager
    basis.StartBasisChange()
    fluid_pkg = basis.FluidPackages.Add("Probe_StreamPkg")
    fluid_pkg.PropertyPackageName = config.PROPERTY_PACKAGE_NAME
    for component_id in config.STANDARD_COMPONENT_IDS:
        fluid_pkg.Components.Add(component_id)
    basis.EndBasisChange()
    report(
        f"  已创建临时 Case/Basis：FP=Probe_StreamPkg，"
        f"PP={config.PROPERTY_PACKAGE_NAME}，"
        f"组分数={len(config.STANDARD_COMPONENT_IDS)}"
    )
    return case, fluid_pkg


def component_names(fluid_pkg) -> list[str]:
    components = item_list(fluid_pkg.Components, "FluidPackage.Components")
    names = []
    for component in components:
        try:
            names.append(str(component.Name))
        except Exception:
            names.append(repr(component))
    return names


def set_tpf(stream) -> None:
    stream.Temperature.SetValue(FEED_T_K, config.HYSYS_TEMPERATURE_UNIT)
    stream.Pressure.SetValue(FEED_P_KPA, config.HYSYS_PRESSURE_UNIT)
    stream.MolarFlow.SetValue(
        FEED_FLOW_KGMOLEH, config.HYSYS_MOLAR_FLOW_UNIT
    )


def read_composition(stream, label: str) -> list | None:
    try:
        value = stream.ComponentMolarFractionValue
    except Exception as exc:  # noqa: BLE001
        report(f"  {label}.ComponentMolarFractionValue 读取失败：{com_detail(exc)}")
        return None

    try:
        array = list(value) if value is not None else None
    except Exception as exc:  # noqa: BLE001
        report(f"  {label} 组成数组转换失败：{com_detail(exc)}；raw={value!r}")
        return None

    if array is None:
        report(f"  {label}.ComponentMolarFractionValue = None")
        return None

    report(
        f"  {label}.ComponentMolarFractionValue = {array} "
        f"(len={len(array)}, sum={sum(v for v in array if isinstance(v, (int, float))):.6f})"
    )
    return array


def read_stream_state(stream, label: str) -> None:
    for attribute in (
        "State", "VapourFractionValue", "TemperatureValue",
        "PressureValue", "MolarFlowValue", "ComponentMolarFractionValue",
    ):
        safe_read(stream, attribute, label)
    for attribute, unit in (
        ("Temperature", config.HYSYS_TEMPERATURE_UNIT),
        ("Pressure", config.HYSYS_PRESSURE_UNIT),
        ("MolarFlow", config.HYSYS_MOLAR_FLOW_UNIT),
    ):
        try:
            value = getattr(stream, attribute).GetValue(unit)
            report(f"  {label}.{attribute}.GetValue({unit!r}) = {value!r}")
        except Exception as exc:  # noqa: BLE001
            report(
                f"  {label}.{attribute}.GetValue({unit!r}) 失败：{com_detail(exc)}"
            )


def dump_api(case, stream, fluid_pkg) -> list[str]:
    report("\n" + "=" * 78)
    report("第 1 部分：ProcessStream / Solver / ComponentMolarFraction 成员表")
    report("=" * 78)

    names = component_names(fluid_pkg)
    report(f"  流体包组分顺序（{len(names)} 个）：{names}")

    dump_type_info(
        stream, "ProcessStream",
        filters=(
            "fraction", "composition", "molar", "flash", "tp",
            "state", "flow", "vapour", "vapor",
        ),
    )
    read_composition(stream, "FEED")
    read_stream_state(stream, "FEED")

    try:
        comp_obj = stream.ComponentMolarFraction
    except Exception as exc:  # noqa: BLE001
        report(f"\n  ProcessStream.ComponentMolarFraction 读取失败：{com_detail(exc)}")
    else:
        report(f"\n  ProcessStream.ComponentMolarFraction = {comp_obj!r}")
        dump_type_info(comp_obj, "ComponentMolarFraction", filters=None)
        for attribute in ("Count", "Values", "Names"):
            safe_read(comp_obj, attribute, "ComponentMolarFraction")

    try:
        solver = case.Solver
    except Exception as exc:  # noqa: BLE001
        report(f"\n  SimulationCase.Solver 读取失败：{com_detail(exc)}")
    else:
        dump_type_info(
            solver, "Solver",
            filters=("can", "solve", "solving", "run", "stop", "state", "integrator"),
        )
        safe_read(solver, "CanSolve", "Solver")
        safe_read(solver, "IsSolving", "Solver")

    dump_type_info(
        case, "SimulationCase",
        filters=("active", "visible", "solver", "state"),
    )
    return names


def try_composition_method(
    case, names: list[str], tag: str, method: str
) -> tuple[object | None, bool]:
    report("\n" + "-" * 78)
    report(f"组成写入测试 [{tag}] method={method}")
    stream = None
    try:
        stream = case.Flowsheet.MaterialStreams.Add(f"COMP_{tag}")
        set_tpf(stream)
        report(f"  已创建流股 COMP_{tag}，T/P/F 已设置")
    except Exception as exc:  # noqa: BLE001
        report(f"  创建流股失败：{com_detail(exc)}")
        return None, False

    if "Methane" not in names:
        report("  流体包中不存在 Methane，无法构造纯组分数组")
        return stream, False
    composition = [0.0] * len(names)
    composition[names.index("Methane")] = 1.0
    report(
        f"  目标组成（纯 Methane，按流体包顺序）：{composition}"
    )

    success = False
    try:
        if method == "calculate":
            stream.ComponentMolarFraction.Calculate(tuple(composition))
            report("  ComponentMolarFraction.Calculate(tuple) 调用成功")
            success = True
        elif method == "values":
            stream.ComponentMolarFraction.Values = tuple(composition)
            report("  ComponentMolarFraction.Values = tuple 赋值成功")
            success = True
        elif method == "molarfractions":
            stream.MolarFractionsValue = tuple(composition)
            report("  MolarFractionsValue = tuple 赋值成功")
            tp_flash = getattr(stream, "TPFlash", None)
            if callable(tp_flash):
                tp_flash()
                report("  TPFlash() 调用成功")
            else:
                report("  ProcessStream 上没有可调用的 TPFlash()")
            success = True
        elif method == "direct":
            stream.ComponentMolarFractionValue = tuple(composition)
            report("  ComponentMolarFractionValue = tuple 赋值成功（意外）")
            success = True
        else:
            report(f"  未知 method={method}")
    except Exception as exc:  # noqa: BLE001
        report(f"  {method} 失败：{com_detail(exc)}")
        # 对 Calculate / Values 再试一次 list 而不是 tuple
        if method in ("calculate", "values"):
            try:
                if method == "calculate":
                    stream.ComponentMolarFraction.Calculate(list(composition))
                else:
                    stream.ComponentMolarFraction.Values = list(composition)
                report(f"  {method} 使用 list 重试成功")
                success = True
            except Exception as exc2:  # noqa: BLE001
                report(f"  {method} 使用 list 重试失败：{com_detail(exc2)}")

    read_composition(stream, f"COMP_{tag}")
    read_stream_state(stream, f"COMP_{tag}")
    return stream, success


def solver_test(case, stream, tag: str) -> None:
    report("\n" + "=" * 78)
    report(f"第 3 部分：求解测试（流股 COMP_{tag}）")
    report("=" * 78)
    if stream is None:
        report("  没有可用的流股，跳过求解测试")
        return

    try:
        solver = case.Solver
    except Exception as exc:  # noqa: BLE001
        report(f"  SimulationCase.Solver 不可用：{com_detail(exc)}")
        return

    for attribute in ("CanSolve", "IsSolving"):
        safe_read(solver, attribute, "Solver")

    try:
        case.Solver.CanSolve = False
        report("  Solver.CanSolve = False 成功")
    except Exception as exc:  # noqa: BLE001
        report(f"  Solver.CanSolve = False 失败：{com_detail(exc)}")

    try:
        case.Solver.CanSolve = True
        report("  Solver.CanSolve = True 成功")
    except Exception as exc:  # noqa: BLE001
        report(f"  Solver.CanSolve = True 失败：{com_detail(exc)}")

    try:
        run = getattr(solver, "Run", None)
        if callable(run):
            run()
            report("  Solver.Run() 调用成功")
        else:
            report("  Solver 上没有可调用的 Run()")
    except Exception as exc:  # noqa: BLE001
        report(f"  Solver.Run() 失败：{com_detail(exc)}")

    deadline = time.time() + 15
    while time.time() < deadline:
        try:
            is_solving = bool(case.Solver.IsSolving)
        except Exception as exc:  # noqa: BLE001
            report(f"  读取 Solver.IsSolving 失败：{com_detail(exc)}")
            break
        if not is_solving:
            report("  Solver.IsSolving = False（已结束）")
            break
        time.sleep(0.2)
    else:
        report("  等待 15s 后 Solver.IsSolving 仍为 True")

    read_stream_state(stream, f"COMP_{tag}")


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

        case, fluid_pkg = create_basis(app)
        feed = case.Flowsheet.MaterialStreams.Add("FEED")
        set_tpf(feed)
        report("  已创建 FEED（T/P/F 已设置，尚未设组成）")

        names = dump_api(case, feed, fluid_pkg)

        report("\n" + "=" * 78)
        report("第 2 部分：物流组成写入方式测试")
        report("=" * 78)
        successful_stream = None
        successful_tag = None
        for method in ("calculate", "values", "molarfractions", "direct"):
            stream, ok = try_composition_method(case, names, method, method)
            if ok and successful_stream is None:
                successful_stream = stream
                successful_tag = method

        if successful_stream is not None:
            solver_test(case, successful_stream, successful_tag or "unknown")
        else:
            report("\n  所有组成写入方式均未成功，跳过求解测试")
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
