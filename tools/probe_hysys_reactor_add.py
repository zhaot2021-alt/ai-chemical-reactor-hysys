# -*- coding: utf-8 -*-
"""
probe_hysys_reactor_add.py — 快速定位 HYSYS V15 反应器 Add 类型字符串
================================================================
只做一件事：
  Flowsheet.Operations.Add("名称", "类型字符串")
  用多种大小写/命名候选，找出 V15 接受的 reactor 类型串。

不配置反应、不连流股、不保存；测试完关闭临时 Case。
运行：
    python tools/probe_hysys_reactor_add.py
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
REPORT_PATH = ROOT / "output" / "probe_hysys_reactor_add_report.txt"

PROBE_COMPONENTS = ("Methane", "H2O")

REACTOR_TYPE_CANDIDATES = (
    # Gibbs —— CamelCase / 小写 / 短名 / 接口名 / 带空格
    "GibbsReactor", "gibbsreactor", "Gibbs", "gibbs",
    "GibbsOp", "gibbsop", "Gibbs Reactor",
    # Conversion
    "ConversionReactor", "conversionreactor", "Conversion", "conversion",
    "ConversionOp", "conversionop", "Conversion Reactor",
    # Equilibrium
    "EquilibriumReactor", "equilibriumreactor", "Equilibrium", "equilibrium",
    "EquilibriumOp", "equilibriumop", "Equilibrium Reactor",
    # 通用 reactor
    "Reactor", "reactor", "ReactorOp", "reactorop",
)

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
            fd = type_info.GetFuncDesc(index)
            names = type_info.GetNames(fd.memid)
        except Exception:
            continue
        name = names[0] if names else "?"
        if filters and not any(token.lower() in name.lower() for token in filters):
            continue
        report(f"  {name:42s} | params={list(names[1:])}")
    for index in range(attr.cVars):
        try:
            vd = type_info.GetVarDesc(index)
            names = type_info.GetNames(vd.memid)
        except Exception:
            continue
        name = names[0] if names else "?"
        if filters and not any(token.lower() in name.lower() for token in filters):
            continue
        report(f"  VAR {name:38s}")


def setup_case(app):
    name = f"ReactorAddProbe_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}"
    try:
        case = app.SimulationCases.Add(name)
    except Exception:
        case = app.SimulationCases.Add()
    try:
        case.Visible = True
    except Exception:
        pass

    basis = case.BasisManager
    basis.StartBasisChange()
    fluid_pkg = basis.FluidPackages.Add("Probe_ReactorPkg")
    fluid_pkg.PropertyPackageName = config.PROPERTY_PACKAGE_NAME
    for component_id in PROBE_COMPONENTS:
        fluid_pkg.Components.Add(component_id)
    basis.EndBasisChange()

    feed = case.Flowsheet.MaterialStreams.Add("FEED")
    feed.Temperature.SetValue(298.15, config.HYSYS_TEMPERATURE_UNIT)
    feed.Pressure.SetValue(100.0, config.HYSYS_PRESSURE_UNIT)
    feed.MolarFlow.SetValue(10.0, config.HYSYS_MOLAR_FLOW_UNIT)
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


def dump_operation(operation, label: str) -> None:
    report(f"\n--- [{label}] 创建成功对象 ---")
    report(f"  对象：{operation!r}")
    for attribute in ("Name", "TypeName", "VisibleTypeName", "TaggedName"):
        try:
            report(f"  {attribute} = {getattr(operation, attribute)!r}")
        except Exception as exc:  # noqa: BLE001
            report(f"  {attribute} 读取失败：{com_detail(exc)}")
    dump_type_info(
        operation, label,
        filters=(
            "reaction", "feed", "product", "stream", "inlet", "outlet",
            "conversion", "duty", "pressure", "temperature", "vapour", "vapor",
        ),
    )


def try_add(collection, type_name: str, index: int):
    operation_name = f"Probe_Reactor_{index}"
    attempts = (
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
            report(f"  ✓ {description} 成功 -> {operation!r}")
            return operation
        except Exception as exc:  # noqa: BLE001
            report(f"  ✗ {description} 失败：{com_detail(exc)}")
    return None


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

        case = setup_case(app)
        flowsheet = case.Flowsheet
        operations = flowsheet.Operations
        report(f"\nFlowsheet.Operations = {operations!r}")
        dump_type_info(
            operations, "Flowsheet.Operations",
            filters=("add", "item", "count", "names", "name"),
        )

        successes = 0
        for index, type_name in enumerate(REACTOR_TYPE_CANDIDATES, start=1):
            report(f"\n[{index}/{len(REACTOR_TYPE_CANDIDATES)}] 类型 {type_name!r}")
            operation = try_add(operations, type_name, index)
            if operation is not None:
                successes += 1
                dump_operation(operation, f"Reactor {type_name}")

        report(f"\n成功创建反应器次数：{successes}")
        if successes == 0:
            report("没有候选类型串成功；下一步改用模板 Case 方案或查 V15 类型库枚举。")
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
