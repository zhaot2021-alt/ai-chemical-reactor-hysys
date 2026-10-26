# -*- coding: utf-8 -*-
"""
read_open_operations.py — 只读导出当前 HYSYS 中的物流与单元操作
================================================================
用途：
  用户在 HYSYS GUI 里手动建一个带 Gibbs/Conversion 反应器的参考 Case 后，
  本脚本读出操作的真实 TypeName、TaggedName、成员表，以及物流名称/参数。

不创建、不修改、不保存任何 Case，也不会 Quit HYSYS。

运行：
    python tools/read_open_operations.py
"""

from __future__ import annotations

import gc
import sys

import pythoncom
import pywintypes
import win32com.client

PROGID = "HYSYS.Application"
REPORT_PATH = None

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


def safe_get(obj, attribute: str, label: str):
    try:
        value = getattr(obj, attribute)
        report(f"      {label}.{attribute} = {value!r}")
        return value
    except Exception as exc:  # noqa: BLE001
        report(f"      {label}.{attribute} 读取失败：{com_detail(exc)}")
        return None


def collection_items(collection) -> list:
    try:
        count = int(collection.Count)
    except Exception as exc:  # noqa: BLE001
        report(f"  集合 Count 读取失败：{com_detail(exc)}")
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


def dump_type_info(obj, label: str, filters: tuple[str, ...]) -> None:
    report(f"\n--- [{label}] COM 成员（过滤：{filters}）---")
    try:
        type_info = obj._oleobj_.GetTypeInfo()
        attr = type_info.GetTypeAttr()
    except Exception as exc:  # noqa: BLE001
        report(f"  GetTypeInfo 失败：{com_detail(exc)}")
        return
    for index in range(attr.cFuncs):
        try:
            fd = type_info.GetFuncDesc(index)
            names = type_info.GetNames(fd.memid)
        except Exception:
            continue
        name = names[0] if names else "?"
        if not any(token.lower() in name.lower() for token in filters):
            continue
        report(f"  {name:42s} | params={list(names[1:])}")


def describe_stream(stream, index: int) -> None:
    report(f"\n  Stream[{index}]")
    safe_get(stream, "Name", "Stream")
    safe_get(stream, "TypeName", "Stream")
    safe_get(stream, "VisibleTypeName", "Stream")
    for attribute, unit in (
        ("Temperature", "K"),
        ("Pressure", "kPa"),
        ("MolarFlow", "kgmole/h"),
    ):
        try:
            value = getattr(stream, attribute).GetValue(unit)
            report(f"      {attribute}.GetValue({unit!r}) = {value!r}")
        except Exception as exc:  # noqa: BLE001
            report(f"      {attribute} 读取失败：{com_detail(exc)}")
    try:
        composition = list(stream.ComponentMolarFractionValue)
        report(f"      ComponentMolarFractionValue = {composition}")
    except Exception as exc:  # noqa: BLE001
        report(f"      ComponentMolarFractionValue 读取失败：{com_detail(exc)}")
    try:
        report(f"      VapourFractionValue = {stream.VapourFractionValue!r}")
    except Exception as exc:  # noqa: BLE001
        report(f"      VapourFractionValue 读取失败：{com_detail(exc)}")


def main() -> int:
    global REPORT_PATH
    from pathlib import Path

    root = Path(__file__).resolve().parents[1]
    REPORT_PATH = root / "output" / "read_open_operations_report.txt"

    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    pythoncom.CoInitialize()
    app = None
    try:
        try:
            app = win32com.client.GetActiveObject(PROGID)
        except Exception as exc:  # noqa: BLE001
            report(f"没有运行中的 HYSYS：{com_detail(exc)}")
            return 1

        cases = collection_items(app.SimulationCases)
        if not cases:
            try:
                active = app.ActiveObject
            except Exception as exc:  # noqa: BLE001
                report(f"没有打开 Case，且 ActiveObject 不可用：{com_detail(exc)}")
                return 0
            if active is not None:
                cases = [active]

        if not cases:
            report("没有打开的 Case")
            return 0

        for case_index, case in enumerate(cases):
            report("\n" + "=" * 72)
            report(f"Case[{case_index}]")
            report("=" * 72)
            safe_get(case, "Name", "SimulationCase")

            try:
                flowsheet = case.Flowsheet
            except Exception as exc:  # noqa: BLE001
                report(f"Flowsheet 读取失败：{com_detail(exc)}")
                continue

            try:
                streams = collection_items(flowsheet.MaterialStreams)
            except Exception as exc:  # noqa: BLE001
                report(f"MaterialStreams 读取失败：{com_detail(exc)}")
            else:
                report(f"\n  MaterialStreams 数量：{len(streams)}")
                for stream_index, stream in enumerate(streams):
                    describe_stream(stream, stream_index)

            try:
                operations = flowsheet.Operations
            except Exception as exc:  # noqa: BLE001
                report(f"\nOperations 读取失败：{com_detail(exc)}")
                continue

            operation_list = collection_items(operations)
            report(f"\n  Operations 数量：{len(operation_list)}")
            for operation_index, operation in enumerate(operation_list):
                report(f"\n  Operation[{operation_index}]")
                safe_get(operation, "Name", "Operation")
                safe_get(operation, "TypeName", "Operation")
                safe_get(operation, "VisibleTypeName", "Operation")
                safe_get(operation, "TaggedName", "Operation")
                dump_type_info(
                    operation, f"Operation[{operation_index}]",
                    filters=(
                        "reaction", "feed", "product", "stream",
                        "inlet", "outlet", "conversion", "temperature",
                        "pressure", "duty", "vapour", "vapor",
                    ),
                )

        report("\n完成：只读操作，未修改/关闭任何 Case。")
        return 0
    finally:
        gc.collect()
        app = None
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        if REPORT_PATH is not None:
            try:
                REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
                REPORT_PATH.write_text("\n".join(_REPORT_LINES), encoding="utf-8")
                print(f"\n完整报告已写入：{REPORT_PATH}", flush=True)
            except Exception as exc:  # noqa: BLE001
                print(f"报告写入失败：{exc}", flush=True)


if __name__ == "__main__":
    raise SystemExit(main())
