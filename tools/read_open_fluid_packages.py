# -*- coding: utf-8 -*-
"""
read_open_fluid_packages.py — 只读导出当前 HYSYS 中所有 Case / 流体包信息
================================================================
用途：
  让用户手动在 HYSYS GUI 里建一个 Peng-Robinson 流体包，保持打开，
  然后运行本脚本，读出 V15 实际使用的 PropertyPackageName / TypeName。

不创建、不修改、不保存任何 Case，也不会 Quit HYSYS。

运行：
    python tools/read_open_fluid_packages.py
"""

from __future__ import annotations

import gc
import sys

import pythoncom
import pywintypes
import win32com.client

PROGID = "HYSYS.Application"


def com_detail(exc: BaseException) -> str:
    if isinstance(exc, pywintypes.com_error):
        hresult = getattr(exc, "hresult", None)
        excepinfo = getattr(exc, "excepinfo", None)
        hr_text = f"0x{hresult & 0xFFFFFFFF:08X}" if isinstance(hresult, int) else "N/A"
        return f"HRESULT={hr_text}, excepinfo={excepinfo!r}"
    return f"{type(exc).__name__}: {exc}"


def safe_get(obj, attribute: str, label: str, default: str = "<unavailable>"):
    try:
        value = getattr(obj, attribute)
        print(f"      {label}.{attribute} = {value!r}")
        return value
    except Exception as exc:  # noqa: BLE001
        print(f"      {label}.{attribute} 读取失败：{com_detail(exc)}")
        return default


def collection_items(collection, label: str) -> list:
    """兼容 0-based / 1-based 索引读取 COM 集合。"""
    items: list = []
    try:
        count = int(collection.Count)
    except Exception as exc:  # noqa: BLE001
        print(f"  {label}.Count 读取失败：{com_detail(exc)}")
        return items

    # HYSYS 文档通常使用 Item(0)；若 0 基失败，再尝试 1 基。
    for start in (0, 1):
        items = []
        for index in range(start, start + count):
            try:
                items.append(collection.Item(index))
            except Exception:
                items = []
                break
        if items:
            break
    return items


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    pythoncom.CoInitialize()
    try:
        try:
            app = win32com.client.GetActiveObject(PROGID)
        except Exception as exc:  # noqa: BLE001
            print(f"没有运行中的 HYSYS：{com_detail(exc)}")
            print("请先打开 HYSYS，或手动新建一个含 Peng-Robinson 的 Case。")
            return 1

        print("=" * 70)
        print("当前打开的 SimulationCases / FluidPackages")
        print("=" * 70)

        cases = collection_items(app.SimulationCases, "SimulationCases")
        if not cases:
            try:
                active = app.ActiveObject
                if active is not None:
                    cases = [active]
                    print("  SimulationCases 为空，改用 Application.ActiveObject")
            except Exception as exc:  # noqa: BLE001
                print(f"  Application.ActiveObject 读取失败：{com_detail(exc)}")

        if not cases:
            print("  没有找到打开的 Case。请先在 HYSYS 中新建/打开一个 Case。")
            return 0

        for case_index, case in enumerate(cases):
            print(f"\nCase[{case_index}]")
            safe_get(case, "Name", "SimulationCase")
            safe_get(case, "Title", "SimulationCase")

            try:
                basis = case.BasisManager
            except Exception as exc:  # noqa: BLE001
                print(f"  BasisManager 读取失败：{com_detail(exc)}")
                continue

            fluid_packages = collection_items(basis.FluidPackages, "FluidPackages")
            print(f"  FluidPackages 数量：{len(fluid_packages)}")

            for fp_index, fluid_pkg in enumerate(fluid_packages):
                print(f"\n  FluidPackage[{fp_index}]")
                safe_get(fluid_pkg, "Name", "FluidPackage")
                safe_get(fluid_pkg, "PropertyPackageName", "FluidPackage")

                try:
                    prop_pkg = fluid_pkg.PropertyPackage
                except Exception as exc:  # noqa: BLE001
                    print(f"      FluidPackage.PropertyPackage 读取失败：{com_detail(exc)}")
                else:
                    print(f"      FluidPackage.PropertyPackage = {prop_pkg!r}")
                    safe_get(prop_pkg, "Name", "PropertyPackage")
                    safe_get(prop_pkg, "TypeName", "PropertyPackage")
                    safe_get(prop_pkg, "VisibleTypeName", "PropertyPackage")

                try:
                    components = collection_items(
                        fluid_pkg.Components, "FluidPackage.Components"
                    )
                except Exception as exc:  # noqa: BLE001
                    print(f"      Components 读取失败：{com_detail(exc)}")
                else:
                    names = []
                    for component in components:
                        try:
                            names.append(str(component.Name))
                        except Exception:
                            names.append(repr(component))
                    print(f"      Components 数量：{len(names)}")
                    print(f"      Components：{names}")

        print("\n完成：只读操作，未修改/关闭任何 Case。")
        return 0
    finally:
        app = None
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
