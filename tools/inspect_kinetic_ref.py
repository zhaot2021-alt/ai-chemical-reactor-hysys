# -*- coding: utf-8 -*-
r"""
inspect_kinetic_ref.py — 只读导出 kinetic PFR/CSTR 参考 Case 的真实配置
=======================================================================
在 HYSYS 工作站运行（不写任何参数、不求解）：

    python tools\inspect_kinetic_ref.py [可选: .hsc 路径 ...]

默认读取 output\kinetic_pfr_ref.hsc 和 output\kinetic_cstr_ref.hsc。
输出 output\inspect_kinetic_ref_report.txt，请回传本机。
"""

from __future__ import annotations

import gc
import sys
import time
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

PROGID = "HYSYS.Application"
REPORT_PATH = ROOT / "output" / "inspect_kinetic_ref_report.txt"
DEFAULT_PATHS = (
    ROOT / "output" / "kinetic_pfr_ref.hsc",
    ROOT / "output" / "kinetic_cstr_ref.hsc",
)
REPORT: list[str] = []


def log(message: str = "") -> None:
    print(message, flush=True)
    REPORT.append(str(message))


def err(exc: BaseException) -> str:
    if isinstance(exc, pywintypes.com_error):
        hr = getattr(exc, "hresult", None)
        return f"HRESULT={hr and f'0x{hr & 0xFFFFFFFF:08X}'} {getattr(exc, 'excepinfo', None)!r}"
    return f"{type(exc).__name__}: {exc}"


def pv(obj: Any, attr: str, units: tuple[str, ...]) -> str:
    """读属性：优先 GetValue 多单位尝试，失败打印 repr。"""
    try:
        value = getattr(obj, attr)
    except Exception as exc:  # noqa: BLE001
        return f"GET-FAIL({err(exc)})"
    if hasattr(value, "GetValue"):
        for unit in units:
            try:
                return f"{value.GetValue(unit)} [{unit}]"
            except Exception:
                continue
        try:
            return f"{float(value)} [default]"
        except Exception as exc:  # noqa: BLE001
            return f"raw({err(exc)})"
    return repr(value)


def name_of(obj: Any) -> str:
    try:
        return str(obj.Name)
    except Exception as exc:  # noqa: BLE001
        return f"<{err(exc)}>"


def item_names(collection: Any, label: str) -> list[str]:
    names: list[str] = []
    try:
        count = int(collection.Count)
    except Exception as exc:  # noqa: BLE001
        log(f"  {label}: Count failed {err(exc)}")
        return names
    for start in (0, 1):
        names = []
        for index in range(start, start + count):
            try:
                names.append(name_of(collection.Item(index)))
            except Exception:
                names = []
                break
        if names:
            log(f"  {label} ({count}): {names}")
            return names
    log(f"  {label}: items unreadable, count={count}")
    return names


def read_reaction(rxn: Any, index: int) -> None:
    log(f"\n  --- reaction[{index}] {name_of(rxn)} ---")
    for attr in ("TypeName", "VisibleTypeName"):
        try:
            log(f"    {attr} = {getattr(rxn, attr)!r}")
        except Exception as exc:  # noqa: BLE001
            log(f"    {attr} GET failed {err(exc)}")
    log(f"    ReactionPhase = {pv(rxn, 'ReactionPhase', ())}")
    log(f"    Basis = {pv(rxn, 'Basis', ())}")
    log(f"    BasisUnits2 = {pv(rxn, 'BasisUnits2', ())}")
    log(f"    RateUnits2 = {pv(rxn, 'RateUnits2', ())}")
    log(f"    ReactantStoichCoefValue = {pv(rxn, 'ReactantStoichCoefValue', ())}")
    log(f"    ComponentForwardOrderValue = {pv(rxn, 'ComponentForwardOrderValue', ())}")
    log(f"    ComponentReverseOrderValue = {pv(rxn, 'ComponentReverseOrderValue', ())}")
    log(f"    ForwardActivationEnergy = "
        f"{pv(rxn, 'ForwardActivationEnergy', ('kJ/kgmole', 'J/mol', 'kcal/kgmole', 'kJ/kmol'))}")
    log(f"    ForwardFrequencyFactor = "
        f"{pv(rxn, 'ForwardFrequencyFactor', ('m3/kgmole-s', 'm3/kmol-s', 'L/mol-s', '1/s', 'kgmole/m3-s'))}")
    log(f"    ReverseActivationEnergy = "
        f"{pv(rxn, 'ReverseActivationEnergy', ('kJ/kgmole', 'J/mol', 'kcal/kgmole', 'kJ/kmol'))}")
    log(f"    ReverseFrequencyFactor = "
        f"{pv(rxn, 'ReverseFrequencyFactor', ('m3/kgmole-s', 'm3/kmol-s', 'L/mol-s', '1/s', 'kgmole/m3-s'))}")
    log(f"    ForwardTemperatureExponent = {pv(rxn, 'ForwardTemperatureExponent', ())}")
    log(f"    ReverseTemperatureExponent = {pv(rxn, 'ReverseTemperatureExponent', ())}")
    log(f"    MinTemperature = {pv(rxn, 'MinTemperatureValue', ('K', 'C'))}")
    log(f"    MaxTemperature = {pv(rxn, 'MaxTemperatureValue', ('K', 'C'))}")
    log(f"    HeatOfReaction = {pv(rxn, 'HeatOfReactionValue', ('kJ/kgmole',))}")
    item_names(getattr(rxn, "Reactants"), "Reactants")


def read_reaction_sets(manager: Any) -> None:
    sets = manager.ReactionSets
    log("\n=== reaction sets ===")
    try:
        count = int(sets.Count)
    except Exception as exc:  # noqa: BLE001
        log(f"  ReactionSets.Count failed {err(exc)}")
        return
    log(f"  ReactionSets.Count = {count}")
    for index in range(count):
        try:
            rset = sets.Item(index)
        except Exception as exc:  # noqa: BLE001
            log(f"  set[{index}] failed {err(exc)}")
            continue
        log(f"  set[{index}] Name={name_of(rset)!r} TypeName={pv(rset, 'TypeName', ())!r} "
            f"VisibleTypeName={pv(rset, 'VisibleTypeName', ())!r}")
        active = getattr(rset, "ActiveReactions")
        item_names(active, "ActiveReactions")
        try:
            active_count = int(active.Count)
        except Exception as exc:  # noqa: BLE001
            log(f"    ActiveReactions.Count failed {err(exc)}")
            continue
        for rxn_index in range(active_count):
            try:
                read_reaction(active.Item(rxn_index), rxn_index)
            except Exception as exc:  # noqa: BLE001
                log(f"    reaction[{rxn_index}] failed {err(exc)}")


def read_reactor(op: Any, label: str) -> None:
    type_name = pv(op, "TypeName", ())
    log(f"\n=== operation {label} TypeName={type_name!r} ===")
    try:
        feeds = op.AttachedFeeds
    except Exception:
        feeds = getattr(op, "Feeds", None)
    if feeds is not None:
        item_names(feeds, "Feeds")
    if "pfr" in str(type_name).lower():
        log(f"    Product = {name_of(op.Product)}")
    else:
        log(f"    VapourProduct = {pv(op, 'VapourProduct', ())}")
        log(f"    LiquidProduct = {pv(op, 'LiquidProduct', ())}")
    log(f"    EnergyStream = {pv(op, 'EnergyStream', ())}")
    try:
        log(f"    ReactionSet = {name_of(op.ReactionSet)}")
    except Exception as exc:  # noqa: BLE001
        log(f"    ReactionSet failed {err(exc)}")
    log(f"    PressureDrop = {pv(op, 'PressureDropValue', ('kPa', 'bar'))}")
    if "pfr" in str(type_name).lower():
        log(f"    TotalVolume = {pv(op, 'TotalVolumeValue', ('m3',))}")
        log(f"    TubeLength = {pv(op, 'TubeLengthValue', ('m',))}")
        log(f"    TubeDiameter = {pv(op, 'TubeDiameterValue', ('m',))}")
        log(f"    PhaseType = {pv(op, 'PhaseType', ())}")
        log(f"    HoldupInitType = {pv(op, 'HoldupInitType', ())}")
    else:
        log(f"    Volume = {pv(op, 'VolumeValue', ('m3',))}")
        log(f"    LiquidLevel = {pv(op, 'LiquidLevelValue', ('%',))}")
        log(f"    HoldupInitType = {pv(op, 'HoldupInitType', ())}")
    log(f"    HeatFlow = {pv(op, 'HeatFlowValue', ('kJ/h', 'kW'))}")


def read_stream(stream: Any, label: str) -> None:
    log(f"  --- {label} {name_of(stream)} ---")
    for attr, unit in (("Temperature", "K"), ("Pressure", "kPa"),
                       ("MolarFlow", "kgmole/h"), ("MassFlow", "kg/h")):
        try:
            log(f"    {attr} = {getattr(stream, attr).GetValue(unit)} [{unit}]")
        except Exception as exc:  # noqa: BLE001
            log(f"    {attr} failed {err(exc)}")


def inspect_case(app: Any, path: Path) -> None:
    log(f"\n{'=' * 68}\nOPEN {path}\n{'=' * 68}")
    case = app.SimulationCases.Open(str(path))
    try:
        basis = case.BasisManager
        log("\n=== fluid packages ===")
        item_names(basis.FluidPackages, "FluidPackages")
        try:
            fp = basis.FluidPackages.Item(0)
        except Exception as exc:  # noqa: BLE001
            log(f"  FluidPackages.Item(0) failed {err(exc)}")
            fp = None
        if fp is not None:
            log(f"  FP Name={name_of(fp)} PropertyPackageName={pv(fp, 'PropertyPackageName', ())}")
            item_names(fp.Components, "Components")

        read_reaction_sets(basis.ReactionPackageManager)

        log("\n=== operations ===")
        try:
            ops = case.Flowsheet.Operations
            count = int(ops.Count)
        except Exception as exc:  # noqa: BLE001
            log(f"  Operations failed {err(exc)}")
            count = 0
        for index in range(count):
            try:
                read_reactor(ops.Item(index), f"op[{index}]")
            except Exception as exc:  # noqa: BLE001
                log(f"  op[{index}] failed {err(exc)}")

        log("\n=== material streams ===")
        try:
            streams = case.Flowsheet.MaterialStreams
            stream_count = int(streams.Count)
        except Exception as exc:  # noqa: BLE001
            log(f"  MaterialStreams failed {err(exc)}")
            stream_count = 0
        for index in range(stream_count):
            try:
                read_stream(streams.Item(index), f"stream[{index}]")
            except Exception as exc:  # noqa: BLE001
                log(f"  stream[{index}] failed {err(exc)}")
    finally:
        try:
            case.Close(False)
        except Exception:
            pass


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    paths = [Path(item).resolve() for item in sys.argv[1:]] or list(DEFAULT_PATHS)

    app = None
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
        for path in paths:
            if not path.exists():
                log(f"SKIP missing file: {path}")
                continue
            try:
                inspect_case(app, path)
            except Exception as exc:  # noqa: BLE001
                log(f"inspect {path} failed: {err(exc)}")
    except Exception as exc:  # noqa: BLE001
        log(f"EXCEPTION: {err(exc)}")
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
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        REPORT_PATH.write_text("\n".join(REPORT), encoding="utf-8")
        print(f"\nreport: {REPORT_PATH}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
