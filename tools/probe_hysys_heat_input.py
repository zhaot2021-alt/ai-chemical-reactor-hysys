# -*- coding: utf-8 -*-
"""
probe_hysys_heat_input.py — 场景1热输入 API 探针
================================================================
目标：FEED 保持 520℃，通过能量流/加热器提供重整热。
只验证：
  1. EnergyStreams.Add / MaterialStreams.Add + IsEnergyStream
  2. EnergyStream.HeatFlow.GetValue/SetValue
  3. GibbsReactor.EnergyStream / DutyControlValve
  4. 候选 heater Operation 类型名
不修改/保存用户 Case；测试完关闭临时 Case。
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
REPORT_PATH = ROOT / "output" / "probe_hysys_heat_input_report.txt"

COMPONENTS = ("Methane", "H2O", "CO", "CO2", "Hydrogen")
FEED_T_K = 793.15
FEED_P_KPA = 1350.0
FEED_FLOW = 1000.0
FEED_COMP = {
    "Methane": 0.2703, "H2O": 0.7297,
    "CO": 0.0, "CO2": 0.0, "Hydrogen": 0.0,
}
DUTY_VALUES_KJ_H = (0.0, 1.0e7, 5.0e7, 1.0e8, 2.0e8, -1.0e8)

_REPORT_LINES: list[str] = []


def report(message: str = "") -> None:
    print(message, flush=True)
    _REPORT_LINES.append(str(message))


def com_detail(exc: BaseException) -> str:
    if isinstance(exc, pywintypes.com_error):
        hresult = getattr(exc, "hresult", None)
        excepinfo = getattr(exc, "excepinfo", None)
        hr = f"0x{hresult & 0xFFFFFFFF:08X}" if isinstance(hresult, int) else "N/A"
        return f"HRESULT={hr}, excepinfo={excepinfo!r}"
    return f"{type(exc).__name__}: {exc}"


def type_info(obj, label: str, filters: tuple[str, ...]) -> None:
    report(f"\n[{label}]")
    try:
        ti = obj._oleobj_.GetTypeInfo()
        ta = ti.GetTypeAttr()
    except Exception as exc:  # noqa: BLE001
        report(f"  type info failed: {com_detail(exc)}")
        return
    for i in range(ta.cFuncs):
        try:
            fd = ti.GetFuncDesc(i)
            names = ti.GetNames(fd.memid)
        except Exception:
            continue
        name = names[0] if names else "?"
        if not any(f.lower() in name.lower() for f in filters):
            continue
        report(f"  {name:40s} params={list(names[1:])}")


def safe(obj, attr: str, label: str):
    try:
        value = getattr(obj, attr)
        report(f"  {label}.{attr} = {value!r}")
        return value
    except Exception as exc:  # noqa: BLE001
        report(f"  {label}.{attr} failed: {com_detail(exc)}")
        return None


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    report(f"time: {datetime.now().isoformat(timespec='seconds')}")

    app = None
    case = None
    owns = False
    com_init = False
    try:
        pythoncom.CoInitialize()
        com_init = True
        try:
            app = win32com.client.GetActiveObject(PROGID)
            report("attached existing HYSYS (will not Quit)")
        except Exception:
            app = win32com.client.Dispatch(PROGID)
            owns = True
            report("started new HYSYS (will Quit)")
        try:
            app.Visible = True
        except Exception:
            pass

        case_name = f"HeatProbe_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        try:
            case = app.SimulationCases.Add(case_name)
        except Exception:
            case = app.SimulationCases.Add()
        basis = case.BasisManager
        basis.StartBasisChange()
        fp = basis.FluidPackages.Add("Probe_HeatPkg")
        fp.PropertyPackageName = config.PROPERTY_PACKAGE_NAME
        for cid in COMPONENTS:
            fp.Components.Add(cid)
        basis.EndBasisChange()

        feed = case.Flowsheet.MaterialStreams.Add("FEED")
        feed.Temperature.SetValue(FEED_T_K, "K")
        feed.Pressure.SetValue(FEED_P_KPA, "kPa")
        feed.MolarFlow.SetValue(FEED_FLOW, "kgmole/h")
        feed.ComponentMolarFractionValue = tuple(
            FEED_COMP[cid] for cid in COMPONENTS
        )
        case.Solver.CanSolve = True

        vapour = case.Flowsheet.MaterialStreams.Add("VAPOUR")
        liquid = case.Flowsheet.MaterialStreams.Add("LIQUID")
        ops = case.Flowsheet.Operations
        reactor = ops.Add("R-GIBBS", "gibbsreactorop")
        reactor.Feeds.Add(feed)
        reactor.VapourProduct = vapour
        reactor.LiquidProduct = liquid
        reactor.PressureDropValue = 0.0
        case.Solver.CanSolve = True
        report(f"initial VAPOUR T = {vapour.Temperature.GetValue('K'):.4f} K")

        # ---- 候选 heater 类型 ----
        for candidate in ("heaterop", "heater", "HeaterOp"):
            try:
                heater = ops.Add(f"Probe_Heater_{candidate}", candidate)
                report(f"heater Add({candidate!r}) OK -> {heater!r}")
                type_info(heater, f"Heater {candidate}",
                          ("feed", "product", "stream", "duty", "temperature", "pressure", "energy"))
                break
            except Exception as exc:  # noqa: BLE001
                report(f"heater Add({candidate!r}) failed: {com_detail(exc)}")

        # ---- 能量流 ----
        q = None
        energy_streams = case.Flowsheet.EnergyStreams
        type_info(energy_streams, "EnergyStreams", ("add", "item", "count", "names"))
        try:
            q = energy_streams.Add("Probe_Q")
            report(f"EnergyStreams.Add OK -> {q!r}")
        except Exception as exc:  # noqa: BLE001
            report(f"EnergyStreams.Add failed: {com_detail(exc)}")
        if q is None:
            try:
                q = case.Flowsheet.MaterialStreams.Add("Probe_Q")
                q.IsEnergyStream = True
                report(f"MaterialStreams.Add + IsEnergyStream OK -> {q!r}")
            except Exception as exc:  # noqa: BLE001
                report(f"create energy stream fallback failed: {com_detail(exc)}")

        if q is None:
            report("no energy stream created; stop")
        else:
            type_info(q, "EnergyStream", ("heat", "duty", "flow", "name", "value"))
            safe(q, "IsEnergyStream", "EnergyStream")
            heat_flow = safe(q, "HeatFlow", "EnergyStream")
            if heat_flow is not None:
                try:
                    report(f"  HeatFlow.GetValue('kJ/h') = {heat_flow.GetValue('kJ/h')!r}")
                except Exception as exc:  # noqa: BLE001
                    report(f"  HeatFlow.GetValue failed: {com_detail(exc)}")
            try:
                reactor.EnergyStream = q
                report("reactor.EnergyStream = q OK")
            except Exception as exc:  # noqa: BLE001
                report(f"reactor.EnergyStream = q failed: {com_detail(exc)}")

            safe(reactor, "DutyControlValve", "Reactor")
            try:
                old = reactor.DutyControlValve
                report(f"  DutyControlValve repr/type: {old!r} {type(old).__name__}")
                if isinstance(old, bool):
                    reactor.DutyControlValve = True
                    report("  DutyControlValve = True OK")
            except Exception as exc:  # noqa: BLE001
                report(f"  DutyControlValve handling failed: {com_detail(exc)}")

            for duty in DUTY_VALUES_KJ_H:
                try:
                    q.HeatFlow.SetValue(duty, "kJ/h")
                    report(f"  HeatFlow.SetValue({duty}, 'kJ/h') OK")
                except Exception as exc:  # noqa: BLE001
                    report(f"  HeatFlow.SetValue({duty}) failed: {com_detail(exc)}")
                    break
                try:
                    case.Solver.CanSolve = True
                    outlet_t = vapour.Temperature.GetValue("K")
                    report(f"  duty={duty:g} kJ/h -> VAPOUR T = {outlet_t:.4f} K")
                except Exception as exc:  # noqa: BLE001
                    report(f"  read after duty={duty} failed: {com_detail(exc)}")

        try:
            case.Close(False)
        except Exception:
            pass
        case = None
    except Exception as exc:  # noqa: BLE001
        report(f"EXCEPTION: {com_detail(exc)}")
        report(traceback.format_exc())
    finally:
        gc.collect()
        if case is not None:
            try:
                case.Close(False)
            except Exception:
                pass
        if app is not None and owns:
            try:
                app.Quit()
                report("Quit owned HYSYS")
                time.sleep(2)
            except Exception as exc:  # noqa: BLE001
                report(f"Quit failed: {com_detail(exc)}")
        app = None
        gc.collect()
        if com_init:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        REPORT_PATH.write_text("\n".join(_REPORT_LINES), encoding="utf-8")
        print(f"\nreport: {REPORT_PATH}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
