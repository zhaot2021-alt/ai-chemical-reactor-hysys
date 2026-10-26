# -*- coding: utf-8 -*-
r"""
probe_hysys_kinetic_api.py — CSTR/PFR + Kinetic Reaction COM API 侦查
=====================================================================
在 HYSYS 工作站运行（可见窗口便于处理可能的弹窗）：

    python tools\probe_hysys_kinetic_api.py --visible

做三件事：
  1. 枚举已注册的 HYSYS/Aspen 类型库，打印与 reactor / reaction / kinetic
     / CSTR / PFR 有关的类型名、成员名和帮助文本（不猜接口）；
  2. 在临时 Case 中实际尝试 ReactionSets.Add / Reactions.Add / Operations.Add
     的候选类型串，确认哪些可用；
  3. 若找到 kinetic 反应与 CSTR/PFR，尝试创建、连接、设参数、求解，
     用真实结果判断接口是否闭环。

结果：
  output\probe_hysys_kinetic_api_report.txt   精简报告（优先回传）
  output\hysys_kinetic_type_dump.txt          完整类型/成员明细（备用）

不保存 .hsc；结束时关闭临时 Case，本次启动的 HYSYS 实例会 Quit。
"""

from __future__ import annotations

import gc
import re
import sys
import time
import traceback
import winreg
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import pythoncom
import pywintypes
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
import config  # noqa: E402

PROGID = "HYSYS.Application"
REPORT_PATH = ROOT / "output" / "probe_hysys_kinetic_api_report.txt"
TYPE_DUMP_PATH = ROOT / "output" / "hysys_kinetic_type_dump.txt"

COMPONENTS = ("Methane", "H2O", "CO", "CO2", "Hydrogen")
INVKIND = {1: "FUNC", 2: "GET", 4: "PUT", 8: "PUTREF", 16: "PUTREF"}
IGNORE_MEMBERS = {
    "QueryInterface", "AddRef", "Release", "GetTypeInfoCount",
    "GetTypeInfo", "GetIDsOfNames", "Invoke",
}

# 类型名/帮助文本里出现这些词就进入侦查范围
KEY_TYPE_RE = re.compile(
    r"(cstr|pfr|plug.?flow|kinetic|reaction|reactor|arrhen|activation|"
    r"pre.?exp|rate.?constant|reaction.?set)",
    re.IGNORECASE,
)
# 需要导出全部成员名的强相关类型
STRONG_TYPE_RE = re.compile(
    r"^(I)?(CSTR|PFR|Plug.?Flow|Kinetic|Reaction|Reactor)",
    re.IGNORECASE,
)
# 成员中与速率/设计参数有关的关键词
KEY_MEMBER_RE = re.compile(
    r"(rate|kinetic|arrhen|activation|pre.?exp|frequency|forward|reverse|"
    r"order|phase|basis|stoich|reactant|feed|product|reaction|volume|"
    r"length|diameter|duty|temperature|pressure|vapour|liquid|energy|"
    r"composition|conversion|hold.?up)",
    re.IGNORECASE,
)

_REPORT: list[str] = []
_FULL: list[str] = []
_FULL_TYPE_LINES: list[str] = []


def report(message: str = "") -> None:
    print(message, flush=True)
    _REPORT.append(str(message))


def full(message: str = "") -> None:
    _FULL.append(str(message))


def com_detail(exc: BaseException) -> str:
    if isinstance(exc, pywintypes.com_error):
        hr = getattr(exc, "hresult", None)
        info = getattr(exc, "excepinfo", None)
        htxt = f"0x{hr & 0xFFFFFFFF:08X}" if isinstance(hr, int) else "N/A"
        return f"HRESULT={htxt}, excepinfo={info!r}"
    return f"{type(exc).__name__}: {exc}"


def safe_get(obj: Any, attr: str, label: str) -> Any:
    try:
        value = getattr(obj, attr)
        report(f"  {label}.{attr} = {value!r}")
        return value
    except Exception as exc:  # noqa: BLE001
        report(f"  {label}.{attr} GET failed: {com_detail(exc)}")
        return None


# ==================== 类型库枚举 ====================
def _type_doc(lib: Any, index: int) -> tuple[str, str]:
    try:
        name, doc = lib.GetDocumentation(index)[:2]
        return str(name or ""), str(doc or "")
    except Exception:  # noqa: BLE001
        return "", ""


def _containing_type_lib(app: Any) -> tuple[Any, str] | None:
    try:
        type_info = app._oleobj_.GetTypeInfo()
        lib, _index = type_info.GetContainingTypeLib()
        return lib, "Application containing type library"
    except Exception as exc:  # noqa: BLE001
        report(f"  containing type library failed: {com_detail(exc)}")
        return None


def _scan_typelib_root(
    root: Any,
    found: list[tuple[str, Any]],
    seen_paths: set[str],
) -> None:
    try:
        guid_count = winreg.QueryInfoKey(root)[0]
    except OSError:
        return
    for guid_i in range(guid_count):
        try:
            guid_text = winreg.EnumKey(root, guid_i)
            guid_key = winreg.OpenKey(root, guid_text)
        except OSError:
            continue
        try:
            version_count = winreg.QueryInfoKey(guid_key)[0]
        except OSError:
            version_count = 0
        for ver_i in range(version_count):
            try:
                version = winreg.EnumKey(guid_key, ver_i)
                ver_key = winreg.OpenKey(guid_key, version)
                try:
                    lib_desc = str(winreg.QueryValueEx(ver_key, "")[0])
                except OSError:
                    lib_desc = ""
            except OSError:
                continue
            for sub in ("0", "win32", "win64"):
                try:
                    path_key = winreg.OpenKey(ver_key, sub)
                    path = str(winreg.QueryValueEx(path_key, "")[0])
                except OSError:
                    continue
                haystack = f"{lib_desc} {path} {guid_text}".lower()
                if "hysys" not in haystack or path in seen_paths:
                    continue
                seen_paths.add(path)
                lib = None
                try:
                    lib = pythoncom.LoadTypeLib(path)
                except Exception:
                    try:
                        major, _, minor = version.partition(".")
                        lib = pythoncom.LoadRegTypeLib(
                            pythoncom.IID(guid_text),
                            int(major or 0),
                            int((minor or "0").split(".")[0] or 0),
                            0,
                        )
                    except Exception as exc:  # noqa: BLE001
                        report(
                            f"  registered type lib load failed: "
                            f"{lib_desc or guid_text} @ {path}: "
                            f"{com_detail(exc)}"
                        )
                if lib is not None:
                    label = f"{lib_desc or guid_text} @ {path}"
                    found.append((label, lib))
                    report(f"  loaded registered type lib: {label}")


def _registered_type_libs() -> list[tuple[str, Any]]:
    """从注册表找 HYSYS 类型库并加载；找不到就返回空列表。"""
    roots = (
        (winreg.HKEY_CLASSES_ROOT, r"TypeLib"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\TypeLib"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Classes\TypeLib"),
    )
    found: list[tuple[str, Any]] = []
    seen_paths: set[str] = set()
    for hive, subkey in roots:
        for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
            try:
                root = winreg.OpenKey(
                    hive, subkey, 0, winreg.KEY_READ | view
                )
            except OSError:
                continue
            try:
                _scan_typelib_root(root, found, seen_paths)
            finally:
                try:
                    winreg.CloseKey(root)
                except Exception:
                    pass
    return found


def _member_lines(lib: Any, index: int) -> tuple[str, list[str]]:
    """返回 (类型头, 成员行列表)；调用方决定完整/过滤展示。"""
    name, doc = _type_doc(lib, index)
    try:
        type_info = lib.GetTypeInfo(index)
        attr = type_info.GetTypeAttr()
    except Exception as exc:  # noqa: BLE001
        return f"TYPE {name} <no type info: {com_detail(exc)}>", []

    header = (
        f"TYPE {name} | kind={getattr(attr, 'typekind', '?')} "
        f"| funcs={getattr(attr, 'cFuncs', 0)} "
        f"| vars={getattr(attr, 'cVars', 0)} | doc={doc!r}"
    )
    lines: list[str] = []
    for func_i in range(int(getattr(attr, "cFuncs", 0))):
        try:
            func_desc = type_info.GetFuncDesc(func_i)
            names = type_info.GetNames(func_desc.memid)
            member_name = str(names[0]) if names else "?"
            if member_name in IGNORE_MEMBERS:
                continue
            try:
                member_doc = str(type_info.GetDocumentation(func_desc.memid)[1] or "")
            except Exception:
                member_doc = ""
            invkind = INVKIND.get(func_desc.invkind, str(func_desc.invkind))
            params = ", ".join(str(item) for item in names[1:])
            line = f"  {member_name:42s} | {invkind:6s} | ({params})"
            if member_doc:
                line += f"  :: {member_doc}"
        except Exception as exc:  # noqa: BLE001
            line = f"  <function {func_i} failed: {com_detail(exc)}>"
        lines.append(line)

    for var_i in range(int(getattr(attr, "cVars", 0))):
        try:
            var_desc = type_info.GetVarDesc(var_i)
            names = type_info.GetNames(var_desc.memid)
            member_name = str(names[0]) if names else "?"
            try:
                member_doc = str(type_info.GetDocumentation(var_desc.memid)[1] or "")
            except Exception:
                member_doc = ""
            line = f"  [var] {member_name:36s} :: {member_doc}"
        except Exception as exc:  # noqa: BLE001
            line = f"  <var {var_i} failed: {com_detail(exc)}>"
        lines.append(line)

    return header, lines


def inspect_type_libraries(app: Any) -> list[str]:
    """枚举类型库；返回强相关类型名列表。"""
    report("\n=== 1. HYSYS type libraries / reaction-reactor types ===")
    libs: list[tuple[str, Any]] = []
    containing = _containing_type_lib(app)
    if containing:
        libs.append(containing)
    libs.extend(_registered_type_libs())

    seen_lib_labels: set[str] = set()
    all_types: list[tuple[str, Any, int, str, str]] = []
    for label, lib in libs:
        if label in seen_lib_labels:
            continue
        seen_lib_labels.add(label)
        try:
            count = int(lib.GetTypeInfoCount())
        except Exception as exc:  # noqa: BLE001
            report(f"  type lib {label}: count failed: {com_detail(exc)}")
            continue
        report(f"  type lib: {label} (types={count})")
        for index in range(count):
            name, doc = _type_doc(lib, index)
            if not name:
                continue
            _FULL_TYPE_LINES.append(f"{label} | {index} | {name} | {doc}")
            if KEY_TYPE_RE.search(name) or KEY_TYPE_RE.search(doc):
                all_types.append((label, lib, index, name, doc))

    # 按名称去重，保留第一次出现
    unique: dict[str, tuple[str, Any, int, str, str]] = {}
    for item in all_types:
        unique.setdefault(item[3].lower(), item)

    report(f"\n--- matching types ({len(unique)}) ---")
    strong_names: list[str] = []
    for label, _lib, index, name, doc in sorted(
        unique.values(), key=lambda x: x[3].lower()
    ):
        report(f"  {name:48s} :: {doc}")
        if STRONG_TYPE_RE.search(name):
            strong_names.append(name)

    report("\n--- member names of strongly-related types ---")
    dumped = 0
    for key, (label, lib, index, name, doc) in sorted(
        unique.items(), key=lambda item: item[1][3].lower()
    ):
        if not STRONG_TYPE_RE.search(name):
            continue
        header, lines = _member_lines(lib, index)
        full("")
        full(f"[{label}]")
        full(header)
        full("\n".join(lines))
        if dumped < 24:
            report("")
            report(f"{header}  [{label}]")
            for line in lines:
                if KEY_MEMBER_RE.search(line):
                    report(line)
        dumped += 1
        if dumped >= 40:
            break

    report(f"\n  dumped member details for {dumped} strong types; "
           f"full list: {TYPE_DUMP_PATH.name}")
    return strong_names


# ==================== 候选类型串 ====================
def _variants(name: str) -> list[str]:
    base = re.sub(r"^I(?=[A-Z])", "", name)
    values = {name, base, base.lower(), base.upper()}
    if base.lower().endswith("op"):
        values.add(base[:-2])
        values.add(base[:-2].lower())
    if not base.lower().endswith("op"):
        values.add(base + "op")
        values.add((base + "op").lower())
    return [item for item in values if item]


def _candidate_operations(type_names: Iterable[str]) -> list[str]:
    hardcoded = [
        "cstrop", "cstr", "CSTROp", "CSTR", "cstrreactorop", "cstrreactor",
        "pfrreactorop", "pfrreactor", "pfrop", "pfr", "PFROp", "PFR",
        "plugflowreactorop", "plugflowreactor", "plugflowop", "plugflow",
        "kineticcstrop", "kineticpfrop",
    ]
    candidates = list(hardcoded)
    for name in type_names:
        if re.search(r"(cstr|pfr|plug.?flow)", name, re.IGNORECASE):
            candidates.extend(_variants(name))
    return list(dict.fromkeys(candidates))


def _candidate_reaction_types(type_names: Iterable[str]) -> list[str]:
    hardcoded = [
        "kineticrxn", "kineticsrxn", "kinetic", "Kinetic", "Kinetics",
        "kineticreaction", "KineticReaction", "rateRxn", "ratereaction",
        "rateReaction", "reactionrxn",
    ]
    candidates = list(hardcoded)
    for name in type_names:
        if re.search(r"kinetic|rate|reaction", name, re.IGNORECASE):
            candidates.extend(_variants(name))
    return list(dict.fromkeys(candidates))


def _candidate_set_types(type_names: Iterable[str]) -> list[str]:
    hardcoded = [
        "Kinetic", "kinetic", "Kinetics", "kinetics", "Rate", "rate",
        "Reaction", "reaction", "Equilibrium", "Conversion",
    ]
    candidates = list(hardcoded)
    for name in type_names:
        if re.search(r"kinetic|rate|reaction.?set", name, re.IGNORECASE):
            candidates.extend(_variants(name))
    return list(dict.fromkeys(candidates))


# ==================== Case / 物流 ====================
def _wait_solver(case: Any, timeout_s: float = 20.0) -> None:
    try:
        case.Solver.CanSolve = True
    except Exception as exc:  # noqa: BLE001
        report(f"  Solver.CanSolve=True failed: {com_detail(exc)}")
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            if not bool(case.Solver.IsSolving):
                return
        except Exception as exc:  # noqa: BLE001
            report(f"  Solver.IsSolving failed: {com_detail(exc)}")
            return
        time.sleep(0.2)
    report(f"  WARNING: solver still busy after {timeout_s:.0f}s")


def _component_names(fp: Any) -> list[str]:
    names: list[str] = []
    count = int(fp.Components.Count)
    for start in (0, 1):
        names = []
        for index in range(start, start + count):
            try:
                names.append(str(fp.Components.Item(index).Name))
            except Exception:
                names = []
                break
        if names:
            return names
    return []


def _read_stream(stream: Any, component_names: list[str], label: str) -> None:
    report(f"  --- {label} ---")
    for attr, unit in (
        ("Temperature", "K"), ("Pressure", "kPa"),
        ("MolarFlow", "kgmole/h"), ("MassFlow", "kg/h"),
    ):
        try:
            value = getattr(stream, attr).GetValue(unit)
            report(f"    {attr:16s} = {float(value):.6g} {unit}")
        except Exception as exc:  # noqa: BLE001
            report(f"    {attr:16s} GET failed: {com_detail(exc)}")
    try:
        report(f"    VapourFraction   = {float(stream.VapourFractionValue):.6g}")
    except Exception as exc:  # noqa: BLE001
        report(f"    VapourFraction GET failed: {com_detail(exc)}")
    try:
        raw = [float(value) for value in stream.ComponentMolarFractionValue]
        shown = {
            component_names[index]: round(value, 6)
            for index, value in enumerate(raw[: len(component_names)])
        }
        report(f"    mole fractions   = {shown}")
    except Exception as exc:  # noqa: BLE001
        report(f"    mole fractions GET failed: {com_detail(exc)}")


def _setup_case(app: Any) -> tuple[Any, Any, Any, list[str]]:
    report("\n=== 2. temporary case + feed ===")
    try:
        case = app.SimulationCases.Add(
            f"KineticProbe_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
    except Exception:
        case = app.SimulationCases.Add()
    basis = case.BasisManager
    basis.StartBasisChange()
    try:
        fp = basis.FluidPackages.Add("KineticProbePkg")
        fp.PropertyPackageName = config.PROPERTY_PACKAGE_NAME
        for component_id in COMPONENTS:
            fp.Components.Add(component_id)
    finally:
        basis.EndBasisChange()
    report(f"  case={safe_get(case, 'Name', 'Case')!r}")
    report(f"  fluid package={safe_get(fp, 'Name', 'FP')!r}")

    feed = case.Flowsheet.MaterialStreams.Add("FEED")
    feed.Temperature.SetValue(config.c_to_k(500.0), "K")
    feed.Pressure.SetValue(config.bar_to_kpa(10.0), "kPa")
    feed.MolarFlow.SetValue(100.0, "kgmole/h")
    components = _component_names(fp)
    fractions = {
        "CO2": 0.5, "Hydrogen": 0.5,
        "Methane": 0.0, "H2O": 0.0, "CO": 0.0,
    }
    feed.ComponentMolarFractionValue = tuple(
        float(fractions.get(name, 0.0)) for name in components
    )
    _wait_solver(case)
    _read_stream(feed, components, "FEED")
    return case, fp, feed, components


# ==================== 反应创建 ====================
def _try_reaction_sets(manager: Any, type_names: list[str]) -> tuple[Any, str]:
    report("\n=== 3. ReactionSets.Add candidates ===")
    for index, type_string in enumerate(
        _candidate_set_types(type_names)[:60], start=1
    ):
        try:
            rset = manager.ReactionSets.Add(f"ProbeSet_{index}", type_string)
        except Exception as exc:  # noqa: BLE001
            report(f"  FAIL ReactionSets.Add({type_string!r}): {com_detail(exc)}")
            continue
        report(f"  OK   ReactionSets.Add({type_string!r}) -> {rset!r}")
        safe_get(rset, "Name", "rset")
        safe_get(rset, "TypeName", "rset")
        safe_get(rset, "VisibleTypeName", "rset")
        return rset, type_string
    return None, ""


def _dump_object_members(obj: Any, label: str, prefix: str) -> None:
    try:
        type_info = obj._oleobj_.GetTypeInfo()
        attr = type_info.GetTypeAttr()
    except Exception as exc:  # noqa: BLE001
        report(f"  {label} member dump failed: {com_detail(exc)}")
        return
    report(f"  --- {label} members (funcs={attr.cFuncs}, vars={attr.cVars}) ---")
    shown = 0
    for index in range(attr.cFuncs):
        try:
            func_desc = type_info.GetFuncDesc(index)
            names = type_info.GetNames(func_desc.memid)
            member_name = str(names[0]) if names else "?"
            if member_name in IGNORE_MEMBERS or not KEY_MEMBER_RE.search(member_name):
                continue
            try:
                member_doc = str(type_info.GetDocumentation(func_desc.memid)[1] or "")
            except Exception:
                member_doc = ""
            invkind = INVKIND.get(func_desc.invkind, str(func_desc.invkind))
            line = f"    {prefix}{member_name:40s} | {invkind:6s}"
            if member_doc:
                line += f" :: {member_doc}"
            report(line)
            shown += 1
        except Exception:
            continue
    if shown == 0:
        report("    (no key members matched)")


def _configure_kinetic_reaction(rxn: Any, components: list[str], label: str) -> None:
    report(f"\n--- configure {label} ---")
    _dump_object_members(rxn, label, "")
    try:
        for component_name in components:
            rxn.Reactants.Add(component_name)
        coefficients = {
            "CO2": -1.0, "Hydrogen": -1.0, "CO": 1.0, "H2O": 1.0,
            "Methane": 0.0,
        }
        rxn.ReactantStoichCoefValue = tuple(
            coefficients.get(name, 0.0) for name in components
        )
        report("  Reactants.Add + ReactantStoichCoefValue OK (RWGS probe values)")
    except Exception as exc:  # noqa: BLE001
        report(f"  Reactants/stoich failed: {com_detail(exc)}")

    # 只读取速率相关属性，避免在不了解单位时乱写整组参数
    for attr in (
        "ForwardRateConstant", "ForwardRateConstantValue",
        "ReverseRateConstant", "ReverseRateConstantValue",
        "ActivationEnergy", "ActivationEnergyValue",
        "PreExponentialFactor", "PreExponentialFactorValue",
        "FrequencyFactor", "FrequencyFactorValue",
        "ForwardOrderValue", "ReverseOrderValue",
        "ReactionPhase", "RateBasis", "ConcentrationBasis",
        "RateUnits", "Basis", "KineticBasis",
    ):
        try:
            value = getattr(rxn, attr)
            report(f"  {attr} = {value!r}")
            if hasattr(value, "_oleobj_"):
                _dump_object_members(value, f"{label}.{attr}", "  ")
        except Exception as exc:  # noqa: BLE001
            report(f"  {attr} GET failed: {com_detail(exc)}")

    # 尝试按常见属性名写入；成功与否都记录
    set_attempts = (
        ("ActivationEnergyValue", (50000.0,), ("kJ/kgmole", "J/mol", "kJ/kmol")),
        ("PreExponentialFactorValue", (1.0e8,), ()),
        ("FrequencyFactorValue", (1.0e8,), ()),
        ("ForwardRateConstantValue", (1.0e-3,), ()),
        ("ReactionPhase", (0,), ()),
    )
    for attr, values, units in set_attempts:
        try:
            original = getattr(rxn, attr)
        except Exception:
            continue
        for value in values:
            if hasattr(original, "SetValue") and units:
                for unit in units:
                    try:
                        original.SetValue(value, unit)
                        report(f"  SET {attr} = {value} [{unit}] OK")
                        break
                    except Exception as exc:  # noqa: BLE001
                        report(f"  SET {attr} = {value} [{unit}] failed: {com_detail(exc)}")
            else:
                try:
                    setattr(rxn, attr, value)
                    report(f"  SET {attr} = {value!r} OK")
                except Exception as exc:  # noqa: BLE001
                    report(f"  SET {attr} = {value!r} failed: {com_detail(exc)}")


def _try_kinetic_reactions(
    manager: Any,
    type_names: list[str],
    components: list[str],
    reaction_set: Any,
) -> Any:
    report("\n=== 4. Reactions.Add candidates ===")
    for index, type_string in enumerate(
        _candidate_reaction_types(type_names)[:80], start=1
    ):
        try:
            rxn = manager.Reactions.Add(f"ProbeRxn_{index}", type_string)
        except Exception as exc:  # noqa: BLE001
            report(f"  FAIL Reactions.Add({type_string!r}): {com_detail(exc)}")
            continue
        report(f"  OK   Reactions.Add({type_string!r}) -> {rxn!r}")
        safe_get(rxn, "Name", "rxn")
        safe_get(rxn, "TypeName", "rxn")
        safe_get(rxn, "VisibleTypeName", "rxn")
        _configure_kinetic_reaction(rxn, components, f"rxn[{type_string}]")

        if reaction_set is not None:
            try:
                reaction_set.ActiveReactions.Add(rxn)
                report("  ActiveReactions.Add OK")
            except Exception as exc:  # noqa: BLE001
                report(f"  ActiveReactions.Add failed: {com_detail(exc)}")
        return rxn
    return None


# ==================== CSTR/PFR ====================
def _try_set_process_value(obj: Any, attr: str, value: float, units: Iterable[str]) -> None:
    try:
        target = getattr(obj, attr)
    except Exception:
        return
    for unit in units:
        try:
            if hasattr(target, "SetValue"):
                target.SetValue(value, unit)
            else:
                setattr(obj, attr, value)
            report(f"  SET {attr} = {value} [{unit}] OK")
            return
        except Exception as exc:  # noqa: BLE001
            report(f"  SET {attr} = {value} [{unit}] failed: {com_detail(exc)}")


def _try_reactor(
    case: Any,
    feed: Any,
    type_string: str,
    index: int,
    components: list[str],
    reaction_set: Any,
) -> bool:
    name = f"ProbeReactor_{index}"
    report(f"\n--- Operations.Add({name!r}, {type_string!r}) ---")
    try:
        op = case.Flowsheet.Operations.Add(name, type_string)
    except Exception as exc:  # noqa: BLE001
        report(f"  FAIL: {com_detail(exc)}")
        return False
    report(f"  OK -> {op!r}")
    report(f"  TypeName={safe_get(op, 'TypeName', 'op')!r} "
           f"VisibleTypeName={safe_get(op, 'VisibleTypeName', 'op')!r}")
    _dump_object_members(op, name, "")

    vapour = case.Flowsheet.MaterialStreams.Add(f"VAP_{index}")
    liquid = case.Flowsheet.MaterialStreams.Add(f"LIQ_{index}")
    try:
        op.Feeds.Add(feed)
        report("  Feeds.Add(FEED) OK")
    except Exception as exc:  # noqa: BLE001
        report(f"  Feeds.Add failed: {com_detail(exc)}")
    for attr, stream in (("VapourProduct", vapour), ("LiquidProduct", liquid)):
        try:
            setattr(op, attr, stream)
            report(f"  {attr} OK")
        except Exception as exc:  # noqa: BLE001
            report(f"  {attr} failed: {com_detail(exc)}")
    if reaction_set is not None:
        try:
            op.ReactionSet = reaction_set
            report("  ReactionSet OK")
        except Exception as exc:  # noqa: BLE001
            report(f"  ReactionSet failed: {com_detail(exc)}")

    # 设计参数：只写常见成员；真实可用性由日志判断
    for attr, value, units in (
        ("Volume", 1.0, ("m3",)),
        ("VolumeValue", 1.0, ("m3",)),
        ("Length", 10.0, ("m",)),
        ("LengthValue", 10.0, ("m",)),
        ("Diameter", 0.5, ("m",)),
        ("DiameterValue", 0.5, ("m",)),
        ("Duty", 0.0, ("kJ/h",)),
        ("DutyValue", 0.0, ("kJ/h",)),
        ("OutletTemperature", config.c_to_k(500.0), ("K",)),
        ("OutletTemperatureValue", config.c_to_k(500.0), ("K",)),
        ("PressureDropValue", 0.0, ("kPa",)),
    ):
        _try_set_process_value(op, attr, value, units)

    _wait_solver(case)
    for stream, label in ((vapour, f"VAP_{index}"), (liquid, f"LIQ_{index}")):
        _read_stream(stream, components, label)
    return True


def try_reactors(
    case: Any,
    feed: Any,
    type_names: list[str],
    components: list[str],
    reaction_set: Any,
) -> None:
    report("\n=== 5. CSTR / PFR creation attempts ===")
    found_kinds: set[str] = set()
    succeeded = 0
    for index, type_string in enumerate(
        _candidate_operations(type_names)[:80], start=1
    ):
        lowered = type_string.lower()
        kind = (
            "pfr" if ("pfr" in lowered or "plug" in lowered)
            else "cstr" if "cstr" in lowered
            else ""
        )
        if kind and kind in found_kinds:
            continue
        if _try_reactor(
            case, feed, type_string, index, components, reaction_set
        ):
            succeeded += 1
            if kind:
                found_kinds.add(kind)
            if found_kinds == {"cstr", "pfr"} or succeeded >= 4:
                break
    report(f"\n  successful reactor type strings: {succeeded}; "
           f"kinds found: {sorted(found_kinds) or 'unknown'}")
    if not succeeded:
        report("  no CSTR/PFR operation type string succeeded")


# ==================== main ====================
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
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        try:
            app = win32com.client.GetActiveObject(PROGID)
            report("attached existing HYSYS")
        except Exception:
            app = win32com.client.Dispatch(PROGID)
            owns = True
            report("started new HYSYS")
        try:
            app.Visible = True
        except Exception:
            pass
        safe_get(app, "Version", "Application")
        safe_get(app, "Build", "Application")

        type_names = inspect_type_libraries(app)
        case, fp, feed, components = _setup_case(app)

        manager = case.BasisManager.ReactionPackageManager
        report("\n--- ReactionPackageManager members ---")
        _dump_object_members(manager, "ReactionPackageManager", "")
        reaction_set, set_type = _try_reaction_sets(manager, type_names)
        if reaction_set is not None:
            report(f"\nusing reaction set type={set_type!r}: {reaction_set!r}")
            try:
                reaction_set.AssociateFluidPackage(fp)
                fp.ReactionPackage.ReactionSets.Add(reaction_set)
                report("  reaction set associated to FP")
            except Exception as exc:  # noqa: BLE001
                report(f"  associate FP failed: {com_detail(exc)}")
        _try_kinetic_reactions(manager, type_names, components, reaction_set)

        try_reactors(case, feed, type_names, components, reaction_set)

        report("\n=== 6. summary ===")
        report("  type library inspection: see section 1")
        report("  If CSTR/PFR and kinetic reaction type strings appear above,")
        report("  send this report (and the full type dump if possible).")
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
            case = None
        if app is not None and owns:
            try:
                app.Quit()
                time.sleep(2)
            except Exception:
                pass
        app = None
        gc.collect()
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        REPORT_PATH.write_text("\n".join(_REPORT), encoding="utf-8")
        TYPE_DUMP_PATH.write_text(
            "\n".join(_FULL_TYPE_LINES + [""] + _FULL), encoding="utf-8"
        )
        print(f"\nreport: {REPORT_PATH}", flush=True)
        print(f"full type dump: {TYPE_DUMP_PATH}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
