# -*- coding: utf-8 -*-
"""
nlp_parser.py — 自然语言解析层（阶段1：规则占位）
================================================================
职责：
  1. 接收用户自然语言文本
  2. 提取进料组分、温度、压力、摩尔流量
  3. 按考核规则给出反应器类型“提示”（阶段1不创建反应器，不驱动 HYSYS）
  4. 输出结构化字典 + missing_info + notes

边界：
  - 不导入、不调用任何 HYSYS COM / win32com
  - 不依赖 pywin32、不访问网络
  - 只依赖 config（别名表、单位换算、反应器规则数据）

【本地可调试】本模块可在无 HYSYS、无 pywin32 的机器上直接运行。
阶段2接入 LLM 时，保持 parse() 输出 schema 不变，替换内部实现即可。
"""

from __future__ import annotations

import json
import re
from typing import TypedDict

import config


class ParsedRequest(TypedDict):
    """parse() 的固定输出结构（供 main / 后续 LLM 复用）。"""

    raw_text: str
    components: list[str]                 # HYSYS 标准组分 ID
    temperature_c: float | None           # 外部单位 ℃
    pressure_bar: float | None            # 外部单位 bar（MPa/kPa/atm 已换算）
    molar_flow_kmolh: float | None        # 外部单位 kmol/h
    reactor_type_hint: str | None         # 仅提示，阶段1不创建反应器
    reactor_reason: str | None
    missing_info: list[str]               # 缺失字段名列表
    notes: list[str]                      # 解析注意事项 / 歧义提示
    parser: str                           # 解析器版本标识


# ==================== 正则：数值与单位 ====================
# 温度：外部标准 ℃；支持 ℃ / °C / 摄氏度
_TEMP_RE = re.compile(
    r"(?P<value>-?\d+(?:\.\d+)?)\s*(?:℃|°\s*C|摄氏度)",
    re.IGNORECASE,
)

# 压力：支持 bar / kPa / MPa / atm，统一换算为 bar；不支持表压（barg）
_PRESSURE_RE = re.compile(
    r"(?P<value>-?\d+(?:\.\d+)?)\s*"
    r"(?P<unit>bar|巴|kpa|千帕|mpa|兆帕|atm|大气压)"
    r"(?![A-Za-z])",
    re.IGNORECASE,
)

# 摩尔流量：支持 kmol/h、kmol/hr、kmol·h-1、mol/h
_MOLAR_FLOW_RE = re.compile(
    r"(?P<value>-?\d+(?:\.\d+)?)\s*"
    r"(?P<unit>kmol\s*/\s*h(?:r)?|kmol\s*[·⋅]\s*h\s*-?\s*1|mol\s*/\s*h(?:r)?)",
    re.IGNORECASE,
)

# 质量流量：仅用于提示“暂不支持换算”，不参与数值计算
_MASS_FLOW_RE = re.compile(
    r"(?P<value>-?\d+(?:\.\d+)?)\s*"
    r"(?P<unit>kg\s*/\s*h(?:r)?|t\s*/\s*h(?:r)?|吨\s*/\s*小时|千克\s*/\s*小时)",
    re.IGNORECASE,
)

# 纯 ASCII 别名需要词边界，避免 CO 命中 CO2、H2 命中 H2O
_ASCII_ALIAS_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9 ._\-]*$")


def _build_component_patterns() -> tuple[tuple[str, str, re.Pattern[str]], ...]:
    """
    由 config.COMPONENT_ALIASES 构建匹配规则。
    按别名长度降序，保证“邻二甲苯”先于“甲苯”/“苯”、“H2O”先于“H2”。
    """
    entries: list[tuple[str, str, re.Pattern[str]]] = []
    for alias, component_id in config.COMPONENT_ALIASES.items():
        if _ASCII_ALIAS_RE.match(alias):
            pattern = re.compile(
                r"(?<![A-Za-z0-9])" + re.escape(alias) + r"(?![A-Za-z0-9])",
                re.IGNORECASE,
            )
        else:
            pattern = re.compile(re.escape(alias), re.IGNORECASE)
        entries.append((alias, component_id, pattern))

    entries.sort(key=lambda item: len(item[0]), reverse=True)
    return tuple(entries)


_COMPONENT_PATTERNS = _build_component_patterns()


# ==================== 各字段提取 ====================
def _extract_components(text: str) -> tuple[list[str], list[str]]:
    """提取 HYSYS 标准组分 ID，返回 (按出现顺序去重的 ID 列表, notes)。"""
    occupied = [False] * len(text)   # 防止“甲苯”内部再命中“苯”
    found: list[tuple[int, str]] = []
    seen: set[str] = set()

    for _alias, component_id, pattern in _COMPONENT_PATTERNS:
        for match in pattern.finditer(text):
            start, end = match.span()
            if end <= start or any(occupied[start:end]):
                continue
            for index in range(start, end):
                occupied[index] = True
            if component_id not in seen:
                seen.add(component_id)
                found.append((start, component_id))

    found.sort(key=lambda item: item[0])
    components = [component_id for _, component_id in found]

    notes: list[str] = []
    lower_text = text.lower()
    for alias, candidates in config.AMBIGUOUS_COMPONENT_ALIASES.items():
        if alias.lower() in lower_text and not any(item in seen for item in candidates):
            notes.append(
                f"检测到未指定异构体的“{alias}”，不能臆测为 {list(candidates)} 中的某一个；"
                "请明确写出邻/间/对二甲苯（o-/m-/p-Xylene）"
            )

    return components, notes


def _extract_temperature(text: str) -> tuple[float | None, list[str]]:
    """提取温度（℃）。多值时取第一个并给出提示。"""
    notes: list[str] = []
    values: list[tuple[int, float]] = []
    for match in _TEMP_RE.finditer(text):
        try:
            values.append((match.start(), float(match.group("value"))))
        except (TypeError, ValueError):
            continue

    if not values:
        return None, notes

    if len(values) > 1:
        shown = [value for _, value in values]
        notes.append(
            f"检测到多个温度值 {shown}（℃）；规则解析取第一个，"
            "阶段2需由 LLM 区分进料温度与出口/设备温度"
        )
    return values[0][1], notes


def _extract_pressure(text: str) -> tuple[float | None, list[str]]:
    """提取压力，统一换算为 bar。多值时取第一个并给出提示。"""
    notes: list[str] = []
    values: list[tuple[int, float]] = []
    for match in _PRESSURE_RE.finditer(text):
        raw_value = match.group("value")
        unit = match.group("unit")
        try:
            bar_value = config.pressure_to_bar(float(raw_value), unit)
        except (TypeError, ValueError) as exc:
            notes.append(str(exc))
            continue
        values.append((match.start(), bar_value))

    if not values:
        return None, notes

    if len(values) > 1:
        shown = [round(value, 6) for _, value in values]
        notes.append(
            f"检测到多个压力值 {shown}（bar）；规则解析取第一个，"
            "阶段2需由 LLM 区分进料压力与设备/出口压力"
        )
    return values[0][1], notes


def _extract_molar_flow(text: str) -> tuple[float | None, list[str]]:
    """提取摩尔流量并统一转换为 kmol/h；质量流量只提示、不猜测换算。"""
    notes: list[str] = []
    values: list[tuple[int, float]] = []

    for match in _MOLAR_FLOW_RE.finditer(text):
        raw_value = match.group("value")
        unit = match.group("unit")
        try:
            flow = config.molar_flow_to_kmolh(float(raw_value), unit)
        except (TypeError, ValueError) as exc:
            notes.append(str(exc))
            continue
        values.append((match.start(), flow))

    mass_matches = list(_MASS_FLOW_RE.finditer(text))
    if mass_matches and not values:
        match = mass_matches[0]
        notes.append(
            f"检测到质量流量 {match.group('value')} {match.group('unit')}；"
            "阶段1规则解析不把 kg/h 换算为 kmol/h（需要分子量与组成），"
            "因此 molar_flow_kmolh 缺失，阶段2可结合物性数据由 LLM 完成"
        )
    elif mass_matches and values:
        notes.append("同时检测到质量流量与摩尔流量；阶段1采用摩尔流量")

    if not values:
        return None, notes

    if len(values) > 1:
        shown = [round(value, 6) for _, value in values]
        notes.append(f"检测到多个摩尔流量值 {shown}（kmol/h）；规则解析取第一个")
    return values[0][1], notes


def _contains_keyword(text: str, keyword: str) -> bool:
    """ASCII 关键词用词边界，中文关键词用子串匹配。"""
    if re.fullmatch(r"[A-Za-z0-9]+", keyword):
        pattern = r"(?<![A-Za-z0-9])" + re.escape(keyword) + r"(?![A-Za-z0-9])"
        return re.search(pattern, text, re.IGNORECASE) is not None
    return keyword in text


def _infer_reactor_hint(text: str) -> tuple[str | None, str | None]:
    """
    按 config.REACTOR_RULES 的顺序做关键词匹配，返回反应器类型“提示”。
    仅为阶段1占位：不创建反应器、不调用 COM；阶段2将由 LLM 完成判断。
    """
    for rule in config.REACTOR_RULES:
        for keyword in rule["keywords"]:
            if _contains_keyword(text, keyword):
                reason = f"{rule['reason']}（命中关键词：{keyword}）"
                return rule["reactor_type"], reason
    return None, None


# ==================== 对外唯一入口 ====================
def parse(text: str) -> ParsedRequest:
    """
    解析自然语言建模需求，返回结构化字典。

    注意：
      - 输出单位固定为外部标准：℃ / bar / kmol·h⁻¹
      - 温度/压力/摩尔流量缺失时对应值为 None，并写入 missing_info
      - 反应器类型只是提示字段，阶段1的 main 不会据此创建反应器
    """
    if not isinstance(text, str) or not text.strip():
        raise ValueError("输入文本为空，无法解析")

    raw_text = text.strip()
    components, component_notes = _extract_components(raw_text)
    temperature_c, temperature_notes = _extract_temperature(raw_text)
    pressure_bar, pressure_notes = _extract_pressure(raw_text)
    molar_flow_kmolh, flow_notes = _extract_molar_flow(raw_text)
    reactor_type_hint, reactor_reason = _infer_reactor_hint(raw_text)

    missing_info: list[str] = []
    if not components:
        missing_info.append("components")
    if temperature_c is None:
        missing_info.append("temperature_c")
    if pressure_bar is None:
        missing_info.append("pressure_bar")
    if molar_flow_kmolh is None:
        missing_info.append("molar_flow_kmolh")

    notes = component_notes + temperature_notes + pressure_notes + flow_notes
    if reactor_type_hint:
        notes.append("reactor_type_hint 仅为规则占位提示，阶段1不会创建反应器")

    return ParsedRequest(
        raw_text=raw_text,
        components=components,
        temperature_c=temperature_c,
        pressure_bar=pressure_bar,
        molar_flow_kmolh=molar_flow_kmolh,
        reactor_type_hint=reactor_type_hint,
        reactor_reason=reactor_reason,
        missing_info=missing_info,
        notes=notes,
        parser="rule_placeholder_v1",
    )


# ==================== 本地自检（不涉及 HYSYS） ====================
if __name__ == "__main__":
    # Windows 控制台默认可能是 GBK，避免中文 JSON 乱码
    try:
        import sys

        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    sample = (
        "甲苯进料流量 10000 kg/h，进料温度 380℃，压力 2.5MPa；"
        "产物含苯、邻二甲苯、间二甲苯、对二甲苯。"
    )
    print(json.dumps(parse(sample), ensure_ascii=False, indent=2))
