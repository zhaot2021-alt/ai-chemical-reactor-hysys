# -*- coding: utf-8 -*-
"""
config.py — 全局配置层（纯配置，无业务逻辑、无 COM 调用）
================================================================
职责：
  1. 项目路径、日志、HYSYS ProgID、COM 生命周期参数
  2. HYSYS 标准组分 ID 与自然语言别名映射
  3. 单位换算函数（外部 ℃ / bar / kmol·h⁻¹ ↔ HYSYS 内部 K / kPa / kgmole·h⁻¹）
  4. LLM 配置占位（阶段1不使用，后续接入）
  5. 反应器选择规则配置（纯数据，阶段2由 nlp_parser / LLM 消费）

【本地可调试】本文件只依赖 Python 标准库，无 HYSYS / pywin32 也能导入和测试。
禁止在此文件写：COM 对象操作、文本解析、反应器创建、文件保存等业务逻辑。
"""

from __future__ import annotations

import os
from pathlib import Path

# ==================== 项目路径 ====================
SRC_DIR = Path(__file__).resolve().parent          # .../chem_ai_model/src
PROJECT_ROOT = SRC_DIR.parent                      # .../chem_ai_model
OUTPUT_DIR = PROJECT_ROOT / "output"               # .hsc 默认输出目录

# ==================== HYSYS COM 连接 ====================
HYSYS_PROGID = "HYSYS.Application"   # 若连接失败，请核对注册表 HKEY_CLASSES_ROOT 下的 ProgID
HYSYS_CONNECT_RETRIES = 3            # Dispatch 启动 hysys.exe 的重试次数
HYSYS_RETRY_INTERVAL_S = 3.0         # 每次重试间隔（秒）
HYSYS_APP_VISIBLE = True             # 是否显示 HYSYS 主窗口
HYSYS_ATTACH_EXISTING = True         # 优先依附已运行的 HYSYS，避免重复占用许可证
HYSYS_QUIT_ATTACHED = False          # 依附已有实例时，默认不 Quit（防止误杀用户会话）
HYSYS_SHUTDOWN_WAIT_S = 2.0          # Quit 后等待进程退出再检查
HYSYS_TASKLIST_TIMEOUT_S = 15.0      # tasklist 检查超时

# ==================== HYSYS 内部单位（驱动层只接受这些单位） ====================
HYSYS_TEMPERATURE_UNIT = "K"
HYSYS_PRESSURE_UNIT = "kPa"
HYSYS_MOLAR_FLOW_UNIT = "kgmole/h"   # HYSYS 的 kgmole ≡ kmol，1 kgmole = 1000 mol

# ==================== 流体包 / 物性包 ====================
# 目标流体包名称。说明：FluidPackages.Add() 公开用法不带名称参数，
# 驱动层会 best-effort 设置 FluidPackage.Name；若该属性不可写，HYSYS 将使用默认名（如 Basis-1）。
FLUID_PACKAGE_NAME = "PR_FluidPkg"

# HYSYS V15 实测（2026-10-26，手动 PR Case + tools/read_open_fluid_packages.py）：
#   FluidPackage.PropertyPackageName (读) = 'Peng-Robinson'   ← 显示名
#   FluidPackage.PropertyPackage.TypeName   = 'pengrob'       ← 内部 token
#   PropertyPackageName = 'pengrob' 才是 V15 Setter 接受的内部名；
#   'PengRobinson' / 'Peng-Robinson' / 'PR' / 'HysysPR' 均被拒绝。
# 另：FluidPackages.Add('名称') 可创建时直接命名；组分必须在 EndBasisChange 前加入。
PROPERTY_PACKAGE_NAME = "pengrob"                   # Setter 内部 token
PROPERTY_PACKAGE_DISPLAY_NAME = "Peng-Robinson"     # 读回/报告用显示名
PROPERTY_PACKAGE_ALIASES = {                        # 兼容调用方传入显示名
    "pengrob": "pengrob",
    "PengRobinson": "pengrob",
    "Peng-Robinson": "pengrob",
}

# ==================== HYSYS 标准组分 ID ====================
# 注意：必须使用 HYSYS 组分库中的标准 Simulation Name：
#   STEAM 不在标准 ID 中 → 用 H2O（HYSYS 中 Water 是 H2O 的同义名）
#   H2 → Hydrogen；XYLENE-O/M/P → o-Xylene / m-Xylene / p-Xylene
STANDARD_COMPONENT_IDS = (
    "Methane",
    "H2O",
    "CO",
    "CO2",
    "Hydrogen",
    "Toluene",
    "Benzene",
    "o-Xylene",
    "m-Xylene",
    "p-Xylene",
)

# 自然语言别名 → HYSYS 标准组分 ID（只做映射，不做解析逻辑）
COMPONENT_ALIASES = {
    # 甲烷
    "甲烷": "Methane",
    "methane": "Methane",
    "ch4": "Methane",
    # 水 / 蒸汽（HYSYS 标准 Sim Name 为 H2O）
    "水蒸气": "H2O",
    "水蒸汽": "H2O",
    "蒸汽": "H2O",
    "水": "H2O",
    "steam": "H2O",
    "water": "H2O",
    "h2o": "H2O",
    # 合成气
    "一氧化碳": "CO",
    "carbon monoxide": "CO",
    "co": "CO",
    "二氧化碳": "CO2",
    "carbon dioxide": "CO2",
    "co2": "CO2",
    "氢气": "Hydrogen",
    "hydrogen": "Hydrogen",
    "h2": "Hydrogen",
    # 芳烃
    "甲苯": "Toluene",
    "toluene": "Toluene",
    "苯": "Benzene",
    "benzene": "Benzene",
    "邻二甲苯": "o-Xylene",
    "o-xylene": "o-Xylene",
    "间二甲苯": "m-Xylene",
    "m-xylene": "m-Xylene",
    "对二甲苯": "p-Xylene",
    "p-xylene": "p-Xylene",
}

# 语义有歧义、不能擅自展开为某个标准 ID 的别名。
# nlp_parser 检测到它们时只写 notes，不猜测组分。
AMBIGUOUS_COMPONENT_ALIASES = {
    "二甲苯": ("o-Xylene", "m-Xylene", "p-Xylene"),
    "xylene": ("o-Xylene", "m-Xylene", "p-Xylene"),
}

# 组分添加失败时可按顺序尝试的“同义名 / 全名”（非标准 Sim Name）。
# 仅 H2O 存在官方同义名 Water；标准 ID 仍以 H2O 为准。
COMPONENT_ID_FALLBACKS = {
    "H2O": ("Water",),
}

# ==================== 阶段1 默认值 ====================
CASE_NAME_PREFIX = "ChemAI_Stage1"
FEED_STREAM_NAME = "FEED"
HSC_EXTENSION = ".hsc"

# 外部输入单位：温度 ℃，压力 bar，摩尔流量 kmol·h⁻¹
DEFAULT_FEED_CONDITIONS = {
    "temperature_c": 25.0,
    "pressure_bar": 1.0,
    "molar_flow_kmolh": 100.0,
}

# ==================== 单位换算（统一收口在本文件） ====================
KELVIN_OFFSET = 273.15
BAR_TO_KPA = 100.0
MPA_TO_BAR = 10.0
ATM_TO_BAR = 1.01325
KMOL_PER_KGMOLE = 1.0   # 1 kmol = 1 kgmole


def c_to_k(celsius: float) -> float:
    """摄氏度 → 开尔文。"""
    return float(celsius) + KELVIN_OFFSET


def k_to_c(kelvin: float) -> float:
    """开尔文 → 摄氏度。"""
    return float(kelvin) - KELVIN_OFFSET


def bar_to_kpa(bar: float) -> float:
    """bar → kPa（HYSYS 内部压力单位）。"""
    return float(bar) * BAR_TO_KPA


def kpa_to_bar(kpa: float) -> float:
    """kPa → bar。"""
    return float(kpa) / BAR_TO_KPA


def mpa_to_bar(mpa: float) -> float:
    """MPa → bar。"""
    return float(mpa) * MPA_TO_BAR


def atm_to_bar(atm: float) -> float:
    """标准大气压 → bar。"""
    return float(atm) * ATM_TO_BAR


def kmolh_to_kgmoleh(kmol_h: float) -> float:
    """kmol/h → HYSYS kgmole/h（数值相同，显式转换以表达单位语义）。"""
    return float(kmol_h) * KMOL_PER_KGMOLE


def kgmoleh_to_kmolh(kgmole_h: float) -> float:
    """HYSYS kgmole/h → kmol/h。"""
    return float(kgmole_h) / KMOL_PER_KGMOLE


def _normalize_unit_token(unit: str) -> str:
    """统一单位字符串：小写、去空格、统一乘点/上标负号。"""
    token = str(unit).strip().lower().replace(" ", "")
    token = token.replace("⋅", "·").replace("⁻¹", "-1").replace("−", "-")
    return token


# 解析层用于把用户文本中的多种压力单位统一为 bar
_PRESSURE_UNIT_TO_BAR = {
    "bar": 1.0,
    "巴": 1.0,
    "kpa": 1.0 / 100.0,
    "千帕": 1.0 / 100.0,
    "mpa": 10.0,
    "兆帕": 10.0,
    "atm": ATM_TO_BAR,
    "大气压": ATM_TO_BAR,
}

# 解析层支持的摩尔流量单位 → kmol/h
_MOLAR_FLOW_UNIT_TO_KMOLH = {
    "kmol/h": 1.0,
    "kmol/hr": 1.0,
    "kmol·h-1": 1.0,
    "kmolh-1": 1.0,
    "mol/h": 1.0 / 1000.0,
    "mol/hr": 1.0 / 1000.0,
}


def pressure_to_bar(value: float, unit: str) -> float:
    """任意受支持压力单位 → bar（外部标准单位）。"""
    token = _normalize_unit_token(unit)
    if token not in _PRESSURE_UNIT_TO_BAR:
        raise ValueError(
            f"不支持的压力单位: {unit!r}；支持: {sorted(_PRESSURE_UNIT_TO_BAR)}"
        )
    return float(value) * _PRESSURE_UNIT_TO_BAR[token]


def molar_flow_to_kmolh(value: float, unit: str) -> float:
    """受支持的摩尔流量单位 → kmol/h（外部标准单位）。"""
    token = _normalize_unit_token(unit)
    if token not in _MOLAR_FLOW_UNIT_TO_KMOLH:
        raise ValueError(
            f"不支持的摩尔流量单位: {unit!r}；支持: {sorted(_MOLAR_FLOW_UNIT_TO_KMOLH)}"
        )
    return float(value) * _MOLAR_FLOW_UNIT_TO_KMOLH[token]


# ==================== LLM 配置（阶段1仅占位，不发起任何请求） ====================
LLM_CONFIG = {
    "enabled": False,
    "provider": "openai-compatible",                    # 兼容 OpenAI 协议的任意服务
    "base_url": os.getenv("CHEM_AI_LLM_BASE_URL", ""),  # 从环境变量读取，避免写死
    "api_key_env": "CHEM_AI_LLM_API_KEY",               # 只保存环境变量名，不保存密钥
    "model": os.getenv("CHEM_AI_LLM_MODEL", ""),
    "temperature": 0.0,
    "timeout_s": 60,
}

# ==================== 反应器选择规则（纯数据，阶段2使用） ====================
# 规则来自考核说明《选择逻辑》：
#   1) 有动力学参数 → 动力学反应器（液相 CSTR / 气相 PFR）
#   2) 无动力学但给出转化率/收率/选择性 → Conversion
#   3) 提供 Ka / 吉布斯自由能等平衡信息 → Equilibrium
#   4) 产物分布未知、复杂高温体系 → Gibbs
# 注意：阶段1只在 nlp_parser 中输出提示，不创建任何反应器，驱动层无反应器代码。
REACTOR_RULES = (
    {
        "reactor_type": "kinetic_CSTR_or_PFR",
        "label": "动力学反应器（CSTR / PFR，按相态选择）",
        "keywords": ("动力学参数", "反应速率常数", "速率常数", "活化能",
                     "CSTR", "PFR", "釜式反应器", "管式反应器"),
        "reason": "用户提供动力学或设备参数；阶段2需按相态选 CSTR（液相）或 PFR（气相）",
    },
    {
        "reactor_type": "conversion",
        "label": "Conversion 转化率反应器",
        "keywords": ("转化率", "收率", "选择性"),
        "reason": "未提供动力学，但给出了转化率/收率/选择性等工艺指标",
    },
    {
        "reactor_type": "equilibrium",
        "label": "Equilibrium 平衡反应器",
        "keywords": ("平衡常数", "平衡数据", "Ka"),
        "reason": "提供平衡常数 / 平衡数据，可按化学平衡计算反应终点",
    },
    {
        "reactor_type": "gibbs",
        "label": "Gibbs 吉布斯反应器",
        "keywords": ("产物分布未知", "吉布斯反应器", "吉布斯自由能",
                     "自由能最小化", "燃烧", "无需定义反应"),
        "reason": "产物分布未知或复杂高温体系，适合吉布斯自由能最小化",
    },
)

# ==================== 日志 ====================
LOG_LEVEL = "INFO"
LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(name)s | %(message)s"
LOG_DATEFMT = "%Y-%m-%d %H:%M:%S"
