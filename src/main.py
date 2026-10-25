# -*- coding: utf-8 -*-
"""
main.py — 程序入口 / 调度层
================================================================
职责：
  1. 接收用户输入文本
  2. 调用 nlp_parser.parse() 得到结构化参数
  3. 处理缺失参数（strict 模式报错；非 strict 模式用 config 默认值并告警）
  4. 用 config 中的单位换算函数把外部单位转成 HYSYS 内部单位
  5. 调用 HYSYSDriver 完成阶段1建模并返回摘要
  6. 提供 CLI 与 dry-run

边界：
  - 不写底层 COM 细节（全部在 hysys_driver）
  - 不写解析规则（全部在 nlp_parser + config）
  - 不创建反应器（阶段1只建 Case / 流体包 / 组分 / 物流）

【本地可调试】parse / 参数校验 / dry-run 不导入 hysys_driver、不连接 HYSYS。
【仅远程可运行】run_stage1() 中标注的 HYSYS 驱动调用需要 HYSYS V15 + COM + 许可证。
"""

from __future__ import annotations

import argparse
import json
import logging
import math
import re
import sys
import time
from pathlib import Path
from typing import Any

import config
import nlp_parser


class Stage1Error(RuntimeError):
    """调度层错误：参数缺失、数值非法、驱动不可用等。"""


logger = logging.getLogger("main")

_FEED_LABELS = {
    "temperature_c": "进料温度(℃)",
    "pressure_bar": "进料压力(bar)",
    "molar_flow_kmolh": "摩尔流量(kmol/h)",
}


# ==================== 基础设施 ====================
def _configure_console() -> None:
    """Windows 控制台默认可能是 GBK；统一输出为 UTF-8，避免中文/argparse 帮助乱码。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001 - 非 TTY / 旧环境忽略
            pass


def setup_logging(level: str | None = None) -> None:
    """配置控制台日志（幂等：重复调用不会叠加 handler）。"""
    _configure_console()

    root = logging.getLogger()
    if not root.handlers:
        handler = logging.StreamHandler(sys.stdout)
        handler.setFormatter(
            logging.Formatter(config.LOG_FORMAT, config.LOG_DATEFMT)
        )
        root.addHandler(handler)

    level_name = (level or config.LOG_LEVEL).upper()
    root.setLevel(getattr(logging, level_name, logging.INFO))


def _load_driver_class() -> type:
    """
    延迟导入驱动层：
      - 保证 --dry-run / parser 自测在无 pywin32、无 HYSYS 的机器上可运行
      - 真正需要 COM 时才引入 win32com
    """
    try:
        from hysys_driver import HYSYSDriver
    except Exception as exc:  # noqa: BLE001
        raise Stage1Error(
            "无法导入 HYSYS 驱动层（hysys_driver）。请确认："
            "1) 当前为 Windows 且已安装 pywin32 311；"
            "2) 在 src 目录下运行脚本；"
            f"3) HYSYS V15 安装完整。原始错误：{type(exc).__name__}: {exc}"
        ) from exc
    return HYSYSDriver


def _safe_filename(name: str) -> str:
    """把 Case 名称变成 Windows 合法文件名。"""
    sanitized = re.sub(r'[\\/:*?"<>|]+', "_", name).strip().strip(".")
    return sanitized or "case"


def _normalize_save_path(path: str | Path) -> Path:
    """转绝对路径并确保 .hsc 后缀。"""
    target = Path(path).expanduser()
    if not target.is_absolute():
        target = (Path.cwd() / target).resolve()
    if target.suffix.lower() != config.HSC_EXTENSION:
        target = target.with_suffix(config.HSC_EXTENSION)
    return target


# ==================== 参数处理（本地逻辑） ====================
def _resolve_feed(
    parsed: nlp_parser.ParsedRequest,
    strict: bool,
) -> tuple[dict[str, float], list[str]]:
    """
    校验解析结果并补默认值。
      - 未识别到组分：一律报错（组分是建流体包的前提，不能臆测）
      - strict=True：任何 missing_info 都报错
      - strict=False：温度/压力/摩尔流量用 config 默认值，并记录 defaults_used
    返回 (外部单位进料字典, 使用默认值的字段列表)。
    """
    if not parsed["components"]:
        raise Stage1Error(
            "未识别到任何进料组分（missing_info 含 components）。"
            "请明确写出组分名称/化学式，例如“甲烷、水蒸气、一氧化碳、二氧化碳、氢气”。"
        )

    if strict and parsed["missing_info"]:
        labels = [
            _FEED_LABELS.get(field, field)
            for field in parsed["missing_info"]
        ]
        raise Stage1Error(
            "strict 模式：参数不完整，缺少 "
            + "、".join(labels)
            + "。请补充后重试，或去掉 --strict 使用默认值做连通性测试。"
        )

    feed: dict[str, float] = {}
    defaults_used: list[str] = []

    for field, default_value in config.DEFAULT_FEED_CONDITIONS.items():
        value = parsed.get(field)  # type: ignore[typeddict-item]
        if value is None:
            value = float(default_value)
            defaults_used.append(field)
            logger.warning(
                "缺少 %s，阶段1使用 config 默认值 %s（仅用于连通性测试，不代表工艺设计）",
                _FEED_LABELS.get(field, field),
                default_value,
            )
        feed[field] = float(value)

    for field, label in _FEED_LABELS.items():
        if not math.isfinite(feed[field]):
            raise Stage1Error(f"{label} 必须是有限数值，收到 {feed[field]!r}")

    if config.c_to_k(feed["temperature_c"]) <= 0:
        raise Stage1Error(
            f"进料温度必须高于绝对零度（-273.15℃），收到 {feed['temperature_c']}℃"
        )
    if feed["pressure_bar"] <= 0:
        raise Stage1Error(f"进料压力必须为正，收到 {feed['pressure_bar']} bar")
    if feed["molar_flow_kmolh"] <= 0:
        raise Stage1Error(
            f"摩尔流量必须为正，收到 {feed['molar_flow_kmolh']} kmol/h"
        )

    return feed, defaults_used


def _to_internal_units(feed: dict[str, float]) -> dict[str, float]:
    """外部 ℃ / bar / kmol·h⁻¹ → HYSYS K / kPa / kgmole·h⁻¹（换算函数全部来自 config）。"""
    temperature_k = config.c_to_k(feed["temperature_c"])
    pressure_kpa = config.bar_to_kpa(feed["pressure_bar"])
    molar_flow_kgmoleh = config.kmolh_to_kgmoleh(feed["molar_flow_kmolh"])

    logger.info(
        "单位换算：%.4f ℃ → %.4f K；%.6f bar → %.4f kPa；"
        "%.4f kmol/h → %.4f kgmole/h",
        feed["temperature_c"],
        temperature_k,
        feed["pressure_bar"],
        pressure_kpa,
        feed["molar_flow_kmolh"],
        molar_flow_kgmoleh,
    )

    return {
        "temperature_k": temperature_k,
        "pressure_kpa": pressure_kpa,
        "molar_flow_kgmoleh": molar_flow_kgmoleh,
    }


# ==================== 阶段1主流程 ====================
def run_stage1(
    text: str,
    *,
    case_name: str | None = None,
    save_path: str | Path | None = None,
    visible: bool | None = None,
    strict: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    """
    阶段1完整链路：解析 → 校验/补默认值 → 单位换算 → HYSYS 建模 → 保存。

    dry_run=True 时只走到“驱动调用计划”，不导入 hysys_driver、不连接 HYSYS，
    可在无 HYSYS 的本地机器执行。
    """
    parsed = nlp_parser.parse(text)
    logger.info(
        "解析结果：%s", json.dumps(parsed, ensure_ascii=False, sort_keys=False)
    )
    if parsed["missing_info"]:
        logger.warning("缺失字段：%s", parsed["missing_info"])
    for note in parsed["notes"]:
        logger.info("解析提示：%s", note)
    if parsed["reactor_type_hint"]:
        logger.info(
            "反应器类型提示（阶段1不创建反应器）：%s；理由：%s",
            parsed["reactor_type_hint"],
            parsed["reactor_reason"],
        )

    feed, defaults_used = _resolve_feed(parsed, strict)
    internal = _to_internal_units(feed)

    final_case_name = (
        case_name
        or f"{config.CASE_NAME_PREFIX}_{time.strftime('%Y%m%d_%H%M%S')}"
    ).strip()
    if not final_case_name:
        raise Stage1Error("Case 名称不能为空")

    if save_path is None:
        final_save_path = (
            config.OUTPUT_DIR
            / f"{_safe_filename(final_case_name)}{config.HSC_EXTENSION}"
        )
    else:
        final_save_path = _normalize_save_path(save_path)

    # ---------- 本地可调试：dry-run 到此处返回 ----------
    if dry_run:
        summary: dict[str, Any] = {
            "mode": "dry-run",
            "stage": "stage1",
            "parsed": parsed,
            "feed_external": feed,
            "feed_internal": internal,
            "defaults_used": defaults_used,
            "driver_plan": {
                "case_name": final_case_name,
                "fluid_package": config.FLUID_PACKAGE_NAME,
                "property_package": config.PROPERTY_PACKAGE_NAME,
                "property_package_display": config.PROPERTY_PACKAGE_DISPLAY_NAME,
                "components": parsed["components"],
                "feed_stream": {
                    "name": config.FEED_STREAM_NAME,
                    **internal,
                },
                "save_path": str(final_save_path),
            },
            "note": (
                "dry-run 未导入 hysys_driver、未连接 HYSYS，"
                "可在无 HYSYS 的本地机器执行"
            ),
        }
        logger.info("dry-run 完成：未连接 HYSYS")
        return summary

    # ---------- 【仅远程可运行】以下需要 HYSYS V15 + pywin32 + 许可证 ----------
    HYSYSDriver = _load_driver_class()
    driver = HYSYSDriver(visible=visible)
    live_summary: dict[str, Any] | None = None

    try:
        driver.connect()
        actual_case_name = driver.new_case(final_case_name)
        # 组分必须在 create_fluid_package 的同一个 Basis 变更中加完；
        # 否则 EndBasisChange 会因为 Basis 为空而失败。
        fluid_package_name = driver.create_fluid_package(
            config.FLUID_PACKAGE_NAME,
            property_package=config.PROPERTY_PACKAGE_NAME,
            components=parsed["components"],
        )
        added_components = driver.component_ids
        stream_state = driver.create_material_stream(
            config.FEED_STREAM_NAME,
            temperature_k=internal["temperature_k"],
            pressure_kpa=internal["pressure_kpa"],
            molar_flow_kgmoleh=internal["molar_flow_kgmoleh"],
        )
        # 阶段2 扩展位：在这里依据 reactor_spec 分发表调用
        # driver.create_conversion_reactor / create_equilibrium_reactor / create_gibbs_reactor。
        # 阶段1 明确不创建反应器，上面的 reactor_type_hint 只做展示。
        saved_path = driver.save_case(final_save_path)

        live_summary = {
            "mode": "live",
            "stage": "stage1",
            "case_name": actual_case_name,
            "fluid_package_name": fluid_package_name,
            "property_package_name": (
                driver.property_package_name
                or config.PROPERTY_PACKAGE_DISPLAY_NAME
            ),
            "components_added": list(added_components),
            "feed_stream": stream_state,
            "hsc_path": saved_path,
            "feed_external": feed,
            "feed_internal": internal,
            "defaults_used": defaults_used,
            "missing_info": parsed["missing_info"],
            "notes": parsed["notes"],
            "reactor_type_hint": parsed["reactor_type_hint"],
            "limitations": [
                "阶段1仅验证连通性：未设置物流组成、未求解",
                "未创建任何反应器（Conversion / Equilibrium / Gibbs 留给阶段2）",
            ],
        }
        logger.info("阶段1执行完成：Case=%s，文件=%s", actual_case_name, saved_path)
    finally:
        # 无论成功、异常，都显式释放 COM，避免 hysys.exe 残留
        driver.shutdown()

    if live_summary is None:  # 理论上不可达
        raise Stage1Error("阶段1未产生结果摘要")
    return live_summary


# ==================== CLI ====================
def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="AI 化工反应器建模 — 阶段1：HYSYS 连通性与基础建模",
    )
    parser.add_argument(
        "--text",
        type=str,
        default=None,
        help="自然语言建模需求；省略则进入交互输入",
    )
    parser.add_argument("--case-name", type=str, default=None, help="Case 名称")
    parser.add_argument(
        "--save",
        type=str,
        default=None,
        help=".hsc 保存路径（默认 output/<CaseName>.hsc）",
    )
    parser.add_argument(
        "--visible",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="是否显示 HYSYS 窗口（--visible / --no-visible）",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="参数缺失时直接报错，不使用 config 默认值",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只解析和校验，不连接 HYSYS（本地可运行）",
    )
    parser.add_argument(
        "--log-level",
        choices=("DEBUG", "INFO", "WARNING", "ERROR"),
        default=config.LOG_LEVEL,
        help="日志级别",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    _configure_console()  # 必须先于 argparse：否则 --help 中文在 GBK 控制台乱码
    parser = _build_arg_parser()
    args = parser.parse_args(argv)
    setup_logging(args.log_level)

    text = args.text
    if not text:
        print("请输入自然语言建模需求（示例：进料为甲烷和水蒸气，"
              "温度 520℃，压力 1 bar，摩尔流量 100 kmol/h），回车结束：")
        try:
            text = input("> ").strip()
        except EOFError:
            text = ""

    if not text:
        logger.error("未提供输入文本")
        return 2

    try:
        summary = run_stage1(
            text,
            case_name=args.case_name,
            save_path=args.save,
            visible=args.visible,
            strict=args.strict,
            dry_run=args.dry_run,
        )
    except Stage1Error as exc:
        logger.error("阶段1失败：%s", exc)
        return 1
    except Exception as exc:  # noqa: BLE001 - CLI 顶层兜底，保留 traceback 便于排查 COM 问题
        logger.exception("执行失败：%s", exc)
        return 1

    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
