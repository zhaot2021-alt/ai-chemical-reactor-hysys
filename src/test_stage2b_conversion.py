# -*- coding: utf-8 -*-
"""
test_stage2b_conversion.py — 甲苯歧化 Conversion 反应器（模板驱动）
================================================================
使用用户手动建好的结构模板 output/ref_conversion.hsc：
  - 5 个组分、Conversion 反应器 CRV-100、反应集 Set-1（50%）
  - Feed1 / Prod1 / Liq1 已连接

运行时用 COM 覆写为真实工况：
  380℃ / 2.5 MPa / 10000 kg/h，纯甲苯进料
然后求解并读回产物组成，保存结果 .hsc。

不修改模板中的反应/转化率；只改进料条件和组成。
运行：
    python src\test_stage2b_conversion.py --visible --log-level DEBUG
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path

import config

logger = logging.getLogger("test_stage2b_conversion")

TEMPLATE_PATH = config.OUTPUT_DIR / "ref_conversion.hsc"
FEED_STREAM = "Feed1"
VAPOUR_STREAM = "Prod1"
LIQUID_STREAM = "Liq1"

FEED_T_K = 653.15            # 380℃
FEED_P_KPA = 2500.0          # 2.5 MPa
FEED_MASS_KG_H = 10000.0
TOLUENE_MW_KG_KMOL = 92.1402  # 纯甲苯进料，kg/h -> kmol/h
FEED_MOLAR_KMOLH = FEED_MASS_KG_H / TOLUENE_MW_KG_KMOL
FEED_COMPOSITION = {"Toluene": 1.0}


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    parser = argparse.ArgumentParser(description="阶段2B Conversion 模板联调")
    parser.add_argument(
        "--visible",
        action=argparse.BooleanOptionalAction,
        default=None,
    )
    parser.add_argument(
        "--log-level",
        choices=("DEBUG", "INFO", "WARNING", "ERROR"),
        default=config.LOG_LEVEL,
    )
    parser.add_argument(
        "--template",
        default=str(TEMPLATE_PATH),
        help="ref_conversion.hsc 路径",
    )
    args = parser.parse_args(argv)
    logging.basicConfig(
        level=getattr(logging, args.log_level, logging.INFO),
        format=config.LOG_FORMAT,
        datefmt=config.LOG_DATEFMT,
    )

    template = Path(args.template)
    if not template.exists():
        logger.error(
            "找不到模板 %s；请先在 HYSYS 里把 Conversion 参考 Case 另存为此路径",
            template,
        )
        return 1

    try:
        from hysys_driver import HYSYSDriver
    except Exception as exc:  # noqa: BLE001
        logger.error("无法导入 hysys_driver：%s", exc)
        return 1

    driver = HYSYSDriver(visible=args.visible)
    try:
        driver.connect()
        driver.open_case(template)
        logger.info("模板物流：%s", driver.list_streams())
        logger.info("模板操作：%s", driver.list_operations())

        # 模板的流量规格是 MolarFlow；MassFlow.SetValue 在 V15 返回 E_ACCESSDENIED。
        driver.set_stream_conditions(
            FEED_STREAM,
            temperature_k=FEED_T_K,
            pressure_kpa=FEED_P_KPA,
            molar_flow_kgmoleh=FEED_MOLAR_KMOLH,
        )
        driver.set_stream_composition(FEED_STREAM, FEED_COMPOSITION)
        driver.solve_case()

        feed = driver.read_stream_results(FEED_STREAM)
        vapour = driver.read_stream_results(VAPOUR_STREAM)
        liquid = driver.read_stream_results(LIQUID_STREAM)
        save_path = driver.save_case(
            Path(config.OUTPUT_DIR) / "conversion_380C_25bar.hsc"
        )

        summary = {
            "template": str(template),
            "feed": feed,
            "vapour_product": vapour,
            "liquid_product": liquid,
            "hsc_path": save_path,
        }
        print(json.dumps(summary, ensure_ascii=False, indent=2))

        # 计算甲苯转化率（进料纯甲苯；产物两相合并）
        feed_flow = float(feed["molar_flow_kgmoleh"] or 0.0)
        feed_tol = feed_flow * float(feed["mole_fractions"]["Toluene"])
        product_tol = 0.0
        for product in (vapour, liquid):
            flow = float(product.get("molar_flow_kgmoleh") or 0.0)
            fractions = product.get("mole_fractions") or {}
            product_tol += flow * float(fractions.get("Toluene", 0.0))
        conversion = (
            (feed_tol - product_tol) / feed_tol if feed_tol > 0 else 0.0
        )
        logger.info(
            "甲苯摩尔流量：进料 %.4f，产物 %.4f，转化率 %.2f%%",
            feed_tol,
            product_tol,
            conversion * 100.0,
        )

        if not vapour.get("solved") and not liquid.get("solved"):
            logger.error("两个产物均未求解")
            return 1
        if abs(conversion - 0.5) > 0.02:
            logger.error("甲苯转化率偏离 50%%：%.2f%%", conversion * 100.0)
            return 1
        print("\n✓ Conversion 场景联调完成（380℃/2.5MPa/10000kg/h，50% 转化率）。")
        return 0
    except Exception as exc:  # noqa: BLE001
        logger.exception("Conversion 场景失败：%s", exc)
        return 1
    finally:
        driver.shutdown()


if __name__ == "__main__":
    raise SystemExit(main())
