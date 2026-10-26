# -*- coding: utf-8 -*-
"""
web_ui.py — 极简智能体界面（零新依赖，标准库 http.server）
============================================================
在远程 HYSYS 工作站运行（该机需有 LLM key 环境变量 + HYSYS）：

    python src\\web_ui.py [--port 8000]

浏览器打开 http://localhost:8000：
  输入自然语言 → /api/plan（LLM 选型+生成 plan）→ 页面显示理由/假设
  → 点击执行 → /api/run（HYSYS 真实建模求解）→ 页面显示结果表格

复用 llm_planner / scenario_runner / hysys_driver，不改动任何已验收代码。
HYSYS 运行串行执行（全局锁），驱动每次新建并在 finally 中 shutdown。
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SRC_DIR))

import agent  # noqa: E402  提供 _check_plan_requirements
import config  # noqa: E402
import llm_planner  # noqa: E402

INDEX_HTML = (SRC_DIR / "index.html").read_text(encoding="utf-8")
_RUN_LOCK = threading.Lock()
logger = logging.getLogger("web_ui")


def _json_response(handler: BaseHTTPRequestHandler, payload: dict, status: int = 200) -> None:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def _handle_plan(body: dict) -> dict:
    text = str(body.get("text") or "").strip()
    if not text:
        return {"ok": False, "error": "请输入反应过程描述"}
    try:
        plan = llm_planner.plan_request(text)
    except Exception as exc:  # noqa: BLE001 - 给页面看原始原因
        return {"ok": False, "error": f"LLM 规划失败：{type(exc).__name__}: {exc}"}
    missing = agent._check_plan_requirements(plan)
    return {
        "ok": True,
        "plan": plan,
        "missing": missing,
        "notes": plan.get("notes") or [],
    }


def _handle_run(plan_dict: dict) -> dict:
    if not isinstance(plan_dict, dict) or not plan_dict:
        return {"ok": False, "error": "plan 为空，请先生成方案"}
    # 先用原始 dict 做缺参检查（比 normalize 更宽松，能给出友好提示）
    missing = agent._check_plan_requirements(plan_dict)
    if missing:
        return {"ok": False, "error": "信息不足：" + "；".join(missing)}
    try:
        plan = llm_planner._normalize_plan(plan_dict)
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": f"plan 不合法：{exc}"}
    missing = agent._check_plan_requirements(plan)
    if missing:
        return {"ok": False, "error": "信息不足：" + "；".join(missing)}

    reactor_type = str(plan.get("reactor_type", "")).lower()
    try:
        from hysys_driver import HYSYSDriver
        from scenario_runner import (
            run_conversion,
            run_equilibrium,
            run_gibbs,
            run_kinetic,
        )
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": f"导入 HYSYS 驱动失败：{exc}"}

    acquired = _RUN_LOCK.acquire(timeout=5)
    if not acquired:
        return {"ok": False, "error": "已有任务在执行，请稍候"}
    try:
        driver = HYSYSDriver(visible=True)
        try:
            driver.connect()
            if reactor_type == "gibbs":
                summary = run_gibbs(plan, driver)
            elif reactor_type == "conversion":
                summary = run_conversion(plan, driver)
            elif reactor_type == "equilibrium":
                summary = run_equilibrium(plan, driver)
            elif reactor_type in ("pfr", "cstr"):
                summary = run_kinetic(plan, driver)
            else:
                return {"ok": False, "error": f"不支持的反应器类型 {reactor_type!r}"}
        finally:
            driver.shutdown()
    except Exception as exc:  # noqa: BLE001
        logger.exception("执行失败：%s", exc)
        return {"ok": False, "error": f"执行失败：{type(exc).__name__}: {exc}"}
    finally:
        _RUN_LOCK.release()
    return {"ok": True, "summary": summary}


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:  # noqa: N802
        if self.path in ("/", "/index.html"):
            body = INDEX_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        _json_response(self, {"ok": False, "error": "not found"}, 404)

    def do_POST(self) -> None:  # noqa: N802
        try:
            length = int(self.headers.get("Content-Length") or 0)
            body = json.loads(self.rfile.read(length).decode("utf-8") or "{}")
        except Exception:  # noqa: BLE001
            return _json_response(self, {"ok": False, "error": "请求体不是合法 JSON"}, 400)
        if self.path == "/api/plan":
            _json_response(self, _handle_plan(body))
        elif self.path == "/api/run":
            _json_response(self, _handle_run(body.get("plan") or body))
        else:
            _json_response(self, {"ok": False, "error": "not found"}, 404)

    def log_message(self, fmt: str, *args: object) -> None:  # noqa: A003
        logger.info("%s - %s", self.address_string(), fmt % args)


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    logging.basicConfig(level=logging.INFO, format=config.LOG_FORMAT)
    parser = argparse.ArgumentParser(description="AI 化工反应器建模 Web 界面")
    parser.add_argument("--port", type=int, default=8000)
    args = parser.parse_args(argv)

    server = ThreadingHTTPServer(("0.0.0.0", args.port), Handler)
    logger.info("打开浏览器访问：http://localhost:%d", args.port)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
