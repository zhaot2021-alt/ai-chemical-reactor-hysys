# -*- coding: utf-8 -*-
"""
llm_planner.py — LLM 结构化规划层（本地运行，密钥只从环境变量读取）
================================================================
输入：自然语言需求
输出：plan dict（反应器类型、组分/组成、工况、出口温度、转化率）

密钥约定：
  - 优先 CHEM_AI_LLM_API_KEY + CHEM_AI_LLM_BASE_URL + CHEM_AI_LLM_MODEL
  - 其次 DEEPSEEK_API_KEY / DEEPSEEK_BASE_URL / DEEPSEEK_MODEL
  - 再次 OPENAI_API_KEY / DASHSCOPE_API_KEY
  - 不把任何密钥写入仓库或日志

远程 HYSYS 工作站不需要密钥：本模块在本地生成 plan.json，
远程只执行 plan（见 agent.py run）。
"""

from __future__ import annotations

import json
import logging
import os
import re
import urllib.request
from typing import Any

logger = logging.getLogger("llm_planner")


class LLMPlanError(RuntimeError):
    pass


_SYSTEM_PROMPT = """你是 Aspen HYSYS 反应器建模助手。只输出一个 JSON 对象，不要输出任何解释或 Markdown。
按考核规则选择反应器类型：
1) 有动力学/设备参数：液相 CSTR、气相 PFR；
2) 无动力学但给出转化率/收率/选择性：Conversion；
3) 提供平衡常数 Ka 或吉布斯自由能：Equilibrium；
4) 产物分布未知、复杂高温体系：Gibbs。
HYSYS 标准组分 ID 只能是：
Methane,H2O,CO,CO2,Hydrogen,Toluene,Benzene,o-Xylene,m-Xylene,p-Xylene。
输出 JSON 字段：
{
 "reactor_type": "gibbs|conversion|equilibrium",
 "reason": "中文理由",
 "components": ["标准ID", "..."],
 "feed_composition": {"标准ID": 摩尔分数, "...": 0.0},
 "feed_temperature_c": 数字或null,
 "pressure_bar": 数字或null,
 "molar_flow_kmolh": 数字或null,
 "mass_flow_kg_h": 数字或null,
 "outlet_temperatures_c": [数字, ...],
 "conversion_percent": 数字或null,
 "reactions": [
   {
     "name": "Rxn-1",
     "stoichiometry": {"组分ID": 系数},
     "base_component": "组分ID",
     "conversion_percent": 数字
   }
 ],
 "notes": ["..."]
}
规则：
- Conversion 必须根据已配平的化学方程式给出 reactions，反应物系数为负、产物为正；
  若产物异构体比例未指定（如邻/间/对二甲苯），按等摩尔拆分并在 notes 中说明假设；
- Gibbs 的 reactions 可以为空数组；
- 不得编造用户未提供的数值；缺失用 null。甲烷蒸汽重整需包含产物组分 CO/CO2/Hydrogen。"""


def _candidate_providers() -> list[dict[str, str]]:
    return [
        {
            "name": "custom",
            "api_key_env": "CHEM_AI_LLM_API_KEY",
            "base_url_env": "CHEM_AI_LLM_BASE_URL",
            "model_env": "CHEM_AI_LLM_MODEL",
            "default_base": "https://api.deepseek.com/v1",
            "default_model": "deepseek-chat",
        },
        {
            "name": "deepseek",
            "api_key_env": "DEEPSEEK_API_KEY",
            "base_url_env": "DEEPSEEK_BASE_URL",
            "model_env": "DEEPSEEK_MODEL",
            "default_base": "https://api.deepseek.com/v1",
            "default_model": "deepseek-chat",
        },
        {
            "name": "openai",
            "api_key_env": "OPENAI_API_KEY",
            "base_url_env": "OPENAI_BASE_URL",
            "model_env": "OPENAI_MODEL",
            "default_base": "https://api.openai.com/v1",
            "default_model": "gpt-4o-mini",
        },
        {
            "name": "dashscope",
            "api_key_env": "DASHSCOPE_API_KEY",
            "base_url_env": "DASHSCOPE_BASE_URL",
            "model_env": "DASHSCOPE_MODEL",
            "default_base": "https://dashscope.aliyuncs.com/compatible-mode/v1",
            "default_model": "qwen-plus",
        },
    ]


def resolve_provider() -> tuple[str, str, str, str]:
    """返回 (provider, api_key, base_url, model)；密钥不打印。"""
    for item in _candidate_providers():
        api_key = os.getenv(item["api_key_env"], "").strip()
        if not api_key:
            continue
        base_url = (
            os.getenv(item["base_url_env"], "").strip()
            or item["default_base"]
        ).rstrip("/")
        model = (
            os.getenv(item["model_env"], "").strip()
            or item["default_model"]
        )
        return item["name"], api_key, base_url, model
    envs = [item["api_key_env"] for item in _candidate_providers()]
    raise LLMPlanError(
        "未找到 LLM key。请在本地设置以下任一环境变量："
        + "、".join(envs)
        + "（远程只需执行 plan，不需要 key）"
    )


def _extract_json(content: str) -> dict[str, Any]:
    text = content.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start >= 0 and end > start:
            try:
                return json.loads(text[start : end + 1])
            except json.JSONDecodeError:
                pass
    raise LLMPlanError(f"LLM 未返回合法 JSON：{content[:300]}")


def _normalize_plan(plan: dict[str, Any]) -> dict[str, Any]:
    reactor_type = str(plan.get("reactor_type", "")).strip().lower()
    if reactor_type not in ("gibbs", "conversion", "equilibrium"):
        raise LLMPlanError(f"不支持的反应器类型：{reactor_type!r}")
    plan["reactor_type"] = reactor_type

    components = plan.get("components") or []
    if not isinstance(components, list) or not components:
        raise LLMPlanError("plan.components 为空")
    plan["components"] = [str(item).strip() for item in components]

    composition = plan.get("feed_composition")
    if composition is not None:
        if not isinstance(composition, dict):
            raise LLMPlanError("plan.feed_composition 必须是对象或 null")
        plan["feed_composition"] = {
            str(key).strip(): float(value)
            for key, value in composition.items()
        }

    targets = plan.get("outlet_temperatures_c") or []
    if not isinstance(targets, list):
        raise LLMPlanError("plan.outlet_temperatures_c 必须是数组")
    plan["outlet_temperatures_c"] = [float(value) for value in targets]

    reactions = plan.get("reactions") or []
    if not isinstance(reactions, list):
        raise LLMPlanError("plan.reactions 必须是数组")
    normalized_reactions = []
    for reaction in reactions:
        if not isinstance(reaction, dict):
            raise LLMPlanError("plan.reactions 的元素必须是对象")
        stoich = reaction.get("stoichiometry") or {}
        if not isinstance(stoich, dict) or not stoich:
            raise LLMPlanError("reaction.stoichiometry 为空")
        normalized_reactions.append({
            "name": str(reaction.get("name") or "Rxn-1"),
            "stoichiometry": {
                str(key): float(value) for key, value in stoich.items()
            },
            "base_component": str(reaction.get("base_component") or ""),
            "conversion_percent": float(reaction.get("conversion_percent"))
            if reaction.get("conversion_percent") is not None else None,
        })
    plan["reactions"] = normalized_reactions
    return plan


def plan_request(text: str) -> dict[str, Any]:
    """调用 LLM，返回结构化 plan。密钥只从环境变量读取。"""
    if not text or not text.strip():
        raise LLMPlanError("输入文本为空")

    provider, api_key, base_url, model = resolve_provider()
    url = (
        base_url
        if base_url.endswith("/chat/completions")
        else base_url + "/chat/completions"
    )
    logger.info("LLM planner：provider=%s model=%s", provider, model)

    body = json.dumps(
        {
            "model": model,
            "temperature": 0,
            "messages": [
                {"role": "system", "content": _SYSTEM_PROMPT},
                {"role": "user", "content": text.strip()},
            ],
        }
    ).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=body,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except Exception as exc:  # noqa: BLE001
        raise LLMPlanError(f"LLM 请求失败：{type(exc).__name__}: {exc}") from exc

    try:
        content = payload["choices"][0]["message"]["content"]
    except Exception as exc:  # noqa: BLE001
        raise LLMPlanError(f"LLM 响应结构异常：{payload!r}") from exc

    return _normalize_plan(_extract_json(content))


if __name__ == "__main__":
    import sys

    result = plan_request(" ".join(sys.argv[1:]))
    print(json.dumps(result, ensure_ascii=False, indent=2))
