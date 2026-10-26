# -*- coding: utf-8 -*-
"""
hysys_driver.py — HYSYS COM 底层驱动层（阶段1）
================================================================
职责（只做这些）：
  - 连接 / 启动 / 退出 HYSYS
  - 新建 Simulation Case
  - 创建流体包 + 设置物性包（Peng-Robinson）
  - 添加标准组分
  - 新建物料流并设置温度 / 压力 / 摩尔流量
  - 保存 .hsc
  - 释放 COM 对象，防止 hysys.exe 后台残留

边界：
  - 禁止写 RAG、LLM 解析、反应器类型判断、反应器创建代码
  - 上层拿不到原始 COM 对象：所有公开方法只返回 str / float / dict / tuple
  - 本层只接受 HYSYS 内部单位：K / kPa / kgmole·h⁻¹
    （外部 ℃ / bar 的换算统一由 config + main 完成）

API 可信度约定：
  - 标注“公开文档/工程代码印证”的调用可直接使用
  - 标注【需要在Excel VBA Object Browser核对】的调用有依据但 V15 签名可能不同，
    全部做了 try/except 或回退；运行失败时按提示核对该接口

【仅远程可运行】本模块导入即依赖 pywin32；真正 COM 操作需要
Windows + Aspen HYSYS V15 + 可用许可证。本地只能做语法检查 / 不导入本模块的 dry-run。
"""

from __future__ import annotations

import gc
import logging
import math
import subprocess
import time
from collections.abc import Iterator, Mapping, Sequence
from contextlib import contextmanager
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
import win32com.client

import config


class HYSYSDriverError(RuntimeError):
    """驱动层统一异常：封装 COM 错误、状态错误和保存失败。"""


# ==================== 模块级工具 ====================
def _safe_repr(obj: Any) -> str:
    """打印 COM 对象用；repr 失败时不打断主流程。"""
    try:
        return repr(obj)
    except Exception as exc:  # noqa: BLE001 - COM repr 可能抛任意异常
        return f"<repr 失败: {exc!r}>"


def _com_error_detail(exc: BaseException) -> str:
    """提取 COM 错误的 HRESULT / excepinfo，便于定位问题。"""
    if isinstance(exc, pywintypes.com_error):
        hresult = getattr(exc, "hresult", None)
        excepinfo = getattr(exc, "excepinfo", None)
        hr_text = (
            f"0x{hresult & 0xFFFFFFFF:08X}"
            if isinstance(hresult, int)
            else "N/A"
        )
        return f"HRESULT={hr_text}, excepinfo={excepinfo!r}"
    return f"{type(exc).__name__}: {exc}"


def _unique(items: Sequence[str]) -> list[str]:
    """去重并保持顺序。"""
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        if item and item not in seen:
            seen.add(item)
            result.append(item)
    return result


def _positive_float(value: Any, label: str) -> float:
    """校验内部单位数值：必须有限且大于 0。"""
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} 必须是数值，收到: {value!r}") from exc
    if not math.isfinite(number) or number <= 0:
        raise ValueError(f"{label} 必须是有限正数，收到: {value!r}")
    return number


def _hysys_process_running() -> bool | None:
    """用 tasklist 检查 hysys.exe；无法检查时返回 None。"""
    try:
        completed = subprocess.run(
            ["tasklist"],
            capture_output=True,
            text=True,
            errors="replace",
            timeout=config.HYSYS_TASKLIST_TIMEOUT_S,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:  # noqa: BLE001 - 进程检查失败不应影响释放流程
        return None
    output = (completed.stdout or "").lower()
    return "hysys" in output


# ==================== 驱动类 ====================
class HYSYSDriver:
    """
    阶段1 HYSYS COM 驱动。

    典型用法::

        with HYSYSDriver(visible=True) as drv:
            case = drv.new_case("Case1")
            fp = drv.create_fluid_package()
            drv.add_components(["Methane", "H2O"])
            drv.create_material_stream("FEED", 298.15, 101.325, 100.0)
            drv.save_case("output/Case1.hsc")

    退出 with 时会自动关闭 Case、退出 HYSYS（仅本次启动的实例）并反初始化 COM。
    注意：单个驱动实例面向单线程、单 Case；不要跨线程共享。
    """

    def __init__(
        self,
        visible: bool | None = None,
        attach_existing: bool | None = None,
        quit_attached: bool | None = None,
    ) -> None:
        self._visible = config.HYSYS_APP_VISIBLE if visible is None else bool(visible)
        self._attach_existing = (
            config.HYSYS_ATTACH_EXISTING
            if attach_existing is None
            else bool(attach_existing)
        )
        self._quit_attached = (
            config.HYSYS_QUIT_ATTACHED
            if quit_attached is None
            else bool(quit_attached)
        )

        self._app: Any = None
        self._case: Any = None
        self._fluid_pkg: Any = None
        self._streams: dict[str, Any] = {}

        self._case_name: str | None = None
        self._fluid_pkg_name: str | None = None
        self._property_package_name: str | None = None
        self._component_ids: list[str] = []

        self._owns_app = False          # True = 本次 Dispatch 启动的实例
        self._com_initialized = False
        self._shutdown_done = False

        self.logger = logging.getLogger("hysys_driver")

    # ---------- 只读信息（只暴露字符串/数值，不暴露 COM） ----------
    @property
    def is_connected(self) -> bool:
        return self._app is not None

    @property
    def case_name(self) -> str | None:
        return self._case_name

    @property
    def fluid_package_name(self) -> str | None:
        return self._fluid_pkg_name

    @property
    def property_package_name(self) -> str | None:
        return self._property_package_name

    @property
    def component_ids(self) -> tuple[str, ...]:
        return tuple(self._component_ids)

    @property
    def stream_names(self) -> tuple[str, ...]:
        return tuple(self._streams.keys())

    # ---------- 生命周期 ----------
    def __enter__(self) -> "HYSYSDriver":
        self.connect()
        return self

    def __exit__(self, exc_type: Any, exc: Any, tb: Any) -> bool:
        self.shutdown()
        return False  # 不吞异常

    def connect(self) -> None:
        """
        连接已运行的 HYSYS；没有则 Dispatch 启动新实例。
        GetActiveObject 附着到的已有实例默认不属于本次调用，shutdown 时不 Quit。
        """
        if self._app is not None:
            self.logger.debug("HYSYS 已连接，跳过重复 connect()")
            return

        self._shutdown_done = False
        self.logger.info("连接 HYSYS COM（ProgID=%s）...", config.HYSYS_PROGID)

        try:
            pythoncom.CoInitialize()
            self._com_initialized = True
        except Exception as exc:  # noqa: BLE001 - COM 可能已被宿主线程初始化
            self.logger.debug("CoInitialize 返回异常（可能已初始化）: %s", exc)

        app: Any = None
        owns = False

        if self._attach_existing:
            try:
                app = win32com.client.GetActiveObject(config.HYSYS_PROGID)
                self.logger.info("已连接到运行中的 HYSYS 实例")
            except Exception:
                self.logger.info("未发现运行中的 HYSYS 实例，准备启动新实例")

        if app is None:
            last_error: BaseException | None = None
            for attempt in range(1, config.HYSYS_CONNECT_RETRIES + 1):
                try:
                    app = win32com.client.Dispatch(config.HYSYS_PROGID)
                    owns = True
                    self.logger.info("HYSYS 实例启动成功")
                    break
                except Exception as exc:  # noqa: BLE001
                    last_error = exc
                    self.logger.warning(
                        "启动 HYSYS 失败（第 %d/%d 次）：%s",
                        attempt,
                        config.HYSYS_CONNECT_RETRIES,
                        _com_error_detail(exc),
                    )
                    if attempt < config.HYSYS_CONNECT_RETRIES:
                        time.sleep(config.HYSYS_RETRY_INTERVAL_S)

            if app is None:
                self._release_com_library()
                raise HYSYSDriverError(
                    "无法启动 HYSYS COM 服务"
                    f"（ProgID={config.HYSYS_PROGID}）：{_com_error_detail(last_error)}。"
                    "请检查：1) HYSYS V15 已安装并成功激活；"
                    "2) ProgID 是否正确（注册表 HKEY_CLASSES_ROOT）；"
                    "3) 是否有许可证/启动对话框需要先手动关闭。"
                ) from last_error

        try:
            app.Visible = self._visible
        except Exception as exc:  # noqa: BLE001
            self.logger.warning(
                "设置 Application.Visible=%s 失败（不影响后台运行）：%s",
                self._visible,
                _com_error_detail(exc),
            )

        self._app = app
        self._owns_app = owns
        self.logger.debug("Application COM 对象：%s", _safe_repr(app))
        self.logger.info("HYSYS 连接完成（本次启动的实例=%s）", owns)

        if not owns and not self._quit_attached:
            self.logger.info(
                "当前依附已有实例：shutdown 时会关闭本 Case，但保留 hysys.exe 进程"
                "（如需强制退出，将 quit_attached=True 或 HYSYS_QUIT_ATTACHED=True）"
            )

    def shutdown(self) -> None:
        """
        显式释放所有 COM 资源（幂等，可重复调用）：
          1) 释放物流 / 流体包引用
          2) Close 本 Case（不保存）
          3) 本次启动的实例或 quit_attached=True 时 Quit
          4) CoUninitialize
          5) tasklist 检查 hysys.exe 是否残留
        """
        if self._shutdown_done:
            return
        self._shutdown_done = True

        self.logger.info("开始释放 COM 资源 ...")
        should_quit = self._app is not None and (self._owns_app or self._quit_attached)

        # 1) 先释放子对象 COM 引用，再关闭 Case
        self._streams.clear()
        self._fluid_pkg = None
        gc.collect()

        # 2) 关闭 Case（本 Case 已由 main 显式 SaveAs 后才正常走到这里）
        self._close_case()

        # 3) 退出 HYSYS
        app = self._app
        self._app = None
        if app is not None:
            if should_quit:
                try:
                    app.Quit()
                    self.logger.info("已调用 Application.Quit()")
                except Exception as exc:  # noqa: BLE001
                    self.logger.warning(
                        "Application.Quit 失败：%s", _com_error_detail(exc)
                    )
            else:
                self.logger.info(
                    "依附已有 HYSYS 实例，未执行 Quit（quit_attached=False）"
                )
            del app
        gc.collect()

        # 4) 反初始化 COM（本次 connect 中 CoInitialize 的配对调用）
        self._release_com_library()

        # 5) 检查后台残留（只有主动 Quit 的实例才检查）
        if should_quit:
            time.sleep(config.HYSYS_SHUTDOWN_WAIT_S)
            running = _hysys_process_running()
            if running is None:
                self.logger.warning(
                    "无法通过 tasklist 检查 hysys.exe；如仍见到后台进程，"
                    "请手动执行：taskkill /F /IM hysys.exe /T"
                )
            elif running:
                self.logger.warning(
                    "检测到 hysys.exe 仍在运行！可手动执行："
                    "taskkill /F /IM hysys.exe /T"
                )
            else:
                self.logger.info("hysys.exe 已正常退出，无后台残留")

        # 6) 清理状态，允许同一对象再次 connect（重新初始化为全新会话）
        self._owns_app = False
        self._component_ids.clear()
        self._case_name = None
        self._fluid_pkg_name = None
        self._property_package_name = None
        self.logger.info("COM 资源释放完成")

    # ---------- 阶段1 能力 ----------
    def new_case(self, name: str | None = None) -> str:
        """新建空白 Simulation Case，返回 Case 名称。"""
        self._require_app()
        if self._case is not None:
            raise HYSYSDriverError(
                f"当前驱动已关联 Case {self._case_name!r}；"
                "请先 shutdown / 新建驱动实例，再创建新 Case"
            )

        case_name = (
            name or f"{config.CASE_NAME_PREFIX}_{time.strftime('%Y%m%d_%H%M%S')}"
        ).strip()
        if not case_name:
            raise ValueError("Case 名称不能为空")

        self.logger.info("新建 Simulation Case：%s", case_name)
        try:
            # 公开 Python/VBA 模板使用 SimulationCases.Add(name)。
            # 【需要在Excel VBA Object Browser核对】Add(name) 在 V15 的参数签名；
            # 失败时回退无参 Add()，并尝试读取实际名称。
            try:
                case = self._app.SimulationCases.Add(case_name)
            except Exception as first_exc:  # noqa: BLE001
                self.logger.warning(
                    "SimulationCases.Add(name) 失败，回退无参 Add()：%s",
                    _com_error_detail(first_exc),
                )
                case = self._app.SimulationCases.Add()
                case_name = self._read_com_name(case, case_name)
        except Exception as exc:  # noqa: BLE001
            raise HYSYSDriverError(
                f"新建 Case 失败：{_com_error_detail(exc)}"
            ) from exc

        try:
            case.Visible = True
        except Exception as exc:  # noqa: BLE001
            self.logger.debug("设置 Case.Visible=True 失败（可忽略）：%s", exc)

        self._case = case
        self._case_name = case_name
        self.logger.debug("SimulationCase COM 对象：%s", _safe_repr(case))
        self.logger.info("Case 已创建：%s", case_name)
        return case_name

    def create_fluid_package(
        self,
        name: str | None = None,
        property_package: str | None = None,
        components: Sequence[str] | None = None,
    ) -> str:
        """
        创建流体包 + 设置物性包 + 在**同一个 Basis 变更**中加入组分。

        HYSYS V15 实测要点（2026-10-26）：
          - FluidPackages.Add("名称") 可按名称创建（官方 FAQ 示例 Add "Steam"）
          - PropertyPackageName 的 Setter 使用内部 token：
              Peng-Robinson -> "pengrob"；读回值则是显示名 "Peng-Robinson"
          - 组分必须在 EndBasisChange 之前加入；空 Basis 的 End 会 E_FAIL
          - FluidPackage.Name 不要在设置物性包之前修改（V15 会破坏 Setter 状态）

        components 必须非空（阶段1的流体包至少需要一个组分）。
        """
        self._require_case()
        if self._fluid_pkg is not None:
            raise HYSYSDriverError(
                f"流体包 {self._fluid_pkg_name!r} 已存在，禁止重复创建"
            )

        target_name = (name or config.FLUID_PACKAGE_NAME).strip()
        requested_pp = (property_package or config.PROPERTY_PACKAGE_NAME).strip()
        pp_token = config.PROPERTY_PACKAGE_ALIASES.get(requested_pp, requested_pp)

        if not components or isinstance(components, str):
            raise ValueError(
                "create_fluid_package 必须提供非空的 components；"
                "HYSYS 不接受没有组分的空 Basis（EndBasisChange 会失败）"
            )
        component_ids = self._validate_component_ids(components)

        self.logger.info(
            "创建流体包（名称=%s，物性包 token=%s，组分=%d 个）...",
            target_name,
            pp_token,
            len(component_ids),
        )

        added: list[str] = []
        failures: list[str] = []
        actual_pp_display = config.PROPERTY_PACKAGE_DISPLAY_NAME

        with self._basis_change("create_fluid_package"):
            fluid_pkg = self._add_fluid_package(target_name)

            # 先设物性包内部 token，再在同一 Basis 中加组分
            # V15 Peng-Robinson 的 Setter token 为 "pengrob"
            try:
                fluid_pkg.PropertyPackageName = pp_token
            except Exception as exc:  # noqa: BLE001
                raise HYSYSDriverError(
                    f"设置物性包 token={pp_token!r} 失败："
                    f"{_com_error_detail(exc)}。"
                    "请确认 config.PROPERTY_PACKAGE_NAME（Peng-Robinson 应为 'pengrob'）；"
                    "可运行 tools/read_open_fluid_packages.py 对照手动 PR Case。"
                ) from exc

            actual_pp_display = self._read_property_package_display(fluid_pkg)
            self.logger.info(
                "物性包已设置：token=%s，读回显示名=%s",
                pp_token,
                actual_pp_display,
            )

            added, failures = self._add_components_uncommitted(
                fluid_pkg, component_ids
            )
            self.logger.debug(
                "FluidPackage COM 对象：%s", _safe_repr(fluid_pkg)
            )

        # 到这里 EndBasisChange 已成功（_basis_change 会在 End 失败时抛异常）

        if failures:
            raise HYSYSDriverError(self._component_failure_message(failures))

        # 说明：此处不调用 case.Flowsheet.FluidPackage = fp 做 Make Current。
        # 公开资料显示 Flowsheet.FluidPackage 主要用于读取当前流体包，可写性未确认；
        # 阶段1只有一个流体包，HYSYS 会将其关联到流程。
        # 后续多流体包场景如需 Make Current，请先
        # 【需要在Excel VBA Object Browser核对】对应接口，再扩展本方法。

        self._fluid_pkg = fluid_pkg
        self._component_ids.extend(added)
        self._property_package_name = actual_pp_display
        self._fluid_pkg_name = self._read_com_name(fluid_pkg, target_name)
        self.logger.info(
            "流体包已创建：%s（物性包=%s，已加组分=%s）",
            self._fluid_pkg_name,
            self._property_package_name,
            added,
        )
        return self._fluid_pkg_name

    def add_components(self, component_ids: Sequence[str]) -> tuple[str, ...]:
        """
        向已创建的流体包追加 HYSYS 标准组分（阶段2/后续扩展用）。
        返回本次新添加的标准 ID 元组；重复 ID 自动跳过。

        注意：阶段1 的主流程把组分放在 create_fluid_package(components=...)
        的同一个 Basis 变更里添加；本方法用于之后单独追加组分。

        component_ids 必须是 config.STANDARD_COMPONENT_IDS 中的标准 ID，
        不接受任意别名（别名 → 标准 ID 的映射由 nlp_parser 完成）。
        """
        self._require_fluid_pkg()
        requested = self._validate_component_ids(component_ids)
        pending = [
            item for item in requested if item not in self._component_ids
        ]
        if not pending:
            self.logger.info("所有组分均已添加，跳过：%s", requested)
            return ()

        self.logger.info("追加组分：%s", pending)
        added: list[str] = []
        failures: list[str] = []

        with self._basis_change("add_components"):
            added, failures = self._add_components_uncommitted(
                self._fluid_pkg, pending
            )

        self._component_ids.extend(added)
        if failures:
            raise HYSYSDriverError(self._component_failure_message(failures))

        self.logger.info("组分添加完成：%s", added)
        return tuple(added)

    def _add_fluid_package(self, target_name: str) -> Any:
        """FluidPackages.Add(name) 优先；失败回退 Add()。"""
        try:
            fluid_pkg = self._case.BasisManager.FluidPackages.Add(target_name)
            self.logger.info("FluidPackages.Add(name) 成功：%s", target_name)
            return fluid_pkg
        except Exception as first_exc:  # noqa: BLE001
            self.logger.warning(
                "FluidPackages.Add(name) 失败，回退 Add()：%s",
                _com_error_detail(first_exc),
            )
            try:
                return self._case.BasisManager.FluidPackages.Add()
            except Exception as second_exc:  # noqa: BLE001
                raise HYSYSDriverError(
                    "创建流体包失败（Add(name) 与 Add() 均失败）："
                    f"{_com_error_detail(second_exc)}"
                ) from second_exc

    def _validate_component_ids(
        self, component_ids: Sequence[str]
    ) -> list[str]:
        if isinstance(component_ids, str) or not component_ids:
            raise ValueError("component_ids 必须是非空的组分 ID 序列")

        requested = [str(item).strip() for item in component_ids]
        if any(not item for item in requested):
            raise ValueError("component_ids 不能包含空字符串")

        unknown = [
            item for item in requested
            if item not in config.STANDARD_COMPONENT_IDS
        ]
        if unknown:
            raise HYSYSDriverError(
                f"检测到非标准组分 ID：{unknown}；"
                "允许的标准 ID 见 config.STANDARD_COMPONENT_IDS："
                f"{list(config.STANDARD_COMPONENT_IDS)}"
            )
        return _unique(requested)

    def _add_components_uncommitted(
        self, fluid_pkg: Any, component_ids: Sequence[str]
    ) -> tuple[list[str], list[str]]:
        """
        在当前已打开的 Basis 变更中添加组分；不调用 Start/EndBasisChange。
        返回 (成功添加的 ID 列表, 失败描述列表)。
        """
        added: list[str] = []
        failures: list[str] = []

        for component_id in component_ids:
            if component_id in self._component_ids:
                continue

            candidates = (
                component_id,
                *config.COMPONENT_ID_FALLBACKS.get(component_id, ()),
            )
            last_error: BaseException | None = None
            success = False

            for candidate in candidates:
                try:
                    # 官方 Customization Guide：FluidPackage.Components.Add(组分名)
                    component_obj = fluid_pkg.Components.Add(candidate)
                    added.append(component_id)
                    success = True
                    if candidate != component_id:
                        self.logger.warning(
                            "组分 %s 以同义名 %s 添加成功（标准 Sim Name 仍为 %s）",
                            component_id,
                            candidate,
                            component_id,
                        )
                    else:
                        self.logger.info("已添加组分：%s", component_id)
                    self.logger.debug(
                        "Components.Add 返回对象：%s", _safe_repr(component_obj)
                    )
                    component_obj = None  # 不保留组分 COM 引用
                    break
                except Exception as exc:  # noqa: BLE001
                    last_error = exc
                    self.logger.debug(
                        "组分 %s（尝试名 %s）添加失败：%s",
                        component_id,
                        candidate,
                        _com_error_detail(exc),
                    )

            if not success:
                failures.append(
                    f"{component_id}（尝试 {list(candidates)}，最后错误："
                    f"{_com_error_detail(last_error) if last_error else '未知'}）"
                )

        return added, failures

    def _component_failure_message(self, failures: Sequence[str]) -> str:
        return (
            "部分组分添加失败：" + " | ".join(failures)
            + "。请在 HYSYS 组分库 GUI 中确认 Sim Name"
            "（如 H2O / Hydrogen / o-Xylene）"
        )

    def _read_property_package_display(self, fluid_pkg: Any) -> str:
        """读回物性包显示名；失败时退回 PropertyPackage.TypeName / 配置值。"""
        try:
            value = fluid_pkg.PropertyPackageName
            if value:
                return str(value)
        except Exception as exc:  # noqa: BLE001
            self.logger.debug(
                "读取 PropertyPackageName 失败：%s", _com_error_detail(exc)
            )
        try:
            value = fluid_pkg.PropertyPackage.TypeName
            if value:
                return str(value)
        except Exception as exc:  # noqa: BLE001
            self.logger.debug(
                "读取 PropertyPackage.TypeName 失败：%s",
                _com_error_detail(exc),
            )
        return config.PROPERTY_PACKAGE_DISPLAY_NAME

    def create_material_stream(
        self,
        name: str,
        temperature_k: float,
        pressure_kpa: float,
        molar_flow_kgmoleh: float,
    ) -> dict[str, Any]:
        """
        新建物料流并设置参数（只接受 HYSYS 内部单位 K / kPa / kgmole·h⁻¹）。
        返回安全的状态字典（不含 COM 对象）。

        阶段1不设置组分组成：物流处于未求解状态，仅用于验证 COM 建流能力。
        阶段2再补 MolarFraction / MassFraction 与实际求解。
        """
        self._require_case()
        self._require_fluid_pkg()

        stream_name = str(name).strip()
        if not stream_name:
            raise ValueError("物流名称不能为空")
        if stream_name in self._streams:
            raise HYSYSDriverError(f"物流 {stream_name!r} 已存在，禁止重复创建")

        temp_k = _positive_float(temperature_k, "temperature_k")
        press_kpa = _positive_float(pressure_kpa, "pressure_kpa")
        flow = _positive_float(molar_flow_kgmoleh, "molar_flow_kgmoleh")

        self.logger.info(
            "新建物料流 %s（T=%.4f K, P=%.4f kPa, F=%.4f kgmole/h）...",
            stream_name,
            temp_k,
            press_kpa,
            flow,
        )

        try:
            # MaterialStreams.Add(name) 有公开 VBA 示例印证
            stream = self._case.Flowsheet.MaterialStreams.Add(stream_name)
        except Exception as exc:  # noqa: BLE001
            raise HYSYSDriverError(
                f"新建物料流 {stream_name!r} 失败：{_com_error_detail(exc)}"
            ) from exc

        self.logger.debug("MaterialStream COM 对象：%s", _safe_repr(stream))

        try:
            self._set_stream_value(
                stream, "Temperature", temp_k,
                config.HYSYS_TEMPERATURE_UNIT, stream_name,
            )
            self._set_stream_value(
                stream, "Pressure", press_kpa,
                config.HYSYS_PRESSURE_UNIT, stream_name,
            )
            self._set_stream_value(
                stream, "MolarFlow", flow,
                config.HYSYS_MOLAR_FLOW_UNIT, stream_name,
            )
        except Exception:
            # 物流已创建但参数不完整；交给上层异常处理 + shutdown 清理
            raise

        self._streams[stream_name] = stream
        state = self.read_stream_values(stream_name)
        self.logger.info("物料流已创建：%s", state)
        self.logger.info(
            "阶段1未设置物流组成，物流处于未求解状态（仅验证连通性）"
        )
        return state

    def read_stream_values(self, name: str) -> dict[str, Any]:
        """
        读取物流的温度/压力/摩尔流量，返回只含数值的安全字典。
        读取失败时对应值为 None，仅记录 warning。
        """
        stream = self._streams.get(name)
        if stream is None:
            raise HYSYSDriverError(
                f"物流 {name!r} 不存在；现有物流：{list(self._streams)}"
            )

        values: dict[str, Any] = {
            "name": name,
            "temperature_k": None,
            "pressure_kpa": None,
            "molar_flow_kgmoleh": None,
        }
        read_specs = (
            ("temperature_k", "Temperature", config.HYSYS_TEMPERATURE_UNIT),
            ("pressure_kpa", "Pressure", config.HYSYS_PRESSURE_UNIT),
            ("molar_flow_kgmoleh", "MolarFlow", config.HYSYS_MOLAR_FLOW_UNIT),
        )
        for key, attribute, unit in read_specs:
            try:
                # ProcessVariable.GetValue(unit) 有公开文档/论文示例印证
                process_variable = getattr(stream, attribute)
                values[key] = float(process_variable.GetValue(unit))
            except Exception as exc:  # noqa: BLE001
                self.logger.warning(
                    "读取物流 %s 的 %s 失败：%s",
                    name,
                    attribute,
                    _com_error_detail(exc),
                )

        return values

    def set_stream_composition(
        self,
        name: str,
        mole_fractions: Mapping[str, float],
    ) -> dict[str, Any]:
        """
        设置物流摩尔组成（HYSYS V15 实测：ComponentMolarFractionValue 可直接 PUT）。

        参数：
          name           物流名（通常是 FEED）
          mole_fractions {HYSYS 标准组分 ID: 摩尔分数}；可只给部分组分，
                         未给出的组分按 0 处理；总和会归一化。

        返回：
          {"name", "requested_mole_fractions", "mole_fractions", "normalized"}

        说明：
          - 组分数组顺序按 HYSYS 流体包实际 Components 顺序构造
          - 写入后立即读回校验；明显不一致会抛 HYSYSDriverError
          - 本方法不调用 solve；由 main 明确调用 solve_case()
        """
        self._require_fluid_pkg()
        stream = self._streams.get(name)
        if stream is None:
            raise HYSYSDriverError(
                f"物流 {name!r} 不存在；现有物流：{list(self._streams)}"
            )
        if not isinstance(mole_fractions, Mapping) or not mole_fractions:
            raise ValueError("mole_fractions 必须是非空的 {组分ID: 摩尔分数} 映射")

        component_names = self._fluid_package_component_names()
        if not component_names:
            raise HYSYSDriverError("无法确定流体包组分顺序，不能设置组成")

        unknown = [
            component_id for component_id in mole_fractions
            if component_id not in component_names
        ]
        if unknown:
            raise HYSYSDriverError(
                f"组成中包含流体包不存在的组分：{unknown}；"
                f"当前流体包组分：{component_names}"
            )

        raw: dict[str, float] = {}
        for component_id in component_names:
            value = mole_fractions.get(component_id, 0.0)
            try:
                number = float(value)
            except (TypeError, ValueError) as exc:
                raise ValueError(
                    f"组分 {component_id!r} 的摩尔分数必须是数值，收到 {value!r}"
                ) from exc
            if not math.isfinite(number) or number < 0:
                raise ValueError(
                    f"组分 {component_id!r} 的摩尔分数必须是非负有限数，收到 {value!r}"
                )
            raw[component_id] = number

        total = sum(raw.values())
        if total <= 0:
            raise ValueError("摩尔分数总和必须大于 0")

        normalized = abs(total - 1.0) <= config.COMPOSITION_TOLERANCE
        if not normalized:
            self.logger.warning(
                "物流 %s 的组成总和为 %.8f，将归一化到 1.0", name, total
            )
        target = {
            component_id: raw[component_id] / total
            for component_id in component_names
        }
        array = tuple(target[component_id] for component_id in component_names)

        self.logger.info(
            "设置物流 %s 组成：%s",
            name,
            {key: round(value, 8) for key, value in target.items()},
        )

        try:
            # V15 实测：直接 PUT 成功；若失败，回退到 ComponentMolarFraction.Calculate
            stream.ComponentMolarFractionValue = array
        except Exception as exc:  # noqa: BLE001
            self.logger.warning(
                "ComponentMolarFractionValue 直接赋值失败，尝试 Calculate：%s",
                _com_error_detail(exc),
            )
            try:
                stream.ComponentMolarFraction.Calculate(array)
            except Exception as calc_exc:  # noqa: BLE001
                raise HYSYSDriverError(
                    f"设置物流 {name!r} 组成失败："
                    f"直接赋值={_com_error_detail(exc)}；"
                    f"Calculate={_com_error_detail(calc_exc)}"
                ) from calc_exc

        readback = self._read_stream_composition(stream, name)
        if readback is None:
            raise HYSYSDriverError(
                f"物流 {name!r} 组成写入后读回为空（-32767 或读取失败）"
            )

        mismatches = []
        for component_id in component_names:
            actual = readback.get(component_id, 0.0)
            if abs(actual - target[component_id]) > 1e-4:
                mismatches.append(
                    (component_id, target[component_id], actual)
                )
        if mismatches:
            raise HYSYSDriverError(
                f"物流 {name!r} 组成读回与目标不一致：{mismatches}"
            )

        self.logger.info(
            "物流 %s 组成设置成功（读回和=%.8f）",
            name,
            sum(readback.values()),
        )
        return {
            "name": name,
            "requested_mole_fractions": target,
            "mole_fractions": readback,
            "input_sum": total,
            "normalization_applied": not normalized,
        }

    def solve_case(self) -> None:
        """将 Solver.CanSolve 置 True，并等待 IsSolving 结束。"""
        self._require_case()
        try:
            # V15 实测：CanSolve 的 GET/PUT 都存在，且可写
            self._case.Solver.CanSolve = True
            self.logger.info("Solver.CanSolve = True")
        except Exception as exc:  # noqa: BLE001
            self.logger.warning(
                "设置 Solver.CanSolve=True 失败，继续等待：%s",
                _com_error_detail(exc),
            )
        self._wait_for_solver()

    def read_stream_results(self, name: str) -> dict[str, Any]:
        """
        读取物流的求解结果：T/P/F、摩尔组成、汽化率、是否已求解。
        返回纯 Python 值，不暴露 COM 对象。
        """
        stream = self._streams.get(name)
        if stream is None:
            raise HYSYSDriverError(
                f"物流 {name!r} 不存在；现有物流：{list(self._streams)}"
            )

        values = self.read_stream_values(name)
        composition = self._read_stream_composition(stream, name)
        vapour_fraction = None
        try:
            candidate = float(stream.VapourFractionValue)
            if candidate > -30000:
                vapour_fraction = candidate
        except Exception as exc:  # noqa: BLE001
            self.logger.warning(
                "读取物流 %s 的 VapourFractionValue 失败：%s",
                name,
                _com_error_detail(exc),
            )

        values["mole_fractions"] = composition
        values["vapour_fraction"] = vapour_fraction
        values["solved"] = composition is not None and vapour_fraction is not None
        return values

    def _fluid_package_component_names(self) -> list[str]:
        """读取流体包实际组分顺序（HYSYS 组成数组的顺序）。"""
        self._require_fluid_pkg()
        try:
            count = int(self._fluid_pkg.Components.Count)
        except Exception as exc:  # noqa: BLE001
            self.logger.warning(
                "读取流体包组分数量失败，退回 driver 记录顺序：%s",
                _com_error_detail(exc),
            )
            return list(self._component_ids)

        # HYSYS 集合通常是 0-based；兼容 1-based。
        for start in (0, 1):
            names: list[str] = []
            for index in range(start, start + count):
                try:
                    names.append(str(self._fluid_pkg.Components.Item(index).Name))
                except Exception:
                    names = []
                    break
            if names:
                return names

        self.logger.warning("无法读取流体包组分顺序，退回 driver 记录顺序")
        return list(self._component_ids)

    def _read_stream_composition(
        self, stream: Any, name: str
    ) -> dict[str, float] | None:
        """读取物流组分摩尔分数；未设置时 HYSYS 返回 -32767。"""
        try:
            raw = stream.ComponentMolarFractionValue
            values = [float(value) for value in raw] if raw is not None else []
        except Exception as exc:  # noqa: BLE001
            self.logger.warning(
                "读取物流 %s 的 ComponentMolarFractionValue 失败：%s",
                name,
                _com_error_detail(exc),
            )
            return None

        if not values or all(value <= -30000 for value in values):
            return None

        component_names = self._fluid_package_component_names()
        if len(values) != len(component_names):
            self.logger.warning(
                "物流 %s 组成数组长度 %d 与组分数量 %d 不一致",
                name,
                len(values),
                len(component_names),
            )
        count = min(len(values), len(component_names))
        return {
            component_names[index]: values[index]
            for index in range(count)
        }

    def _wait_for_solver(self) -> None:
        solver = self._case.Solver
        deadline = time.time() + config.HYSYS_SOLVE_TIMEOUT_S
        while True:
            try:
                is_solving = bool(solver.IsSolving)
            except Exception as exc:  # noqa: BLE001
                self.logger.warning(
                    "读取 Solver.IsSolving 失败，等待 0.5s 后继续：%s",
                    _com_error_detail(exc),
                )
                time.sleep(0.5)
                return

            if not is_solving:
                self.logger.info("求解器空闲（IsSolving=False）")
                return
            if time.time() >= deadline:
                raise HYSYSDriverError(
                    f"等待 HYSYS 求解超时（>{config.HYSYS_SOLVE_TIMEOUT_S:.0f}s）"
                )
            time.sleep(config.HYSYS_SOLVE_POLL_S)

    def save_case(self, path: str | Path | None = None) -> str:
        """保存 .hsc，返回绝对路径。默认保存到 config.OUTPUT_DIR。"""
        self._require_case()

        if path is None:
            file_name = f"{self._case_name or 'case'}{config.HSC_EXTENSION}"
            target = config.OUTPUT_DIR / file_name
        else:
            target = Path(path).expanduser()
            if not target.is_absolute():
                target = (Path.cwd() / target).resolve()
            if target.suffix.lower() != config.HSC_EXTENSION:
                target = target.with_suffix(config.HSC_EXTENSION)

        target.parent.mkdir(parents=True, exist_ok=True)
        self.logger.info("保存 Case 到：%s", target)

        try:
            # SimulationCase.SaveAs(path) 有公开模板印证
            self._case.SaveAs(str(target))
        except Exception as exc:  # noqa: BLE001
            raise HYSYSDriverError(
                f"保存 Case 失败：{target}；{_com_error_detail(exc)}"
            ) from exc

        # SaveAs 通常同步完成；加短暂等待提高健壮性
        for _ in range(5):
            if target.exists():
                break
            time.sleep(0.2)

        if not target.exists():
            raise HYSYSDriverError(
                f"SaveAs 已调用但未生成文件：{target}。"
                "请检查 HYSYS 是否弹出保存对话框或路径是否可写。"
            )

        self.logger.info(
            "Case 已保存：%s（%d 字节）", target, target.stat().st_size
        )
        return str(target)

    # ---------- 阶段2 扩展位（当前故意留空，不生成任何反应器代码） ----------
    # 约定：
    #   1) 新增 create_conversion_reactor / create_equilibrium_reactor / create_gibbs_reactor 时，
    #      复用 _basis_change()、_require_fluid_pkg()、HYSYSDriverError 与 shutdown 生命周期；
    #   2) 反应器的 Add 接口、属性名、单位必须先对照 Object Model Browser，禁止猜测；
    #   3) 不在本层写“根据自然语言选择反应器类型”的判断，该职责属于 nlp_parser / 未来 LLM。

    # ---------- 内部辅助 ----------
    def _require_app(self) -> None:
        if self._app is None:
            raise HYSYSDriverError(
                "尚未连接 HYSYS；请先 connect() 或使用 with HYSYSDriver(...)"
            )

    def _require_case(self) -> None:
        self._require_app()
        if self._case is None:
            raise HYSYSDriverError("尚未创建 Simulation Case；请先调用 new_case()")

    def _require_fluid_pkg(self) -> None:
        self._require_case()
        if self._fluid_pkg is None:
            raise HYSYSDriverError(
                "尚未创建流体包；请先调用 create_fluid_package()"
            )

    @contextmanager
    def _basis_change(self, purpose: str) -> Iterator[None]:
        """
        Basis 修改通知：StartBasisChange / EndBasisChange。
        Aspen 文档要求修改流体包/组分等 Basis 对象时通知 HYSYS，
        否则可能弹出交互对话框或修改不生效。
        """
        self._require_case()
        try:
            self._case.BasisManager.StartBasisChange()
        except Exception as exc:  # noqa: BLE001
            self.logger.debug(
                "[%s] StartBasisChange 不可用（继续执行，EndBasisChange 仍会调用）：%s",
                purpose,
                _com_error_detail(exc),
            )

        had_original_error = False
        try:
            yield
        except BaseException:
            had_original_error = True
            raise
        finally:
            try:
                self._case.BasisManager.EndBasisChange()
                self.logger.debug("[%s] EndBasisChange 完成", purpose)
            except Exception as exc:  # noqa: BLE001
                self.logger.warning(
                    "[%s] EndBasisChange 失败：%s",
                    purpose,
                    _com_error_detail(exc),
                )
                if not had_original_error:
                    raise HYSYSDriverError(
                        f"EndBasisChange 失败（{purpose}）："
                        f"{_com_error_detail(exc)}；"
                        "Basis 可能仍处于 changing 状态，本 Case 不应继续使用"
                    ) from exc

    def _set_stream_value(
        self,
        stream: Any,
        attribute: str,
        value: float,
        unit: str,
        stream_name: str,
    ) -> None:
        """ProcessVariable.SetValue(value, unit) 的薄封装。"""
        try:
            process_variable = getattr(stream, attribute)
            process_variable.SetValue(value, unit)
        except Exception as exc:  # noqa: BLE001
            raise HYSYSDriverError(
                f"设置物流 {stream_name!r} 的 {attribute} 失败"
                f"（value={value!r}, unit={unit!r}）：{_com_error_detail(exc)}"
            ) from exc

    def _read_com_name(self, obj: Any, fallback: str) -> str:
        """best-effort 读取 COM 对象 Name；失败时返回 fallback。"""
        try:
            name = obj.Name
            if name is not None and str(name).strip():
                return str(name)
        except Exception as exc:  # noqa: BLE001
            self.logger.debug("读取 COM 对象 Name 失败：%s", _com_error_detail(exc))
        return fallback

    def _close_case(self) -> None:
        """关闭本驱动创建的 Case（已保存时不会弹保存确认）。"""
        case = self._case
        self._case = None
        if case is None:
            return

        try:
            try:
                # Close(False) = 不提示保存。
                # 【需要在Excel VBA Object Browser核对】SimulationCase.Close 的参数签名；
                # 回退 Close() 时若 Case 未保存可能弹窗，正常流程在 shutdown 前已 SaveAs。
                case.Close(False)
                self.logger.info("Case 已关闭（不保存）")
            except Exception as first_exc:  # noqa: BLE001
                self.logger.warning(
                    "Close(False) 失败，回退 Close()：%s",
                    _com_error_detail(first_exc),
                )
                case.Close()
                self.logger.info("Case 已关闭（Close()）")
        except Exception as exc:  # noqa: BLE001
            self.logger.warning("关闭 Case 失败：%s", _com_error_detail(exc))
        finally:
            del case

    def _release_com_library(self) -> None:
        """与 connect() 中的 CoInitialize 配对；只在本驱动初始化过 COM 时执行。"""
        if not self._com_initialized:
            return
        try:
            pythoncom.CoUninitialize()
        except Exception as exc:  # noqa: BLE001
            self.logger.debug("CoUninitialize 异常：%s", exc)
        finally:
            self._com_initialized = False
