# HYSYS AI 反应器建模 — 阶段1 架构说明

> 范围：**阶段1 仅打通** HYSYS 连接 → 新建 Case → 流体包（Peng-Robinson）→ 10 个标准组分 → FEED 物流 → 保存 `.hsc` → COM 释放。
> **本阶段不创建任何反应器**（Conversion / Equilibrium / Gibbs 留到阶段2），也不设置物流组成、不求解。

---

## ① 模块依赖图（文字版）

```text
                         ┌──────────────────────┐
                         │      config.py       │
                         │ 纯配置 + 单位换算     │  ← 无本地模块依赖
                         └──────────┬───────────┘
                                    │ 被所有模块导入（单向）
              ┌─────────────────────┴─────────────────────┐
              ▼                                           ▼
   ┌────────────────────┐                    ┌─────────────────────────┐
   │   nlp_parser.py    │                    │    hysys_driver.py      │
   │ 规则解析（本地可跑）│                    │ COM 驱动（仅远程可跑）   │
   │ 不 import COM      │                    │ 不 import parser/main   │
   └─────────┬──────────┘                    └────────────▲────────────┘
             │                                            │
             └────────────────────┬───────────────────────┘
                                  ▼
                        ┌──────────────────┐
                        │     main.py      │  调度 / 校验 / 单位换算 / CLI
                        └────────┬─────────┘
                                 │
                        ┌────────▼─────────┐
                        │  test_stage1.py  │  固定文本联调入口（非 pytest）
                        └──────────────────┘
```

依赖规则：

| 约束 | 说明 |
|---|---|
| 单向依赖 | `main → parser`、`main → driver`；`parser` 与 `driver` **互不依赖** |
| 公共配置 | 所有模块可导入 `config`；`config` 不导入任何本地模块 |
| COM 封闭 | 只有 `hysys_driver` 导入 `win32com`；上层永远拿不到原始 COM 对象 |
| 本地优先 | `parser` / `main --dry-run` / `test_stage1.py --dry-run` 不导入驱动层 |

---

## ② 模块职责与对外接口

### 2.1 `src/config.py` — 配置层

只放常量、数据映射和单位换算，不写业务判断、不碰 COM。

| 类别 | 对外内容 |
|---|---|
| 路径 / 日志 | `PROJECT_ROOT`、`OUTPUT_DIR`、`LOG_LEVEL`、`LOG_FORMAT` |
| COM 生命周期 | `HYSYS_PROGID`、`HYSYS_ATTACH_EXISTING`、`HYSYS_QUIT_ATTACHED`、`HYSYS_SHUTDOWN_WAIT_S` |
| HYSYS 内部单位 | `HYSYS_TEMPERATURE_UNIT="K"`、`HYSYS_PRESSURE_UNIT="kPa"`、`HYSYS_MOLAR_FLOW_UNIT="kgmole/h"` |
| 标准组分 | `STANDARD_COMPONENT_IDS`（10 个）、`COMPONENT_ALIASES`、`COMPONENT_ID_FALLBACKS` |
| 流体包 | `FLUID_PACKAGE_NAME`、`PROPERTY_PACKAGE_NAME="PengRobinson"`、备用拼写 |
| 默认进料 | `DEFAULT_FEED_CONDITIONS`（25℃ / 1 bar / 100 kmol·h⁻¹） |
| 单位换算 | `c_to_k`、`k_to_c`、`bar_to_kpa`、`kpa_to_bar`、`mpa_to_bar`、`kmolh_to_kgmoleh`、`pressure_to_bar`、`molar_flow_to_kmolh` |
| LLM 占位 | `LLM_CONFIG`（阶段1不发起网络请求） |
| 反应器规则 | `REACTOR_RULES`（纯数据，阶段2消费；阶段1无反应器代码） |

**组分 ID 对照（必须使用 HYSYS 标准 Sim Name）**

| 用户写法 | HYSYS 标准 ID | 说明 |
|---|---|---|
| METHANE | `Methane` | |
| STEAM | `H2O` | HYSYS 无 “STEAM”；`Water` 只是 H2O 同义名 |
| H2 | `Hydrogen` | |
| CO / CO2 | `CO` / `CO2` | |
| TOLUENE / BENZENE | `Toluene` / `Benzene` | |
| XYLENE-O/M/P | `o-Xylene` / `m-Xylene` / `p-Xylene` | |

### 2.2 `src/nlp_parser.py` — 自然语言解析层

阶段1为**正则规则占位**；阶段2可替换为 LLM，但保持输出 schema 不变。

| 对外接口 | 说明 |
|---|---|
| `parse(text: str) -> ParsedRequest` | 唯一入口；返回结构化字典（`dict` 子类语义） |

输出字段：

| 字段 | 类型 | 单位 / 取值 |
|---|---|---|
| `components` | `list[str]` | HYSYS 标准 ID，按文中出现顺序去重 |
| `temperature_c` | `float \| None` | **℃**（外部标准） |
| `pressure_bar` | `float \| None` | **bar**（MPa/kPa/atm 已换算） |
| `molar_flow_kmolh` | `float \| None` | **kmol·h⁻¹** |
| `reactor_type_hint` | `str \| None` | 仅提示，**阶段1不消费** |
| `reactor_reason` | `str \| None` | 提示理由 |
| `missing_info` | `list[str]` | 缺失字段名，如 `molar_flow_kmolh` |
| `notes` | `list[str]` | 歧义、单位未支持、多值等提醒 |
| `parser` | `str` | 当前为 `rule_placeholder_v1` |

已知限制：

- 多个温度/压力时取**第一个**，并写入 `notes`。场景1同时有“进料 520℃ / 出口 710℃”时必须由阶段2 LLM 做语义区分。
- 质量流量（kg/h）暂不换算为摩尔流量（需分子量/组成），只提示。
- 通用“二甲苯”有歧义，不擅自映射为某一异构体，写入 `notes`。
- 不支持表压（barg）、N m³/h 等；需要时在 `config` 增加换算后再扩展解析。

### 2.3 `src/hysys_driver.py` — COM 驱动层

只封装与 HYSYS 的基础交互；**无 RAG、无 LLM、无反应器类型判断、无反应器创建**。

| 公开成员 | 签名 / 返回值 | 说明 |
|---|---|---|
| `HYSYSDriver(...)` | `visible`、`attach_existing`、`quit_attached` | 构造时不连接 COM |
| `connect()` | `None` | 优先附着已运行实例；否则 Dispatch + 重试 |
| `shutdown()` | `None` | 幂等释放：关 Case → Quit（视归属）→ CoUninitialize → tasklist 检查 |
| `new_case(name=None)` | `str` | 新建空白 Case，返回名称 |
| `create_fluid_package(name=None, property_package=None)` | `str` | 创建流体包 + 设置 Peng-Robinson |
| `add_components(ids)` | `tuple[str, ...]` | 添加标准组分，重复自动跳过 |
| `create_material_stream(name, T_K, P_kPa, F_kgmole/h)` | `dict` | 建物流并设置内部单位参数，返回安全状态字典 |
| `read_stream_values(name)` | `dict` | 只读温度/压力/摩尔流量，**不暴露 COM** |
| `save_case(path=None)` | `str` | 保存 `.hsc`，返回绝对路径 |
| `is_connected` / `case_name` / `fluid_package_name` / `property_package_name` / `component_ids` / `stream_names` | 只读属性 | 仅供调度层展示 |
| `HYSYSDriverError` | 异常类 | 封装 COM 错误、状态错误、保存错误 |

上下文管理：

```python
with HYSYSDriver(visible=True) as drv:
    drv.new_case("Case1")
    ...
# 退出 with 自动 shutdown()
```

**COM 对象封装原则**：`_app`、`_case`、`_fluid_pkg`、`_streams` 全部私有；公开方法只返回字符串、数值、元组、字典。这样上层不会误持有/误操作 COM 引用，也避免释放顺序出错。

### 2.4 `src/main.py` — 程序入口 / 调度层

| 对外接口 | 说明 |
|---|---|
| `run_stage1(text, *, case_name, save_path, visible, strict, dry_run) -> dict` | 阶段1完整链路；`dry_run=True` 不连接 HYSYS |
| `main(argv=None) -> int` | CLI，返回进程退出码 |
| `Stage1Error` | 参数缺失 / 校验失败 / 驱动不可用 |

职责：

1. 调 `nlp_parser.parse()`；
2. 校验缺失参数：`--strict` 报错，否则使用 `config` 默认值并告警；
3. 用 `config` 的函数把 **℃/bar/kmol·h⁻¹ → K/kPa/kgmole·h⁻¹**；
4. 在 `try/finally` 中调用驱动，**无论如何都 `shutdown()`**；
5. 返回可序列化摘要（供 CLI、测试、后续 Web 层使用）。

不在此层写：COM 调用、解析正则、反应器创建。

### 2.5 `src/test_stage1.py` — 联调测试入口

固定测试文本覆盖 10 个组分与三个工艺参数；先本地校验解析结果，再调用 `main.run_stage1()`：

- `--dry-run`：本地执行，不导入驱动层；
- 不带 `--dry-run`：远程工作站真实建 Case / 流体包 / 组分 / FEED / `.hsc`。

---

## ③ 关键数据流

```text
用户文本
  │
  ▼
nlp_parser.parse()                    ← 本地；输出 ℃ / bar / kmol·h⁻¹
  │  ParsedRequest + missing_info
  ▼
main._resolve_feed()                  ← 本地；缺失补默认值 / strict 报错
  │
  ▼
main._to_internal_units()             ← 仅调用 config 换算函数
  │  K / kPa / kgmole·h⁻¹
  ▼
HYSYSDriver（仅远程）
  ├─ connect()
  ├─ new_case()
  ├─ create_fluid_package("PengRobinson")
  ├─ add_components([...10 个标准 ID...])
  ├─ create_material_stream("FEED", T_K, P_kPa, F)
  ├─ save_case("output/xxx.hsc")
  └─ shutdown()  ← finally 保证执行
  │
  ▼
摘要 dict（JSON 可序列化）
```

---

## ④ 单位策略（唯一收口）

| 层 | 使用的单位 | 说明 |
|---|---|---|
| 用户 / `nlp_parser` 输出 | ℃、bar、kmol·h⁻¹ | 外部标准；MPa/kPa/atm 由 parser 经 `config.pressure_to_bar` 归一为 bar |
| `main` | 调 `config` 换算 | `c_to_k` / `bar_to_kpa` / `kmolh_to_kgmoleh` |
| `hysys_driver` | K、kPa、kgmole·h⁻¹ | 只接受内部单位，不自行换单位；`SetValue(..., "K"/"kPa"/"kgmole/h")` |
| HYSYS | K、kPa、kgmole·h⁻¹ | `GetValue` 读回验证 |

> HYSYS 的 **kgmole ≡ kmol**（1 kgmole = 1000 mol），数值相同，但代码里用显式函数表达语义，避免单位歧义。

---

## ⑤ COM API 可信度清单

### 已在公开文档 / 工程代码中印证

| API | 用途 |
|---|---|
| `Dispatch("HYSYS.Application")` / `GetActiveObject` | 启动 / 附着 HYSYS |
| `SimulationCases.Add(name)`、`Case.Visible` | 新建并显示 Case |
| `BasisManager.StartBasisChange()` / `EndBasisChange()` | Basis 修改通知（官方 Customization Guide 要求） |
| `BasisManager.FluidPackages.Add()` | 新建流体包（**不带名称参数**） |
| `FluidPackage.PropertyPackageName` | 物性包，内部名 `PengRobinson` |
| `FluidPackage.Components.Add("Methane")` | 添加标准组分 |
| `Flowsheet.MaterialStreams.Add("FEED")` | 新建物流（VBA 公开示例） |
| `Stream.Temperature/Pressure/MolarFlow.SetValue(value, unit)` / `.GetValue(unit)` | 带单位读写 |
| `SimulationCase.SaveAs(path)` | 保存 `.hsc` |
| `Application.Quit()` | 退出 HYSYS |

### 有依据但需核对（代码中已用 try/except 隔离并给回退）

| 待核对项 | 当前处理 | 标记 |
|---|---|---|
| `SimulationCases.Add(name)` 的参数签名 | 先试带名，再回退无参 `Add()` | 【需要在Excel VBA Object Browser核对】 |
| `FluidPackage.Name` 是否可写（流体包改名） | best-effort；失败则保留 HYSYS 默认名（如 `Basis-1`） | 【需要在Excel VBA Object Browser核对】 |
| `FluidPackage.PropertyPackageName` 可写性 / V15 内部名 | 单次尝试 `PengRobinson`，且必须**先于** `FluidPackage.Name` 改名；失败则提示运行 `tools/probe_hysys_property_package.py` | 【需要在Excel VBA Object Browser核对】 |
| `SimulationCase.Close(False)` 参数签名 | 失败回退 `Close()`；正常流程先 `SaveAs`，避免弹保存框 | 【需要在Excel VBA Object Browser核对】 |
| `Flowsheet.FluidPackage = fp`（Make Current） | **未实现**：公开资料显示该属性主要用于读取；阶段1只有一个流体包，HYSYS 自动关联 | 待确认后再扩展 |
| `Components.Add("H2O")` 对水的接受度 | 标准 ID 用 `H2O`；失败时尝试官方同义名 `Water` 并告警 | 回退策略，不改变标准 ID |
| `tasklist` 中进程名 | 以包含 `hysys` 判断，残留时提示 `taskkill /F /IM hysys.exe /T` | 进程名为常识性检查，非 COM |

**不使用的猜测性接口**：不调用未确认的 `PropertyPackages.Add`、不设置未确认可写的 `Flowsheet.FluidPackage`、不假设反应器对象模型。

---

## ⑥ 异常处理与 COM 释放

```text
main.run_stage1()
  try:
      driver = HYSYSDriver()
      driver.connect()
      ... 阶段1 建模 ...
      driver.save_case()
  finally:
      driver.shutdown()
```

`shutdown()` 顺序：

1. 释放 `_streams` / `_fluid_pkg` 引用，`gc.collect()`；
2. `Case.Close(False)`（回退 `Close()`）；
3. 若**本次启动的实例**（或 `quit_attached=True`）→ `Application.Quit()`；附着用户已有实例则保留；
4. `pythoncom.CoUninitialize()`（与 `connect()` 的 `CoInitialize()` 配对）；
5. `sleep` 后用 `tasklist` 检查；残留则打印 `taskkill /F /IM hysys.exe /T`；
6. 清理内部状态，允许同一对象重新 `connect()`（新会话）。

异常处理：

- 驱动层把 COM 错误统一包装为 `HYSYSDriverError`，消息含 `HRESULT` 与 `excepinfo`；
- 调度层 `Stage1Error` 负责参数/输入问题；
- CLI 顶层捕获并返回非 0 退出码；
- `finally` 始终释放，避免后台 `hysys.exe` 残留。

> 附着已有实例时**故意不 Quit**（`HYSYS_QUIT_ATTACHED=False`），防止误杀你正在使用的 HYSYS 会话。需要“本次运行结束必退”时，把该配置改为 `True`，或确保运行时没有已有实例（本次 Dispatch 启动的实例会被 Quit）。

---

## ⑦ 【本地可调试】与【仅远程可运行】

| 能力 | 运行位置 | 依赖 |
|---|---|---|
| `python src/nlp_parser.py` | 任意机器 | 仅 Python + config |
| `python src/test_stage1.py --dry-run` | 任意机器 | 仅 Python + config + parser |
| `python src/main.py --dry-run --text "..."` | 任意机器 | 同上（驱动层延迟导入） |
| `import hysys_driver` | 任意 Windows | pywin32（不连接 HYSYS） |
| 不带 `--dry-run` 的联调 | **HYSYS 工作站 / RDP** | HYSYS V15 + pywin32 311 + 许可证 |

代码中以注释显式标注：

- 【本地可调试】：`config`、`nlp_parser`、`main` 的解析/校验/dry-run；
- 【仅远程可运行】：`hysys_driver` 全部 COM 调用、`main.run_stage1` 的 live 分支。

---

## ⑧ 扩展点（阶段2+ 不改动现有连通性代码）

```text
阶段2新增（规划，当前未实现）：
  src/reactor_specs.py      反应器参数数据类（转化率、平衡常数、动力学参数…）
  hysys_driver 新方法       create_conversion_reactor / create_equilibrium_reactor / create_gibbs_reactor
                            （必须先在 Object Browser 核对 Add 接口与参数设置方式）
  nlp_parser               ParsedRequest 增加 reactor_spec 字段；内部换成 LLM，保持 schema
  main                     依据 reactor_type 做分发表，调用对应驱动方法
  src/evaluator.py          读取产物组成/温度/收敛状态，计算评估指标
  src/report.py             输出 Markdown / JSON / PDF 报表
```

扩展原则：

1. **只加不改**：新增反应器方法复用 `_basis_change`、`_require_*`、异常与释放机制，不改阶段1的 `connect/new_case/create_fluid_package/add_components/save_case`；
2. **解析与执行解耦**：LLM 或规则解析只输出同一个 `ParsedRequest` schema，调度层不感知解析实现；
3. **新增组分**：只改 `config.STANDARD_COMPONENT_IDS` / `COMPONENT_ALIASES`；
4. **新增评估/报表**：作为独立模块消费驱动层返回的安全字典，不访问 COM；
5. **反应器规则**：规则文本/关键词仍放 `config`，判断逻辑放 parser / LLM。

> 阶段1**没有**任何 Conversion / Equilibrium / Gibbs 创建代码；上述仅为架构预留说明。

---

## ⑨ Git 分阶段提交建议

当前目录尚未初始化 Git；按功能拆分提交便于答辩展示演进过程：

```bash
git init
git add .gitignore requirements.txt
git commit -m "chore: scaffold repo and python dependencies"

git add src/config.py
git commit -m "chore(config): add constants, component IDs and unit conversion"

git add src/hysys_driver.py
git commit -m "feat(driver): add stage-1 HYSYS COM driver with safe cleanup"

git add src/nlp_parser.py
git commit -m "feat(nlp): add rule-based parser placeholder"

git add src/main.py
git commit -m "feat(main): orchestrate stage-1 pipeline and CLI"

git add src/test_stage1.py
git commit -m "test(stage1): add fixed-text dry-run and live integration entry"

git add doc/ARCHITECTURE.md README.md
git commit -m "docs: add stage-1 architecture and usage"
```

> 提交前先本地跑 `python src/test_stage1.py --dry-run`，远程跑通后再提交驱动层。

---

## ⑩ 运行方式

```bat
cd E:\desktop\chem_ai_model

:: 本地：解析自检
python src\nlp_parser.py

:: 本地：端到端 dry-run（不连接 HYSYS）
python src\test_stage1.py --dry-run

:: 本地：CLI dry-run
python src\main.py --dry-run --text "进料为甲烷和水蒸气，温度 520℃，压力 1 bar，摩尔流量 100 kmol/h"

:: 远程工作站：真实联调
python src\test_stage1.py --no-visible
python src\main.py --text "进料为甲烷和水蒸气，温度 520℃，压力 1 bar，摩尔流量 100 kmol/h" --save output\case.hsc

:: 需要打印每一步 COM 对象 repr（便于答辩/排查）
python src\test_stage1.py --no-visible --log-level DEBUG
```

真实联调成功标准：

1. HYSYS 弹出/连接成功，生成 Case；
2. 流体包创建，物性包为 Peng-Robinson；
3. 10 个标准组分无失败；
4. `FEED` 物流的 T/P/F 读回与输入一致（298.15 K / 100 kPa / 100 kgmole/h 的测试值）；
5. `output\*.hsc` 存在；
6. 本次启动的 HYSYS 进程退出，`tasklist` 无残留（附着已有实例时进程保留是预期行为）。

---

## ⑪ 阶段1 已知限制

- 物流**未设置组成、未求解**；只验证 COM 能力，不代表工艺结果；
- 未创建任何反应器；反应器类型字段仅规则占位提示；
- 规则解析对多温度场景（进料 vs 出口）无法区分，阶段2必须用 LLM 语义解析；
- 公斤/小时流量未换算；
- 流体包名称 best-effort，失败时 HYSYS 默认命名（如 `Basis-1`）；
- 仅支持单一流体包；多流体包 Make Current 待核对；
- HYSYS COM 部分接口签名需在 V15 Object Browser 核对（见 ⑤）。

---

## ⑫ 远程联调记录（2026-10-26）

已在远程 HYSYS V15 工作站（Python 3.12.4 64 位、pywin32 311）完成：

1. 本地/远程 dry-run 通过（10 标准组分、单位换算正确）；
2. COM 连接、`SimulationCases.Add`、`FluidPackages.Add`、`FluidPackage.Name` 成功；
3. 在 `PropertyPackageName = "PengRobinson"/"Peng-Robinson"` 处返回
   `HRESULT=0x80020009`、`excepinfo scode=0x80070057 (E_INVALIDARG)`，
   随后 `EndBasisChange` 返回 `E_FAIL`。

观察到改名发生在设置物性包之前。当前处理：

- 驱动改为**先设置物性包、成功后再 best-effort 改名**；
- 物性包名只尝试一次，避免第二次失败把 Basis 状态进一步污染；
- 新增 `tools/probe_hysys_property_package.py`：
  - 导出 `BasisManager` / `FluidPackages` / `FluidPackage` 的 COM 成员表
    （含 `GET` / `PUT` / `PUTREF`，判断属性是否可写）；
  - 在独立临时 Case 中测试多种设置顺序和内部名
    （`PengRobinson` / `PR` / `HysysPR` / `SRK`），输出到
    `output/probe_hysys_property_package_report.txt`。

修复顺序只是当前假设，需要远程复测确认；若仍失败，以探针输出为准修正。

---

## 参考资料

- AspenTech HYSYS Customization Guide（BasisManager / FluidPackage / Components / ProcessStream）
- OptiMaL-PSE-Lab / Sketch2Simulation `hysys_template_base.py`（Python win32com 基础模板）
- Aspen HYSYS 2004.2 手册（Automation 中物性包须使用内部名称的说明）
- Cheresources 论坛 VBA 示例（`Flowsheet.MaterialStreams.Add("Name")`）
- NTNU / UFRJ 等公开论文中的 `ProcessVariable.SetValue/GetValue` 示例
