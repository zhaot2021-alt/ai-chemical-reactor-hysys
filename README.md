<p align="center">
  <img alt="Python" src="https://img.shields.io/badge/Python-3.12%20x64-3776AB?logo=python&logoColor=white">
  <img alt="Aspen HYSYS" src="https://img.shields.io/badge/Aspen%20HYSYS-V15-FF6F00">
  <img alt="Stages 1 &amp; 2A" src="https://img.shields.io/badge/Stages%201%20%26%202A-Verified-2EA043">
  <img alt="COM" src="https://img.shields.io/badge/Boundary-pywin32%20311-4B8BBE">
</p>

# AI 化工反应器建模（Aspen HYSYS V15）

面向「AI 驱动的化工反应器智能建模」考核项目的分层实现。目标是把自然语言需求转化为 HYSYS 中可执行的反应器建模方案，并返回真实的模拟计算结果。

## 里程碑

| 阶段 | 能力 | 状态 |
|---|---|---|
| **Stage 1 — 连通性基础层** | COM 连接、Simulation Case、PR 流体包、10 个标准组分、FEED 物流、`.hsc`、COM 释放 | **2026-10-26 远程验收通过** |
| **Stage 2A — 进料组成与求解** | 摩尔组成写入、Solver、T/P/F/汽化率/组成读回 | **2026-10-26 远程验收通过** |
| **Stage 2B — 反应器建模** | Gibbs 710/600℃、Conversion 380℃/25bar/50% | **2026-10-26 统一入口联调通过** |
| **Stage 3 — AI 编排** | 本地 LLM 生成 plan，远程执行 plan | **2026-10-26 两段式联调通过** |

阶段1/2A/2B 与两段式 AI 入口均已完成；两个必测场景已通过统一入口远程联调。

## 架构概览

```text
自然语言
   │
   ▼
nlp_parser.py         规则解析占位（后续替换为 LLM，schema 不变）
   │
   ▼
main.py               参数校验 / 缺失处理 / 单位换算 / 调度 / CLI
   │
   ▼
hysys_driver.py       唯一 COM 边界：Case、流体包、组分、物流、组成、求解、释放
   │
   ▼
Aspen HYSYS V15  →  output/*.hsc
```

设计原则：

- **单向依赖**：`main → parser / driver`；parser 与 driver 互不依赖；配置集中在 `config.py`。
- **COM 封闭**：只有 `hysys_driver` 导入 `win32com`，上层拿不到原始 COM 对象。
- **可本地调试**：parser、dry-run 不导入驱动层；真实 COM 仅在 HYSYS 工作站执行。
- **单位收口**：外部 ℃/bar/kmol·h⁻¹ ↔ HYSYS K/kPa/kgmole·h⁻¹ 的换算全部在 `config.py`。

详细分层、接口和扩展点见 **[doc/ARCHITECTURE.md](doc/ARCHITECTURE.md)**。

## 阶段1验收记录（2026-10-26）

环境：Windows + **Python 3.12.4 x64** + **pywin32 311** + **Aspen HYSYS V15**。

验收命令：

```bat
python src\test_stage1.py --visible --log-level DEBUG
```

| 验收项 | 实际结果 |
|---|---|
| HYSYS 实例 | 脚本启动新实例，结束时正常 `Quit` |
| Case | `ChemAI_Stage1_20261025_235816` |
| 流体包 | `PR_FluidPkg` |
| 物性包 | Setter token `pengrob` → 读回 `Peng-Robinson` |
| 组分 | 10 个标准组分全部添加成功 |
| FEED | `298.15 K / 100 kPa / 100 kgmole/h` |
| `.hsc` | 已保存（160,165 字节） |
| 后台进程 | `hysys.exe` 正常退出，无残留 |

### 三个关键工程结论

1. **V15 的 `PropertyPackageName` Setter 用内部 token，不是显示名**  
   Peng-Robinson 的内部 token 是 **`pengrob`**；读回值才是显示名 `Peng-Robinson`。
   旧写法 `PengRobinson` / `Peng-Robinson` / `PR` / `HysysPR` 均返回 `E_INVALIDARG`。
2. **流体包在创建时命名**  
   `FluidPackages.Add("PR_FluidPkg")` 成功；不要先 `Add()` 再改 `Name`。
3. **组分必须与流体包在同一个 Basis 变更中完成**  
   `StartBasisChange → Add(name) → PropertyPackageName="pengrob" → Components.Add(...) → EndBasisChange`。  
   空 Basis 直接 `EndBasisChange` 会失败。

## 阶段2A验收记录（2026-10-26）

验收命令：

```bat
python src\test_stage2.py --visible --log-level DEBUG
```

测试工况：10 个标准组分，等摩尔组成，`25℃ / 1 bar / 100 kmol·h⁻¹`。

| 验收项 | 实际结果 |
|---|---|
| Case | `ChemAI_Stage1_20261026_003648` |
| 组成 | 10 × 0.1，读回总和 = 1.0 |
| 求解 | `solved=True`，汽化率 `0.42919` |
| FEED | `298.15 K / 100 kPa / 100 kgmole/h` |
| `.hsc` | 已保存（162,725 字节） |
| 后台进程 | HYSYS 正常 `Quit`，无残留 |

### 阶段2A关键接口结论

1. **`ProcessStream.ComponentMolarFractionValue` 在 V15 可直接写入**  
   GET/PUT 都存在；传入 `tuple` 即可，数组顺序与 `FluidPackage.Components` 一致。
   组成的 `ComponentMolarFraction.Calculate(...)` / `.Values = ...` 也可用，
   `MolarFractionsValue` 不可写。
2. **求解状态用 Solver 与汽化率判断**  
   `SimulationCase.Solver.CanSolve` 可读写，`Solver.IsSolving` 可读，
   V15 该对象没有 `Run()`；流股没有 `State` 属性，用 `VapourFractionValue` 判断求解结果。
3. **原始 `*Value` 属性使用 HYSYS 默认单位**  
   例如 `MolarFlowValue` 是 kgmole/s（100 kgmole/h 显示为 0.02778）；
   读写一律使用 `GetValue/SetValue` 带明确单位。
4. **组成解析仍是规则占位**  
   parser 目前支持“单一组分 → 纯组分”和“各组分等摩尔”；
   明确的摩尔分数/摩尔比交给后续 LLM。

## 目录结构

```text
chem_ai_model/
├── src/
│   ├── config.py          # 配置、标准组分 ID、单位换算（无业务逻辑）
│   ├── hysys_driver.py    # HYSYS COM 驱动（唯一 win32com 边界）
│   ├── nlp_parser.py      # 自然语言解析（阶段1/2A规则占位，可本地运行）
│   ├── main.py            # 入口 / 调度 / 校验 / CLI
│   ├── test_stage1.py     # 阶段1连通模式（无组成）
│   └── test_stage2.py     # 阶段2A组成 + 求解 + 读回
├── tools/
│   ├── probe_hysys_property_package.py  # 物性包 COM API 诊断探针
│   ├── probe_hysys_stream_composition.py # 物流组成/求解 COM 探针
│   ├── probe_hysys_reactors.py           # 反应器对象模型探针
│   ├── probe_hysys_reactor_add.py        # 反应器 Add 类型串快速探针
│   ├── read_open_operations.py           # 只读导出参考 Case 的物流/操作成员
│   └── read_open_fluid_packages.py       # 只读导出 Case/流体包/真实物性包名
├── plans/                 # LLM 生成的 plan.json（两个必测场景）
├── doc/
│   └── ARCHITECTURE.md    # 分层架构、接口、COM 可信度、扩展点
├── output/                # .hsc 输出目录
└── README.md
```

## 快速开始

### 本地（无需 HYSYS）

```bat
python src\nlp_parser.py
python src\test_stage1.py --dry-run
python src\test_stage2.py --dry-run
python src\main.py --dry-run --text "进料为甲烷和水蒸气，各组分等摩尔，温度 520℃，压力 1 bar，摩尔流量 100 kmol/h"
```

### HYSYS 工作站（远程）

环境：Windows + Python 3.12 64 位 + pywin32 311 + Aspen HYSYS V15 + 可用许可证。

```bat
pip install -r requirements.txt

:: 阶段1：连通性（无组成）
python src\test_stage1.py --no-visible

:: 阶段2A：组成 + 求解 + 读回
python src\test_stage2.py --no-visible
```

真实运行会创建 `PR_FluidPkg`、10 个标准组分与 `FEED`；阶段2A 还会设置组成、求解并读回 T/P/F/汽化率，最后保存到 `output/`。建议先关闭已有 HYSYS，使本次实例由脚本启动，结束时自动 `Quit` 并检查残留；若附着已有实例，则默认保留该进程。

需要观察每一步 COM 对象 repr 时，加 `--log-level DEBUG`。

## 两段式 AI 使用

**本地（LLM key 只在本地环境变量；不写入仓库，也不需要放到远程）**

```bat
python src\agent.py plan --text "自然语言需求" --out output\plan.json
```

已预生成两个必测场景的 plan：

```text
plans\gibbs_reformer.json
plans\toluene_disproportionation.json
```

**远程 HYSYS 工作站（不需要 LLM key）**

```bat
python src\agent.py run --plan plans\gibbs_reformer.json --visible --log-level DEBUG
python src\agent.py run --plan plans\toluene_disproportionation.json --visible --log-level DEBUG
```

> Gibbs 场景保持 FEED 520℃，通过 `Q-GIBBS` 能量流补热到 710℃/600℃；Conversion 场景由 plan.reactions 动态创建反应集/反应/反应器，需要 LLM 给出已配平反应式。

### 联合验收结果（2026-10-26）

| 场景 | 关键结果 |
|---|---|
| 甲烷蒸汽重整（Gibbs） | FEED 520℃；能量流 duty 38.87 / 19.63 ×10⁶ kJ/h；710℃：CH4 9.63% / H2O 38.46% / CO 4.56% / CO2 6.73% / H2 40.61%；600℃：CH4 16.17% / H2O 49.75% / CO 1.17% / CO2 5.88% / H2 27.03% |
| 甲苯歧化（Conversion） | 380℃ / 2.5 MPa / 10000 kg/h，实际转化率 **50.00%**；苯 25%，邻/间/对二甲苯各约 8.3%，甲苯 50% |

两个场景均由 `agent.py run` 读取 LLM plan 执行，远程不需要 LLM key；HYSYS 正常 Quit，无后台残留。

## 单位约定

| 位置 | 单位 |
|---|---|
| 用户输入 / 解析输出 | ℃、bar、kmol·h⁻¹ |
| 驱动层 / HYSYS | K、kPa、kgmole·h⁻¹（kgmole ≡ kmol） |

换算函数全部集中在 `src/config.py`。

## 诊断工具

- `tools/read_open_fluid_packages.py`：只读列出当前打开 Case 的流体包、`PropertyPackageName`、`PropertyPackage.TypeName` 和组分。
- `tools/probe_hysys_property_package.py`：导出 COM 类型成员表，验证 `Add(name)`、组分/物性包顺序矩阵和 Peng-Robinson token。
- `tools/probe_hysys_stream_composition.py`：验证物流组成写入方式、`Solver.CanSolve`、`IsSolving` 和求解后的流股状态。
- `tools/probe_hysys_reactors.py`：导出 Flowsheet/Operations/ReactionPackage 成员，并尝试创建 Gibbs / Conversion / Equilibrium 反应器。
- `tools/probe_hysys_reactor_add.py`：用多种大小写/短名快速定位 V15 的 reactor Add 类型字符串。
- `tools/read_open_operations.py`：只读导出手动参考 Case 中的物流参数、操作 TypeName 和反应器成员。

阶段1/2A排查中已确认 `pengrob` token、Basis 提交顺序、`ComponentMolarFractionValue` 写入和 Solver 状态接口；工具保留用于后续版本或新环境复测。

## 当前边界（阶段2B）

- 阶段1测试故意不含组成，用于验证 COM 连通；阶段2A测试含等摩尔组成并完成求解。
- Gibbs/Conversion 反应器已可创建/连接/求解；Equilibrium 尚未接入执行器。
- Gibbs 通过 `Q-GIBBS` 能量流补热，FEED 保持 520℃；duty 由目标出口温度反算。
- Conversion 执行器支持 plan.reactions 中的组分配平、BaseComponent 与转化率；LLM 未给出反应式时拒绝猜测。
- LLM 规划在当前机器生成 plan；远程执行不依赖 key。
- HYSYS COM 无完整公开文档：仅使用有公开文档/工程代码/远程实测印证的接口；不确定项在代码中标注 **【需要在Excel VBA Object Browser核对】** 并带回退。
- 默认优先附着已运行的 HYSYS 实例；本次启动的实例在退出时会 `Quit` 并检查 `hysys.exe` 残留。
