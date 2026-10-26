<p align="center">
  <img alt="Python" src="https://img.shields.io/badge/Python-3.12%20x64-3776AB?logo=python&logoColor=white">
  <img alt="Aspen HYSYS" src="https://img.shields.io/badge/Aspen%20HYSYS-V15-FF6F00">
  <img alt="Stage 1" src="https://img.shields.io/badge/Stage%201-Verified-2EA043">
  <img alt="COM" src="https://img.shields.io/badge/Boundary-pywin32%20311-4B8BBE">
</p>

# AI 化工反应器建模（Aspen HYSYS V15）

面向「AI 驱动的化工反应器智能建模」考核项目的分层实现。目标是把自然语言需求转化为 HYSYS 中可执行的反应器建模方案，并返回真实的模拟计算结果。

## 里程碑

| 阶段 | 能力 | 状态 |
|---|---|---|
| **Stage 1 — 连通性基础层** | COM 连接、Simulation Case、PR 流体包、10 个标准组分、FEED 物流、`.hsc`、COM 释放 | **2026-10-26 远程验收通过** |
| Stage 2 — 物流与反应器 | 物流组成/求解、Conversion / Equilibrium / Gibbs 反应器、结果读取 | 规划中 |
| Stage 3 — AI 编排 | LLM 解析、反应器规则判断、评估指标、报告输出 | 规划中 |

阶段1的物流**尚未设置组成、未求解**；反应器代码尚未开始。

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
hysys_driver.py       唯一 COM 边界：Case、流体包、组分、物流、保存、释放
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

## 目录结构

```text
chem_ai_model/
├── src/
│   ├── config.py          # 配置、标准组分 ID、单位换算（无业务逻辑）
│   ├── hysys_driver.py    # HYSYS COM 驱动（唯一 win32com 边界）
│   ├── nlp_parser.py      # 自然语言解析（阶段1规则占位，可本地运行）
│   ├── main.py            # 入口 / 调度 / 校验 / CLI
│   └── test_stage1.py     # 阶段1联调入口（dry-run / 真实 HYSYS）
├── tools/
│   ├── probe_hysys_property_package.py  # 物性包 COM API 诊断探针
│   └── read_open_fluid_packages.py      # 只读导出 Case/流体包/真实物性包名
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
python src\main.py --dry-run --text "进料为甲烷和水蒸气，温度 520℃，压力 1 bar，摩尔流量 100 kmol/h"
```

### HYSYS 工作站（远程）

环境：Windows + Python 3.12 64 位 + pywin32 311 + Aspen HYSYS V15 + 可用许可证。

```bat
pip install -r requirements.txt
python src\test_stage1.py --no-visible
```

真实运行会创建 `PR_FluidPkg` 流体包、10 个标准组分与 `FEED` 物流，并保存到 `output/`。建议先关闭已有 HYSYS，使本次实例由脚本启动，结束时自动 `Quit` 并检查残留；若附着已有实例，则默认保留该进程。

需要观察每一步 COM 对象 repr 时，加 `--log-level DEBUG`。

## 单位约定

| 位置 | 单位 |
|---|---|
| 用户输入 / 解析输出 | ℃、bar、kmol·h⁻¹ |
| 驱动层 / HYSYS | K、kPa、kgmole·h⁻¹（kgmole ≡ kmol） |

换算函数全部集中在 `src/config.py`。

## 诊断工具

- `tools/read_open_fluid_packages.py`：只读列出当前打开 Case 的流体包、`PropertyPackageName`、`PropertyPackage.TypeName` 和组分。
- `tools/probe_hysys_property_package.py`：导出 COM 类型成员表，验证 `Add(name)`、组分/物性包顺序矩阵和 Peng-Robinson token。

阶段1排查中已确认 V15 的 `pengrob` token 与 Basis 提交顺序；这两个工具保留用于后续版本或新环境复测。

## 阶段1边界

- 物流不含组成、不求解；本阶段只验证 COM 连通与基础建模。
- 未创建任何反应器；反应器类型字段仅为规则占位提示。
- HYSYS COM 无完整公开文档：仅使用有公开文档/工程代码/远程实测印证的接口；不确定项在代码中标注 **【需要在Excel VBA Object Browser核对】** 并带回退。
- 默认优先附着已运行的 HYSYS 实例；本次启动的实例在退出时会 `Quit` 并检查 `hysys.exe` 残留。
