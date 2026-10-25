# AI 化工反应器建模（Aspen HYSYS V15）

面向「AI 驱动的化工反应器智能建模」考核项目的分层实现。目标是由自然语言驱动 HYSYS 完成反应器建模并返回真实计算结果。

## 当前进度：阶段1（连通性基础层）

```text
自然语言 → nlp_parser（规则占位）→ main 调度/校验/单位换算
        → hysys_driver（HYSYS COM）→ .hsc
```

阶段1 已实现：

- 连接 / 启动 HYSYS，异常安全退出
- 新建 Simulation Case
- 创建流体包并设置 Peng-Robinson（内部名 `PengRobinson`）
- 添加 10 个 HYSYS 标准组分
- 新建物料流并设置温度 / 压力 / 摩尔流量
- 保存 `.hsc`，显式释放 COM（防 `hysys.exe` 残留）

**阶段1 未实现（阶段2 再做）**：物流组成与求解、Conversion / Equilibrium / Gibbs 反应器、结果评估与报表、LLM 接入。

## 目录结构

```text
chem_ai_model/
├── src/
│   ├── config.py          # 配置 + 标准组分 ID + 单位换算（无业务逻辑）
│   ├── hysys_driver.py    # HYSYS COM 驱动（仅此文件导入 win32com）
│   ├── nlp_parser.py      # 自然语言解析（阶段1规则占位，本地可跑）
│   ├── main.py            # 入口 / 调度层 + CLI
│   └── test_stage1.py     # 阶段1 联调入口（dry-run / 真实 HYSYS）
├── tools/
│   └── probe_hysys_property_package.py  # 物性包 COM API 诊断探针
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

真实运行会新建 Case、创建 `PR_FluidPkg`（名称 best-effort）与 `FEED` 物流，并保存到 `output/`。

需要查看每一步的 COM 对象 repr 时，加 `--log-level DEBUG`：

```bat
python src\test_stage1.py --no-visible --log-level DEBUG
```

如果在物性包设置处遇到 COM 错误（如 `0x80070057`），运行诊断探针并把
`output/probe_hysys_property_package_report.txt` 发回来：

```bat
python tools\probe_hysys_property_package.py
```

## 单位约定

| 位置 | 单位 |
|---|---|
| 用户输入 / 解析输出 | ℃、bar、kmol·h⁻¹ |
| 驱动层 / HYSYS | K、kPa、kgmole·h⁻¹（kgmole ≡ kmol） |

换算函数全部集中在 `src/config.py`。

## 重要说明

- HYSYS COM 无完整公开文档：仅使用有公开文档/工程代码印证的接口；
  不确定的调用在代码中标注 **【需要在Excel VBA Object Browser核对】** 并带 try/except 回退。
- 默认优先附着已运行的 HYSYS 实例；若为本次启动的实例，退出时会 `Quit` 并检查 `hysys.exe` 残留。
- 阶段1 的物流不含组成、不求解；本阶段不创建任何反应器。

详细设计见 [doc/ARCHITECTURE.md](doc/ARCHITECTURE.md)。
