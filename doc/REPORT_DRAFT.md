# AI 驱动的 HYSYS 反应器建模 — 技术报告草稿

## 1. 目标与范围

系统接收自然语言反应描述，选择反应器类型并在 Aspen HYSYS V15 中实际创建、配置、求解，返回真实产物结果。

当前覆盖：
- 五类反应器：**Gibbs、Conversion、Equilibrium（数值 Ka）、PFR（平推流）、CSTR（全混流）**；
- 两个必测场景：甲烷蒸汽重整（710/600℃）、甲苯歧化（380℃/2.5MPa/10000kg/h，50%）；
- 31 个 HYSYS 标准组分（组分 ID 来自本机 V15 组分列表实测输出）。

## 2. 技术选型

| 决策 | 选择 | 理由 |
|---|---|---|
| 自动化接口 | HYSYS COM + Python 3.12 x64 + pywin32 311 | 官方 Automation 接口，能真实创建/求解 |
| COM 封装 | 独立 `hysys_driver.py`，上层拿不到 COM 对象 | 隔离误操作，统一生命周期与释放 |
| LLM | DeepSeek（OpenAI 兼容），本地生成 plan.json | key 只留本地，远程 HYSYS 工作站不需要 key |
| 执行方式 | 两段式：本地 LLM plan → 远程执行 plan | 解决远程环境无 LLM key，同时保留自然语言入口 |
| 架构 | 单向依赖 config → parser/driver → main/agent | 分层清晰，便于答辩与扩展 |

## 3. 架构

```text
用户文本
  →（本地）agent.py plan → llm_planner（LLM）→ plans/*.json
  →（远程）agent.py run → scenario_runner → hysys_driver（唯一 COM 边界）
  → HYSYS → .hsc + 结果 JSON
```

模块：`config`（常量/单位/组分/规则）、`nlp_parser`（本地规则占位）、`llm_planner`、`hysys_driver`、`scenario_runner`、`main`/`agent`、诊断工具。

所有 COM 调用只出现在 `hysys_driver`；公开方法只返回字符串/数值/字典。

## 4. 探索过程与关键问题

HYSYS V15 COM 文档不完整，主要问题均通过黑盒探针 + 类型库成员枚举解决：

1. **物性包设置**：`PropertyPackageName` Setter 必须用内部 token **`pengrob`**；读回是显示名 `Peng-Robinson`；旧值 `PengRobinson`/`PR` 均被拒绝。
2. **Basis 提交**：流体包必须与组分在同一个 Basis 变更中完成；空 Basis 的 `EndBasisChange` 会失败。
3. **反应器创建**：`Flowsheet.Operations.Add(name, TypeName)` 真实 TypeName 为 **`gibbsreactorop` / `conversionreactorop` / `equilibriumreactorop`**；CamelCase 名称均失败。
4. **物流组成与求解**：`ComponentMolarFractionValue` 可直接 PUT；`Solver.CanSolve` 可写，`IsSolving` 可读；V15 没有 `Run()`。
5. **热输入**：Gibbs 反应器通过 `EnergyStream` + `HeatFlow.SetValue` 补热，FEED 保持 520℃。
6. **反应配置**：`conversionrxn` / `equilibriumrxn` + `Reactants.Add` + `ReactantStoichCoefValue` + `BaseComponent` / `EquilibriumConstant` + `AssociateFluidPackage`；反应器 `ReactionSet` 绑定后求解。
7. **单位坑**：原始 `*Value` 属性用 HYSYS 默认单位（如 `MolarFlowValue` 是 kgmole/s），代码统一用 `GetValue/SetValue`。

## 5. 必测场景结果

### 甲烷蒸汽重整（Gibbs，FEED 520℃ / 13.5 bar / 1:2.7）

| 出口 | 能量流 duty (×10⁶ kJ/h) | CH4 / H2O / CO / CO2 / H2 |
|---|---|---|
| 710℃ | 38.87 | 9.63% / 38.46% / 4.56% / 6.73% / 40.61% |
| 600℃ | 19.63 | 16.17% / 49.75% / 1.17% / 5.88% / 27.03% |

### 甲苯歧化（Conversion，380℃ / 25 bar / 10000 kg/h）

实际转化率 **50.00%**；苯 25%，邻/间/对二甲苯各 8.33%，甲苯 50%。反应集、化学计量、转化率由代码/LLM plan 动态创建，不使用固定模板。

## 6. 通用能力

- Gibbs：按 plan 的组分/组成/目标出口温度动态创建，能量流 duty 二分反算；
- Conversion：按 plan.reactions 的已配平反应式创建 `conversionrxn`，支持任意转化率；
- Equilibrium：按 plan.reactions 的数值 Ka 创建 `equilibriumrxn`，已用 RWGS（Ka=1）验证闭环；
- Kinetic PFR/CSTR：按 plan.reactions[].kinetics（AE/频率因子/反应级数）与 reactor_spec（体积/管长管径/液位）
  动态创建 `kineticrxn` + `PFReactorop`/`KineticReactorop`；LLM 按考核规则"液相 CSTR、气相 PFR"选型；
- 31 个标准组分，LLM 可从自然语言提取组分/工况/反应式；质量流量按组分库分子量自动换算摩尔流量。

### kinetic 动力学验证（2026-10-26 远程实测）

RWGS：CO₂ + H₂ → CO + H₂O，AE=50000 kJ/kgmole、A=1000 m³/kgmole·s（浓度基准 kgmole/m³），
进料 573 K / 10 bar / 100 kmol·h⁻¹（CO₂:H₂=1:1），V=1 m³：

| 反应器 | 出口温度 | CO/H₂O 摩尔分数 | CO₂ 转化率 | 手算对照 |
|---|---|---|---|---|
| PFR | 562.5 K（吸热降温） | 0.989% / 0.989% | 1.98% | 二级反应解析解 ≈2.1% ✓ |
| CSTR（液位 50%） | 568.0 K | 0.498% / 0.498% | 1.00% | 气相停留体积减半后解析解 ≈1.0% ✓ |

PFR 转化率约为同体积 CSTR 的 2 倍，符合正级数反应器理论；吸热反应出口降温且与转化量正相关，
验证了 HYSYS 求解结果与 COM 写入参数的物理一致性。

## 7. 限制（如实说明）

- Equilibrium 仅支持数值 Ka，尚未接入吉布斯自由能形式；
- kinetic 仅支持浓度基准（Basis=3）幂律动力学；LHHW/异相催化未实现；
- LLM 在本地运行、远程只执行 plan；新需求需本地 key 生成 plan；
- 复杂/不常见组分可能不在 31 个库内；
- 未做交互式缺参澄清，plan 缺字段时直接报错；
- Gibbs 为绝热反应器 + 能量流补热，不是直接温度 setpoint。

## 8. AI 协作

全程使用 AI 编程助手：COM 探针设计、类型库成员枚举、驱动/执行器/LLM schema、测试脚本、文档与 git 提交。开发过程与问题定位保留在 commit history 中。

## 9. Demo 步骤

```text
本地：
  python src\agent.py plan --text "自然语言需求" --out output\plan.json
远程：
  python src\agent.py run --plan plans\gibbs_reformer.json --visible
  python src\agent.py run --plan plans\toluene_disproportionation.json --visible
```
