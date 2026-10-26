# AI 开发记录（Claude Code 会话导出）

- 项目：AI 驱动的 HYSYS 反应器建模（Aspen HYSYS V15）
- 导出时间：2026-10-26
- 内容：开发过程中与 AI 编程助手的协作记录摘要
- 安全说明：本文件不含任何 API key、密码或远程凭据。

## 开发过程摘要

1. 需求：先打通 HYSYS COM 最小闭环（连接、Case、Peng-Robinson、组分、FEED、保存、释放）。
2. 搭建分层架构：`config` / `nlp_parser` / `hysys_driver` / `main` / `agent`，COM 仅存在于驱动层。
3. 通过 HYSYS V15 类型库与黑盒探针确认：
   - 物性包 Setter 使用内部 token `pengrob`；
   - 流体包与组分必须在同一个 Basis 变更中提交；
   - 反应器真实 TypeName：`gibbsreactorop` / `conversionreactorop` / `equilibriumreactorop`；
   - `ComponentMolarFractionValue` 可直接写、`Solver.CanSolve`/`IsSolving` 可用；
   - Gibbs 通过能量流 `HeatFlow.SetValue` 补热；
   - `conversionrxn` / `equilibriumrxn` 可动态创建并配置化学计量、转化率、数值 Ka。
4. 实现两段式 AI：本地 LLM 生成 `plan.json`，远程只执行 plan，不需要远程 LLM key。
5. 验证两个必测场景及三种通用反应器路径。

## 两个必测场景结果

### 甲烷蒸汽重整（Gibbs，FEED 520℃ / 13.5 bar / 摩尔比 1:2.7）

| 出口 | 能量流 duty (×10⁶ kJ/h) | CH4 / H2O / CO / CO2 / H2 |
|---|---|---|
| 710℃ | 38.87 | 9.63% / 38.46% / 4.56% / 6.73% / 40.61% |
| 600℃ | 19.63 | 16.17% / 49.75% / 1.17% / 5.88% / 27.03% |

### 甲苯歧化（Conversion，380℃ / 2.5 MPa / 10000 kg/h）

- 实际转化率：**50.00%**
- 产物（摩尔分数）：苯 25%，邻/间/对二甲苯各 8.33%，甲苯 50%
- 反应集、化学计量与转化率由 plan/代码动态创建，不使用固定模板。

## 通用能力与边界

- 已覆盖：Gibbs、Conversion、Equilibrium（数值 Ka）。
- 已覆盖：31 个 HYSYS 标准组分。
- 未实现：CSTR/PFR 动力学；Equilibrium 的吉布斯自由能形式。
- LLM 在本地运行，远程只执行 plan；新需求需本地 key 生成 plan。
- 信息不足时主要报错/说明假设，尚未做交互式澄清。

## 主要提交

```text
详见 `git log --oneline`；包括架构、COM 驱动、物流组成/求解、
反应器创建、Gibbs 能量流、Conversion/Equilibrium 通用配置、
两段式 LLM plan、组分库扩展与报告。
```

## 完整会话与验证证据路径

原始 Claude Code 会话（JSONL，完整内容，未嵌入本文档）：

- Session A（架构 / COM / 阶段1-2B 探索）：
  `C:\Users\HONOR\.claude\projects\E--desktop-hysys\7f3af75e-e758-4b16-a267-327b23c1121f.jsonl`
- Session B（Kinetic / 报告收尾）：
  `C:\Users\HONOR\.claude\projects\E--desktop-chem-ai-model\fc0355d6-2b85-4ffa-8723-9093b379b77b.jsonl`

output 下的探针 / 验证记录（txt，均为本项目远程运行结果）：

- `E:\desktop\chem_ai_model\output\probe_hysys_property_package_report.txt`
- `E:\desktop\chem_ai_model\output\probe_hysys_stream_composition_report.txt`
- `E:\desktop\chem_ai_model\output\probe_hysys_reactors_report.txt`
- `E:\desktop\chem_ai_model\output\probe_hysys_reactor_add_report.txt`
- `E:\desktop\chem_ai_model\output\probe_hysys_reactor_add_report.txt`
- `E:\desktop\chem_ai_model\output\read_open_operations_report.txt`
- `E:\desktop\chem_ai_model\output\probe_hysys_generic_conversion_report.txt`
- `E:\desktop\chem_ai_model\output\probe_hysys_heat_input_report.txt`
- `E:\desktop\chem_ai_model\output\probe_hysys_equilibrium_report.txt`
- `E:\desktop\chem_ai_model\output\probe_hysys_equilibrium_solve_report.txt`
- `E:\desktop\chem_ai_model\output\probe_hysys_kinetics_report.txt`
- `E:\desktop\chem_ai_model\output\probe_hysys_kinetic_api_report.txt`
- `E:\desktop\chem_ai_model\output\probe_hysys_kinetic_config_report.txt`
- `E:\desktop\chem_ai_model\output\probe_kinetic_diag_report.txt`

> 原始 JSONL 与探针报告可能包含完整命令输出；对外提交前请再次脱敏，
> 不要把 API key、密码、远程主机信息放入公开仓库。

## 说明

本文件是交付物“AI 开发记录”的摘要版；完整原始会话见上方路径。
