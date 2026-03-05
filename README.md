# 工作手册（WORK_MANUAL）

## 1. 项目概览

### 1.1 用途
- 读取台账与参数（Excel）
- 计算价格、供电边界、燃料与碳排参数矩阵
- 调用调度优化模块分配机组供电量
- 计算财务、税务、现金流、债务滚动结果
- 导出结果矩阵到 `result_output.xlsx`

### 1.2 输入与输出
- 输入目录：`data/`
- 核心输入文件：
  - `data/01_基础静态台账.xlsx`
  - `data/02_类型级动态参数.xlsx`
  - `data/03_机组级动态参数.xlsx`
  - `data/04_燃料与碳排参数.xlsx`
- 配置文件：`data_config.py`
- 输出文件：`result_output.xlsx`（位于项目根目录）

### 1.3 主要流程
1. `main.py` 读取 4 个 Excel 文件中的各 sheet。
2. 组装 `unit_general_ledger`，补全字段并初始化固定资产/借款状态。
3. 基于规划年序列生成动态矩阵：
   - 价格矩阵
   - 供电边界矩阵
   - 燃耗/燃料成本矩阵
   - 碳排放与碳捕集强度矩阵
4. 调用 `Dispatch.solve_optimal_unit_generation_allocation` 求解供电量分配。
5. 计算年度经营结果（收入、成本、净值、税项）。
6. 按年份循环做时间耦合结转（税务、利润、现金流、债务）。
7. 汇总结果矩阵并导出到 Excel 多 sheet。

---

## 2. 环境与依赖

### 2.1 Python 版本
从二进制模块文件名 `Dispatch.cp313-win_amd64.pyd` 推断：
- 目标运行环境为 **Python 3.13（Windows x64）**。

### 2.2 关键依赖包
- `pandas`
- `numpy`
- `pathlib`（标准库）
- 本地扩展模块：`Dispatch.cp313-win_amd64.pyd`

### 2.3 安装方式（基于现有文件）
参考安装命令（示例）：
```bash
pip install numpy pandas
```

---

## 3. 目录结构说明

当前 workspace 关键结构：

- `main.py`：程序入口，串联全流程。
- `data_config.py`：全局配置（规划年限、补贴机组类型、惩罚因子）。
- `tool_preprocess.py`：前处理与矩阵构建函数。
- `tool_postprocess.py`：后处理、税务利润结转、现金债务结转函数。
- `Dispatch.cp313-win_amd64.pyd`：调度优化求解模块（二进制扩展）。
- `data/`：输入数据目录（四个主输入 Excel）。
- `result_output.xlsx`：运行后输出结果文件。
- `data/data_static.xlsx`：在当前 `main.py` 中未被引用。

---

## 4. 快速上手

### 4.1 最小可运行步骤
1. 准备环境（建议 Python 3.13 x64）。
2. 安装 `numpy`、`pandas`。
3. 确认项目根目录存在：
   - `Dispatch.cp313-win_amd64.pyd`
   - `data/01_基础静态台账.xlsx`
   - `data/02_类型级动态参数.xlsx`
   - `data/03_机组级动态参数.xlsx`
   - `data/04_燃料与碳排参数.xlsx`
4. 在项目根目录执行：
```bash
python main.py
```
5. 检查输出：`result_output.xlsx`。

### 4.2 输入数据准备要点
- 各 Excel 的 sheet 名称必须与代码中的 `sheet_name` 完全一致。
- 关键列名必须与代码索引字段一致（见第 10 章字段说明）。

---

## 5. 配置说明

### 5.1 `data_config.py`
配置项来源：`data_config` 字典。
- `规划时间跨度`：决定规划年份数 `T`。
- `规划起始年份`：决定起始年份。
- `存量带补贴机组类型`：影响电价补贴逻辑。
- `持续性补贴机组类型`：影响电价补贴逻辑。
- `缺电惩罚因子`：调度优化中的缺电惩罚。
- `弃电惩罚因子`：调度优化中的弃电惩罚。

影响范围：
- 直接影响 `planning_year_sequence`、`is_in_service`、价格矩阵、调度结果与全部后续财务矩阵。

### 5.2 Excel 配置型数据
- 多数业务参数在 4 个 Excel 中维护（按 `sheet_name` 读取）。
- 修改 Excel 等同于修改模型输入配置。

---

## 6. 核心模块说明

### 6.1 `main.py`（流程编排）
职责：
- 读取数据
- 调用前处理函数
- 调用调度优化
- 调用后处理函数
- 做时间循环结转并导出

关键中间数据结构：
- `years_in_operation`: `ndarray (N_units, N_years)`
- `is_in_service`: `ndarray (N_units, N_years)`
- 各类价格/成本/碳排矩阵：`ndarray (N_units, N_years)`

### 6.2 `tool_preprocess.py`（前处理）
函数与职责：
- `initialize_fixed_assets_and_debt_status(row)`
  - 输入：单机组 `pd.Series`
  - 输出：长度 7 的 `pd.Series`
  - 作用：初始化固定资产、借款、残值、折旧等静态财务状态。
- `calculate_power_heat_price_matrix(...)`
  - 输出：`on_grid_power_price/capacity_power_price/heat_price`，shape 均为 `(N_units, N_years)`。
- `calculate_power_generation_bounds_matrix(...)`
  - 输出：`min_power_generation/max_power_generation`，shape `(N_units, N_years)`。
- `calculate_fuel_consumption_and_fuel_cost_matrix(...)`
  - 输出：燃耗率字典、燃料价格矩阵、供电/供热燃料成本矩阵。
- `calculate_carbon_emission_and_capture_intensity_matrix(...)`
  - 输出：6 个碳相关矩阵，shape 均为 `(N_units, N_years)`。

### 6.3 `Dispatch.cp313-win_amd64.pyd`（调度优化）
对外使用接口：
- `solve_optimal_unit_generation_allocation(data_config, unit_power_fuel_cost, min_power_generation, max_power_generation, local_planned_total_generation)`

返回：
- `unit_power_generation`
- `local_power_shortage`
- `local_power_curtailment`

说明：该模块为二进制扩展，内部实现不可直接在源码中查看。

### 6.4 `tool_postprocess.py`（后处理/结转）
函数与职责：
- `calculate_annual_revenue_cost_nbv_and_input_vat(...)`
  - 输出：收入、成本、净值、折旧、可抵扣/不可抵扣进项税矩阵。
- `calculate_output_vat_and_loan_principal_interest(...)`
  - 输出：销项税、长期借款本息、流资借款利息矩阵。
- `carry_forward_single_year_tax_and_profit(...)`
  - 输入输出均为单年向量 `(N_units,)`，用于税务利润结转。
- `carry_forward_single_year_cashflow_and_debt(...)`
  - 输入输出均为单年向量 `(N_units,)`，用于现金流与债务滚动。

### 6.5 数据结构与单位约定
- 统一 shape 约定：
  - 行维：机组（`N_units`）
  - 列维：年份（`N_years`）
- 单年循环切片统一使用 `[:, t]`，进入结转函数后为一维 `(N_units,)`。
- 价格/成本/税率单位未在代码中做统一注释文档，需与 Excel 模板保持一致。

---

## 7. 常见修改场景（重点）

### 7.1 新增一个参数/变量
建议步骤：
1. 确定参数来源（`data_config.py` 或某个 Excel sheet/列）。
2. 在 `main.py` 的读取阶段补充读取逻辑。
3. 在对应函数签名中透传参数（`tool_preprocess.py` 或 `tool_postprocess.py`）。
4. 保持 shape 对齐：
   - 常量按 `(N_units, 1)` 或 `(1, N_years)` 扩维后广播。
5. 若参数需要输出，加入 `result_matrices` 并导出。

不建议做法：
- 在函数内部硬编码魔法数字且不在 `data_config`/Excel 可追溯。

### 7.2 修改一个计算公式
推荐改动位置：
- 价格相关：`tool_preprocess.py`
- 收入成本税务：`tool_postprocess.py`
- 时间耦合滚动：`carry_forward_single_year_*`

必须同步检查：
1. 公式输入变量的 shape 是否一致。
2. 是否仍保留 `is_in_service` 掩码。
3. 公式改动是否影响下游现金流/债务滚动。
4. 导出结果是否仍可解释。

### 7.3 新增一个机组类型/发电类型
应修改/核对：
1. `01_基础静态台账.xlsx`：机组台账新增类型。
2. `02_类型级动态参数.xlsx`：新增类型列（电价、小时数、负荷率、可用率等）。
3. `04_燃料与碳排参数.xlsx`：若为燃料型，需补齐燃耗率、碳排基准、燃料价格映射。
4. `data_config.py`：若涉及补贴策略，更新机组类型列表。
5. 校验 `reindex/get_indexer` 处是否仍无缺失（否则会抛 `KeyError`）。

### 7.4 处理 Excel 表结构变更（sheet/列名变化）
影响点：
- `main.py` 的 `sheet_name="..."` 读取处。
- 所有 `DataFrame["列名"]` 字段访问处。

处理流程：
1. 先改读取映射。
2. 全局搜索旧列名（建议 `rg`）。
3. 逐模块跑通到最终导出。
4. 对比关键输出 sheet 是否仍齐全。

---

## 8. 运行校验与排错

### 8.1 常见报错与处理
- 缺文件：`FileNotFoundError`
  - 检查 `data/` 下四个输入 Excel 是否存在。
- 缺年份/缺机组类型：`KeyError`
  - 触发位置：`tool_preprocess.py` 中多个 `get_indexer` 与 `reindex` 检查。
  - 处理：补齐对应 sheet 的年份索引和类型列。
- 参数非法：`ValueError`
  - 触发位置：`initialize_fixed_assets_and_debt_status`（建设周期/还款年限/折旧年限 <= 0）。
- shape 不一致：
  - 高风险点：`local_planned_total_generation` 长度与 `规划时间跨度` 不一致。
- NaN 扩散：
  - 代码中已有 `.ffill()`，但若源表首行即空仍可能保留 NaN，需在输入侧修复。

### 8.2 建议校验点（可加断言/日志）
建议在 `main.py` 增加：
- `planning_horizon_years == local_planned_total_generation.shape[0]`
- 所有核心矩阵 shape 与 `(N_units, N_years)` 一致
- 导出前检查是否存在 NaN/inf

当前已存在的防御性检查：
- `tool_preprocess.py` 中对索引对齐的 `KeyError`
- `tool_preprocess.py` 中对周期参数合法性的 `ValueError`

---

## 9. 书写规范

### 9.1 修改前检查清单
- 不随意改业务字符串 key（尤其是 Excel 列名、sheet 名）。
- 不破坏函数接口输入输出 shape。
- 保持税率、价格、容量等单位口径一致。
- 新增字段时保证全链路透传（读取 -> 计算 -> 输出）。
- 修改后至少验证能跑通到 `result_output.xlsx` 导出。

### 9.2 代码风格
- 变量命名以业务语义为主，常配中英文注释。
- 矩阵计算优先使用 NumPy 广播。
- 分阶段注释结构：`# %% --- 模块 ---`。

---

## 10. 附录

### 10.1 关键术语表（中文术语 + 代码名）
- 进项税留抵余额：`input_vat_credit_balance`
- 累计未弥补亏损：`cumulative_uncovered_loss`
- 期末未分配利润：`ending_undistributed_profit`
- 累计提取法定盈余公积金：`cumulative_legal_surplus_reserve`
- 机组建设期借款余额：`unit_construction_loan_balance`
- 机组运行期借款余额：`unit_operation_loan_balance`
- 机组累计现金余额：`unit_cumulative_cash_balance`
- 资产总额：`total_assets`
- 净资产：`net_assets`
- 负债率：`debt_ratio`

### 10.2 常用命令汇总
```bash
# 运行主程序
python main.py

# 查看项目文件
rg --files

# 搜索关键字段（示例）
rg "sheet_name=" main.py
rg "unit_general_ledger\[" main.py tool_preprocess.py tool_postprocess.py
```

