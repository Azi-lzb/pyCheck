# 子项目：检查程序（check-program）

## 角色与范围

- **名称**：检查程序开发 Agent
- **职责**：只在本目录（`check-program/`）内开发与维护 **调用参照表与数据** 的 Python 程序，实现策略、配置与结果输出。
- **数据依赖**：参照表与 Excel 源文件由 `excel-data/` 子项目维护；本程序通过约定路径或 `config.xlsx` 的 mapping 表读取这些数据，不直接修改源表结构。

## 目录结构

- **main/**：主程序入口、配置与结果
  - `build_strategy1.py`：策略一主程序
  - `reset_config_sheet.py`：重置 config 工作表的工具
  - `config.xlsx`：运行期维护（参照表/映射表路径等在 mapping 表）
  - `策略一核查结果*.xlsx`：策略一输出（运行后生成）
  - `使用说明.md`：策略一主程序使用说明（含 config/clue、核查逻辑、FAQ）
- **scripts/**：与主流程无关的临时、调试、一次性脚本（如 `debug_hts06.py`、`find_7517.py`、`fix_ref_from_2017.py`）

## 技术栈与约定

- **语言**：Python 3
- **典型依赖**：pandas、openpyxl 等，用于读取 excel-data 产出的表；业务逻辑（策略一、策略二、配置重置等）在本目录实现。
- **参照表路径**：优先从 `main/config.xlsx` 的 mapping 表（如 A2 参照表、B2 映射表）读取；未配置时可回退到默认路径（如 `../excel-data/data/` 下文件）。

## 开发与运行

- 运行前确认参照表/映射表路径指向 `excel-data/data/` 等约定位置，或已在 config 的 mapping 中正确配置。
- 修改“如何读取/解析参照表”时，与 excel-data 的列名、编码、文件格式约定保持一致；不擅自改表结构，若需改表应到 excel-data 子项目处理。

## 测试与质量

- 修改策略或配置逻辑后，用现有或新增脚本做一次小范围运行（如单条/单表）验证结果。
- 使用说明、策略说明文档放在 `main/` 或本目录，便于 Agent 与人工查阅。

## 与 excel-data 的边界

- **本子项目**：实现“用参照表做检查、算结果、写配置”的程序逻辑。
- **excel-data 子项目**：负责参照表与 Excel 源文件的准备、修复与校验；本程序只读其产出，不在此处改表。
