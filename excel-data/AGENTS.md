# 子项目：Excel 参照表与数据处理（excel-data）

## 角色与范围

- **名称**：Excel/参照表数据处理 Agent
- **职责**：只在本目录（`excel-data/`）内工作，负责参照表、Excel 源文件的准备、修复、校验与产出。
- **不负责**：不修改 `check-program/` 下的业务程序；业务逻辑与“如何调用数据”由检查程序子项目负责。

## 目录结构

- **data/**：参照表等源文件（如国民经济行业分类映射表、五篇大文章与行业分类对应参照表等）
- **fakedata/**：测试/假数据（可选）
- **scripts/**：处理脚本；当前包含 **fixExcelScripts/** 子目录，用于修复、校验、生成
  - 修复类：`fix_reference_table.py`、`fix_industry_mapping_by_small.py`、`fix_section_letter_by_big_class.py` 等
  - 校验类：`verify_*.py`、`final_verify.py`、`check_*.py` 等

## 技术栈与约定

- **语言**：Python 3
- **典型依赖**：openpyxl、pandas、xlrd（按需），读写 Excel/CSV 时注意编码与表头一致性。
- **数据位置**：
  - 参照表等源文件放在 `excel-data/data/`
  - 产出物（修正后的表、校验报告）可放在 `excel-data/output/` 或与需求方约定的路径；校验报告也可与脚本同目录（如 `fix_reference_report.txt`）

## 开发与脚本规范

- 脚本放在 `excel-data/scripts/`（含现有 `fixExcelScripts/` 等子目录）
- 新增脚本需有清晰用途注释；修改已有参照表格式时需在 README 或本 AGENTS.md 中说明列名、编码、版本
- 校验/修复脚本应支持幂等或可重复执行，并输出明确成功/失败或报告路径

## 测试与质量

- 修改参照表结构或脚本逻辑后，运行相关校验脚本（如 `verify_*.py`）确保无回归
- 大批量改表前先备份或使用副本，避免破坏唯一数据源

## 与 check-program 的边界

- 本子项目**产出**：可被程序读取的参照表文件（如 CSV/Excel）、字段与编码约定说明；check-program 通过 config 的 mapping 或默认路径读取 `excel-data/data/` 等位置。
- 本子项目**不实现**：策略一/策略二、检查报告生成、配置表重置等业务逻辑——这些在 `check-program/` 中实现并**调用**本处产出的数据。
