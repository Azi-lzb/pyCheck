# Excel 参照表与数据处理（excel-data）

本子项目负责**参照表及 Excel 源文件**的维护与处理，不包含“检查程序”的业务逻辑。

## 目录结构

```
excel-data/
├── AGENTS.md          # Agent 说明（勿删）
├── README.md          # 本说明
├── data/              # 参照表等源文件（如国民经济行业分类映射表、五篇大文章对应参照表等）
├── fakedata/          # 测试/假数据（可选）
├── output/            # 校验报告、修正后表（可选）
└── scripts/           # 处理脚本
    └── fixExcelScripts/   # 修复、校验、生成脚本
        ├── fix_reference_table.py
        ├── fix_industry_mapping_by_small.py
        ├── fix_section_letter_by_big_class.py
        ├── verify_*.py / final_verify.py
        ├── check_*.py
        └── fix_reference_report.txt   # 示例：校验/修复报告
```

- **data/**：参照表源文件（CSV/Excel），供 check-program 通过 config 或默认路径读取。
- **scripts/fixExcelScripts/**：修复参照表（列名、编码、版本统一）、校验表结构与数据一致性、按需生成报表或中间文件。

## 与 check-program 的协作

- check-program 主程序通过 `main/config.xlsx` 的 mapping 表或默认路径读取本目录 `data/` 下的参照表与映射表。
- 变更表结构、列名或编码时，需在 README 或 AGENTS.md 中说明，并与 check-program 的读取约定保持一致。

---

在 `excel-data/` 下工作时，Cursor 会优先采用本目录的 Agent 设定，侧重 Excel/表结构与数据质量。
