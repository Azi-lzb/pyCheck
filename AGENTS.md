# pyCheckCursor 项目总览

本仓库为 **monorepo**，包含两个子项目。请根据当前打开/编辑的文件所在目录，遵循对应子项目下的 `AGENTS.md`。

## 仓库结构

```
pyCheckCursor/
├── AGENTS.md           # 本文件：项目总览与 Agent 路由
├── excel-data/         # 子项目：参照表与 Excel 数据处理
│   ├── AGENTS.md
│   ├── README.md
│   ├── data/           # 参照表等源文件
│   ├── fakedata/       # 测试/假数据（可选）
│   └── scripts/        # 修复、校验、生成脚本（含 fixExcelScripts/）
└── check-program/      # 子项目：检查程序
    ├── AGENTS.md
    ├── README.md
    ├── main/           # 主程序、config、策略一结果与说明
    └── scripts/        # 临时/调试/辅助脚本
```

## 子项目划分

| 子项目       | 目录              | 职责简述 |
|--------------|-------------------|----------|
| Excel 参照表 | `excel-data/`     | 参照表与 Excel 源文件的处理、修复、校验与产出 |
| 检查程序     | `check-program/`  | 调用参照表与数据的 Python 程序，实现策略与结果输出 |

## 通用约定

- **边界**：在 `excel-data/` 下只改表与表处理脚本；在 `check-program/` 下只改程序逻辑与对数据的调用方式；不跨子项目随意改对方目录下的文件。
- **文档**：项目级说明可放在根目录或 `docs/`（若存在）；子项目说明放在各自目录的 README 与 AGENTS.md。
- **语言**：两子项目均为 Python，注意依赖与运行环境一致（如 Python 版本、虚拟环境）。

## 当前打开目录对应的 Agent

- 正在编辑 **excel-data/** 内文件 → 遵循 `excel-data/AGENTS.md`（侧重表结构、Excel、数据质量）
- 正在编辑 **check-program/** 内文件 → 遵循 `check-program/AGENTS.md`（侧重程序逻辑、策略、配置）
