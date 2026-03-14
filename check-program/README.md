# 检查程序（check-program）

本子项目是**调用参照表与数据**的 Python 程序，负责策略执行、配置与结果输出，不负责参照表/Excel 源文件的制作与修复。

## 目录结构

```
check-program/
├── AGENTS.md
├── README.md           # 本说明
├── main/               # 主程序与配置
│   ├── build_strategy1.py      # 策略一主程序入口
│   ├── reset_config_sheet.py   # 重置 config 工作表
│   ├── config.xlsx             # 运行期生成/维护（mapping 表配置参照表/映射表路径）
│   ├── 策略一核查结果*.xlsx     # 策略一输出（运行后生成）
│   ├── 使用说明.md
│   └── README.md
└── scripts/            # 与主程序无关的临时/辅助脚本
    ├── debug_hts06.py
    ├── find_7517.py
    ├── fix_ref_from_2017.py
    └── README.md
```

- **main/**：主程序入口、`config.xlsx`、策略一结果文件及使用说明。
- **scripts/**：临时、调试、一次性脚本，不参与主流程。

## 数据依赖

- 参照表、映射表等 Excel 源文件由 **excel-data** 子项目维护（通常位于 `excel-data/data/`）。
- 主程序通过 `main/config.xlsx` 的 **mapping** 表（如 A2 参照表路径、B2 映射表路径）或默认路径读取这些文件。

## 运行方式

在项目根目录或本目录下执行，例如：

```bash
python check-program/main/build_strategy1.py
```

或在 `check-program/main/` 下：

```bash
python build_strategy1.py
```

配置与结果文件会落在 `main/` 目录。

---

在 `check-program/` 下工作时，Cursor 会采用本目录的 Agent 设定，侧重程序逻辑与对参照表数据的调用方式。
