# main - 主程序与配置

本目录存放**主程序**入口与配置文件。

## 内容

- **build_strategy1.py**：策略一主程序入口
- **reset_config_sheet.py**：重置 config 工作表的工具（依赖 build_strategy1）
- **config.xlsx**：运行期生成/维护的配置文件（参照表、映射表路径等）
- **策略一核查结果.xlsx**：策略一输出结果（运行后生成）
- **使用说明.md**：策略一主程序使用说明（工具说明、菜单、结果列、config/clue 配置、核查逻辑、FAQ，推荐先阅）

## 运行方式

在项目根目录下：

```bash
python check-program/main/build_strategy1.py
```

或在 `check-program/main/` 下：

```bash
python build_strategy1.py
```

配置与结果文件会落在本目录（main/）。
