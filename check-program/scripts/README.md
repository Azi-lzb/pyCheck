# scripts - 临时与辅助脚本

本目录存放与主程序**无关**的临时、调试或一次性脚本，不参与主流程。

## 当前脚本

- **debug_hts06.py**：调试用，分析某条借据 HTS06 未填报原因（读取 main/ 下结果文件）
- **find_7517.py**：在参照表中查找含 7517 的行
- **fix_ref_from_2017.py**：从 2017 注释生成/修正参照表（依赖 excel-data/data 下源文件）

运行时可从项目根执行，例如：

```bash
python check-program/scripts/find_7517.py
```
