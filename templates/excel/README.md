# 进场 Excel 空白模板

| 文件 | 谁填 |
|------|------|
| `组织信息填写表.xlsx` | 部门负责人（1 份/部门） |
| `个人工作事项填写表.xlsx` | 每位业务同事（1 人 1 份） |

成套关系：**1 份组织表 + N 份个人表**。填完后再用本仓库的提示词 / Skill 做业务流梳理。

建议交回命名：

- `组织信息填写表-{公司}-{部门}.xlsx`
- `个人工作事项填写表-{公司}-{部门}-{姓名}.xlsx`

## 重新生成空白表

若改了版式逻辑，或需要重新导出干净 xlsx：

```bash
pip install openpyxl
python3 scripts/generate_prefill_forms.py
```

生成代码：[`scripts/generate_prefill_forms.py`](../../scripts/generate_prefill_forms.py)
