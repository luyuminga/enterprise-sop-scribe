---
name: enterprise-sop-scribe
description: >-
  企业 SOP 标准化：进场表少问多补，但业务流细节调研清楚；生成开发可读的 As-Is 流程图（节点=角色·系统·动作+字段说明）
  与作业指导书，Word 先图后文供业务员确认，再派生 FDE 触点与知识库条目。适用于豆包/DeepSeek/Cursor 等。
  Enterprise SOP from Excels — diagram-first, developer-ready flowcharts, employee confirmation, FDE & KB prep.
  Triggers: SOP, 流程图开发, 作业指导书, FDE, 知识库, 工作事项表, 流程确认.
---

# Enterprise SOP Scribe（企业 SOP 梳理 · 跨平台 Skill）

开源目标：**企业 SOP 标准化 · 辅佐 FDE 调研 · 知识库资料准备**。  
（输入仍是个人工作事项表；对外名强调企业 SOP，不叫 personal。）

你正在执行本 Skill。**先完整阅读并遵守**同目录：

1. `01-系统提示词-主规则.md`（行为与阶段机，必遵）
2. 需要时再读：`04-公开方法与对标.md`、`03-输出模板/`、`FDE/Word导出约定.md`
3. 空白进场表：`templates/excel/`

## 平台无关硬规则

- 输入优先：组织表 + 个人表；公司名公网补行业/业务范围
- 少问多补；As-Is；对话与 Word 均 **先流程图、后业务背景与步骤**
- 未业务员确认：不定稿、不派生 `03/04/05` 定稿
- 有文件写入能力时：按 `03-输出模板/命名与目录规范.md` 建 `{公司}-{部门}-{姓名}-业务流梳理/flows/…`
- 有 docx/脚本能力时：用 `FDE/build_sop_confirm_docx.py` 导出确认稿；否则先交齐 MD + 说明由 FDE 出 Word
- 脱敏：往来方A、单号***；禁止账号密码

## 开场

用户提供表格后：检索公司上下文 → 出示清单与推荐 → 只问「先做哪一条？序号或按推荐」。

## 详细平台贴法

见 `05-多平台使用指南.md`。
