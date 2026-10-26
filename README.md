# 企业 SOP 梳理包 · Enterprise SOP Scribe

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](./LICENSE)

开源、可拉取即用的 **企业 SOP 标准化辅助包**：用进场表格 + 多平台 AI Agent，把「个人工作事项」整理成**可确认的流程图与作业指导书**，并给 FDE / 知识库留下结构化素材。

> Skill / 仓库名：`enterprise-sop-scribe`  
> 适用：豆包工作台 · DeepSeek Harness · Cursor · 普通大模型对话  

**一句话（GitHub Description）：**  
Enterprise SOP from work-item Excels — diagram-first confirmation, FDE discovery, knowledge-base prep. 企业SOP标准化 · FDE调研 · 知识库准备

---

## 三大目标

| # | 目标 | 本包如何服务 |
|---|------|----------------|
| 1 | **企业 SOP 标准化** | 每人每条事项 → As-Is 流程图 + SOP；业务员 Word **先图后文**确认后定稿 |
| 2 | **辅佐 FDE 团队调研** | 统一目录与触点/自动化分析；进场少空转，收件可校验 |
| 3 | **企业知识库资料准备** | 确认后输出知识库条目卡与可入库 SOP，避免未确认内容入库 |

---

## 30 秒理解

```
组织表 + 个人工作事项表
        ↓
  AI（少问多补 + 公司名公网补上下文）
        ↓
  对话呈现：先流程图，后背景步骤
        ↓
  MD 主稿 → Word 确认稿（业务员确认）
        ↓
  定稿后：系统触点 / 自动化候选 / 知识库条目
```

公开方法（SIPOC、Process Doc、RPA 评估等）只提供「写全什么」；本包坚持：**进场表先行、少问多补、确认门禁、多 Harness 一套规则**。详见 [04-公开方法与对标.md](./04-公开方法与对标.md)。

---

## 拉取使用

```bash
git clone git@github.com:luyuminga/enterprise-sop-scribe.git
# 或：git clone https://github.com/luyuminga/enterprise-sop-scribe.git
cd enterprise-sop-scribe
```

仓库名与 Skill 名一致：`enterprise-sop-scribe`。宜先 **Private**，稳定后再改为 **Public**。

### 1. 先填表

空白模板在 [`templates/excel/`](./templates/excel/)：

- `组织信息填写表.xlsx` — 部门负责人，每部门 1 份  
- `个人工作事项填写表.xlsx` — 业务同事，每人 1 份  

### 2. 再挂到 AI

| 环境 | 做法 |
|------|------|
| **通用对话** | 复制 [`01-系统提示词-主规则.md`](./01-系统提示词-主规则.md) 作首条；上传两张表 |
| **豆包工作台** | 工作伙伴指令用 [`豆包工作伙伴-指令摘要.md`](./豆包工作伙伴-指令摘要.md)；上下文挂本仓库；详见 [05](./05-多平台使用指南.md) |
| **DeepSeek Harness** | 将本仓库加入 skills，加载 `enterprise-sop-scribe`（[`SKILL.md`](./SKILL.md)） |
| **Cursor** | 项目内打开本仓库，或拷贝到 `~/.cursor/skills/enterprise-sop-scribe/` |
| **项目约定** | 使用根目录 [`AGENTS.md`](./AGENTS.md) |

完整平台说明：[05-多平台使用指南.md](./05-多平台使用指南.md)  
业务员 5 分钟上手：[00-业务员使用说明.md](./00-业务员使用说明.md)

### 3. 出 Word 确认稿（FDE / 有终端时）

```bash
pip install python-docx pillow
# 建议：npm install -g @mermaid-js/mermaid-cli
python3 FDE/build_sop_confirm_docx.py --flow-dir "/path/to/flows/01-事项短名"
```

Word **固定先「一、业务流程图」，再业务背景与步骤**——方便业务员一眼确认，也方便 FDE 建立整体认识。

---

## 仓库结构

```
.
├── README.md                 # 本说明（开源入口）
├── LICENSE                   # MIT
├── CONTRIBUTING.md
├── SKILL.md                  # Agent Skill 清单项
├── AGENTS.md
├── 00-业务员使用说明.md
├── 01-系统提示词-主规则.md     # 核心行为规则
├── 02-阶段补丁/
├── 03-输出模板/
├── 04-公开方法与对标.md
├── 05-多平台使用指南.md
├── 豆包工作伙伴-指令摘要.md
├── templates/excel/          # 进场空白表
├── FDE/                      # 收件清单 + Word 导出
├── _fixtures/                # 脱敏试跑样例
└── docs/                     # 设计/实现文档
```

---

## 给三类用户

| 角色 | 从哪开始 |
|------|----------|
| **企业业务员** | `00` + 填个人表 + 按平台挂规则 → 确认 Word 里的图 |
| **FDE / 实施顾问** | `05` 挂载 + `FDE/收件校验清单.md` + 导出脚本；用确认后的流做调研与自动化候选 |
| **知识库 / 数字化负责人** | 只收「已确认」的 `02` + `05-知识库条目卡`；未确认不入库 |

---

## 本包明确不做

- 软件项目报价与技术方案（请用贵司自己的评估流程）  
- 替代企业级咨询总报告  
- 未经业务员确认就把内容写入生产知识库  

---

## 设计文档

- [docs/2026-10-26-个人业务流梳理-设计文档.md](./docs/2026-10-26-个人业务流梳理-设计文档.md)  
- [docs/2026-10-26-个人业务流梳理-实现计划.md](./docs/2026-10-26-个人业务流梳理-实现计划.md)  

---

## License

[MIT](./LICENSE)

欢迎 Issue / PR。提交前请确认样例已脱敏，勿上传真实客户数据。
