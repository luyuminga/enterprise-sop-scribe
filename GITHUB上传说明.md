# 上传 GitHub（维护者备忘）

## 仓库名（与 Skill 一致）

**`enterprise-sop-scribe`**

- Skill frontmatter：`name: enterprise-sop-scribe`
- 建议路径：`https://github.com/luyuminga/enterprise-sop-scribe`
- 策略：**先 Private，再 Public**（Settings → Danger Zone → Change repository visibility）

## 你手动建库时

1. GitHub → New repository  
2. Repository name：`enterprise-sop-scribe`（不要改名）  
3. 选 **Private**  
4. **不要**勾选 Add a README / .gitignore / license（本地已有）  
5. Create repository  

然后在本目录执行：

```bash
cd "/Users/mac/Desktop/数谷/合资计划/FDE-Excel表单/业务流梳理包"

git remote add origin git@github.com:luyuminga/enterprise-sop-scribe.git
# 若已存在：git remote set-url origin git@github.com:luyuminga/enterprise-sop-scribe.git

git push -u origin main
```

推送成功后，他人（公开后）拉取：

```bash
git clone git@github.com:luyuminga/enterprise-sop-scribe.git
```

## GitHub About 描述（建库 / Settings → General → Description）

推荐（中英短句，About 栏可直接粘贴）：

```text
Enterprise SOP from work-item Excels — diagram-first confirmation, FDE discovery, knowledge-base prep. 企业SOP标准化 · FDE调研 · 知识库准备
```

更短版（若字数紧）：

```text
企业SOP标准化：进场表→流程图/SOP确认→FDE与知识库素材 | Excel-first SOP scribe for FDE & KB
```

## Topics（公开时可加）

`sop` `enterprise-sop` `rpa` `knowledge-base` `fde` `agent-skills` `deepseek` `doubao` `enterprise-sop-scribe`
