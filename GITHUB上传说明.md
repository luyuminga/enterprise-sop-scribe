# 上传 GitHub（维护者备忘）

## 仓库名（与 Skill 一致）

**`personal-workflow-scribe`**

- Skill frontmatter：`name: personal-workflow-scribe`
- 建议路径：`https://github.com/luyuminga/personal-workflow-scribe`
- 策略：**先 Private，再 Public**（Settings → Danger Zone → Change repository visibility）

## 你手动建库时

1. GitHub → New repository  
2. Repository name：`personal-workflow-scribe`（不要改名）  
3. 选 **Private**  
4. **不要**勾选 Add a README / .gitignore / license（本地已有）  
5. Create repository  

然后在本目录执行：

```bash
cd "/Users/mac/Desktop/数谷/合资计划/FDE-Excel表单/业务流梳理包"

git remote add origin git@github.com:luyuminga/personal-workflow-scribe.git
# 若已存在：git remote set-url origin git@github.com:luyuminga/personal-workflow-scribe.git

git push -u origin main
```

推送成功后，他人（公开后）拉取：

```bash
git clone git@github.com:luyuminga/personal-workflow-scribe.git
```

## Topics（公开时可加）

`sop` `rpa` `knowledge-base` `fde` `agent-skills` `deepseek` `doubao` `personal-workflow-scribe`
