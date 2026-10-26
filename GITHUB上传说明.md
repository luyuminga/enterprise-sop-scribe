# 上传 GitHub 开源（维护者备忘）

本目录即可作为仓库根目录推送。

## 建议仓库名

- `personal-workflow-scribe`（推荐，与 Skill name 一致）  
- 或 `enterprise-sop-scribe`

## 建议步骤

```bash
cd "/path/to/业务流梳理包"   # 即以本包为根
git init
git add .
git status   # 确认无真实客户 xlsx / 未脱敏样例
git commit -m "feat: open-source personal workflow scribe for SOP, FDE, and KB prep"
# 在 GitHub 新建空仓库后：
git remote add origin git@github.com:<org>/personal-workflow-scribe.git
git branch -M main
git push -u origin main
```

## 推送后

1. 把根 `README.md` 里的 `git clone` URL 换成真实地址  
2. GitHub About 描述可写：`Enterprise SOP standardization · FDE discovery · Knowledge base prep`  
3. Topics 建议：`sop` `rpa` `knowledge-base` `fde` `agent-skills` `deepseek` `doubao`

## 注意

- `_fixtures` 仅含虚构样例；禁止提交真实客户表  
- `LICENSE` 为 MIT；若公司要求换许可证，推送前改 LICENSE 与 README 徽章  
