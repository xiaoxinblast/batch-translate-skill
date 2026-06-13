# batch-translate Skill

Claude Code 用户级 Skill：mqxliff/docx/xlsx/txt 批量翻译 + 校对，全自动循环。

## 安装

```bash
mkdir -p ~/.claude/skills/batch-translate
cp SKILL.md ~/.claude/skills/batch-translate/
```

## 依赖

首次触发时自动从 GitHub 安装工具包：https://github.com/xiaoxinblast/batch-translate

## 使用

在 Claude Code 对话中说：

> "开始批量翻译"

Skill 自动完成：安装工具包 → 生成风格指南/术语库 → 全量语境分析 → 分批翻译 → 逐批校对 → 写回。
