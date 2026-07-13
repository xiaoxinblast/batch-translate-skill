# batch-translate Skill

Reasonix 全局 Skill：mqxliff/docx/xlsx/txt 批量翻译 + 校对，全自动循环。

## 安装

```bash
# Reasonix 中安装
/reasonix install-skill https://github.com/xiaoxinblast/batch-translate-skill
```

## 依赖

首次触发时自动从 GitHub 安装工具包：https://github.com/xiaoxinblast/batch-translate

## 使用

在 Reasonix 对话中说：

> "开始批量翻译"

Skill 自动完成：安装工具包 → 扫描项目文件 → 编译风格指南/术语库 → 全量语境分析 → 分批翻译 → 逐批校对 → 写回。

## 核心 subagent

| subagent | 模型 | 职责 |
|----------|------|------|
| `translator` | deepseek-pro | 日中翻译，按风格指南产出自然中文 |
| `trans-reviewer` | deepseek-pro | 硬性错误检查 + 语言润色 |
| `context-analyzer` | deepseek-flash | 全量语境分析，识别术语缺口 |
