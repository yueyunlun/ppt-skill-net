# Spire.Presentation Skill 格式调整计划

## 当前状态

当前创建的文件不符合 Claude Code skills 的标准格式，需要进行调整。

## 问题分析

1. **文件结构问题**:
   - 缺少 `SKILL.md` 主文件（当前是 `README.md`）
   - 目录名称不符合规范（应使用小写连字符）
   - 缺少可选的子目录结构

2. **元数据问题**:
   - Frontmatter 格式不符合规范
   - description 字段格式错误（应使用第三人称）
   - 缺少必需的 `name` 字段

3. **内容组织问题**:
   - 内容过于庞大，需要采用渐进式披露设计
   - 应将详细内容移至 `references/` 目录
   - 应添加 `examples/` 目录存放代码示例

## 调整计划

### 第一步：重组目录结构

**目标目录结构**:
```
~/.claude/skills/spire-presentation/
├── SKILL.md                      # 主技能定义（核心内容）
├── references/                   # 详细文档
│   ├── 01-getting-started.md
│   ├── 02-basic-operations.md
│   ├── 03-text-content.md
│   ├── 04-shapes-images.md
│   ├── 05-tables.md
│   ├── 06-charts.md
│   ├── 07-smartart.md
│   ├── 08-multimedia.md
│   ├── 09-animations.md
│   ├── 10-hyperlinks.md
│   ├── 11-conversion.md
│   ├── 12-advanced-features.md
│   ├── 13-security.md
│   ├── 14-printing.md
│   └── 15-best-practices.md
├── examples/                     # 代码示例
│   ├── basic/
│   ├── charts/
│   ├── tables/
│   └── advanced/
└── evals/
    └── trigger_eval.json         # 触发评估测试
```

### 第二步：创建主 SKILL.md 文件

**内容要求**:
- YAML frontmatter（必需：name, description）
- 概述部分
- 核心概念和快速参考
- 使用场景
- 指向 references 的链接
- 字数控制在 1,500-2,000 字

**Frontmatter 格式**:
```yaml
---
name: spire-presentation
description: This skill should be used when the user asks to "create a PowerPoint presentation", "edit a PPTX file", "convert PowerPoint to PDF", "add charts to slides", or mentions Spire.Presentation, PowerPoint automation, or .NET presentation processing.
version: 0.1.0
---
```

### 第三步：调整现有文件

将现有的 16 个文档：
1. 移动到 `references/` 目录
2. 保留主要内容
3. 确保 markdown 格式正确
4. 添加相互引用链接

### 第四步：创建示例目录

从文档中提取代码示例：
- 按功能分类存放
- 确保示例完整可运行
- 添加必要的说明

### 第五步：创建评估文件

创建 `evals/trigger_eval.json`：
- 包含应触发技能的查询
- 包含不应触发技能的查询

## 具体执行步骤

1. 创建新的目录结构
2. 编写主 SKILL.md 文件
3. 移动和调整现有文档到 references/
4. 从文档中提取代码示例到 examples/
5. 创建 trigger_eval.json
6. 测试技能加载
7. 清理旧的文件结构

## 输出目标

将整个 skill 移动到正确位置：`~/.claude/skills/spire-presentation/`
