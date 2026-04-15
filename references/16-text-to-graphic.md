---
title: 文本降维与图形化
category: spire-presentation
description: 将文本内容自动转换为流程图或图形化表示
---

# 文本降维与图形化

## 概述

文本降维与图形化功能可以自动将幻灯片中的枯燥文本转换为直观的流程图、层次图或关系图。通过分析文本结构，智能选择合适的 SmartArt 布局，将文字信息可视化，提升演示文稿的可读性和吸引力。

## 使用场景

- 将步骤说明转换为流程图
- 将层次结构文本转换为组织结构图
- 将列表内容转换为图形化列表
- 将复杂逻辑转换为关系图

## 文本分析

### 识别文本结构

文本分析是转换的第一步，需要识别以下特征：

```csharp
// 文本分析示例
public static TextStructure AnalyzeText(string text)
{
    TextStructure structure = new TextStructure();

    // 检测序号
    if (Regex.IsMatch(text, @"\d+\s*[.)、]"))
    {
        structure.HasNumbering = true;
    }

    // 检测关键词
    structure.Keywords = new List<string>();
    string[] flowKeywords = { "开始", "结束", "如果", "那么", "因为", "所以",
                              "then", "else", "if", "because", "therefore" };
    foreach (string keyword in flowKeywords)
    {
        if (text.Contains(keyword))
        {
            structure.Keywords.Add(keyword);
        }
    }

    // 检测分隔符
    string[] separators = { "；", "。", "。", "；", ",", ";", "." };
    foreach (string separator in separators)
    {
        if (text.Contains(separator))
        {
            structure.HasSeparators = true;
            break;
        }
    }

    // 识别步骤
    structure.Steps = ExtractSteps(text);

    return structure;
}
```

### 提取步骤列表

```csharp
// 从文本中提取步骤
public static List<string> ExtractSteps(string text)
{
    List<string> steps = new List<string>();

    // 方法1: 按序号分割
    MatchCollection matches = Regex.Matches(text, @"\d+\s*[.)、]\s*(.*?)(?=\d+\s*[.)、]|$)");
    if (matches.Count > 0)
    {
        foreach (Match match in matches)
        {
            steps.Add(match.Groups[1].Value.Trim());
        }
    }

    // 方法2: 按标点符号分割
    else if (steps.Count == 0)
    {
        string[] parts = text.Split(new[] { "；", "。", "。", "；", ";", "." },
                                     StringSplitOptions.RemoveEmptyEntries);
        foreach (string part in parts)
        {
            if (!string.IsNullOrWhiteSpace(part))
            {
                steps.Add(part.Trim());
            }
        }
    }

    return steps;
}
```

## 图形类型选择

### 选择策略

根据文本分析结果选择最合适的图形类型：

```csharp
// 智能选择图形布局
public static SmartArtLayoutType SelectLayout(TextStructure structure)
{
    // 检测流程关键词
    bool hasFlowKeywords = structure.Keywords.Any(k =>
        k.Contains("如果") || k.Contains("那么") ||
        k.Contains("if") || k.Contains("then"));

    // 检测循环关键词
    bool hasCycleKeywords = structure.Keywords.Any(k =>
        k.Contains("循环") || k.Contains("重复") ||
        k.Contains("loop") || k.Contains("repeat"));

    // 检测层次关键词
    bool hasHierarchyKeywords = structure.Keywords.Any(k =>
        k.Contains("上级") || k.Contains("下级") ||
        k.Contains("parent") || k.Contains("child"));

    // 决策逻辑
    if (hasCycleKeywords || structure.Keywords.Any(k => k.Contains("迭代")))
    {
        return SmartArtLayoutType.BasicCycle;
    }
    else if (hasHierarchyKeywords || structure.HasIndentation)
    {
        return SmartArtLayoutType.Hierarchy;
    }
    else if (hasFlowKeywords || structure.Keywords.Any(k =>
             k.Contains("开始") || k.Contains("结束")))
    {
        return SmartArtLayoutType.BasicProcess;
    }
    else if (structure.Steps.Count <= 5)
    {
        return SmartArtLayoutType.BasicProcess;
    }
    else if (structure.Steps.Count <= 8)
    {
        return SmartArtLayoutType.ChevronProcess;
    }
    else
    {
        return SmartArtLayoutType.BasicBendingProcess;
    }
}
```

### 图形类型对照表

| 文本特征 | 推荐图形 | SmartArt 布局 |
|---------|---------|---------------|
| 顺序步骤（< 5项） | 简单流程图 | BasicProcess |
| 顺序步骤（5-8项） | 箭头流程图 | ChevronProcess |
| 顺序步骤（> 8项） | 弯曲流程图 | BasicBendingProcess |
| 循环/迭代 | 循环图 | BasicCycle, Cycle |
| 层次结构 | 层次图 | Hierarchy |
| 组织结构 | 组织结构图 | OrganizationalChart |
| 并列列表 | 列表 | BasicBendingProcess |
| 因果关系 | 关系图 | Balance, ConvergingRadial |
| 中心辐射 | 辐射图 | RadialCycle |

## 基础转换

### 示例 1: 简单文本转流程图

```csharp
using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using Spire.Presentation.Diagrams;

public static class TextToGraphicConverter
{
    // 将文本转换为流程图
    public static ISmartArt ConvertToFlowchart(
        Presentation presentation,
        ISlide slide,
        string text,
        RectangleF position)
    {
        // 分析文本
        List<string> steps = ExtractSteps(text);

        // 选择布局
        SmartArtLayoutType layout = SelectOptimalLayout(steps);

        // 创建 SmartArt
        ISmartArt smartArt = slide.Shapes.AppendSmartArt(position, layout);

        // 添加节点
        foreach (string step in steps)
        {
            ISmartArtNode node = smartArt.Nodes.AddNode();
            node.TextFrame.Text = TruncateText(step, 20); // 限制文本长度
        }

        // 应用样式
        smartArt.ColorStyle = SmartArtColorType.Colorful;
        smartArt.SmartArtStyle = SmartArtStyleType.WhiteOutline;

        return smartArt;
    }

    // 选择最优布局
    private static SmartArtLayoutType SelectOptimalLayout(List<string> steps)
    {
        int count = steps.Count;

        if (count <= 5)
            return SmartArtLayoutType.BasicProcess;
        else if (count <= 8)
            return SmartArtLayoutType.ChevronProcess;
        else
            return SmartArtLayoutType.BasicBendingProcess;
    }

    // 截断过长文本
    private static string TruncateText(string text, int maxLength)
    {
        if (text.Length <= maxLength)
            return text;
        return text.Substring(0, maxLength - 3) + "...";
    }

    // 提取步骤
    private static List<string> ExtractSteps(string text)
    {
        List<string> steps = new List<string>();

        // 按序号分割
        MatchCollection matches = Regex.Matches(
            text, @"\d+\s*[.)、]\s*(.*?)(?=\d+\s*[.)、]|$)"
        );

        if (matches.Count > 0)
        {
            foreach (Match match in matches)
            {
                steps.Add(match.Groups[1].Value.Trim());
            }
        }
        else
        {
            // 按标点符号分割
            string[] parts = text.Split(new[] { "；", "。", "；", ";", "." },
                                         StringSplitOptions.RemoveEmptyEntries);
            steps.AddRange(parts.Select(p => p.Trim()).Where(p => !string.IsNullOrWhiteSpace(p)));
        }

        return steps;
    }
}
```

### 示例 2: 使用转换器

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 获取要转换的文本
string sourceText = "1. 需求分析 2. 设计方案 3. 开发实现 4. 测试验证 5. 部署上线";

// 转换为流程图
RectangleF position = new RectangleF(50, 50, 700, 300);
ISmartArt flowchart = TextToGraphicConverter.ConvertToFlowchart(
    presentation,
    presentation.Slides[0],
    sourceText,
    position
);

presentation.SaveToFile("with_flowchart.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 高级转换

### 示例 3: 层次文本转层次图

```csharp
using System;
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using Spire.Presentation.Diagrams;

public static class TextToGraphicConverter
{
    // 将层次文本转换为层次图
    public static ISmartArt ConvertToHierarchy(
        Presentation presentation,
        ISlide slide,
        string hierarchicalText,
        RectangleF position)
    {
        ISmartArt hierarchy = slide.Shapes.AppendSmartArt(
            position,
            SmartArtLayoutType.Hierarchy
        );

        // 解析层次结构
        var nodes = ParseHierarchy(hierarchicalText);

        // 构建层次图
        BuildHierarchyTree(hierarchy, nodes);

        // 应用样式
        hierarchy.ColorStyle = SmartArtColorType.GradientLoopAccent1;
        hierarchy.SmartArtStyle = SmartArtStyleType.WhiteOutline;

        return hierarchy;
    }

    // 解析层次文本
    private static List<HierarchyNode> ParseHierarchy(string text)
    {
        List<HierarchyNode> result = new List<HierarchyNode>();

        string[] lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        HierarchyNode currentRoot = null;
        HierarchyNode currentParent = null;
        int previousLevel = 0;

        foreach (string line in lines)
        {
            string trimmed = line.Trim();
            if (string.IsNullOrWhiteSpace(trimmed)) continue;

            int level = GetIndentLevel(line);
            HierarchyNode node = new HierarchyNode
            {
                Text = trimmed.TrimStart(),
                Level = level,
                Children = new List<HierarchyNode>()
            };

            if (level == 0)
            {
                result.Add(node);
                currentRoot = node;
                currentParent = node;
            }
            else if (level > previousLevel)
            {
                currentParent.Children.Add(node);
                currentParent = node;
            }
            else
            {
                // 回退到适当级别
                currentParent = FindParentAtLevel(result, level - 1);
                if (currentParent != null)
                {
                    currentParent.Children.Add(node);
                }
            }

            previousLevel = level;
        }

        return result;
    }

    // 获取缩进级别
    private static int GetIndentLevel(string line)
    {
        int level = 0;
        foreach (char c in line)
        {
            if (c == '\t') level++;
            else if (c == ' ') level += 4;
            else break;
        }
        return level / 4; // 假设每级缩进4个空格
    }

    // 构建层次树
    private static void BuildHierarchyTree(ISmartArt smartArt, List<HierarchyNode> nodes)
    {
        foreach (HierarchyNode node in nodes)
        {
            ISmartArtNode smartNode = smartArt.Nodes.AddNode();
            smartNode.TextFrame.Text = node.Text;
            AddChildNodes(smartNode, node.Children);
        }
    }

    // 递归添加子节点
    private static void AddChildNodes(ISmartArtNode parentNode, List<HierarchyNode> children)
    {
        foreach (HierarchyNode child in children)
        {
            ISmartArtNode childNode = parentNode.ChildNodes.AddNode();
            childNode.TextFrame.Text = child.Text;
            AddChildNodes(childNode, child.Children);
        }
    }
}

public class HierarchyNode
{
    public string Text { get; set; }
    public int Level { get; set; }
    public List<HierarchyNode> Children { get; set; }
}
```

### 示例 4: 使用层次转换

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

string hierarchicalText = @"
公司
    技术部
        开发组
        测试组
    市场部
        销售组
        推广组
    运营部
        产品组
        客服组
";

RectangleF position = new RectangleF(50, 50, 700, 500);
ISmartArt hierarchy = TextToGraphicConverter.ConvertToHierarchy(
    presentation,
    presentation.Slides[0],
    hierarchicalText,
    position
);

presentation.SaveToFile("with_hierarchy.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 5: 智能图形选择

```csharp
using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using Spire.Presentation.Diagrams;

public static class TextToGraphicConverter
{
    // 智能转换（自动选择图形类型）
    public static ISmartArt ConvertSmart(
        Presentation presentation,
        ISlide slide,
        string text,
        RectangleF position)
    {
        // 分析文本
        TextAnalysis analysis = AnalyzeText(text);

        // 选择最佳布局
        SmartArtLayoutType layout = SelectBestLayout(analysis);

        // 创建 SmartArt
        ISmartArt smartArt = slide.Shapes.AppendSmartArt(position, layout);

        // 添加内容
        if (analysis.Type == GraphicType.Hierarchy)
        {
            BuildHierarchy(smartArt, analysis.HierarchyNodes);
        }
        else
        {
            BuildProcess(smartArt, analysis.Steps);
        }

        // 应用推荐样式
        ApplyRecommendedStyle(smartArt, analysis);

        return smartArt;
    }

    // 文本分析
    private static TextAnalysis AnalyzeText(string text)
    {
        TextAnalysis analysis = new TextAnalysis();

        // 检测层次结构
        if (HasIndentation(text))
        {
            analysis.Type = GraphicType.Hierarchy;
            analysis.HierarchyNodes = ParseHierarchy(text);
        }
        else
        {
            analysis.Type = GraphicType.Process;
            analysis.Steps = ExtractSteps(text);
        }

        // 检测关键词
        analysis.Keywords = ExtractKeywords(text);

        // 检测循环
        analysis.HasLoop = analysis.Keywords.Any(k =>
            k.Contains("循环") || k.Contains("重复") ||
            k.Contains("loop") || k.Contains("repeat"));

        return analysis;
    }

    // 检测是否有缩进
    private static bool HasIndentation(string text)
    {
        string[] lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (string line in lines.Skip(1))
        {
            if (line.StartsWith("\t") || line.StartsWith("    "))
            {
                return true;
            }
        }
        return false;
    }

    // 提取关键词
    private static List<string> ExtractKeywords(string text)
    {
        List<string> keywords = new List<string>();
        string[] keywordPatterns = {
            "开始", "结束", "如果", "那么", "因为", "所以", "循环", "重复",
            "start", "end", "if", "then", "else", "loop", "repeat", "because"
        };
        foreach (string pattern in keywordPatterns)
        {
            if (text.Contains(pattern))
            {
                keywords.Add(pattern);
            }
        }
        return keywords;
    }
}

public class TextAnalysis
{
    public GraphicType Type { get; set; }
    public List<string> Steps { get; set; }
    public List<HierarchyNode> HierarchyNodes { get; set; }
    public List<string> Keywords { get; set; }
    public bool HasLoop { get; set; }
}

public enum GraphicType
{
    Process,
    Hierarchy,
    Cycle,
    Relationship
}
```

## 批量转换

### 示例 6: 批量转换幻灯片中的文本

```csharp
using Spire.Presentation;
using System.Drawing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

foreach (ISlide slide in presentation.Slides)
{
    // 查找包含大量文本的形状
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IAutoShape autoShape && autoShape.TextFrame != null)
        {
            string text = autoShape.TextFrame.Text;

            // 检查是否适合转换（超过50个字符且包含步骤）
            if (text.Length > 50 && ContainsSteps(text))
            {
                // 备份原文本到备注
                slide.NotesSlide.NotesTextFrame.Text = $"原文本:\n{text}";

                // 转换为图形
                RectangleF position = new RectangleF(
                    autoShape.X, autoShape.Y,
                    Math.Max(600, autoShape.Width * 1.5f),
                    Math.Max(300, autoShape.Height * 1.5f)
                );

                ISmartArt graphic = TextToGraphicConverter.ConvertToFlowchart(
                    presentation,
                    slide,
                    text,
                    position
                );

                // 删除原文本形状
                slide.Shapes.Remove(shape);

                Console.WriteLine($"已转换幻灯片 {slide.SlideNumber} 的文本");
            }
        }
    }
}

presentation.SaveToFile("converted.pptx", FileFormat.Pptx2010);
presentation.Dispose();

// 辅助方法：检查是否包含步骤
private static bool ContainsSteps(string text)
{
    return System.Text.RegularExpressions.Regex.IsMatch(
        text, @"\d+\s*[.)、]"
    );
}
```

## 布局优化

### 自动调整大小

```csharp
// 根据内容自动调整 SmartArt 大小
public static void AutoSizeSmartArt(ISmartArt smartArt)
{
    int nodeCount = smartArt.Nodes.Count;
    int maxTextLength = smartArt.Nodes.Max(n => n.TextFrame.Text.Length);

    // 根据节点数量计算高度
    float estimatedHeight = 50 * nodeCount + 100;

    // 根据文本长度计算宽度
    float estimatedWidth = Math.Max(300, maxTextLength * 15) + 100;

    // 设置大小
    smartArt.Width = estimatedWidth;
    smartArt.Height = estimatedHeight;
}
```

### 自动居中

```csharp
// 将 SmartArt 居中于幻灯片
public static void CenterSmartArt(ISlide slide, ISmartArt smartArt)
{
    float slideWidth = slide.SlideSize.Size.Width;
    float slideHeight = slide.SlideSize.Size.Height;

    smartArt.X = (slideWidth - smartArt.Width) / 2;
    smartArt.Y = (slideHeight - smartArt.Height) / 2;
}
```

## 样式定制

### 应用颜色主题

```csharp
// 根据幻灯片主题应用颜色
public static void ApplyThemeColors(ISmartArt smartArt, Color accentColor)
{
    // 使用主题强调色
    smartArt.ColorStyle = SmartArtColorType.Accented1;

    // 自定义颜色（需要遍历节点）
    foreach (ISmartArtNode node in smartArt.Nodes)
    {
        node.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid;
        node.TextFrame.TextRange.Fill.SolidColor.Color = accentColor;
    }
}
```

## 完整示例

### 示例 7: 端到端转换

```csharp
using System.Drawing;
using Spire.Presentation;

// 加载演示文稿
Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 获取第一张幻灯片
ISlide slide = presentation.Slides[0];

// 要转换的文本
string processText = @"1. 用户提交申请
2. 系统验证信息
3. 人工审核申请
4. 生成审核结果
5. 通知用户结果";

// 转换为流程图
RectangleF position = new RectangleF(50, 50, 700, 300);
ISmartArt flowchart = TextToGraphicConverter.ConvertSmart(
    presentation,
    slide,
    processText,
    position
);

// 自动调整大小
TextToGraphicConverter.AutoSizeSmartArt(flowchart);

// 居中显示
TextToGraphicConverter.CenterSmartArt(slide, flowchart);

// 应用样式
TextToGraphicConverter.ApplyThemeColors(flowchart, Color.FromArgb(0, 120, 215));

// 保存
presentation.SaveToFile("converted.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 注意事项

1. **文本长度**: 单个节点文本不宜超过20个字符
2. **节点数量**: SmartArt 建议不超过15个节点
3. **布局适配**: 不同布局对节点数量有不同要求
4. **语言支持**: 中英文关键词都能识别
5. **保留原文**: 建议将原文本保存到备注中

## API 参考

### 相关类

| 类 | 描述 |
|----|------|
| `ISmartArt` | SmartArt 图形对象 |
| `ISmartArtNode` | SmartArt 节点 |
| `SmartArtLayoutType` | 布局类型枚举 |
| `SmartArtColorType` | 颜色样式枚举 |
| `SmartArtStyleType` | 形状样式枚举 |

### 主要方法

| 方法 | 描述 |
|------|------|
| `ExtractSteps(text)` | 从文本中提取步骤 |
| `AnalyzeText(text)` | 分析文本结构 |
| `SelectLayout(analysis)` | 选择最佳布局 |
| `ConvertToFlowchart(...)` | 转换为流程图 |
| `ConvertToHierarchy(...)` | 转换为层次图 |
| `ConvertSmart(...)` | 智能转换 |

## 相关功能

- [文本处理](./03-text-content.md) - 文本格式化
- [SmartArt](./07-smartart.md) - SmartArt 基础操作
- [形状处理](./04-shapes-images.md) - 自定义形状
- [动画](./09-animations.md) - 为图形添加动画
