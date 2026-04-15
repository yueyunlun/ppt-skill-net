---
name: text-to-flowchart
description: Complete examples for converting text to graphics and flowcharts
---

# Text to Flowchart Examples

## Example 1: Simple Steps to Flowchart

```csharp
using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using Spire.Presentation.Diagrams;

public static class TextToGraphicConverter
{
    // 将步骤文本转换为流程图
    public static ISmartArt ConvertStepsToFlowchart(
        Presentation presentation,
        ISlide slide,
        string stepText,
        RectangleF position)
    {
        // 提取步骤
        List<string> steps = ExtractSteps(stepText);

        // 选择布局
        SmartArtLayoutType layout = steps.Count <= 5
            ? SmartArtLayoutType.BasicProcess
            : SmartArtLayoutType.ChevronProcess;

        // 创建流程图
        ISmartArt flowchart = slide.Shapes.AppendSmartArt(position, layout);

        // 添加节点
        foreach (string step in steps)
        {
            ISmartArtNode node = flowchart.Nodes.AddNode();
            node.TextFrame.Text = TruncateText(step, 20);
        }

        // 应用样式
        flowchart.ColorStyle = SmartArtColorType.Colorful;
        flowchart.SmartArtStyle = SmartArtStyleType.WhiteOutline;

        return flowchart;
    }

    private static List<string> ExtractSteps(string text)
    {
        List<string> steps = new List<string>();
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
            string[] parts = text.Split(new[] { "；", "。", "；", ";", "." },
                                         StringSplitOptions.RemoveEmptyEntries);
            steps.AddRange(parts.Select(p => p.Trim()).Where(p => !string.IsNullOrWhiteSpace(p)));
        }

        return steps;
    }

    private static string TruncateText(string text, int maxLength)
    {
        return text.Length <= maxLength ? text : text.Substring(0, maxLength - 3) + "...";
    }
}

// 使用示例
Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

string steps = "1. 接收请求 2. 验证数据 3. 处理业务 4. 返回结果";
RectangleF pos = new RectangleF(50, 50, 700, 200);

ISmartArt flowchart = TextToGraphicConverter.ConvertStepsToFlowchart(
    presentation, presentation.Slides[0], steps, pos
);

presentation.SaveToFile("flowchart.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## Example 2: Cycle Process to Cycle Diagram

```csharp
// 将循环过程转换为循环图
public static ISmartArt ConvertCycleToDiagram(
    Presentation presentation,
    ISlide slide,
    string cycleText,
    RectangleF position)
{
    List<string> steps = ExtractSteps(cycleText);

    ISmartArt cycle = slide.Shapes.AppendSmartArt(position, SmartArtLayoutType.BasicCycle);

    foreach (string step in steps)
    {
        ISmartArtNode node = cycle.Nodes.AddNode();
        node.TextFrame.Text = step;
    }

    // 应用循环风格
    cycle.ColorStyle = SmartArtColorType.GradientLoopAccent1;
    cycle.SmartArtStyle = SmartArtStyleType.WhiteOutline;

    return cycle;
}

// 使用示例
string cycle = "1. 收集反馈 2. 分析问题 3. 改进产品 4. 再次收集";
ISmartArt diagram = TextToGraphicConverter.ConvertCycleToDiagram(
    presentation, presentation.Slides[0], cycle, pos
);
```

## Example 3: Hierarchy Text to Organization Chart

```csharp
// 将层次文本转换为组织结构图
public static ISmartArt ConvertHierarchyToOrgChart(
    Presentation presentation,
    ISlide slide,
    string hierarchyText,
    RectangleF position)
{
    ISmartArt orgChart = slide.Shapes.AppendSmartArt(
        position,
        SmartArtLayoutType.OrganizationalChart
    );

    // 解析层次结构
    var nodes = ParseHierarchy(hierarchyText);

    // 构建组织结构
    if (nodes.Count > 0)
    {
        // 设置根节点
        orgChart.Nodes[0].TextFrame.Text = nodes[0].Text;

        // 添加子节点
        foreach (var child in nodes[0].Children)
        {
            ISmartArtNode childNode = orgChart.Nodes[0].ChildNodes.AddNode();
            childNode.TextFrame.Text = child.Text;

            // 添加孙节点
            foreach (var grandchild in child.Children)
            {
                ISmartArtNode grandNode = childNode.ChildNodes.AddNode();
                grandNode.TextFrame.Text = grandchild.Text;
            }
        }
    }

    orgChart.ColorStyle = SmartArtColorType.Colorful;
    orgChart.SmartArtStyle = SmartArtStyleType.WhiteOutline;

    return orgChart;
}

// 使用示例
string hierarchy = @"
总经理
    技术总监
        开发经理
        测试经理
    市场总监
        销售经理
        推广经理
";

ISmartArt orgChart = TextToGraphicConverter.ConvertHierarchyToOrgChart(
    presentation, presentation.Slides[0], hierarchy, new RectangleF(50, 50, 600, 400)
);
```

## Example 4: Smart Auto-Detection and Conversion

```csharp
// 智能检测文本类型并转换
public static ISmartArt AutoConvertText(
    Presentation presentation,
    ISlide slide,
    string text,
    RectangleF position)
{
    TextType type = DetectTextType(text);
    SmartArtLayoutType layout = GetLayoutForType(type);

    ISmartArt graphic = slide.Shapes.AppendSmartArt(position, layout);

    switch (type)
    {
        case TextType.Process:
            List<string> steps = ExtractSteps(text);
            foreach (string step in steps)
            {
                graphic.Nodes.AddNode().TextFrame.Text = step;
            }
            break;

        case TextType.Cycle:
            List<string> cycleSteps = ExtractSteps(text);
            foreach (string step in cycleSteps)
            {
                graphic.Nodes.AddNode().TextFrame.Text = step;
            }
            break;

        case TextType.Hierarchy:
            BuildHierarchy(graphic, text);
            break;
    }

    return graphic;
}

// 检测文本类型
private static TextType DetectTextType(string text)
{
    // 检测循环关键词
    if (text.Contains("循环") || text.Contains("重复") ||
        text.Contains("loop") || text.Contains("repeat"))
    {
        return TextType.Cycle;
    }

    // 检测层次结构
    if (text.Contains("\n    ") || text.Contains("\n\t"))
    {
        return TextType.Hierarchy;
    }

    // 默认为流程
    return TextType.Process;
}

// 根据类型获取布局
private static SmartArtLayoutType GetLayoutForType(TextType type)
{
    return type switch
    {
        TextType.Process => SmartArtLayoutType.BasicProcess,
        TextType.Cycle => SmartArtLayoutType.BasicCycle,
        TextType.Hierarchy => SmartArtLayoutType.Hierarchy,
        _ => SmartArtLayoutType.BasicProcess
    };
}

public enum TextType
{
    Process,
    Cycle,
    Hierarchy
}

// 使用示例
string mixedText = "1. 收集数据 2. 分析数据 3. 生成报告";
ISmartArt result = TextToGraphicConverter.AutoConvertText(
    presentation, presentation.Slides[0], mixedText, pos
);
```

## Example 5: Batch Convert All Text Shapes in Presentation

```csharp
// 批量转换演示文稿中的所有文本形状
public static void BatchConvertAllText(Presentation presentation)
{
    foreach (ISlide slide in presentation.Slides)
    {
        List<IShape> shapesToRemove = new List<IShape>();

        foreach (IShape shape in slide.Shapes)
        {
            if (shape is IAutoShape autoShape)
            {
                string text = autoShape.TextFrame?.Text ?? "";

                // 检查是否适合转换
                if (ShouldConvert(text))
                {
                    // 保存原文本到备注
                    string existingNotes = slide.NotesSlide?.NotesTextFrame?.Text ?? "";
                    slide.NotesSlide.NotesTextFrame.Text = existingNotes + $"\n[原文本]: {text}";

                    // 转换
                    RectangleF newPos = CalculateNewPosition(autoShape);
                    ISmartArt graphic = AutoConvertText(presentation, slide, text, newPos);

                    // 标记删除原形状
                    shapesToRemove.Add(shape);

                    Console.WriteLine($"转换幻灯片 {slide.SlideNumber} 的文本");
                }
            }
        }

        // 删除已转换的形状
        foreach (IShape shape in shapesToRemove)
        {
            slide.Shapes.Remove(shape);
        }
    }
}

// 判断是否应该转换
private static bool ShouldConvert(string text)
{
    // 至少30个字符
    if (text.Length < 30) return false;

    // 包含步骤编号
    if (System.Text.RegularExpressions.Regex.IsMatch(text, @"\d+\s*[.)、]"))
        return true;

    // 包含关键词
    string[] keywords = { "步骤", "流程", "循环", "层次", "第一步", "开始", "结束" };
    foreach (string keyword in keywords)
    {
        if (text.Contains(keyword)) return true;
    }

    return false;
}

// 计算新位置
private static RectangleF CalculateNewPosition(IAutoShape original)
{
    float newWidth = Math.Max(600, original.Width * 1.5f);
    float newHeight = Math.Max(300, original.Height * 1.5f);

    return new RectangleF(
        original.X,
        original.Y,
        newWidth,
        newHeight
    );
}

// 使用示例
Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

TextToGraphicConverter.BatchConvertAllText(presentation);

presentation.SaveToFile("converted_all.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## Example 6: Custom Flowchart with Decision Diamonds

```csharp
// 创建带决策节点的流程图（使用 Shape）
public static void CreateDecisionFlowchart(
    Presentation presentation,
    ISlide slide,
    string condition,
    string trueAction,
    string falseAction,
    RectangleF position)
{
    float y = position.Y;
    float centerX = position.X + position.Width / 2;

    // 1. 开始节点（椭圆）
    IAutoShape start = slide.Shapes.AppendShape(
        ShapeType.Ellipse,
        new RectangleF(centerX - 50, y, 100, 50)
    );
    start.Fill.FillType = FillFormatType.Solid;
    start.Fill.SolidColor.Color = Color.LightGreen;
    start.AppendTextFrame("开始");
    y += 70;

    // 2. 条件节点（菱形）
    IAutoShape decision = slide.Shapes.AppendShape(
        ShapeType.Diamond,
        new RectangleF(centerX - 50, y, 100, 80)
    );
    decision.Fill.FillType = FillFormatType.Solid;
    decision.Fill.SolidColor.Color = Color.Yellow;
    decision.AppendTextFrame(condition);
    y += 100;

    // 3. Yes 分支
    IAutoShape yesBox = slide.Shapes.AppendShape(
        ShapeType.Rectangle,
        new RectangleF(centerX - 150, y, 120, 60)
    );
    yesBox.Fill.FillType = FillFormatType.Solid;
    yesBox.Fill.SolidColor.Color = Color.LightBlue;
    yesBox.AppendTextFrame(trueAction);

    // 4. No 分支
    IAutoShape noBox = slide.Shapes.AppendShape(
        ShapeType.Rectangle,
        new RectangleF(centerX + 30, y, 120, 60)
    );
    noBox.Fill.FillType = FillFormatType.Solid;
    noBox.Fill.SolidColor.Color = Color.Orange;
    noBox.AppendTextFrame(falseAction);

    // 5. 结束节点
    y += 80;
    IAutoShape end = slide.Shapes.AppendShape(
        ShapeType.Ellipse,
        new RectangleF(centerX - 50, y, 100, 50)
    );
    end.Fill.FillType = FillFormatType.Solid;
    end.Fill.SolidColor.Color = Color.Red;
    end.AppendTextFrame("结束");

    // 添加连接线
    // start -> decision
    slide.Shapes.AppendConnector(ConnectorType.Straight,
        new PointF(centerX, position.Y + 50),
        new PointF(centerX, position.Y + 120)
    );

    // decision -> yes
    slide.Shapes.AppendConnector(ConnectorType.Straight,
        new PointF(centerX - 50, position.Y + 160),
        new PointF(centerX - 90, position.Y + 200)
    );

    // decision -> no
    slide.Shapes.AppendConnector(ConnectorType.Straight,
        new PointF(centerX + 50, position.Y + 160),
        new PointF(centerX + 90, position.Y + 200)
    );
}

// 使用示例
TextToGraphicConverter.CreateDecisionFlowchart(
    presentation,
    presentation.Slides[0],
    "年龄 >= 18?",
    "允许访问",
    "拒绝访问",
    new RectangleF(50, 50, 400, 400)
);
```

## Example 7: Apply Theme and Styles

```csharp
// 为转换后的图形应用主题样式
public static void ApplyThemeToSmartArt(
    ISmartArt smartArt,
    Color primaryColor,
    Color accentColor)
{
    // 设置颜色样式
    smartArt.ColorStyle = SmartArtColorType.Colorful;

    // 设置形状样式
    smartArt.SmartArtStyle = SmartArtStyleType.Powder;

    // 自定义节点颜色
    for (int i = 0; i < smartArt.Nodes.Count; i++)
    {
        ISmartArtNode node = smartArt.Nodes[i];
        bool useAccent = (i % 2 == 1);

        node.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid;
        node.TextFrame.TextRange.Fill.SolidColor.Color = useAccent ? accentColor : primaryColor;
        node.TextFrame.TextRange.FontHeight = 16;
    }
}

// 使用示例
ISmartArt flowchart = TextToGraphicConverter.ConvertStepsToFlowchart(
    presentation, presentation.Slides[0], steps, pos
);

TextToGraphicConverter.ApplyThemeToSmartArt(
    flowchart,
    Color.FromArgb(0, 102, 204),
    Color.FromArgb(255, 128, 0)
);
```

## Helper Functions

```csharp
// 解析层次结构
private static List<HierarchyNode> ParseHierarchy(string text)
{
    List<HierarchyNode> result = new List<HierarchyNode>();
    string[] lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
    HierarchyNode currentRoot = null;

    foreach (string line in lines)
    {
        string trimmed = line.Trim();
        if (string.IsNullOrWhiteSpace(trimmed)) continue;

        int level = CountIndent(line);
        HierarchyNode node = new HierarchyNode
        {
            Text = trimmed,
            Level = level,
            Children = new List<HierarchyNode>()
        };

        if (level == 0)
        {
            result.Add(node);
            currentRoot = node;
        }
    }

    return result;
}

// 计算缩进级别
private static int CountIndent(string line)
{
    int count = 0;
    foreach (char c in line)
    {
        if (c == '\t') count++;
        else if (c == ' ') count += 4;
        else break;
    }
    return count / 4;
}

// 构建层次树
private static void BuildHierarchy(ISmartArt smartArt, string text)
{
    var nodes = ParseHierarchy(text);
    if (nodes.Count > 0)
    {
        smartArt.Nodes[0].TextFrame.Text = nodes[0].Text;
        foreach (var child in nodes[0].Children)
        {
            ISmartArtNode childNode = smartArt.Nodes[0].ChildNodes.AddNode();
            childNode.TextFrame.Text = child.Text;
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
