---
title: SmartArt 图形
category: spire-presentation
description: 使用 Spire.Presentation 创建和编辑 SmartArt 图形
---

# SmartArt 图形

## 概述

SmartArt 是 PowerPoint 中用于创建专业图形的强大工具，包括流程图、循环图、层次结构图等。Spire.Presentation 提供了完整的 SmartArt 操作功能。

## SmartArt 布局类型

### 流程图类
- `SmartArtLayoutType.BasicBlockProcess` - 基本块流程
- `SmartArtLayoutType.BasicChevronProcess` - 基本V形流程
- `SmartArtLayoutType.BasicProcess` - 基本流程
- `SmartArtLayoutType.ContinuousArrowProcess` - 连续箭头流程
- `SmartArtLayoutType.AccentProcess` - 强调流程

### 循环图类
- `SmartArtLayoutType.BasicCycle` - 基本循环
- `SmartArtLayoutType.BlockCycle` - 块循环
- `SmartArtLayoutType.Cycle` - 循环

### 层次结构类
- `SmartArtLayoutType.Hierarchy` - 层次结构
- `SmartArtLayoutType.OrganizationalChart` - 组织结构图
- `SmartArtLayoutType.HorizontalHierarchy` - 水平层次结构

### 关系图类
- `SmartArtLayoutType.Balance` - 平衡
- `SmartArtLayoutType.Collage` - 拼贴画
- `SmartArtLayoutType.ConvergingRadial` - 汇聚辐射

### 列表类
- `SmartArtLayoutType.BasicBendingProcess` - 基本弯曲流程
- `SmartArtLayoutType.ChevronProcess` - V形流程
- `SmartArtLayoutType.Process` - 流程

### 其他
- `SmartArtLayoutType.Pyramid` - 金字塔
- `SmartArtLayoutType.RadialCycle` - 辐射循环
- `SmartArtLayoutType.Matrix` - 矩阵

## 示例

### 示例 1: 添加 SmartArt

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Diagrams;

Presentation presentation = new Presentation();

// 添加 SmartArt
RectangleF rect = new RectangleF(50, 50, 500, 300);
ISmartArt smartArt = presentation.Slides[0].Shapes.AppendSmartArt(
    rect,
    SmartArtLayoutType.BasicCycle
);

presentation.SaveToFile("SmartArt.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 2: 添加 SmartArt 节点

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Diagrams;

Presentation presentation = new Presentation();

// 创建 SmartArt
RectangleF rect = new RectangleF(50, 50, 500, 300);
ISmartArt smartArt = presentation.Slides[0].Shapes.AppendSmartArt(
    rect,
    SmartArtLayoutType.BasicProcess
);

// 添加节点
ISmartArtNode node1 = smartArt.Nodes.AddNode();
node1.TextFrame.Text = "步骤 1";

ISmartArtNode node2 = smartArt.Nodes.AddNode();
node2.TextFrame.Text = "步骤 2";

ISmartArtNode node3 = smartArt.Nodes.AddNode();
node3.TextFrame.Text = "步骤 3";

presentation.SaveToFile("SmartArtWithNodes.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 3: 在指定位置添加节点

```csharp
// ... 创建 SmartArt

// 在索引 1 的位置添加节点
ISmartArtNode newNode = smartArt.Nodes.AddNode(1);
newNode.TextFrame.Text = "新插入的步骤";

// 设置节点文本样式
newNode.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid;
newNode.TextFrame.TextRange.Fill.SolidColor.Color = Color.Blue;
```

### 示例 4: 删除节点

```csharp
// 删除索引为 1 的节点
smartArt.Nodes.RemoveNode(1);

// 或删除特定节点引用
ISmartArtNode nodeToRemove = smartArt.Nodes[1];
smartArt.Nodes.RemoveNode(nodeToRemove);
```

### 示例 5: 设置 SmartArt 样式

```csharp
using Spire.Presentation.Drawing;

// ... 创建 SmartArt

// 设置颜色样式
smartArt.ColorStyle = SmartArtColorType.Colorful;

// 设置形状样式
smartArt.SmartArtStyle = SmartArtStyleType.WhiteOutline;

// 或使用具体值
smartArt.ColorStyle = SmartArtColorType.Accent1_2;
smartArt.SmartArtStyle = SmartArtStyleType.Powder;
```

### 示例 6: 访问和修改节点

```csharp
// 遍历所有节点
foreach (ISmartArtNode node in smartArt.Nodes)
{
    Console.WriteLine($"节点文本: {node.TextFrame.Text}");

    // 修改节点文本
    node.TextFrame.Text = "新文本: " + node.TextFrame.Text;

    // 设置节点文本颜色
    node.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid;
    node.TextFrame.TextRange.Fill.SolidColor.Color = Color.DarkBlue;
}
```

### 示例 7: 访问子节点（嵌套）

```csharp
// 访问节点的子节点
ISmartArtNode parentNode = smartArt.Nodes[0];
foreach (ISmartArtNode childNode in parentNode.ChildNodes)
{
    Console.WriteLine($"子节点: {childNode.TextFrame.Text}");
}
```

### 示例 8: 创建组织结构图

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Diagrams;

Presentation presentation = new Presentation();

// 创建组织结构图
RectangleF rect = new RectangleF(50, 50, 600, 400);
ISmartArt orgChart = presentation.Slides[0].Shapes.AppendSmartArt(
    rect,
    SmartArtLayoutType.OrganizationalChart
);

// 添加根节点（CEO）
ISmartArtNode ceo = orgChart.Nodes[0];
ceo.TextFrame.Text = "CEO";

// 添加经理节点（子节点）
ISmartArtNode manager1 = ceo.ChildNodes.AddNode();
manager1.TextFrame.Text = "技术经理";

ISmartArtNode manager2 = ceo.ChildNodes.AddNode();
manager2.TextFrame.Text = "运营经理";

// 添加员工节点
ISmartArtNode dev1 = manager1.ChildNodes.AddNode();
dev1.TextFrame.Text = "开发人员";

ISmartArtNode dev2 = manager1.ChildNodes.AddNode();
dev2.TextFrame.Text = "测试人员";

ISmartArtNode ops1 = manager2.ChildNodes.AddNode();
ops1.TextFrame.Text = "运营专员";

// 设置样式
orgChart.ColorStyle = SmartArtColorType.Colorful;
orgChart.SmartArtStyle = SmartArtStyleType.WhiteOutline;

presentation.SaveToFile("OrganizationChart.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 9: 创建流程图

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Diagrams;

Presentation presentation = new Presentation();

// 创建流程图
RectangleF rect = new RectangleF(50, 50, 600, 200);
ISmartArt flowChart = presentation.Slides[0].Shapes.AppendSmartArt(
    rect,
    SmartArtLayoutType.BasicProcess
);

// 添加流程步骤
string[] steps = { "开始", "分析需求", "设计方案", "开发实现", "测试", "部署", "结束" };
for (int i = 0; i < steps.Length; i++)
{
    ISmartArtNode node = flowChart.Nodes.AddNode();
    node.TextFrame.Text = steps[i];
}

// 设置颜色
flowChart.ColorStyle = SmartArtColorType.GradientLoopAccent1;
flowChart.SmartArtStyle = SmartArtStyleType.WhiteOutline;

presentation.SaveToFile("FlowChart.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 10: 访问特定节点

```csharp
// 通过索引访问
ISmartArtNode nodeByIndex = smartArt.Nodes[2];

// 遍历查找特定节点
ISmartArtNode targetNode = null;
foreach (ISmartArtNode node in smartArt.Nodes)
{
    if (node.TextFrame.Text.Contains("特定文本"))
    {
        targetNode = node;
        break;
    }
}

// 访问父节点
if (smartArt.Nodes[0].ParentNode != null)
{
    ISmartArtNode parent = smartArt.Nodes[0].ParentNode;
}
```

### 示例 11: 设置助理节点

```csharp
// 在组织结构图中添加助理节点
ISmartArtNode manager = orgChart.Nodes[0].ChildNodes[0];

// 添加助理
ISmartArtNode assistant = manager.ChildNodes.AddNode();
assistant.IsAssistant = true;
assistant.TextFrame.Text = "助理";

// 设置助理节点样式
assistant.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid;
assistant.TextFrame.TextRange.Fill.SolidColor.Color = Color.Gray;
```

### 示例 12: 编辑现有 SmartArt

```csharp
// 加载包含 SmartArt 的演示文稿
Presentation presentation = new Presentation();
presentation.LoadFromFile("existing.pptx");

// 获取幻灯片上的 SmartArt
ISmartArt smartArt = presentation.Slides[0].Shapes[0] as ISmartArt;

if (smartArt != null)
{
    // 修改节点文本
    smartArt.Nodes[0].TextFrame.Text = "更新的标题";

    // 添加新节点
    ISmartArtNode newNode = smartArt.Nodes.AddNode();
    newNode.TextFrame.Text = "新增步骤";

    // 更改样式
    smartArt.ColorStyle = SmartArtColorType.Dark1Outline;
}

presentation.SaveToFile("updated_smartart.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 13: 获取 SmartArt 布局信息

```csharp
// 获取 SmartArt 信息
ISmartArt smartArt = presentation.Slides[0].Shapes[0] as ISmartArt;

if (smartArt != null)
{
    Console.WriteLine($"布局类型: {smartArt.LayoutType}");
    Console.WriteLine($"颜色样式: {smartArt.ColorStyle}");
    Console.WriteLine($"形状样式: {smartArt.SmartArtStyle}");
    Console.WriteLine($"节点数量: {smartArt.Nodes.Count}");

    // 获取每个节点的详细信息
    foreach (ISmartArtNode node in smartArt.Nodes)
    {
        Console.WriteLine($"节点: {node.TextFrame.Text}");
        Console.WriteLine($"  子节点数: {node.ChildNodes.Count}");
    }
}
```

## SmartArt 样式和颜色

### 颜色样式 (SmartArtColorType)

| 样式 | 描述 |
|------|------|
| `Colorful` - 彩色 |
| `GradientLoopAccent1` - 渐变循环强调色 |
| `GradientLoopAccent2` - 渐变循环强调色2 |
| `GradientLoopAccent3` - 渐变循环强调色3 |
| `GradientLoopAccent4` - 渐变循环强调色4 |
| `Dark1Outline` - 深色轮廓 |
| `Light1Outline` - 浅色轮廓 |
| `Accented1` - 强调色1 |
| `Accented2` - 强调色2 |

### 形状样式 (SmartArtStyleType)

| 样式 | 描述 |
|------|------|
| `WhiteOutline` - 白色轮廓 |
| `SubtleEffect` - 细微效果 |
| `ModerateEffect` - 中等效果 |
| `IntenseEffect` - 强烈效果 |
| `Powder` - 粉末 |
| `Cartoon` - 卡通 |
| `Intense` - 强烈 |
| `Circle` - 圆形 |
| `Simple` - 简单 |
| `Polished` - 抛光 |

## SmartArt 组件

### 主要属性

| 属性 | 描述 |
|------|------|
| `LayoutType` | 布局类型 |
| `ColorStyle` | 颜色样式 |
| `SmartArtStyle` | 形状样式 |
| `Nodes` | 节点集合 |
| `Width` | 宽度 |
| `Height` | 高度 |

### ISmartArtNode 属性

| 属性 | 描述 |
|------|------|
| `TextFrame` - 文本框 |
| `ChildNodes` - 子节点集合 |
| `ParentNode` - 父节点 |
| `IsAssistant` - 是否为助理节点 |

## 注意事项

1. **布局限制**: 某些布局类型对节点数量有限制
2. **文本长度**: 节点文本不宜过长，否则影响显示效果
3. **样式兼容性**: 部分样式组合可能不兼容，建议测试后使用
4. **性能**: 大量节点可能影响性能，建议合理控制节点数量

## 相关功能

- [文本处理](./03-text-content.md) - SmartArt 节点文本格式化
- [形状处理](./04-shapes-images.md) - SmartArt 作为形状处理
- [动画](./09-animations.md) - 为 SmartArt 添加动画效果
