---
title: 文本和段落
category: spire-presentation
description: 使用 Spire.Presentation 添加和管理文本内容
---

# 文本和段落

## 概述

Spire.Presentation 提供了强大的文本处理功能，包括：
- 添加和编辑文本
- 段落格式化（对齐、缩进、行距）
- 文本样式（字体、颜色、大小）
- 项目符号和编号
- HTML 内容
- 文本框设置

## 添加文本

### 示例 1: 添加简单文本

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();

// 添加形状作为文本框
RectangleF rect = new RectangleF(50, 50, 400, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);

// 设置形状样式
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.White;
shape.ShapeStyle.LineColor.Color = Color.Black;

// 添加文本
shape.AppendTextFrame("Hello, World!");

// 设置文本样式
shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 24;
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.Black;

presentation.SaveToFile("simple_text.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 2: 添加多段文本

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 500, 200);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);

// 添加第一个段落
TextParagraph para1 = new TextParagraph();
para1.Text = "这是第一段文本内容。";
shape.TextFrame.Paragraphs.Append(para1);

// 添加第二个段落
TextParagraph para2 = new TextParagraph();
para2.Text = "这是第二段文本内容。";
shape.TextFrame.Paragraphs.Append(para2);

// 添加第三个段落
TextParagraph para3 = new TextParagraph();
para3.Text = "这是第三段文本内容。";
shape.TextFrame.Paragraphs.Append(para3);

presentation.SaveToFile("multiple_paragraphs.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 3: 在形状中添加文本

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 添加圆形
RectangleF rect = new RectangleF(100, 100, 150, 150);
IAutoShape circle = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Ellipse,
    rect
);
circle.Fill.FillType = FillFormatType.Solid;
circle.Fill.SolidColor.Color = Color.LightBlue;
circle.AppendTextFrame("圆形");

// 添加三角形
RectangleF rect2 = new RectangleF(300, 100, 150, 150);
IAutoShape triangle = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Triangle,
    rect2
);
triangle.Fill.FillType = FillFormatType.Solid;
triangle.Fill.SolidColor.Color = Color.LightGreen;
triangle.AppendTextFrame("三角形");

presentation.SaveToFile("shapes_with_text.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 段落格式化

### 示例 4: 设置对齐方式

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 400, 200);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);

// 添加段落
TextParagraph para1 = new TextParagraph();
para1.Text = "左对齐文本";
para1.Alignment = TextAlignmentType.Left;
shape.TextFrame.Paragraphs.Append(para1);

TextParagraph para2 = new TextParagraph();
para2.Text = "居中对齐文本";
para2.Alignment = TextAlignmentType.Center;
shape.TextFrame.Paragraphs.Append(para2);

TextParagraph para3 = new TextParagraph();
para3.Text = "右对齐文本";
para3.Alignment = TextAlignmentType.Right;
shape.TextFrame.Paragraphs.Append(para3);

TextParagraph para4 = new TextParagraph();
para4.Text = "两端对齐文本";
para4.Alignment = TextAlignmentType.Justify;
shape.TextFrame.Paragraphs.Append(para4);

presentation.SaveToFile("text_alignment.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 5: 设置缩进和行距

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 500, 200);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);

TextParagraph para = new TextParagraph();
para.Text = "这是一个设置了缩进和行距的段落示例。Spire.Presentation 提供了丰富的段落格式化选项。";

// 设置首行缩进
para.Indent = 30;

// 设置行距（单位：磅）
para.LineSpacing = 1.5f;

// 设置行距类型（百分比）
para.LineSpacingType = TextLineSpacingType.Percent;
para.LineSpacing = 150; // 150%

shape.TextFrame.Paragraphs.Append(para);

presentation.SaveToFile("paragraph_formatting.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 6: 设置段落间距

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 400, 200);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);

TextParagraph para1 = new TextParagraph();
para1.Text = "第一段";
para1.SpaceAfter = 20; // 段后间距
para1.SpaceBefore = 10; // 段前间距
shape.TextFrame.Paragraphs.Append(para1);

TextParagraph para2 = new TextParagraph();
para2.Text = "第二段";
para2.SpaceAfter = 20;
shape.TextFrame.Paragraphs.Append(para2);

TextParagraph para3 = new TextParagraph();
para3.Text = "第三段";
shape.TextFrame.Paragraphs.Append(para3);

presentation.SaveToFile("paragraph_spacing.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 文本样式

### 示例 7: 设置字体和颜色

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 400, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.AppendTextFrame("样式化文本");

// 获取文本范围
TextRange textRange = shape.TextFrame.Paragraphs[0].TextRanges[0];

// 设置字体
textRange.LatinFont = new TextFont("Arial");
textRange.EastAsianFont = new TextFont("宋体");
textRange.FontHeight = 32;

// 设置颜色
textRange.Fill.FillType = FillFormatType.Solid;
textRange.Fill.SolidColor.Color = Color.Blue;

// 设置加粗
textRange.IsBold = TriState.True;

// 设置斜体
textRange.IsItalic = TriState.True;

// 设置下划线
textRange.FontUnderlineType = TextUnderlineType.Single;

presentation.SaveToFile("styled_text.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 8: 部分文本样式

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 500, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.AppendTextFrame("这是普通的文本，这是加粗的文本，这是红色的文本");

// 为不同的文本范围设置样式
TextRange textRange = shape.TextFrame.Paragraphs[0].TextRanges[0];

// 设置加粗文本（第10-14个字符）
textRange = shape.TextFrame.Paragraphs[0].TextRanges[0];
textRange.PortionCount = 3;
textRange[0].Fill.FillType = FillFormatType.Solid;
textRange[0].Fill.SolidColor.Color = Color.Black;

textRange[1].Fill.FillType = FillFormatType.Solid;
textRange[1].Fill.SolidColor.Color = Color.Black;
textRange[1].IsBold = TriState.True;

textRange[2].Fill.FillType = FillFormatType.Solid;
textRange[2].Fill.SolidColor.Color = Color.Red;

presentation.SaveToFile("partial_styling.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 9: 设置文本背景

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 400, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.AppendTextFrame("带背景的文本");

// 设置文本背景
TextRange textRange = shape.TextFrame.Paragraphs[0].TextRanges[0];
textRange.Fill.FillType = FillFormatType.Solid;
textRange.Fill.SolidColor.Color = Color.White;

// 设置段落背景
shape.TextFrame.Paragraphs[0].Fill.FillType = FillFormatType.Solid;
shape.TextFrame.Paragraphs[0].Fill.SolidFillColor.Color = Color.LightYellow;

presentation.SaveToFile("text_background.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 项目符号

### 示例 10: 添加项目符号

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 400, 200);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);

// 添加带项目符号的段落
TextParagraph para1 = new TextParagraph();
para1.Text = "第一项";
para1.Bullet.Type = TextBulletType.Symbol;
para1.Bullet.Char = '●'; // 项目符号字符
para1.Bullet.Height = 15;
shape.TextFrame.Paragraphs.Append(para1);

TextParagraph para2 = new TextParagraph();
para2.Text = "第二项";
para2.Bullet.Type = TextBulletType.Symbol;
para2.Bullet.Char = '●';
para2.Bullet.Height = 15;
shape.TextFrame.Paragraphs.Append(para2);

TextParagraph para3 = new TextParagraph();
para3.Text = "第三项";
para3.Bullet.Type = TextBulletType.Symbol;
para3.Bullet.Char = '●';
para3.Bullet.Height = 15;
shape.TextFrame.Paragraphs.Append(para3);

presentation.SaveToFile("bullet_points.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 11: 使用自定义项目符号

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 400, 200);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);

// 添加带自定义项目符号的段落
TextParagraph para1 = new TextParagraph();
para1.Text = "重要事项";
para1.Bullet.Type = TextBulletType.Symbol;
para1.Bullet.Char = '★';
para1.Bullet.Height = 20;
shape.TextFrame.Paragraphs.Append(para1);

TextParagraph para2 = new TextParagraph();
para2.Text = "注意事项";
para2.Bullet.Type = TextBulletType.Symbol;
para2.Bullet.Char = '★';
para2.Bullet.Height = 20;
shape.TextFrame.Paragraphs.Append(para2);

presentation.SaveToFile("custom_bullets.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 12: 添加编号列表

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 400, 200);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);

// 添加编号列表
TextParagraph para1 = new TextParagraph();
para1.Text = "第一步骤";
para1.Bullet.Type = TextBulletType.Numbered;
para1.Bullet.NumberedBulletStyle = NumberedBulletStyle.ArabicPeriod;
shape.TextFrame.Paragraphs.Append(para1);

TextParagraph para2 = new TextParagraph();
para2.Text = "第二步骤";
para2.Bullet.Type = TextBulletType.Numbered;
para2.Bullet.NumberedBulletStyle = NumberedBulletStyle.ArabicPeriod;
shape.TextFrame.Paragraphs.Append(para2);

TextParagraph para3 = new TextParagraph();
para3.Text = "第三步骤";
para3.Bullet.Type = TextBulletType.Numbered;
para3.Bullet.NumberedBulletStyle = NumberedBulletStyle.ArabicPeriod;
shape.TextFrame.Paragraphs.Append(para3);

presentation.SaveToFile("numbered_list.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 文本框设置

### 示例 13: 设置文本框属性

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 400, 200);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);

shape.AppendTextFrame("这是自动调整大小的文本框示例。");

// 设置自动调整大小
shape.TextFrame.AutofitType = TextAutofitType.Normal;

// 设置垂直对齐
shape.TextFrame.VerticalAlignment = VerticalAlignmentType.Top;

// 设置文本框边距
shape.TextFrame.MarginTop = 10;
shape.TextFrame.MarginBottom = 10;
shape.TextFrame.MarginLeft = 10;
shape.TextFrame.MarginRight = 10;

// 设置文本方向（垂直文本）
// shape.TextFrame.TextVerticalType = TextVerticalType.Vertical;

presentation.SaveToFile("textbox_settings.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 14: 文本自动换行

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 200, 150);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);

shape.AppendTextFrame("这是一段很长的文本，它会自动换行以适应文本框的宽度。Spire.Presentation 提供了自动换行功能。");

// 设置自动换行
shape.TextFrame.WrapText = TriState.True;

// 设置文本垂直对齐
shape.TextFrame.VerticalAlignment = VerticalAlignmentType.Top;

presentation.SaveToFile("word_wrap.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## HTML 内容

### 示例 15: 添加 HTML 内容

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 500, 200);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);

// 添加 HTML 内容
string html = "<b>这是粗体文本</b><br><i>这是斜体文本</i><br><u>这是带下划线的文本</u>";
shape.TextFrame.HtmlText = html;

presentation.SaveToFile("html_content.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 边框和底纹

### 示例 16: 设置段落边框

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 400, 150);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);

TextParagraph para = new TextParagraph();
para.Text = "带边框的段落";

// 设置段落边框
para.BorderTop.FillType = FillFormatType.Solid;
para.BorderTop.SolidFillColor.Color = Color.Black;
para.BorderTop.Width = 1;

para.BorderBottom.FillType = FillFormatType.Solid;
para.BorderBottom.SolidFillColor.Color = Color.Black;
para.BorderBottom.Width = 1;

para.BorderLeft.FillType = FillFormatType.Solid;
para.BorderLeft.SolidFillColor.Color = Color.Black;
para.BorderLeft.Width = 1;

para.BorderRight.FillType = FillFormatType.Solid;
para.BorderRight.SolidFillColor.Color = Color.Black;
para.BorderRight.Width = 1;

shape.TextFrame.Paragraphs.Append(para);

presentation.SaveToFile("paragraph_border.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 文本效果

### 示例 17: 设置文本阴影

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();

RectangleF rect = new RectangleF(50, 50, 400, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.AppendTextFrame("带阴影的文本");

// 设置文本阴影
TextRange textRange = shape.TextFrame.Paragraphs[0].TextRanges[0];
textRange.EffectOuterGlow = true;
textRange.EffectOuterGlowColor = Color.Gray;
textRange.EffectOuterGlowSize = 10;

presentation.SaveToFile("text_shadow.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 文本相关类型

### TextAlignmentType

| 对齐方式 | 描述 |
|----------|------|
| `Left` - 左对齐 |
| `Center` - 居中对齐 |
| `Right` - 右对齐 |
| `Justify` - 两端对齐 |

### VerticalAlignmentType

| 对齐方式 | 描述 |
|----------|------|
| `Top` - 顶部对齐 |
| `Middle` - 中部对齐 |
| `Bottom` - 底部对齐 |

### TextLineSpacingType

| 类型 | 描述 |
|------|------|
| `Single` - 单倍行距 |
| `OnePointFive` - 1.5 倍行距 |
| `Double` - 双倍行距 |
| `Percent` - 百分比 |
| `Points` - 磅值 |

### TextAutofitType

| 类型 | 描述 |
|------|------|
| `None` - 不自动调整 |
| `Normal` - 正常调整 |
| `ResizeShapeToFitText` - 调整形状以适应文本 |
| `ResizeTextToFitShape` - 调整文本以适应形状 |

## 注意事项

1. **字体兼容性**: 确保使用的字体在目标系统上可用
2. **文本长度**: 过长的文本可能影响显示效果
3. **编码**: 使用正确的编码处理特殊字符
4. **样式继承**: 文本样式可能继承自模板

## 最佳实践

1. **使用模板**: 定义统一的文本样式模板
2. **文本框大小**: 为文本框预留足够的空间
3. **字体选择**: 使用常用字体以确保兼容性
4. **测试输出**: 在不同 PowerPoint 版本中测试显示效果

## 相关功能

- [形状处理](./04-shapes-images.md) - 文本框作为形状
- [表格](./05-tables.md) - 表格中的文本
- [图表](./06-charts.md) - 图表标签文本
