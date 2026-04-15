---
title: 超链接
category: spire-presentation
description: 使用 Spire.Presentation 添加和管理超链接
---

# 超链接

## 概述

Spire.Presentation 提供了完整的超链接功能，包括：
- 文本超链接
- 图片超链接
- 链接到其他幻灯片
- 链接到外部网站
- 链接到文件
- 修改和删除超链接

## 示例

### 示例 1: 为文本添加超链接

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 添加形状和文本
RectangleF rect = new RectangleF(50, 50, 400, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.White;
shape.AppendTextFrame("访问我们的网站");

// 为文本添加超链接
TextRange textRange = shape.TextFrame.Paragraphs[0].TextRanges[0];
ClickHyperlink hyperlink = new ClickHyperlink("https://www.example.com");
textRange.ClickAction = hyperlink;

// 设置超链接颜色
textRange.ClickAction.Address = "https://www.example.com";
textRange.ClickAction.ActionType = HyperlinkType.Hyperlink;
textRange.ClickAction.TextClickEffect = TextClickEffectType.Color;

presentation.SaveToFile("TextHyperlink.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 2: 为图片添加超链接

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 插入图片
RectangleF imageRect = new RectangleF(100, 100, 200, 150);
IEmbedImage image = presentation.Slides[0].Shapes.AppendEmbedImage(
    ShapeType.Rectangle,
    "logo.png",
    imageRect
);

// 为图片添加超链接
ClickHyperlink hyperlink = new ClickHyperlink("https://www.company.com");
image.Click = hyperlink;

presentation.SaveToFile("ImageHyperlink.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 3: 链接到特定幻灯片

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 添加多张幻灯片
presentation.Slides.Append();
presentation.Slides.Append();
presentation.Slides.Append();

// 在第一张幻灯片添加链接形状
RectangleF rect = new RectangleF(50, 50, 400, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.AppendTextFrame("跳转到第3张幻灯片");

// 创建链接到第3张幻灯片的超链接（索引2）
ClickHyperlink hyperlink = new ClickHyperlink(presentation.Slides[2]);
shape.TextFrame.Paragraphs[0].TextRanges[0].ClickAction = hyperlink;

presentation.SaveToFile("SlideHyperlink.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 4: 链接到最后查看的幻灯片

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 添加幻灯片
presentation.Slides.Append();
presentation.Slides.Append();

// 添加返回按钮
RectangleF rect = new RectangleF(50, 400, 100, 50);
IAutoShape shape = presentation.Slides[1].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.LightGray;
shape.AppendTextFrame("返回");

// 创建链接到最后查看幻灯片的超链接
ClickHyperlink hyperlink = new ClickHyperlink();
hyperlink.ActionType = HyperlinkType.LastSlideViewed;
shape.Click = hyperlink;

presentation.SaveToFile("LastSlideHyperlink.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 5: 链接到文件

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 添加链接形状
RectangleF rect = new RectangleF(50, 50, 400, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.AppendTextFrame("打开详细信息文档");

// 创建链接到文件的超链接
ClickHyperlink hyperlink = new ClickHyperlink();
hyperlink.ActionType = HyperlinkType.OtherFile;
hyperlink.Address = @"C:\Documents\details.docx";
shape.TextFrame.Paragraphs[0].TextRanges[0].ClickAction = hyperlink;

presentation.SaveToFile("FileHyperlink.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 6: 链接到电子邮件

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 添加联系信息形状
RectangleF rect = new RectangleF(50, 50, 400, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.AppendTextFrame("联系我们");

// 创建电子邮件链接
ClickHyperlink hyperlink = new ClickHyperlink();
hyperlink.ActionType = HyperlinkType.Hyperlink;
hyperlink.Address = "mailto:contact@example.com?subject=咨询&body=您好，我想咨询...";
shape.TextFrame.Paragraphs[0].TextRanges[0].ClickAction = hyperlink;

presentation.SaveToFile("EmailHyperlink.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 7: 修改超链接

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 查找包含超链接的形状
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        // 检查形状的超链接
        if (shape.Click != null)
        {
            Console.WriteLine($"找到超链接: {shape.Click.Address}");

            // 修改超链接
            shape.Click.Address = "https://www.new-url.com";
        }

        // 检查文本中的超链接
        if (shape is IAutoShape autoShape)
        {
            foreach (TextParagraph para in autoShape.TextFrame.Paragraphs)
            {
                foreach (TextRange range in para.TextRanges)
                {
                    if (range.ClickAction != null)
                    {
                        Console.WriteLine($"找到文本超链接: {range.ClickAction.Address}");

                        // 修改超链接
                        range.ClickAction.Address = "https://www.new-url.com";
                    }
                }
            }
        }
    }
}

presentation.SaveToFile("ModifiedHyperlinks.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 8: 删除超链接

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 删除所有形状的超链接
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        // 删除形状超链接
        shape.Click = null;

        // 删除文本超链接
        if (shape is IAutoShape autoShape)
        {
            foreach (TextParagraph para in autoShape.TextFrame.Paragraphs)
            {
                foreach (TextRange range in para.TextRanges)
                {
                    range.ClickAction = null;
                }
            }
        }
    }
}

presentation.SaveToFile("HyperlinksRemoved.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 9: 获取所有超链接

```csharp
using Spire.Presentation;
using System.Collections.Generic;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

List<string> hyperlinks = new List<string>();

// 收集所有超链接
foreach (ISlide slide in presentation.Slides)
{
    Console.WriteLine($"幻灯片 {slide.SlideNumber}:");

    foreach (IShape shape in slide.Shapes)
    {
        // 形状超链接
        if (shape.Click != null && !string.IsNullOrEmpty(shape.Click.Address))
        {
            Console.WriteLine($"  形状超链接: {shape.Click.Address}");
            hyperlinks.Add(shape.Click.Address);
        }

        // 文本超链接
        if (shape is IAutoShape autoShape)
        {
            foreach (TextParagraph para in autoShape.TextFrame.Paragraphs)
            {
                foreach (TextRange range in para.TextRanges)
                {
                    if (range.ClickAction != null &&
                        !string.IsNullOrEmpty(range.ClickAction.Address))
                    {
                        Console.WriteLine($"  文本超链接: {range.ClickAction.Address}");
                        Console.WriteLine($"    文本内容: {range.Text}");
                        hyperlinks.Add(range.ClickAction.Address);
                    }
                }
            }
        }
    }
}

Console.WriteLine($"\n共找到 {hyperlinks.Count} 个超链接");
presentation.Dispose();
```

### 示例 10: 为 SmartArt 添加超链接

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

// 添加节点并设置超链接
ISmartArtNode node1 = smartArt.Nodes.AddNode();
node1.TextFrame.Text = "产品介绍";

// 为节点添加超链接
ClickHyperlink hyperlink = new ClickHyperlink("https://www.example.com/products");
node1.TextFrame.TextRange.ClickAction = hyperlink;

presentation.SaveToFile("SmartArtHyperlink.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 11: 设置超链接样式

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();

// 添加文本形状
RectangleF rect = new RectangleF(50, 50, 400, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.AppendTextFrame("访问网站");

// 添加超链接
TextRange textRange = shape.TextFrame.Paragraphs[0].TextRanges[0];
ClickHyperlink hyperlink = new ClickHyperlink("https://www.example.com");
textRange.ClickAction = hyperlink;

// 设置超链接颜色（下划线）
textRange.ClickAction.TextClickEffect = TextClickEffectType.Color;
textRange.Fill.FillType = FillFormatType.Solid;
textRange.Fill.SolidColor.Color = Color.Blue;

// 添加下划线
textRange.FontUnderlineType = TextUnderlineType.Single;

presentation.SaveToFile("StyledHyperlink.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 12: 创建导航菜单

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 添加目标幻灯片
presentation.Slides.Append(); // 幻灯片 1
presentation.Slides.Append(); // 幻灯片 2
presentation.Slides.Append(); // 幻灯片 3

// 在第一张幻灯片创建导航菜单
string[] menuItems = { "首页", "产品", "服务", "联系" };
for (int i = 0; i < menuItems.Length; i++)
{
    RectangleF rect = new RectangleF(50, 50 + i * 40, 200, 35);
    IAutoShape menuItem = presentation.Slides[0].Shapes.AppendShape(
        ShapeType.Rectangle,
        rect
    );
    menuItem.Fill.FillType = FillFormatType.Solid;
    menuItem.Fill.SolidColor.Color = Color.LightBlue;
    menuItem.ShapeStyle.LineColor.Color = Color.DarkBlue;
    menuItem.AppendTextFrame(menuItems[i]);

    // 添加超链接到对应幻灯片
    ClickHyperlink hyperlink = new ClickHyperlink(presentation.Slides[i + 1]);
    menuItem.TextFrame.Paragraphs[0].TextRanges[0].ClickAction = hyperlink;
}

// 在其他幻灯片添加返回按钮
for (int i = 1; i < presentation.Slides.Count; i++)
{
    RectangleF backRect = new RectangleF(50, 400, 100, 35);
    IAutoShape backButton = presentation.Slides[i].Shapes.AppendShape(
        ShapeType.Rectangle,
        backRect
    );
    backButton.Fill.FillType = FillFormatType.Solid;
    backButton.Fill.SolidColor.Color = Color.LightGray;
    backButton.AppendTextFrame("返回");

    // 添加超链接到首页
    ClickHyperlink backLink = new ClickHyperlink(presentation.Slides[0]);
    backButton.TextFrame.Paragraphs[0].TextRanges[0].ClickAction = backLink;
}

presentation.SaveToFile("NavigationMenu.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 13: 屏幕提示（Tooltip）

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 添加形状
RectangleF rect = new RectangleF(50, 50, 400, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.AppendTextFrame("点击查看详细信息");

// 添加超链接
ClickHyperlink hyperlink = new ClickHyperlink("https://www.example.com");
shape.TextFrame.Paragraphs[0].TextRanges[0].ClickAction = hyperlink;

// 设置屏幕提示
hyperlink.Tip = "这将打开公司网站，了解更多信息";

presentation.SaveToFile("TooltipHyperlink.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 14: 批量替换超链接

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 批量替换旧链接为新链接
string oldLink = "https://www.old-domain.com";
string newLink = "https://www.new-domain.com";

int replacedCount = 0;

foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        // 替换形状超链接
        if (shape.Click != null && shape.Click.Address == oldLink)
        {
            shape.Click.Address = newLink;
            replacedCount++;
        }

        // 替换文本超链接
        if (shape is IAutoShape autoShape)
        {
            foreach (TextParagraph para in autoShape.TextFrame.Paragraphs)
            {
                foreach (TextRange range in para.TextRanges)
                {
                    if (range.ClickAction != null &&
                        range.ClickAction.Address == oldLink)
                    {
                        range.ClickAction.Address = newLink;
                        replacedCount++;
                    }
                }
            }
        }
    }
}

Console.WriteLine($"已替换 {replacedCount} 个超链接");
presentation.SaveToFile("LinksReplaced.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 15: 验证超链接

```csharp
using Spire.Presentation;
using System.Net;
using System.Collections.Generic;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 验证所有超链接
List<HyperlinkInfo> links = new List<HyperlinkInfo>();

foreach (ISlide slide in presentation.Slides)
{
    int slideIndex = slide.SlideNumber - 1;

    foreach (IShape shape in slide.Shapes)
    {
        if (shape.Click != null && !string.IsNullOrEmpty(shape.Click.Address))
        {
            bool isValid = ValidateHyperlink(shape.Click.Address);
            links.Add(new HyperlinkInfo
            {
                Slide = slideIndex,
                Type = "形状",
                Url = shape.Click.Address,
                IsValid = isValid
            });
        }

        if (shape is IAutoShape autoShape)
        {
            foreach (TextParagraph para in autoShape.TextFrame.Paragraphs)
            {
                foreach (TextRange range in para.TextRanges)
                {
                    if (range.ClickAction != null &&
                        !string.IsNullOrEmpty(range.ClickAction.Address))
                    {
                        bool isValid = ValidateHyperlink(range.ClickAction.Address);
                        links.Add(new HyperlinkInfo
                        {
                            Slide = slideIndex,
                            Type = "文本",
                            Url = range.ClickAction.Address,
                            IsValid = isValid
                        });
                    }
                }
            }
        }
    }
}

// 输出验证结果
Console.WriteLine("超链接验证结果:");
foreach (var link in links)
{
    Console.WriteLine($"幻灯片 {link.Slide + 1} ({link.Type}): {link.Url} - {(link.IsValid ? "有效" : "无效")}");
}

// 验证函数
bool ValidateHyperlink(string url)
{
    try
    {
        if (!url.StartsWith("http")) return true; // 跳过非HTTP链接

        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
        request.Method = "HEAD";
        using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
        {
            return response.StatusCode == HttpStatusCode.OK;
        }
    }
    catch
    {
        return false;
    }
}

class HyperlinkInfo
{
    public int Slide { get; set; }
    public string Type { get; set; }
    public string Url { get; set; }
    public bool IsValid { get; set; }
}

presentation.Dispose();
```

## 超链接类型

### HyperlinkType

| 类型 | 描述 |
|------|------|
| `Hyperlink` - 普通超链接 |
| `OtherFile` - 链接到文件 |
| `LastSlideViewed` - 最后查看的幻灯片 |
| `FirstSlide` - 第一张幻灯片 |
| `LastSlide` - 最后一张幻灯片 |
| `NextSlide` - 下一张幻灯片 |
| `PreviousSlide` - 上一张幻灯片 |

### TextClickEffectType

| 效果 | 描述 |
|------|------|
| `None` - 无效果 |
| `Color` - 颜色变化 |
| `TextHighlight` - 文本高亮 |

## 注意事项

1. **安全**: 超链接可能导致安全问题，建议验证所有链接
2. **兼容性**: 某些超链接类型在不同 PowerPoint 版本中表现可能不同
3. **相对路径**: 使用相对路径链接文件时要确保路径正确
4. **网络依赖**: 外部链接需要网络连接才能工作

## 最佳实践

1. **使用描述性文本**: 超链接文本应描述目标内容
2. **验证链接**: 定期验证超链接的有效性
3. **提供替代方案**: 为网络链接提供备选联系方式
4. **测试导航**: 测试所有幻灯片间导航链接

## 相关功能

- [文本处理](./03-text-content.md) - 文本超链接格式化
- [形状处理](./04-shapes-images.md) - 形状超链接
- [SmartArt](./07-smartart.md) - SmartArt 超链接
