---
title: 高级功能
category: spire-presentation
description: Spire.Presentation 高级功能包括水印、注释、页眉页脚、备注等
---

# 高级功能

## 概述

Spire.Presentation 提供了许多高级功能来增强演示文稿的专业性和功能性，包括：
- 文本和图片水印
- 注释和备注
- 晔讲者备注自动生成（详见[演讲者备注生成](./17-speaker-notes-generation.md)）
- 全局主题与色调接管（详见[全局主题管理](./18-global-theme-manager.md)）
- 页眉页脚
- SmartArt（详见第7章）
- OLE 对象

## 水印

### 示例 1: 添加文本水印

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 为每张幻灯片添加水印
foreach (ISlide slide in presentation.Slides)
{
    // 插入水印形状（旋转45度）
    RectangleF rect = new RectangleF(
        presentation.SlideSize.Size.Width / 2 - 200,
        presentation.SlideSize.Size.Height / 2 - 50,
        400,
        100
    );

    IAutoShape watermark = slide.Shapes.AppendShape(ShapeType.Rectangle, rect);
    watermark.Rotation = -45;

    // 设置形状样式
    watermark.Fill.FillType = FillFormatType.Solid;
    watermark.Fill.SolidColor.Color = Color.FromArgb(30, Color.Gray); // 半透明
    watermark.ShapeStyle.LineColor.Color = Color.Transparent;
    watermark.LockAspectRatio = false;

    // 添加文本
    watermark.AppendTextFrame("机密文档");

    // 设置文本样式
    watermark.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial");
    watermark.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 36;
    watermark.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
    watermark.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.White;
    watermark.TextFrame.Paragraphs[0].TextRanges[0].IsBold = TriState.True;
    watermark.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;

    // 将水印移到最底层
    watermark.ZOrder(ShapeZOrderType.SendToBack);
}

presentation.SaveToFile("WithWatermark.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 2: 添加图片水印

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 为每张幻灯片添加图片水印
foreach (ISlide slide in presentation.Slides)
{
    RectangleF rect = new RectangleF(
        presentation.SlideSize.Size.Width / 2 - 100,
        presentation.SlideSize.Size.Height / 2 - 100,
        200,
        200
    );

    IEmbedImage watermark = slide.Shapes.AppendEmbedImage(
        ShapeType.Rectangle,
        "watermark.png",
        rect
    );

    // 设置透明度
    watermark.Picture.Fill.PictureTransparency = 0.7f; // 70% 透明

    // 移到最底层
    watermark.ZOrder(ShapeZOrderType.SendToBack);
}

presentation.SaveToFile("WithImageWatermark.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 3: 删除水印

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 删除所有包含特定文本的形状（假设水印是文字）
foreach (ISlide slide in presentation.Slides)
{
    for (int i = slide.Shapes.Count - 1; i >= 0; i--)
    {
        if (slide.Shapes[i] is IAutoShape shape)
        {
            if (shape.TextFrame.Text.Contains("机密") ||
                shape.TextFrame.Text.Contains("水印"))
            {
                slide.Shapes.RemoveAt(i);
            }
        }
    }
}

// 删除所有透明度较高的图片（可能是图片水印）
foreach (ISlide slide in presentation.Slides)
{
    for (int i = slide.Shapes.Count - 1; i >= 0; i--)
    {
        if (slide.Shapes[i] is IEmbedImage image)
        {
            if (image.Picture.Fill.PictureTransparency > 0.5f)
            {
                slide.Shapes.RemoveAt(i);
            }
        }
    }
}

presentation.SaveToFile("WatermarkRemoved.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 注释和备注

### 示例 4: 添加注释

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 添加注释到幻灯片
ISlide slide = presentation.Slides[0];

IComment comment = slide.Comments.AddComment("张三", "这是对内容的说明");
comment.Author = "张三";
comment.Text = "这段内容需要更新";
comment.PositionX = 100;
comment.PositionY = 100;

// 添加带日期的注释
DateTime commentDate = DateTime.Now;
IComment dateComment = slide.Comments.AddComment("李四", $"意见（{commentDate:yyyy-MM-dd}）");
dateComment.Text = "建议添加更多细节";

presentation.SaveToFile("WithComments.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 5: 获取注释信息

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 获取所有幻灯片的注释
foreach (ISlide slide in presentation.Slides)
{
    Console.WriteLine($"幻灯片 {slide.SlideNumber} 的注释:");

    foreach (IComment comment in slide.Comments)
    {
        Console.WriteLine($"  作者: {comment.Author}");
        Console.WriteLine($"  内容: {comment.Text}");
        Console.WriteLine($"  位置: ({comment.PositionX}, {comment.PositionY})");
    }
}

presentation.Dispose();
```

### 示例 6: 删除注释

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

ISlide slide = presentation.Slides[0];

// 删除特定注释
if (slide.Comments.Count > 0)
{
    slide.Comments.RemoveAt(0);
}

// 删除所有注释
slide.Comments.Clear();

presentation.SaveToFile("CommentsRemoved.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 7: 添加备注（演讲者备注）

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

ISlide slide = presentation.Slides[0];

// 设置备注
slide.NotesSlide.NotesTextFrame.Text = "演讲者备注内容：\n1. 首先介绍背景\n2. 然后说明核心问题\n3. 最后给出解决方案";

// 添加多段落备注
slide.NotesSlide.NotesTextFrame.Paragraphs[0].Text = "第一段备注";
TextParagraph para2 = new TextParagraph();
para2.Text = "第二段备注";
slide.NotesSlide.NotesTextFrame.Paragraphs.Append(para2);

presentation.SaveToFile("WithSpeakerNotes.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 8: 获取备注

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

ISlide slide = presentation.Slides[0];

if (slide.NotesSlide != null)
{
    Console.WriteLine("演讲者备注:");
    foreach (TextParagraph para in slide.NotesSlide.NotesTextFrame.Paragraphs)
    {
        Console.WriteLine(para.Text);
    }
}

presentation.Dispose();
```

### 示例 9: 删除备注

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 删除特定幻灯片的备注
presentation.Slides[0].NotesSlide = null;

// 删除所有幻灯片的备注
foreach (ISlide slide in presentation.Slides)
{
    slide.NotesSlide = null;
}

presentation.SaveToFile("SpeakerNotesRemoved.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 页眉页脚

### 示例 10: 设置页眉页脚

```csharp
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 获取母版幻灯片
IMasterSlide masterSlide = presentation.Masters[0];

// 设置日期页脚
masterSlide.DateTimeFormat = "yyyy年MM月dd日";
foreach (ISlide slide in presentation.Slides)
{
    slide.HeadersFooters.IsDateTimeVisible = true;
}

// 设置幻灯片编号
foreach (ISlide slide in presentation.Slides)
{
    slide.HeadersFooters.IsSlideNumberVisible = true;
}

// 设置页脚文本
foreach (ISlide slide in presentation.Slides)
{
    slide.HeadersFooters.IsFooterVisible = true;
    slide.HeadersFooters.FooterText = "公司机密文档";
}

presentation.SaveToFile("WithHeaderFooter.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 11: 设置备注母版页眉页脚

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 获取备注母版
INotesSlide notesMaster = presentation.NotesMaster;

if (notesMaster != null)
{
    // 设置日期
    notesMaster.HeadersFooters.IsDateTimeVisible = true;
    notesMaster.HeadersFooters.DateTimeFormat = "yyyy-MM-dd";

    // 设置页脚
    notesMaster.HeadersFooters.IsFooterVisible = true;
    notesMaster.HeadersFooters.FooterText = "内部使用";

    // 设置页眉
    notesMaster.HeadersFooters.IsHeaderVisible = true;
    notesMaster.HeadersFooters.HeaderText = "备注";
}

presentation.SaveToFile("WithNotesHeaderFooter.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 12: 自定义页眉页脚样式

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 自定义页脚样式
foreach (ISlide slide in presentation.Slides)
{
    if (slide.HeadersFooters.IsFooterVisible)
    {
        // 获取页脚形状并设置样式
        foreach (IShape shape in slide.Shapes)
        {
            if (shape.Name.Contains("Footer"))
            {
                IAutoShape footerShape = shape as IAutoShape;
                if (footerShape != null)
                {
                    footerShape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
                    footerShape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.Gray;
                    footerShape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 10;
                }
            }
        }
    }
}

presentation.SaveToFile("CustomHeaderFooter.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## OLE 对象

### 示例 13: 嵌入 Excel 作为 OLE 对象

```csharp
using System.Drawing;
using Spire.Presentation;
using System.IO;

Presentation presentation = new Presentation();

// 读取 Excel 文件数据
byte[] excelData = File.ReadAllBytes("data.xlsx");

// 嵌入 Excel OLE 对象
RectangleF rect = new RectangleF(50, 50, 500, 300);
IOleObject oleObject = presentation.Slides[0].Shapes.AppendOleObject(
    "Excel.Sheet.12",  // Excel 2007+ 的程序标识符
    excelData,
    rect
);

oleObject.ProgId = "Excel.Sheet.12";

presentation.SaveToFile("WithExcelOLE.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 14: 嵌入 ZIP 文件

```csharp
using System.Drawing;
using Spire.Presentation;
using System.IO;

Presentation presentation = new Presentation();

// 读取 ZIP 文件
byte[] zipData = File.ReadAllBytes("package.zip");

// 嵌入 ZIP OLE 对象
RectangleF rect = new RectangleF(50, 50, 200, 200);
IOleObject oleObject = presentation.Slides[0].Shapes.AppendOleObject(
    "Package",
    zipData,
    rect
);

// 显示为图标
oleObject.ObjectIcon = true;
oleObject.DisplayAsIcon = true;

presentation.SaveToFile("WithZipOLE.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 15: 提取 OLE 对象数据

```csharp
using Spire.Presentation;
using System.IO;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 提取所有 OLE 对象
int oleIndex = 0;
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IOleObject oleObject)
        {
            // 保存 OLE 对象数据
            string extension = GetFileExtension(oleObject.ProgId);
            File.WriteAllBytes($"extracted_ole_{oleIndex}{extension}", oleObject.Data);
            oleIndex++;
        }
    }
}

presentation.Dispose();

// 辅助方法：根据 ProgId 获取文件扩展名
string GetFileExtension(string progId)
{
    switch (progId)
    {
        case "Excel.Sheet.12":
        case "Excel.Sheet.8":
            return ".xlsx";
        case "Word.Document.12":
        case "Word.Document.8":
            return ".docx";
        case "PowerPoint.Show.12":
        case "PowerPoint.Show.8":
            return ".pptx";
        default:
            return ".bin";
    }
}
```

### 示例 16: 修改 OLE 对象

```csharp
using Spire.Presentation;
using System.IO;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 查找并修改 OLE 对象
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IOleObject oleObject)
        {
            // 替换 OLE 对象数据
            byte[] newData = File.ReadAllBytes("new_data.xlsx");
            oleObject.Data = newData;
        }
    }
}

presentation.SaveToFile("ModifiedOLE.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 其他高级功能

### 示例 17: 获取文档主题信息

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 获取主题信息
IMasterSlide master = presentation.Masters[0];

Console.WriteLine($"主题名称: {master.Theme.Name}");
Console.WriteLine($"主题颜色数量: {master.Theme.ColorScheme.Count}");

// 获取主题颜色
foreach (var color in master.Theme.ColorScheme)
{
    Console.WriteLine($"  {color.Key}: {color.Value}");
}

presentation.Dispose();
```

### 示例 18: 设置文档主题

```csharp
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

IMasterSlide master = presentation.Masters[0];

// 修改主题颜色
master.Theme.ColorScheme[SchemeColor.Accent1] = Color.Blue;
master.Theme.ColorScheme[SchemeColor.Accent2] = Color.Green;
master.Theme.ColorScheme[SchemeColor.Accent3] = Color.Orange;

// 修改主题字体
master.Theme.MinorFont.LatinFont = new TextFont("Arial");
master.Theme.MajorFont.LatinFont = new TextFont("Calibri");

presentation.SaveToFile("WithCustomTheme.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 19: 设置幻灯片背景

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

ISlide slide = presentation.Slides[0];

// 设置纯色背景
slide.Background.Type = BackgroundType.Custom;
slide.Background.FillFormat.FillType = FillFormatType.Solid;
slide.Background.FillFormat.SolidFillColor.Color = Color.LightBlue;

// 或设置渐变背景
slide.Background.Type = BackgroundType.Custom;
slide.Background.FillFormat.FillType = FillFormatType.Gradient;
slide.Background.FillFormat.Gradient.GradientStops.Append(0f, KnownColors.LightBlue);
slide.Background.FillFormat.Gradient.GradientStops.Append(1f, KnownColors.DarkBlue);

// 或设置图片背景
slide.Background.Type = BackgroundType.Custom;
slide.Background.FillFormat.FillType = FillFormatType.Picture;
slide.Background.FillFormat.Picture.Fill.PictureFillMode = PictureFillMode.Stretch;
slide.Background.FillFormat.Picture.Fill.Url = "background.jpg";

presentation.SaveToFile("WithCustomBackground.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 20: 获取幻灯片缩略图

```csharp
using System.Drawing.Imaging;
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 为每张幻灯片生成缩略图
for (int i = 0; i < presentation.Slides.Count; i++)
{
    ISlide slide = presentation.Slides[i];
    Bitmap thumbnail = slide.GetThumbnail(1.0f, 1.0f); // 缩放比例 1.0 = 原始大小
    thumbnail.Save($"slide_{i + 1}.png", ImageFormat.Png);
    thumbnail.Dispose();
}

presentation.Dispose();
```

## 注意事项

1. **水印位置**: 水印应放在内容下方（使用 SendToBack）
2. **注释兼容性**: 注释在不同 PowerPoint 版本中显示可能有所不同
3. **OLE 对象**: OLE 对象依赖外部程序，在某些环境中可能无法正常显示
4. **背景图片**: 大型背景图片会增加文件体积，建议优化图片

## 最佳实践

1. **水印设计**: 使用半透明水印，不要遮挡重要内容
2. **备注使用**: 将详细说明放在备注中，保持幻灯片简洁
3. **页脚信息**: 页脚应包含必要的标识信息（公司名称、日期等）
4. **OLE 管理**: OLE 对象会增加文件大小，谨慎使用

## 相关功能

- [文本处理](./03-text-content.md) - 备注文本格式化
- [形状处理](./04-shapes-images.md) - 水印形状设计
- [演讲者备注生成](./17-speaker-notes-generation.md) - 自动生成详细的演讲者备注
- [全局主题管理](./18-global-theme-manager.md) - 统一色调和字体方案
- [安全性](./13-security.md) - 文档保护和加密
