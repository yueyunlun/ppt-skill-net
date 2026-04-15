---
title: 格式转换
category: spire-presentation
description: 使用 Spire.Presentation 将演示文稿转换为各种格式
---

# 格式转换

## 概述

Spire.Presentation 提供了强大的格式转换功能，可以将演示文稿转换为：
- PDF 文档
- SVG 矢量图
- HTML 网页
- TIFF 图像
- XPS 文档
- 图片格式（PNG, JPEG, GIF, BMP）
- OFD 文档
- ODP 格式

## 转换为 PDF

### 示例 1: 将整个演示文稿转换为 PDF

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 转换整个演示文稿为 PDF
presentation.SaveToFile("output.pdf", FileFormat.PDF);

presentation.Dispose();
```

### 示例 2: 转换特定幻灯片为 PDF

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 获取第二张幻灯片
ISlide slide = presentation.Slides[1];

// 仅转换该幻灯片为 PDF
slide.SaveToFile("slide2.pdf", FileFormat.PDF);

presentation.Dispose();
```

### 示例 3: 转换指定范围的幻灯片为 PDF

```csharp
using Spire.Presentation;
using System.Linq;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 选择要转换的幻灯片（第2到第4张）
var selectedSlides = presentation.Slides.Skip(1).Take(3).ToList();

// 创建新演示文稿
Presentation newPresentation = new Presentation();

// 复制选定的幻灯片
foreach (ISlide slide in selectedSlides)
{
    newPresentation.Slides.AppendByTemplate(slide);
}

// 保存为 PDF
newPresentation.SaveToFile("slides_2-4.pdf", FileFormat.PDF);
newPresentation.Dispose();
presentation.Dispose();
```

### 示例 4: 设置 PDF 转换选项

```csharp
using Spire.Presentation;
using Spire.Pdf;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 保存为 PDF 并获取 PdfDocument
presentation.SaveToFile("output.pdf", FileFormat.PDF);

// 如果需要进一步处理 PDF，可以使用 Spire.Pdf
// presentation.SaveToFile("output.pdf", FileFormat.PDF, "password");

presentation.Dispose();
```

## 转换为图片

### 示例 5: 将幻灯片转换为 PNG 图片

```csharp
using Spire.Presentation;
using System.Drawing.Imaging;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 为每张幻灯片生成 PNG 图片
for (int i = 0; i < presentation.Slides.Count; i++)
{
    ISlide slide = presentation.Slides[i];
    Bitmap image = slide.GetThumbnail(1.0f, 1.0f); // 原始大小

    image.Save($"slide_{i + 1}.png", ImageFormat.Png);
    image.Dispose();
}

presentation.Dispose();
```

### 示例 6: 将演示文稿转换为单个 TIFF 文件

```csharp
using Spire.Presentation;
using System.Drawing.Imaging;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 创建 EncoderParameters 指定 TIFF 格式
EncoderParameters encoderParams = new EncoderParameters(2);
encoderParams.Param[0] = new EncoderParameter(
    System.Drawing.Imaging.Encoder.SaveFlag,
    (long)EncoderValue.MultiFrame);
encoderParams.Param[1] = new EncoderParameter(
    System.Drawing.Imaging.Encoder.Compression,
    (long)EncoderValue.CompressionLZW);

// 获取第一张幻灯片作为基础
Bitmap multiFrame = presentation.Slides[0].GetThumbnail(1.0f, 1.0f);

// 添加其他幻灯片
for (int i = 1; i < presentation.Slides.Count; i++)
{
    Bitmap frame = presentation.Slides[i].GetThumbnail(1.0f, 1.0f);
    multiFrame.SaveAdd(frame, encoderParams);
    frame.Dispose();
}

// 保存 TIFF
multiFrame.Save("output.tiff", GetTiffCodecInfo(), encoderParams);
multiFrame.Dispose();
presentation.Dispose();

// 获取 TIFF 编码器
ImageCodecInfo GetTiffCodecInfo()
{
    ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
    foreach (ImageCodecInfo codec in codecs)
    {
        if (codec.FormatID == ImageFormat.Tiff.Guid)
            return codec;
    }
    return null;
}
```

### 示例 7: 指定缩放比例生成图片

```csharp
using Spire.Presentation;
using System.Drawing.Imaging;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 生成缩略图（50%大小）
float scaleX = 0.5f;
float scaleY = 0.5f;

for (int i = 0; i < presentation.Slides.Count; i++)
{
    ISlide slide = presentation.Slides[i];
    Bitmap thumbnail = slide.GetThumbnail(scaleX, scaleY);
    thumbnail.Save($"thumbnail_{i + 1}.png", ImageFormat.Png);
    thumbnail.Dispose();
}

presentation.Dispose();
```

### 示例 8: 转换为 JPEG

```csharp
using Spire.Presentation;
using System.Drawing.Imaging;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

for (int i = 0; i < presentation.Slides.Count; i++)
{
    ISlide slide = presentation.Slides[i];
    Bitmap image = slide.GetThumbnail(1.0f, 1.0f);
    image.Save($"slide_{i + 1}.jpg", ImageFormat.Jpeg);
    image.Dispose();
}

presentation.Dispose();
```

## 转换为 SVG

### 示例 9: 将幻灯片转换为 SVG

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 转换所有幻灯片为 SVG
for (int i = 0; i < presentation.Slides.Count; i++)
{
    ISlide slide = presentation.Slides[i];
    slide.SaveToFile($"slide_{i + 1}.svg", FileFormat.SVG);
}

presentation.Dispose();
```

### 示例 10: 转换特定幻灯片为 SVG

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 只转换第一张幻灯片
presentation.Slides[0].SaveToFile("first_slide.svg", FileFormat.SVG);

presentation.Dispose();
```

## 转换为 HTML

### 示例 11: 将幻灯片转换为 HTML

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 转换所有幻灯片为 HTML
for (int i = 0; i < presentation.Slides.Count; i++)
{
    ISlide slide = presentation.Slides[i];
    slide.SaveToFile($"slide_{i + 1}.html", FileFormat.HTML);
}

presentation.Dispose();
```

### 示例 12: 转换整个演示文稿为单个 HTML 文件

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 保存为 HTML（可能生成多个文件）
presentation.SaveToFile("output.html", FileFormat.HTML);

presentation.Dispose();
```

## 转换为 XPS

### 示例 13: 转换为 XPS

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 转换为 XPS
presentation.SaveToFile("output.xps", FileFormat.XPS);

presentation.Dispose();
```

## 转换为 OFD

### 示例 14: 转换为 OFD

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 转换为 OFD
presentation.SaveToFile("output.ofd", FileFormat.OFD);

presentation.Dispose();
```

## 转换为 ODP

### 示例 15: 转换为 ODP

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 转换为 ODP (OpenDocument Presentation)
presentation.SaveToFile("output.odp", FileFormat.ODP);

presentation.Dispose();
```

### 示例 16: 从 ODP 转换为 PPTX

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();

// 加载 ODP 文件
presentation.LoadFromFile("input.odp");

// 保存为 PPTX
presentation.SaveToFile("output.pptx", FileFormat.Pptx2010);

presentation.Dispose();
```

## PPT 格式转换

### 示例 17: PPT 转换为 PPTX

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();

// 加载旧版 PPT 文件
presentation.LoadFromFile("old.ppt", FileFormat.Ppt);

// 保存为 PPTX
presentation.SaveToFile("new.pptx", FileFormat.Pptx2010);

presentation.Dispose();
```

### 示例 18: 转换为特定版本的 PPTX

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 保存为不同版本的 PPTX
presentation.SaveToFile("ppt2010.pptx", FileFormat.Pptx2010);
presentation.SaveToFile("ppt2013.pptx", FileFormat.Pptx2013);
presentation.SaveToFile("ppt2016.pptx", FileFormat.Pptx2016);

presentation.Dispose();
```

## 批量转换

### 示例 19: 批量转换文件夹中的 PPT 文件为 PDF

```csharp
using Spire.Presentation;
using System.IO;
using System.Linq;

string inputFolder = "input_ppt";
string outputFolder = "output_pdf";

// 确保输出文件夹存在
Directory.CreateDirectory(outputFolder);

// 获取所有 PPT 文件
string[] pptFiles = Directory.GetFiles(inputFolder, "*.pptx")
    .Concat(Directory.GetFiles(inputFolder, "*.ppt"))
    .ToArray();

foreach (string pptFile in pptFiles)
{
    string fileName = Path.GetFileNameWithoutExtension(pptFile);
    string outputPath = Path.Combine(outputFolder, fileName + ".pdf");

    using (Presentation presentation = new Presentation())
    {
        presentation.LoadFromFile(pptFile);
        presentation.SaveToFile(outputPath, FileFormat.PDF);
        Console.WriteLine($"已转换: {fileName}");
    }
}

Console.WriteLine($"完成！共转换 {pptFiles.Length} 个文件");
```

### 示例 20: 批量生成幻灯片预览图

```csharp
using Spire.Presentation;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

string inputFolder = "presentations";
string outputFolder = "previews";
int thumbnailSize = 300; // 缩略图宽度

Directory.CreateDirectory(outputFolder);

string[] pptFiles = Directory.GetFiles(inputFolder, "*.pptx")
    .Concat(Directory.GetFiles(inputFolder, "*.ppt"))
    .ToArray();

foreach (string pptFile in pptFiles)
{
    string fileName = Path.GetFileNameWithoutExtension(pptFile);
    string fileOutputFolder = Path.Combine(outputFolder, fileName);
    Directory.CreateDirectory(fileOutputFolder);

    using (Presentation presentation = new Presentation())
    {
        presentation.LoadFromFile(pptFile);

        // 计算缩放比例
        float scaleX = (float)thumbnailSize / presentation.SlideSize.Size.Width;
        float scaleY = scaleX;

        // 生成每张幻灯片的缩略图
        for (int i = 0; i < presentation.Slides.Count; i++)
        {
            ISlide slide = presentation.Slides[i];
            Bitmap thumbnail = slide.GetThumbnail(scaleX, scaleY);
            thumbnail.Save(Path.Combine(fileOutputFolder, $"slide_{i + 1}.png"), ImageFormat.Png);
            thumbnail.Dispose();
        }

        Console.WriteLine($"已生成预览: {fileName} ({presentation.Slides.Count} 张幻灯片)");
    }
}
```

## 转换选项

### 示例 21: 保留备注转换 PDF

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 保存为 PDF 时保留备注
// 注意：具体选项取决于版本支持
presentation.SaveToFile("with_notes.pdf", FileFormat.PDF);

presentation.Dispose();
```

### 示例 22: 设置输出质量

```csharp
using Spire.Presentation;
using System.Drawing.Imaging;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 获取高质量图片（2倍大小）
for (int i = 0; i < presentation.Slides.Count; i++)
{
    ISlide slide = presentation.Slides[i];
    Bitmap image = slide.GetThumbnail(2.0f, 2.0f);
    image.Save($"slide_{i + 1}_high_res.png", ImageFormat.Png);
    image.Dispose();
}

presentation.Dispose();
```

### 示例 23: 自定义输出大小

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 设置输出幻灯片大小
presentation.SlideSize.Type = SlideSizeType.A4;
presentation.SlideSize.Orientation = SlideOrienation.Landscape;

// 保存为 PDF
presentation.SaveToFile("a4_landscape.pdf", FileFormat.PDF);

presentation.Dispose();
```

## 转换模板

### 示例 24: 使用模板批量生成文档

```csharp
using Spire.Presentation;

// 加载模板
Presentation template = new Presentation();
template.LoadFromFile("template.pptx");

// 修改模板内容
template.Slides[0].Shapes[0].TextFrame.Text = "报告 1";
template.SaveToFile("report1.pptx", FileFormat.Pptx2010);

// 重置模板
template.Slides[0].Shapes[0].TextFrame.Text = "报告 2";
template.SaveToFile("report2.pptx", FileFormat.Pptx2010);

// 转换为 PDF
template.LoadFromFile("report1.pptx");
template.SaveToFile("report1.pdf", FileFormat.PDF);

template.Dispose();
```

## 注意事项

1. **格式支持**: 某些高级特性在转换时可能丢失
2. **文件大小**: PDF 和图片转换可能产生较大的文件
3. **字体问题**: 确保系统中有演示文稿使用的字体
4. **性能**: 大型文件的转换可能需要较长时间

## 最佳实践

1. **批量处理**: 使用批量转换提高效率
2. **测试输出**: 在生产环境前测试转换结果
3. **优化图片**: 使用适当的缩放比例平衡质量和大小
4. **错误处理**: 添加适当的错误处理机制

## 支持的转换格式

| 源格式 | 目标格式 |
|--------|----------|
| PPTX | PDF, SVG, HTML, XPS, TIFF, PNG, JPEG, ODP, OFD |
| PPT | PDF, SVG, HTML, XPS, TIFF, PNG, JPEG, PPTX |
| ODP | PDF, SVG, HTML, XPS, TIFF, PNG, JPEG, PPTX |
| DPS/DPT | PDF, PPTX |

## 相关功能

- [基础操作](./02-basic-operations.md) - 文件加载和保存
- [打印](./14-printing.md) - 打印演示文稿
- [安全性](./13-security.md) - 加密文档的转换
