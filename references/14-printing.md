---
title: 打印功能
category: spire-presentation
description: 使用 Spire.Presentation 打印演示文稿
---

# 打印功能

## 概述

Spire.Presentation 提供了完整的打印功能，包括：
- 直接打印到默认打印机
- 打印到指定打印机
- 打印指定范围的幻灯片
- 多张幻灯片打印在一页
- 自定义打印设置

## 示例

### 示例 1: 打印到默认打印机

```csharp
using Spire.Presentation;
using System.Drawing.Printing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrintDocument
PrintDocument printDoc = new PrintDocument();
printDoc.PrinterSettings.PrinterName = ""; // 空字符串表示使用默认打印机

// 打印
presentation.Print(printDoc);

presentation.Dispose();
```

### 示例 2: 打印到指定打印机

```csharp
using Spire.Presentation;
using System.Drawing.Printing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrintDocument 并指定打印机
PrintDocument printDoc = new PrintDocument();
printDoc.PrinterSettings.PrinterName = "Microsoft Print to PDF";

// 打印
presentation.Print(printDoc);

presentation.Dispose();
```

### 示例 3: 静默打印到默认打印机

```csharp
using Spire.Presentation;
using System.Drawing.Printing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrintDocument
PrintDocument printDoc = new PrintDocument();

// 静默打印（不显示打印对话框）
printDoc.PrintController = new StandardPrintController();

// 打印
presentation.Print(printDoc);

presentation.Dispose();
```

### 示例 4: 打印指定范围的幻灯片

```csharp
using Spire.Presentation;
using System.Drawing.Printing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrintDocument
PrintDocument printDoc = new PrintDocument();

// 设置打印范围（打印第2-4张幻灯片）
printDoc.PrinterSettings.PrintRange = PrintRange.SomePages;
printDoc.PrinterSettings.FromPage = 2;
printDoc.PrinterSettings.ToPage = 4;

// 打印
presentation.Print(printDoc);

presentation.Dispose();
```

### 示例 5: 多张幻灯片打印在一页

```csharp
using Spire.Presentation;
using System.Drawing.Printing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrintDocument
PrintDocument printDoc = new PrintDocument();

// 设置每页打印的幻灯片数量（例如每页2张）
// 这通常需要在打印设置中配置

// 打印
presentation.Print(printDoc);

presentation.Dispose();
```

### 示例 6: 设置打印份数

```csharp
using Spire.Presentation;
using System.Drawing.Printing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrintDocument
PrintDocument printDoc = new PrintDocument();

// 设置打印份数
printDoc.PrinterSettings.Copies = 2;

// 设置是否逐份打印
printDoc.PrinterSettings.Collate = true;

// 打印
presentation.Print(printDoc);

presentation.Dispose();
```

### 示例 7: 使用 PrinterSettings 设置打印选项

```csharp
using Spire.Presentation;
using System.Drawing.Printing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrinterSettings
PrinterSettings printerSettings = new PrinterSettings();

// 设置打印机
printerSettings.PrinterName = "Microsoft Print to PDF";

// 设置打印份数
printerSettings.Copies = 1;

// 设置打印范围
printerSettings.PrintRange = PrintRange.AllPages;

// 创建 PrintDocument
PrintDocument printDoc = new PrintDocument();
printDoc.PrinterSettings = printerSettings;

// 打印
presentation.Print(printDoc);

presentation.Dispose();
```

### 示例 8: 预览打印

```csharp
using Spire.Presentation;
using System.Windows.Forms;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建打印预览对话框
PrintPreviewDialog previewDialog = new PrintPreviewDialog();

// 创建 PrintDocument
System.Drawing.Printing.PrintDocument printDoc = new System.Drawing.Printing.PrintDocument();

// 设置打印页面事件
printDoc.PrintPage += (sender, e) =>
{
    // 获取要打印的幻灯片
    // 注意：这需要更复杂的实现来处理多页

    // 生成幻灯片图像
    Bitmap slideImage = presentation.Slides[0].GetThumbnail(1.0f, 1.0f);

    // 绘制到打印页面
    e.Graphics.DrawImage(slideImage, 0, 0, e.PageBounds.Width, e.PageBounds.Height);

    slideImage.Dispose();
};

// 设置预览文档
previewDialog.Document = printDoc;

// 显示预览
previewDialog.ShowDialog();

presentation.Dispose();
```

### 示例 9: 设置打印方向

```csharp
using Spire.Presentation;
using System.Drawing.Printing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrintDocument
PrintDocument printDoc = new PrintDocument();

// 设置打印方向为横向
printDoc.DefaultPageSettings.Landscape = true;

// 打印
presentation.Print(printDoc);

presentation.Dispose();
```

### 示例 10: 设置打印页面大小

```csharp
using Spire.Presentation;
using System.Drawing.Printing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrintDocument
PrintDocument printDoc = new PrintDocument();

// 设置页面大小为 A4
printDoc.DefaultPageSettings.PaperSize = new PaperSize("A4", 827, 1169);

// 打印
presentation.Print(printDoc);

presentation.Dispose();
```

### 示例 11: 打印前显示打印对话框

```csharp
using Spire.Presentation;
using System.Windows.Forms;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrintDialog
PrintDialog printDialog = new PrintDialog();

// 创建 PrintDocument
System.Drawing.Printing.PrintDocument printDoc = new System.Drawing.Printing.PrintDocument();

// 设置文档
printDialog.Document = printDoc;

// 显示打印对话框
if (printDialog.ShowDialog() == DialogResult.OK)
{
    // 用户点击了打印按钮
    presentation.Print(printDoc);
}

presentation.Dispose();
```

### 示例 12: 使用虚拟打印机

```csharp
using Spire.Presentation;
using System.Drawing.Printing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrintDocument
PrintDocument printDoc = new PrintDocument();

// 设置虚拟打印机（例如 Microsoft Print to PDF）
printDoc.PrinterSettings.PrinterName = "Microsoft Print to PDF";

// 设置输出文件（某些虚拟打印机支持）
printDoc.PrinterSettings.PrintFileName = "output.pdf";

// 打印
presentation.Print(printDoc);

presentation.Dispose();
```

### 示例 13: 双面打印

```csharp
using Spire.Presentation;
using System.Drawing.Printing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrintDocument
PrintDocument printDoc = new PrintDocument();

// 设置双面打印
printDoc.PrinterSettings.Duplex = Duplex.Vertical;

// 打印
presentation.Print(printDoc);

presentation.Dispose();
```

### 示例 14: 打印彩色或黑白

```csharp
using Spire.Presentation;
using System.Drawing.Printing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrintDocument
PrintDocument printDoc = new PrintDocument();

// 设置为黑白打印
printDoc.DefaultPageSettings.Color = false;

// 或设置为彩色打印
// printDoc.DefaultPageSettings.Color = true;

// 打印
presentation.Print(printDoc);

presentation.Dispose();
```

### 示例 15: 打印质量设置

```csharp
using Spire.Presentation;
using System.Drawing.Printing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 创建 PrintDocument
PrintDocument printDoc = new PrintDocument();

// 设置打印质量
printDoc.DefaultPageSettings.PrinterResolution = new PrinterResolution
{
    Kind = PrinterResolutionKind.High
};

// 打印
presentation.Print(printDoc);

presentation.Dispose();
```

## 打印设置类

### PrinterSettings 常用属性

| 属性 | 描述 |
|------|------|
| `PrinterName` | 打印机名称 |
| `Copies` | 打印份数 |
| `Collate` | 是否逐份打印 |
| `PrintRange` | 打印范围 |
| `FromPage` | 起始页 |
| `ToPage` | 结束页 |
| `Duplex` | 双面打印设置 |

### PageSettings 常用属性

| 属性 | 描述 |
|------|------|
| `PaperSize` | 纸张大小 |
| `Landscape` | 是否横向 |
| `Margins` | 页边距 |
| `Color` | 是否彩色打印 |
| `PrinterResolution` | 打印分辨率 |

### Duplex 枚举

| 值 | 描述 |
|----|------|
| `Default` - 默认 |
| `Simplex` - 单面 |
| `Vertical` - 纵向双面 |
| `Horizontal` - 横向双面 |

### PrinterResolutionKind 枚举

| 值 | 描述 |
|----|------|
| `Draft` - 草稿质量 |
| `Low` - 低质量 |
| `Medium` - 中等质量 |
| `High` - 高质量 |

## 注意事项

1. **打印机名称**: 确保打印机名称正确，可以通过 `PrinterSettings.InstalledPrinters` 获取可用打印机列表
2. **权限**: 确保应用程序有打印权限
3. **纸张大小**: 不同打印机支持不同的纸张大小
4. **打印队列**: 大量打印任务可能需要在打印队列中等待

## 最佳实践

1. **预览先打印**: 使用打印预览确认输出效果
2. **测试打印**: 在正式打印前进行测试打印
3. **错误处理**: 添加适当的错误处理以应对打印机不可用等情况
4. **资源管理**: 及时释放打印相关资源

## 获取可用打印机

```csharp
using System.Drawing.Printing;

// 获取所有安装的打印机
foreach (string printer in PrinterSettings.InstalledPrinters)
{
    Console.WriteLine($"打印机: {printer}");

    // 获取打印机详情
    PrinterSettings settings = new PrinterSettings();
    settings.PrinterName = printer;

    Console.WriteLine($"  默认打印机: {settings.IsDefaultPrinter}");
    Console.WriteLine($"  支持彩色: {settings.SupportsColor}");
    Console.WriteLine($"  支持双面: {settings.CanDuplex}");
}
```

## 相关功能

- [转换](./11-conversion.md) - 转换为 PDF 后打印
- [基础操作](./02-basic-operations.md) - 幻灯片管理
