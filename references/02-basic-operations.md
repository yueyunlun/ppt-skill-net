---
title: 基础操作
category: spire-presentation
description: Spire.Presentation 基础操作：创建、保存、加载、幻灯片管理等
---

# 基础操作

## 概述

本章介绍 Spire.Presentation 的基础操作，包括：
- 创建和打开演示文稿
- 保存演示文稿
- 幻灯片管理（添加、删除、移动、克隆）
- 页面设置
- 文档属性

## 创建和打开演示文稿

### 示例 1: 创建新的演示文稿

```csharp
using Spire.Presentation;

// 创建新的空白演示文稿
Presentation presentation = new Presentation();

// 添加新幻灯片
presentation.Slides.Append();

// 保存文件
presentation.SaveToFile("output.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 2: 从模板创建演示文稿

```csharp
using Spire.Presentation;

// 创建新演示文稿
Presentation presentation = new Presentation();

// 加载模板文件
presentation.LoadFromFile("template.pptx");

// 修改内容...

// 保存为新的文件
presentation.SaveToFile("new_presentation.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 3: 打开现有演示文稿

```csharp
using Spire.Presentation;

// 打开 PPTX 文件
Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 打开旧版 PPT 文件
presentation.LoadFromFile("old.ppt", FileFormat.Ppt);

// 打开 ODP 文件
presentation.LoadFromFile("presentation.odp");

// 处理内容...

presentation.SaveToFile("output.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 4: 使用流加载演示文稿

```csharp
using Spire.Presentation;
using System.IO;

// 从文件流加载
using (FileStream stream = new FileStream("input.pptx", FileMode.Open))
{
    Presentation presentation = new Presentation();
    presentation.LoadFromStream(stream);

    // 处理内容...

    presentation.SaveToFile("output.pptx", FileFormat.Pptx2010);
    presentation.Dispose();
}
```

### 示例 5: 打开受密码保护的演示文稿

```csharp
using Spire.Presentation;

// 使用密码打开
Presentation presentation = new Presentation();
presentation.LoadFromFile("protected.pptx", "my_password");

// 处理内容...

presentation.SaveToFile("unlocked.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 保存演示文稿

### 示例 6: 保存为不同格式

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 保存为 PPTX 2010
presentation.SaveToFile("ppt2010.pptx", FileFormat.Pptx2010);

// 保存为 PPTX 2013
presentation.SaveToFile("ppt2013.pptx", FileFormat.Pptx2013);

// 保存为 PPTX 2016
presentation.SaveToFile("ppt2016.pptx", FileFormat.Pptx2016);

// 保存为 PDF
presentation.SaveToFile("document.pdf", FileFormat.PDF);

presentation.Dispose();
```

### 示例 7: 保存到流

```csharp
using Spire.Presentation;
using System.IO;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 保存到内存流
using (MemoryStream stream = new MemoryStream())
{
    presentation.SaveToStream(stream, FileFormat.Pptx2010);

    // 使用流（例如上传到网络）
    UploadToServer(stream);
}

presentation.Dispose();
```

### 示例 8: 检查并创建输出目录

```csharp
using Spire.Presentation;
using System.IO;

string outputPath = @"C:\Output\presentation.pptx";

// 确保输出目录存在
string outputDir = Path.GetDirectoryName(outputPath);
if (!Directory.Exists(outputDir))
{
    Directory.CreateDirectory(outputDir);
}

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");
presentation.SaveToFile(outputPath, FileFormat.Pptx2010);
presentation.Dispose();
```

## 幻灯片管理

### 示例 9: 添加幻灯片

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();

// 添加新幻灯片（使用默认布局）
presentation.Slides.Append();

// 添加指定数量的幻灯片
for (int i = 0; i < 5; i++)
{
    presentation.Slides.Append();
}

Console.WriteLine($"幻灯片总数: {presentation.Slides.Count}");

presentation.SaveToFile("with_slides.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 10: 删除幻灯片

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 删除指定索引的幻灯片（从0开始）
presentation.Slides.RemoveAt(1);  // 删除第二张幻灯片

// 删除最后一张幻灯片
presentation.Slides.RemoveAt(presentation.Slides.Count - 1);

// 删除所有幻灯片
while (presentation.Slides.Count > 0)
{
    presentation.Slides.RemoveAt(0);
}

presentation.SaveToFile("slides_removed.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 11: 克隆幻灯片

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 在同一演示文稿内克隆幻灯片
// 在演示文稿末尾添加克隆
presentation.Slides.AppendByTemplate(presentation.Slides[0]);

// 在指定位置插入克隆
presentation.Slides.Insert(1, presentation.Slides[0]);

presentation.SaveToFile("slides_cloned.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 12: 克隆幻灯片到另一个演示文稿

```csharp
using Spire.Presentation;

// 源演示文稿
Presentation source = new Presentation();
source.LoadFromFile("source.pptx");

// 目标演示文稿
Presentation target = new Presentation();

// 克隆幻灯片
for (int i = 0; i < source.Slides.Count; i++)
{
    target.Slides.AppendByTemplate(source.Slides[i]);
}

target.SaveToFile("merged.pptx", FileFormat.Pptx2010);
source.Dispose();
target.Dispose();
```

### 示例 13: 移动幻灯片

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 将第二张幻灯片移动到第五张的位置
ISlide slideToMove = presentation.Slides[1];
presentation.Slides.Insert(4, slideToMove);
presentation.Slides.RemoveAt(1);  // 删除原位置的幻灯片

// 或使用 MoveTo 方法（如果支持）
// presentation.Slides.MoveTo(1, 4);

presentation.SaveToFile("slides_moved.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 14: 更改幻灯片顺序

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 交换两张幻灯片的位置
ISlide slide1 = presentation.Slides[0];
ISlide slide2 = presentation.Slides[1];

// 重新插入实现交换
presentation.Slides.Insert(0, slide2);
presentation.Slides.RemoveAt(2);
presentation.Slides.Insert(1, slide1);
presentation.Slides.RemoveAt(2);

presentation.SaveToFile("slides_reordered.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 页面设置

### 示例 15: 设置幻灯片大小

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 设置为标准大小
presentation.SlideSize.Type = SlideSizeType.A4;

// 设置为屏幕 4:3
presentation.SlideSize.Type = SlideSizeType.Screen4x3;

// 设置为屏幕 16:9
presentation.SlideSize.Type = SlideSizeType.Screen16x9;

// 设置为屏幕 16:10
presentation.SlideSize.Type = SlideSizeType.Screen16x10;

// 设置为自定义大小
presentation.SlideSize.Type = SlideSizeType.Custom;
presentation.SlideSize.Size = new SizeF(960, 720);

presentation.SaveToFile("resized.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 16: 设置幻灯片方向

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 设置为横向
presentation.SlideSize.Orientation = SlideOrienation.Landscape;

// 设置为纵向
presentation.SlideSize.Orientation = SlideOrienation.Portrait;

presentation.SaveToFile("orientation_set.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 文档属性

### 示例 17: 设置内置属性

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 设置文档属性
presentation.DocumentProperty.Author = "张三";
presentation.DocumentProperty.Title = "项目汇报";
presentation.DocumentProperty.Subject = "2024年度总结";
presentation.DocumentProperty.Company = "ABC公司";
presentation.DocumentProperty.Keywords = "汇报, 总结, 2024";
presentation.DocumentProperty.Comments = "机密文档";
presentation.DocumentProperty.Category = "工作文档";

presentation.SaveToFile("with_properties.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 18: 获取文档属性

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 获取内置属性
Console.WriteLine($"作者: {presentation.DocumentProperty.Author}");
Console.WriteLine($"标题: {presentation.DocumentProperty.Title}");
Console.WriteLine($"主题: {presentation.DocumentProperty.Subject}");
Console.WriteLine($"公司: {presentation.DocumentProperty.Company}");
Console.WriteLine($"关键词: {presentation.DocumentProperty.Keywords}");
Console.WriteLine($"创建时间: {presentation.DocumentProperty.CreatedTime}");
Console.WriteLine($"修改时间: {presentation.DocumentProperty.ModifiedTime}");

presentation.Dispose();
```

### 示例 19: 设置自定义属性

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 设置自定义属性
presentation.CustomDocumentProperties.Add("部门", "技术部");
presentation.CustomDocumentProperties.Add("项目编号", "PRJ-2024-001");
presentation.CustomDocumentProperties.Add("审核人", "李四");

presentation.SaveToFile("with_custom_properties.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 20: 获取自定义属性

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 获取所有自定义属性
foreach (var property in presentation.CustomDocumentProperties)
{
    Console.WriteLine($"{property.Name}: {property.Value}");
}

// 获取特定自定义属性
var department = presentation.CustomDocumentProperties["部门"];
if (department != null)
{
    Console.WriteLine($"部门: {department.Value}");
}

presentation.Dispose();
```

## 节管理

### 示例 21: 添加节

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 添加节
presentation.SectionList.Add("第一章", presentation.Slides[0]);
presentation.SectionList.Add("第二章", presentation.Slides[2]);
presentation.SectionList.Add("第三章", presentation.Slides[4]);

presentation.SaveToFile("with_sections.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 22: 删除节

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 删除所有节
presentation.SectionList.Clear();

// 删除特定节
presentation.SectionList.RemoveAt(0);

presentation.SaveToFile("sections_removed.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 分割演示文稿

### 示例 23: 将演示文稿分割为单个幻灯片文件

```csharp
using Spire.Presentation;
using System.IO;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

string outputDir = "separated_slides";
Directory.CreateDirectory(outputDir);

// 将每张幻灯片保存为单独的文件
for (int i = 0; i < presentation.Slides.Count; i++)
{
    Presentation singleSlide = new Presentation();
    singleSlide.Slides.AppendByTemplate(presentation.Slides[i]);

    string outputFile = Path.Combine(outputDir, $"slide_{i + 1}.pptx");
    singleSlide.SaveToFile(outputFile, FileFormat.Pptx2010);
    singleSlide.Dispose();
}

presentation.Dispose();
```

## 合并演示文稿

### 示例 24: 合并多个演示文稿

```csharp
using Spire.Presentation;
using System.IO;

string[] files = { "part1.pptx", "part2.pptx", "part3.pptx" };
Presentation merged = new Presentation();

foreach (string file in files)
{
    using (Presentation temp = new Presentation())
    {
        temp.LoadFromFile(file);

        // 克隆所有幻灯片
        for (int i = 0; i < temp.Slides.Count; i++)
        {
            merged.Slides.AppendByTemplate(temp.Slides[i]);
        }
    }
}

merged.SaveToFile("merged.pptx", FileFormat.Pptx2010);
merged.Dispose();
```

## 设置演示文稿类型

### 示例 25: 设置为展台模式

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 设置为展台模式（自动循环播放）
presentation.ShowType = ShowShowType.Kiosk;
presentation.ShowLoop = true;

presentation.SaveToFile("kiosk_mode.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 文件格式参考

### FileFormat 枚举

| 格式 | 描述 |
|------|------|
| `PPT` - PowerPoint 97-2003 |
| `PPTX` - PowerPoint 2007+ |
| `PPTX2010` - PowerPoint 2010 |
| `PPTX2013` - PowerPoint 2013 |
| `PPTX2016` - PowerPoint 2016 |
| `PDF` - PDF 文档 |
| `SVG` - SVG 矢量图 |
| `HTML` - HTML 网页 |
| `XPS` - XPS 文档 |
| `ODP` - OpenDocument Presentation |
| `OFD` - OFD 文档 |

### SlideSizeType 枚举

| 大小 | 描述 |
|------|------|
| `A4` - A4 纸张 |
| `A3` - A3 纸张 |
| `Letter` - 信纸 |
| `Screen4x3` - 屏幕 4:3 |
| `Screen16x9` - 屏幕 16:9 |
| `Screen16x10` - 屏幕 16:10 |
| `B4ISO` - B4 ISO |
| `B5ISO` - B5 ISO |
| `Custom` - 自定义 |

## 注意事项

1. **资源管理**: 始终调用 `Dispose()` 或使用 `using` 语句
2. **文件路径**: 使用 `Path.Combine()` 确保跨平台兼容性
3. **异常处理**: 添加适当的异常处理机制
4. **密码保护**: 处理受密码保护的文档时提供正确的密码

## 最佳实践

1. **使用 using**: 始终使用 `using` 语句确保资源释放
2. **检查文件**: 操作前检查文件是否存在
3. **备份**: 修改重要文件前先备份
4. **验证输出**: 保存后验证输出文件

## 相关功能

- [文本处理](./03-text-content.md) - 添加文本内容
- [形状处理](./04-shapes-images.md) - 添加图形元素
- [转换](./11-conversion.md) - 格式转换
