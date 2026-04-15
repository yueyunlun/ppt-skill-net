---
title: 表格处理
category: spire-presentation
description: 使用 Spire.Presentation 创建和编辑表格
---

# 表格处理

## 概述

Spire.Presentation 提供了完整的表格处理功能，包括：
- 创建表格
- 添加/删除行和列
- 单元格样式设置
- 合并/拆分单元格
- 表格样式应用

## 示例

### 示例 1: 创建基本表格

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 定义列宽和行高
double[] widths = { 100, 150, 100, 150 };
double[] heights = { 30, 30, 30, 30 };

// 创建表格（指定位置）
ITable table = presentation.Slides[0].Shapes.AppendTable(
    50,  // X 坐标
    50,  // Y 坐标
    widths,
    heights
);

// 添加数据
table[0, 0].TextFrame.Text = "姓名";
table[1, 0].TextFrame.Text = "年龄";
table[2, 0].TextFrame.Text = "城市";
table[3, 0].TextFrame.Text = "职业";

table[0, 1].TextFrame.Text = "张三";
table[1, 1].TextFrame.Text = "25";
table[2, 1].TextFrame.Text = "北京";
table[3, 1].TextFrame.Text = "工程师";

table[0, 2].TextFrame.Text = "李四";
table[1, 2].TextFrame.Text = "30";
table[2, 2].TextFrame.Text = "上海";
table[3, 2].TextFrame.Text = "设计师";

table[0, 3].TextFrame.Text = "王五";
table[1, 3].TextFrame.Text = "28";
table[2, 3].TextFrame.Text = "深圳";
table[3, 3].TextFrame.Text = "产品经理";

presentation.SaveToFile("BasicTable.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 2: 应用表格样式

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();

// 创建表格
double[] widths = { 100, 150, 100 };
double[] heights = { 30, 30, 30, 30 };
ITable table = presentation.Slides[0].Shapes.AppendTable(
    50, 50, widths, heights
);

// 添加数据...

// 应用预设样式
table.StylePreset = TableStylePreset.LightStyle1;
// 其他可选样式:
// TableStylePreset.LightStyle1Accent1
// TableStylePreset.LightStyle2
// TableStylePreset.MediumStyle1
// TableStylePreset.MediumStyle2
// TableStylePreset.DarkStyle1
// TableStylePreset.DarkStyle2

presentation.SaveToFile("StyledTable.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 3: 设置单元格样式

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();
// 创建表格...

// 设置表头样式
for (int col = 0; col < table.ColumnsCount; col++)
{
    table[col, 0].FillFormat.FillType = FillFormatType.Solid;
    table[col, 0].FillFormat.SolidFillColor.Color = Color.LightBlue;
    table[col, 0].TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
    table[col, 0].TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.White;
    table[col, 0].TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 14;
    table[col, 0].TextFrame.Paragraphs[0].TextRanges[0].IsBold = TriState.True;
    table[col, 0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
}

// 设置单元格边框
table[0, 0].BorderTop.FillType = FillFormatType.Solid;
table[0, 0].BorderTop.SolidFillColor.Color = Color.Black;
table[0, 0].BorderTop.Width = 1;

// 或设置整个表格的边框样式
for (int row = 0; row < table.RowsCount; row++)
{
    for (int col = 0; col < table.ColumnsCount; col++)
    {
        table[col, row].BorderLeft.FillType = FillFormatType.Solid;
        table[col, row].BorderLeft.SolidFillColor.Color = Color.Gray;
        table[col, row].BorderLeft.Width = 1;

        table[col, row].BorderRight.FillType = FillFormatType.Solid;
        table[col, row].BorderRight.SolidFillColor.Color = Color.Gray;
        table[col, row].BorderRight.Width = 1;

        table[col, row].BorderTop.FillType = FillFormatType.Solid;
        table[col, row].BorderTop.SolidFillColor.Color = Color.Gray;
        table[col, row].BorderTop.Width = 1;

        table[col, row].BorderBottom.FillType = FillFormatType.Solid;
        table[col, row].BorderBottom.SolidFillColor.Color = Color.Gray;
        table[col, row].BorderBottom.Width = 1;
    }
}
```

### 示例 4: 添加行

```csharp
Presentation presentation = new Presentation();
presentation.LoadFromFile("table.pptx");

ITable table = presentation.Slides[0].Shapes[0] as ITable;

// 在指定位置添加行
table.AppendRow(2);  // 在第2行后添加

// 或在表格末尾添加
table.AppendRow();

// 或使用克隆方式添加行
int newRow = table.AppendRow();
table[0, newRow].TextFrame.Text = "新数据1";
table[1, newRow].TextFrame.Text = "新数据2";

presentation.SaveToFile("TableRowAdded.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 5: 添加列

```csharp
Presentation presentation = new Presentation();
presentation.LoadFromFile("table.pptx");

ITable table = presentation.Slides[0].Shapes[0] as ITable;

// 在指定位置添加列
table.AppendColumn(2);  // 在第2列后添加

// 或在表格末尾添加
table.AppendColumn();

// 设置新列数据
for (int row = 0; row < table.RowsCount; row++)
{
    table[table.ColumnsCount - 1, row].TextFrame.Text = "新列数据";
}

presentation.SaveToFile("TableColumnAdded.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 6: 删除行

```csharp
Presentation presentation = new Presentation();
presentation.LoadFromFile("table.pptx");

ITable table = presentation.Slides[0].Shapes[0] as ITable;

// 删除指定行
table.RowsList.RemoveAt(1);  // 删除第2行（索引1）

// 或
table.RemoveRow(1);

presentation.SaveToFile("TableRowRemoved.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 7: 删除列

```csharp
Presentation presentation = new Presentation();
presentation.LoadFromFile("table.pptx");

ITable table = presentation.Slides[0].Shapes[0] as ITable;

// 删除指定列
table.ColumnsList.RemoveAt(1);  // 删除第2列（索引1）

// 或
table.RemoveColumn(1);

presentation.SaveToFile("TableColumnRemoved.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 8: 合并单元格

```csharp
Presentation presentation = new Presentation();
// 创建表格...

// 合并单元格：从 (0,0) 开始，跨越 2 列 2 行
table.MergeCells(table[0, 0], table[1, 1]);

// 合并整行
for (int col = 1; col < table.ColumnsCount - 1; col++)
{
    table.MergeCells(table[col, 0], table[col + 1, 0]);
}

// 合并整列
for (int row = 1; row < table.RowsCount - 1; row++)
{
    table.MergeCells(table[0, row], table[0, row + 1]);
}
```

### 示例 9: 拆分单元格

```csharp
// Spire.Presentation 中拆分单元格的有限支持
// 可以通过重新创建表格来实现复杂布局

// 如果需要拆分，通常的做法是：
// 1. 删除原表格
// 2. 创建新的表格结构
// 3. 重新填充数据
```

### 示例 10: 设置列宽和行高

```csharp
Presentation presentation = new Presentation();
presentation.LoadFromFile("table.pptx");

ITable table = presentation.Slides[0].Shapes[0] as ITable;

// 设置特定列宽
table.ColumnsList[0].Width = 120;
table.ColumnsList[1].Width = 180;

// 设置特定行高
table.RowsList[0].Height = 40;
table.RowsList[1].Height = 35;

// 设置所有列宽
for (int i = 0; i < table.ColumnsCount; i++)
{
    table.ColumnsList[i].Width = 100;
}

presentation.SaveToFile("TableResized.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 11: 单元格对齐

```csharp
using Spire.Presentation;

// 设置单元格文本对齐
table[0, 0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;      // 水平居中
table[0, 0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;         // 左对齐
table[0, 0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Right;        // 右对齐
table[0, 0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Justify;      // 两端对齐

// 设置单元格垂直对齐
table[0, 0].TextFrame.VerticalAlignment = VerticalAlignmentType.Top;       // 顶部
table[0, 0].TextFrame.VerticalAlignment = VerticalAlignmentType.Middle;    // 中部
table[0, 0].TextFrame.VerticalAlignment = VerticalAlignmentType.Bottom;    // 底部
```

### 示例 12: 设置单元格边距

```csharp
// 设置单元格内边距
table[0, 0].TextFrame.MarginTop = 5;
table[0, 0].TextFrame.MarginBottom = 5;
table[0, 0].TextFrame.MarginLeft = 5;
table[0, 0].TextFrame.MarginRight = 5;

// 设置缩进
table[0, 0].TextFrame.Paragraphs[0].Indent = 10;
```

### 示例 13: 在单元格中插入图片

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();
// 创建表格...

// 在单元格中插入图片
RectangleF imageRect = new RectangleF(10, 10, 80, 60);
IEmbedImage image = table[0, 0].Shapes.AppendEmbedImage(
    ShapeType.Rectangle,
    "logo.png",
    imageRect
);

image.Line.FillFormat.FillType = FillFormatType.None;
```

### 示例 14: 设置单元格背景色

```csharp
using Spire.Presentation.Drawing;

// 设置单元格背景
table[0, 0].FillFormat.FillType = FillFormatType.Solid;
table[0, 0].FillFormat.SolidFillColor.Color = Color.LightBlue;

// 设置渐变背景
table[0, 0].FillFormat.FillType = FillFormatType.Gradient;
table[0, 0].FillFormat.Gradient.GradientStops.Append(0f, KnownColors.LightBlue);
table[0, 0].FillFormat.Gradient.GradientStops.Append(1f, KnownColors.DarkBlue);
```

### 示例 15: 识别合并的单元格

```csharp
Presentation presentation = new Presentation();
presentation.LoadFromFile("table.pptx");

ITable table = presentation.Slides[0].Shapes[0] as ITable;

// 检查单元格是否合并
if (table[0, 0].FirstRow > -1)
{
    Console.WriteLine($"单元格(0,0)已合并");
    Console.WriteLine($"合并起始: ({table[0, 0].FirstColumn}, {table[0, 0].FirstRow})");
}

// 遍历所有单元格
for (int row = 0; row < table.RowsCount; row++)
{
    for (int col = 0; col < table.ColumnsCount; col++)
    {
        if (table[col, row].FirstRow == -1)
        {
            // 这是一个未合并的单元格或合并区域的起点
            Console.WriteLine($"单元格({col},{row}): {table[col, row].TextFrame.Text}");
        }
    }
}
```

### 示例 16: 克隆行和列

```csharp
Presentation presentation = new Presentation();
presentation.LoadFromFile("table.pptx");

ITable table = presentation.Slides[0].Shapes[0] as ITable;

// 克隆行（创建新行并复制数据）
int newRow = table.AppendRow();
for (int col = 0; col < table.ColumnsCount; col++)
{
    table[col, newRow].TextFrame.Text = table[col, 1].TextFrame.Text;
    table[col, newRow].FillFormat.FillType = table[col, 1].FillFormat.FillType;
    table[col, newRow].FillFormat.SolidFillColor.Color = table[col, 1].FillFormat.SolidFillColor.Color;
}

// 克隆列（创建新列并复制数据）
int newCol = table.AppendColumn();
for (int row = 0; row < table.RowsCount; row++)
{
    table[newCol, row].TextFrame.Text = table[1, row].TextFrame.Text;
}

presentation.SaveToFile("TableCloned.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 17: 锁定表格宽高比

```csharp
// 锁定表格的宽高比
table.LockAspectRatio = true;
```

### 示例 18: 设置表格位置和大小

```csharp
// 设置表格位置
table.Frame.X = 50;
table.Frame.Y = 50;

// 设置表格大小
table.Frame.Width = 600;
table.Frame.Height = 400;
```

### 示例 19: 编辑表格数据

```csharp
Presentation presentation = new Presentation();
presentation.LoadFromFile("table.pptx");

ITable table = presentation.Slides[0].Shapes[0] as ITable;

// 修改单元格数据
table[0, 0].TextFrame.Text = "新的标题";
table[1, 1].TextFrame.Text = "更新的数据";

// 批量修改
for (int row = 0; row < table.RowsCount; row++)
{
    for (int col = 0; col < table.ColumnsCount; col++)
    {
        string text = table[col, row].TextFrame.Text;
        table[col, row].TextFrame.Text = text.ToUpper();
    }
}

presentation.SaveToFile("TableDataUpdated.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 20: 获取表格边框颜色

```csharp
Presentation presentation = new Presentation();
presentation.LoadFromFile("table.pptx");

ITable table = presentation.Slides[0].Shapes[0] as ITable;

// 获取单元格边框颜色
Color topColor = table[0, 0].BorderTop.SolidFillColor.Color;
Color bottomColor = table[0, 0].BorderBottom.SolidFillColor.Color;
Color leftColor = table[0, 0].BorderLeft.SolidFillColor.Color;
Color rightColor = table[0, 0].BorderRight.SolidFillColor.Color;

Console.WriteLine($"上边框颜色: {topColor}");
Console.WriteLine($"下边框颜色: {bottomColor}");
Console.WriteLine($"左边框颜色: {leftColor}");
Console.WriteLine($"右边框颜色: {rightColor}");

presentation.Dispose();
```

## 表格样式参考

### TableStylePreset

| 样式 | 描述 |
|------|------|
| `LightStyle1` - 浅色样式1 |
| `LightStyle1Accent1` - 浅色样式1 强调色1 |
| `LightStyle2` - 浅色样式2 |
| `LightStyle2Accent1` - 浅色样式2 强调色1 |
| `MediumStyle1` - 中等样式1 |
| `MediumStyle1Accent1` - 中等样式1 强调色1 |
| `MediumStyle2` - 中等样式2 |
| `DarkStyle1` - 深色样式1 |
| `DarkStyle1Accent1` - 深色样式1 强调色1 |
| `NoStyle` - 无样式 |

### TextAlignmentType

| 对齐方式 | 描述 |
|----------|------|
| `Left` | 左对齐 |
| `Center` | 居中对齐 |
| `Right` | 右对齐 |
| `Justify` | 两端对齐 |

### VerticalAlignmentType

| 对齐方式 | 描述 |
|----------|------|
| `Top` | 顶部对齐 |
| `Middle` | 中部对齐 |
| `Bottom` | 底部对齐 |

## ITable 主要属性

| 属性 | 类型 | 描述 |
|------|------|------|
| `ColumnsCount` | int | 列数 |
| `RowsCount` | int | 行数 |
| `ColumnsList` | ColumnList | 列集合 |
| `RowsList` | RowList | 行集合 |
| `StylePreset` | TableStylePreset | 表格样式 |
| `Frame` | RectangleF | 表格位置和大小 |
| `LockAspectRatio` | bool | 是否锁定宽高比 |

## 注意事项

1. **单元格索引**: 单元格索引从 0 开始，`table[columnIndex, rowIndex]`
2. **合并单元格**: 合并后，原位置单元格会被覆盖，需要正确处理索引
3. **样式继承**: 修改单元格样式可能会覆盖表格样式
4. **性能**: 大型表格可能影响性能，建议合理控制表格大小

## 最佳实践

1. **合理设计**: 根据内容需求设计合适的表格结构
2. **样式一致**: 在整个演示文稿中使用一致的表格样式
3. **数据验证**: 在填充表格前验证数据的有效性
4. **资源管理**: 及时释放大型表格占用的资源

## 相关功能

- [文本处理](./03-text-content.md) - 单元格文本格式化
- [形状处理](./04-shapes-images.md) - 表格中的图片和形状
- [图表](./06-charts.md) - 从表格数据创建图表
