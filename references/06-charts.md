---
title: 图表处理
category: spire-presentation
description: 使用 Spire.Presentation 创建和编辑各种类型的图表
---

# 图表处理

## 概述

Spire.Presentation 支持创建 20+ 种图表类型，包括柱状图、饼图、折线图、散点图等，并提供丰富的图表格式化选项。

## 支持的图表类型

### 基础图表
- `ChartType.ClusteredColumn` - 聚合柱状图
- `ChartType.Pie` - 饼图
- `ChartType.Line` - 折线图
- `ChartType.Scatter` - 散点图

### 高级图表
- `ChartType.Doughnut` - 环形图
- `ChartType.Bubble` - 气泡图
- `ChartType.StackedBar100Percent` - 100% 堆积柱状图
- `ChartType.Cylinder3DClustered` - 3D 圆柱图
- `ChartType.Funnel` - 漏斗图
- `ChartType.Histogram` - 直方图
- `ChartType.Map` - 地图
- `ChartType.Sunburst` - 旭日图
- `ChartType.TreeMap` - 树状图
- `ChartType.WaterFall` - 瀑布图
- `ChartType.BoxAndWhisker` - 箱线图
- `ChartType.Pareto` - 帕累托图

## 示例

### 示例 1: 创建饼图

```csharp
using System;
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Drawing;

// 创建演示文稿
Presentation presentation = new Presentation();

// 插入饼图
RectangleF rect = new RectangleF(40, 100, 550, 320);
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.Pie, rect, false);

// 设置图表标题
chart.ChartTitle.TextProperties.Text = "销售分布";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

// 设置图表数据
string[] categories = { "第一季度", "第二季度", "第三季度", "第四季度" };
int[] values = { 210, 320, 180, 500 };

chart.ChartData[0, 0].Text = "季度";
chart.ChartData[0, 1].Text = "销售额";

for (int i = 0; i < categories.Length; i++)
{
    chart.ChartData[i + 1, 0].Value = categories[i];
    chart.ChartData[i + 1, 1].Value = values[i];
}

// 设置标签和数据
chart.Series.SeriesLabel = chart.ChartData["B1", "B1"];
chart.Categories.CategoryLabels = chart.ChartData["A2", "A5"];
chart.Series[0].Values = chart.ChartData["B2", "B5"];

// 为每个数据点设置不同颜色
Color[] colors = { Color.RosyBrown, Color.LightBlue, Color.LightPink, Color.MediumPurple };
for (int i = 0; i < chart.Series[0].Values.Count; i++)
{
    ChartDataPoint cdp = new ChartDataPoint(chart.Series[0]);
    cdp.Index = i;
    chart.Series[0].DataPoints.Add(cdp);
    chart.Series[0].DataPoints[i].Fill.FillType = FillFormatType.Solid;
    chart.Series[0].DataPoints[i].Fill.SolidColor.Color = colors[i];
}

// 显示数据标签和百分比
chart.Series[0].DataLabels.LabelValueVisible = true;
chart.Series[0].DataLabels.PercentValueVisible = true;

presentation.SaveToFile("PieChart.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 2: 创建聚合柱状图

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Charts;

Presentation presentation = new Presentation();

// 创建柱状图
RectangleF rect = new RectangleF(50, 50, 500, 400);
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.ClusteredColumn, rect, false);

// 设置数据
chart.ChartData[0, 0].Text = "产品";
chart.ChartData[0, 1].Text = "2023";
chart.ChartData[0, 2].Text = "2024";

string[] products = { "产品A", "产品B", "产品C", "产品D" };
int[][] values = {
    new int[] { 100, 120 },
    new int[] { 200, 180 },
    new int[] { 150, 200 },
    new int[] { 80, 110 }
};

for (int i = 0; i < products.Length; i++)
{
    chart.ChartData[i + 1, 0].Value = products[i];
    chart.ChartData[i + 1, 1].Value = values[i][0];
    chart.ChartData[i + 1, 2].Value = values[i][1];
}

// 设置系列
chart.Series.SeriesLabel = chart.ChartData["B1", "C1"];
chart.Categories.CategoryLabels = chart.ChartData["A2", "A5"];
chart.Series[0].Values = chart.ChartData["B2", "B5"];
chart.Series[1].Values = chart.ChartData["C2", "C5"];

presentation.SaveToFile("ColumnChart.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 3: 创建折线图

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Charts;

Presentation presentation = new Presentation();

// 创建折线图
RectangleF rect = new RectangleF(50, 50, 600, 400);
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.Line, rect, false);

// 设置数据和标签
chart.ChartData[0, 0].Text = "月份";
chart.ChartData[0, 1].Text = "销售额";

string[] months = { "1月", "2月", "3月", "4月", "5月", "6月" };
double[] sales = { 5000, 5500, 4800, 6200, 5800, 7000 };

for (int i = 0; i < months.Length; i++)
{
    chart.ChartData[i + 1, 0].Value = months[i];
    chart.ChartData[i + 1, 1].Value = sales[i];
}

chart.Series.SeriesLabel = chart.ChartData["B1", "B1"];
chart.Categories.CategoryLabels = chart.ChartData["A2", "A7"];
chart.Series[0].Values = chart.ChartData["B2", "B7"];

// 设置线条样式
chart.Series[0].Format.Line.FillFormat.FillType = FillFormatType.Solid;
chart.Series[0].Format.Line.FillFormat.SolidColor.Color = Color.Blue;
chart.Series[0].Format.Line.Weight = 2.5f;

// 显示数据标记
chart.Series[0].Marker = true;
chart.Series[0].MarkerStyle = ChartMarkerStyleType.Circle;
chart.Series[0].MarkerSize = 8;
chart.Series[0].MarkerFillColor.Color = Color.Blue;

presentation.SaveToFile("LineChart.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 4: 创建组合图表

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Charts;

Presentation presentation = new Presentation();

// 创建组合图表（柱状图 + 折线图）
RectangleF rect = new RectangleF(50, 50, 600, 400);
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.Combination, rect, true);

// 设置数据
chart.ChartData[0, 0].Text = "月份";
chart.ChartData[0, 1].Text = "销售量";
chart.ChartData[0, 2].Text = "销售额";

for (int i = 1; i <= 6; i++)
{
    chart.ChartData[i, 0].Value = i + "月";
    chart.ChartData[i, 1].Value = 100 * i;
    chart.ChartData[i, 2].Value = 500 * i;
}

// 设置系列
chart.Series.SeriesLabel = chart.ChartData["B1", "C1"];
chart.Categories.CategoryLabels = chart.ChartData["A2", "A7"];
chart.Series[0].Values = chart.ChartData["B2", "B7"];
chart.Series[1].Values = chart.ChartData["C2", "C7"];

// 设置第一个系列为柱状图
chart.Series[0].Type = ChartType.ClusteredColumn;
// 设置第二个系列为折线图，并使用次坐标轴
chart.Series[1].Type = ChartType.Line;
chart.Series[1].UseSecondaryAxis = true;

presentation.SaveToFile("CombinationChart.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 5: 添加趋势线

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Charts;

Presentation presentation = new Presentation();

// 创建图表
RectangleF rect = new RectangleF(50, 50, 600, 400);
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.Scatter, rect, false);

// 设置数据
// ...（数据设置代码同上）

// 添加趋势线
chart.Series[0].TrendLines.Add(TrendLineType.Polynomial);
chart.Series[0].TrendLines[0].DisplayEquation = true;
chart.Series[0].TrendLines[0].DisplayRSquaredValue = true;

presentation.SaveToFile("TrendLine.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 6: 设置图表样式

```csharp
// ... 创建图表代码

// 设置图表区域样式
chart.ChartArea.Fill.FillType = FillFormatType.Solid;
chart.ChartArea.Fill.SolidColor.Color = Color.WhiteSmoke;

// 设置图例样式
chart.HasLegend = true;
chart.Legend.Position = LegendPositionType.Bottom;
chart.Legend.TextProperties.Paragraphs[0].TextRanges[0].FontHeight = 12;

// 设置坐标轴样式
chart.PrimaryCategoryAxis.HasMajorGridLines = true;
chart.PrimaryCategoryAxis.MajorGridLines.FillFormat.FillType = FillFormatType.Solid;
chart.PrimaryCategoryAxis.MajorGridLines.FillFormat.SolidColor.Color = Color.LightGray;

chart.PrimaryValueAxis.HasMajorGridLines = true;
chart.PrimaryValueAxis.MajorGridLines.FillFormat.FillType = FillFormatType.Solid;
chart.PrimaryValueAxis.MajorGridLines.FillFormat.SolidColor.Color = Color.LightGray;

// 设置数据标签
chart.Series[0].DataLabels.LabelValueVisible = true;
chart.Series[0].DataLabels.Position = LegendDataLabelPositionType.Center;
chart.Series[0].DataLabels.TextProperties.Paragraphs[0].TextRanges[0].FontHeight = 10;
```

### 示例 7: 创建 3D 图表

```csharp
// 创建 3D 柱状图
RectangleF rect = new RectangleF(50, 50, 600, 400);
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.Cylinder3DClustered, rect, false);

// 设置 3D 视角
chart.Rotation3D.XRotation = 15;
chart.Rotation3D.YRotation = 20;
chart.Rotation3D.Perspective = 30;
chart.Rotation3D.RightAngleAxes = false;

// 设置 3D 深度
chart.Depth = 200;
```

### 示例 8: 添加误差线

```csharp
// 添加标准误差线
chart.Series[0].ErrorBars.HasErrorBars = true;
chart.Series[0].ErrorBars.ErrorBarType = ErrorBarType.StandardError;
chart.Series[0].ErrorBars.ErrorBarValueType = ErrorBarValueType.Percentage;
chart.Series[0].ErrorBars.Value = 10;

// 设置误差线样式
chart.Series[0].ErrorBars.LineWidth = 1.5f;
chart.Series[0].ErrorBars.LineColor.Color = Color.Red;
```

### 示例 9: 编辑图表数据

```csharp
// 加载包含图表的演示文稿
Presentation presentation = new Presentation();
presentation.LoadFromFile("chart.pptx");

// 获取第一个幻灯片上的第一个图表
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

// 修改图表数据
chart.ChartData[2, 1].Value = 150;
chart.ChartData[3, 1].Value = 200;

// 修改系列名称
chart.Series[0].SeriesLabel = chart.ChartData["B1", "B1"];

presentation.SaveToFile("updated_chart.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 图表组件

### 主要属性

| 属性 | 描述 |
|------|------|
| `ChartTitle` | 图表标题 |
| `HasLegend` | 是否显示图例 |
| `HasTitle` | 是否有标题 |
| `ChartData` | 图表数据 |
| `Series` | 数据系列集合 |
| `Categories` | 分类标签 |
| `PrimaryValueAxis` | 主值坐标轴 |
| `PrimaryCategoryAxis` | 主分类坐标轴 |
| `SecondaryValueAxis` | 次值坐标轴 |

### 常用方法

| 方法 | 描述 |
|------|------|
| `AppendChart()` | 添加图表 |
| `GetChartData()` | 获取图表数据 |
| `ClearChartData()` | 清除图表数据 |

## 注意事项

1. **数据范围**: 确保数据范围正确设置，否则图表可能显示不完整
2. **性能**: 大量数据点可能影响性能，建议合理控制数据量
3. **兼容性**: 某些高级图表类型可能在旧版本 PowerPoint 中显示异常
4. **格式化**: 建议在设置完所有数据后再进行格式化

## 相关功能

- [文本处理](./03-text-content.md) - 图表标签文本格式化
- [形状处理](./04-shapes-images.md) - 图表作为形状处理
- [表格处理](./05-tables.md) - 从表格数据创建图表
