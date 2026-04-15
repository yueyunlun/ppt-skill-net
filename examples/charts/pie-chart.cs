---
name: chart-pie
description: Create a pie chart in PowerPoint
---

# Pie Chart Example

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();

// Create pie chart
RectangleF rect = new RectangleF(40, 100, 550, 320);
IChart chart = presentation.Slides[0].Shapes.AppendChart(
    ChartType.Pie,
    rect,
    false
);

// Set title
chart.ChartTitle.TextProperties.Text = "Sales by Quarter";
chart.HasTitle = true;

// Add data
string[] quarters = { "Q1", "Q2", "Q3", "Q4" };
int[] sales = { 210, 320, 180, 500 };

chart.ChartData[0, 0].Text = "Quarter";
chart.ChartData[0, 1].Text = "Sales";

for (int i = 0; i < quarters.Length; i++)
{
    chart.ChartData[i + 1, 0].Value = quarters[i];
    chart.ChartData[i + 1, 1].Value = sales[i];
}

chart.Series.SeriesLabel = chart.ChartData["B1", "B1"];
chart.Categories.CategoryLabels = chart.ChartData["A2", "A5"];
chart.Series[0].Values = chart.ChartData["B2", "B5"];

// Add colors
Color[] colors = { Color.RosyBrown, Color.LightBlue, Color.LightPink, Color.MediumPurple };
for (int i = 0; i < chart.Series[0].Values.Count; i++)
{
    ChartDataPoint cdp = new ChartDataPoint(chart.Series[0]);
    cdp.Index = i;
    chart.Series[0].DataPoints.Add(cdp);
    chart.Series[0].DataPoints[i].Fill.FillType = FillFormatType.Solid;
    chart.Series[0].DataPoints[i].Fill.SolidColor.Color = colors[i];
}

// Show labels
chart.Series[0].DataLabels.LabelValueVisible = true;
chart.Series[0].DataLabels.PercentValueVisible = true;

presentation.SaveToFile("pie-chart.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```
