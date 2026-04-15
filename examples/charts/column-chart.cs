---
name: chart-column
description: Create a clustered column chart
---

# Column Chart Example

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();

// Create column chart
RectangleF rect = new RectangleF(50, 50, 500, 400);
IChart chart = presentation.Slides[0].Shapes.AppendChart(
    ChartType.ClusteredColumn,
    rect,
    false
);

// Set title
chart.ChartTitle.TextProperties.Text = "Annual Sales";
chart.HasTitle = true;

// Add data
chart.ChartData[0, 0].Text = "Product";
chart.ChartData[0, 1].Text = "2023";
chart.ChartData[0, 2].Text = "2024";

string[] products = { "A", "B", "C", "D" };
for (int i = 0; i < products.Length; i++)
{
    chart.ChartData[i + 1, 0].Value = products[i];
    chart.ChartData[i + 1, 1].Value = 100 + i * 50;
    chart.ChartData[i + 1, 2].Value = 150 + i * 60;
}

chart.Series.SeriesLabel = chart.ChartData["B1", "C1"];
chart.Categories.CategoryLabels = chart.ChartData["A2", "A5"];
chart.Series[0].Values = chart.ChartData["B2", "B5"];
chart.Series[1].Values = chart.ChartData["C2", "C5"];

presentation.SaveToFile("column-chart.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```
