---
name: table-create
description: Create a table with data
---

# Create Table Example

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// Define dimensions
double[] widths = { 100, 150, 100, 150 };
double[] heights = { 30, 30, 30, 30 };

// Create table
ITable table = presentation.Slides[0].Shapes.AppendTable(
    50, 50,  // X, Y position
    widths,
    heights
);

// Add data
string[,] data = {
    { "Name", "Age", "City", "Occupation" },
    { "张三", "25", "北京", "工程师" },
    { "李四", "30", "上海", "设计师" },
    { "王五", "28", "深圳", "产品经理" }
};

for (int row = 0; row < 4; row++)
{
    for (int col = 0; col < 4; col++)
    {
        table[col, row].TextFrame.Text = data[row, col];
        table[col, row].TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 14;
    }
}

// Apply style
table.StylePreset = TableStylePreset.LightStyle1Accent1;

presentation.SaveToFile("table.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```
