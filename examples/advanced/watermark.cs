---
name: add-watermark
description: Add text watermark to all slides
---

# Add Watermark Example

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// Add watermark to each slide
foreach (ISlide slide in presentation.Slides)
{
    RectangleF rect = new RectangleF(
        presentation.SlideSize.Size.Width / 2 - 200,
        presentation.SlideSize.Size.Height / 2 - 50,
        400,
        100
    );

    IAutoShape watermark = slide.Shapes.AppendShape(
        ShapeType.Rectangle,
        rect
    );
    watermark.Rotation = -45;
    watermark.Fill.FillType = FillFormatType.Solid;
    watermark.Fill.SolidColor.Color = Color.FromArgb(30, Color.Gray);
    watermark.ShapeStyle.LineColor.Color = Color.Transparent;
    watermark.AppendTextFrame("机密文档");

    watermark.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial");
    watermark.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 36;
    watermark.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
    watermark.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.White;
    watermark.TextFrame.Paragraphs[0].TextRanges[0].IsBold = TriState.True;
    watermark.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;

    watermark.ZOrder(ShapeZOrderType.SendToBack);
}

presentation.SaveToFile("watermarked.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```
