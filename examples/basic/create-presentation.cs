---
name: basic-create-ppt
description: Create a new PowerPoint presentation with slides
---

# Basic Presentation Creation

## Create New Presentation

```csharp
using Spire.Presentation;

// Create new presentation
Presentation presentation = new Presentation();

// Add slides
presentation.Slides.Append();
presentation.Slides.Append();

// Save file
presentation.SaveToFile("output.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## Add Text to Slide

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();

// Add shape with text
RectangleF rect = new RectangleF(50, 50, 400, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.White;
shape.AppendTextFrame("Hello, World!");

presentation.SaveToFile("with_text.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```
