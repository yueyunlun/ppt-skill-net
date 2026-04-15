---
name: slide-thumbnails
description: Generate thumbnails for all slides
---

# Generate Slide Thumbnails Example

```csharp
using Spire.Presentation;
using System.Drawing.Imaging;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// Generate thumbnail for each slide
for (int i = 0; i < presentation.Slides.Count; i++)
{
    ISlide slide = presentation.Slides[i];
    Bitmap thumbnail = slide.GetThumbnail(1.0f, 1.0f); // Original size
    thumbnail.Save($"slide_{i + 1}.png", ImageFormat.Png);
    thumbnail.Dispose();
}

presentation.Dispose();
```
