---
name: basic-load-save
description: Load existing presentation and save in different formats
---

# Load and Save Presentations

## Load and Modify

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();

// Load existing file
presentation.LoadFromFile("input.pptx");

// Modify content...
presentation.Slides[0].Shapes[0].TextFrame.Text = "Updated text";

// Save
presentation.SaveToFile("output.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## Convert to PDF

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// Convert to PDF
presentation.SaveToFile("output.pdf", FileFormat.PDF);
presentation.Dispose();
```

## Load Password Protected File

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();

// Load with password
presentation.LoadFromFile("protected.pptx", "password");

// Process...
presentation.SaveToFile("unlocked.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```
