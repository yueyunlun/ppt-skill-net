---
name: spire-presentation
description: This skill should be used when the user asks to "create a PowerPoint presentation", "edit a PPTX file", "convert PowerPoint to PDF", "add charts to slides", "convert text to flowchart", "visualize text as diagram", or mentions Spire.Presentation, PowerPoint automation, .NET presentation processing, or slide manipulation.
version: 0.1.0
---

# Spire.Presentation for .NET

## Overview

Spire.Presentation for .NET is a professional PowerPoint-compatible component that enables developers to create, read, write, modify, convert, and print PowerPoint documents without requiring Microsoft PowerPoint installation.

## When to Use

This skill activates when the user's request involves:
- Creating or editing PowerPoint presentations (PPTX, PPT)
- Converting presentations to other formats (PDF, HTML, SVG, images)
- Adding content like charts, tables, shapes, or SmartArt to slides
- Converting text to flowcharts, diagrams, or graphics
- Visualizing text as SmartArt or structured diagrams
- Working with animations, multimedia, or hyperlinks
- Generating speaker notes for presentations
- Adding detailed演讲者备注 to slides
- Applying global themes and color schemes to presentations
- Changing PPT colors to tech blue with sans-serif fonts
- Unifying presentation styles and themes
- Creating PPT from templates and outlines
- Generating presentation based on AI-created structure
- Translating PowerPoint files to different languages
- Using AI to translate PPT content
- Auto-translating presentation text with AI services
- Automating PowerPoint workflows in .NET applications
- Processing large batches of presentation files

## Quick Reference

### File Operations

```csharp
// Create new presentation
Presentation presentation = new Presentation();
presentation.Slides.Append();
presentation.SaveToFile("output.pptx", FileFormat.Pptx2010);

// Load existing file
presentation.LoadFromFile("input.pptx");

// Convert to PDF
presentation.SaveToFile("output.pdf", FileFormat.PDF);
```

### Common Operations

| Task | Reference |
|------|-----------|
| Installation and setup | [Getting Started](./references/01-getting-started.md) |
| Creating slides | [Basic Operations](./references/02-basic-operations.md) |
| Adding text | [Text Content](./references/03-text-content.md) |
| Working with shapes | [Shapes & Images](./references/04-shapes-images.md) |
| Creating tables | [Tables](./references/05-tables.md) |
| Adding charts | [Charts](./references/06-charts.md) |
| Using SmartArt | [SmartArt](./references/07-smartart.md) |
| Adding multimedia | [Audio & Video](./references/08-multimedia.md) |
| Animations | [Animations](./references/09-animations.md) |
| Hyperlinks | [Hyperlinks](./references/10-hyperlinks.md) |
| Format conversion | [Conversion](./references/11-conversion.md) |
| Advanced features | [Advanced Features](./references/12-advanced-features.md) |
| Security | [Security](./references/13-security.md) |
| Printing | [Printing](./references/14-printing.md) |
| Text to Graphic | [Text to Graphic](./references/16-text-to-graphic.md) |
| Speaker Notes Generation | [Speaker Notes Generation](./references/17-speaker-notes-generation.md) |
| Global Theme Manager | [Global Theme Manager](./references/18-global-theme-manager.md) |
| AI PPT Creation | [AI PPT Creation](./references/19-ai-ppt-creation.md) |
| PPT Translation | [PPT Translation](./references/20-ppt-translation.md) |
| Best practices | [Best Practices](./references/15-best-practices.md) |

## Supported Formats

**Input**: PPTX, PPT, PPS, PPSX, ODP, DPS, DPT
**Output**: PPTX, PPT, PDF, SVG, HTML, XPS, TIFF, PNG, JPEG, GIF, BMP, OFD

## Code Examples

See the [examples](./examples/) directory for complete, runnable code samples organized by functionality.

## Key Concepts

- Always use `using` statements or call `Dispose()` to release resources
- Slide indexes start at 0
- Text formatting applies at the TextRange level
- Charts support 20+ chart types
- Animations apply to shapes and can be sequenced

## Requirements

- .NET Framework 4.0+ or .NET Core 5+
- NuGet package: `Spire.Presentation`
- License required for production use (evaluation watermark applies without license)

## Common Patterns

### Creating from Template
```csharp
using (Presentation presentation = new Presentation())
{
    presentation.LoadFromFile("template.pptx");
    // Modify content...
    presentation.SaveToFile("output.pptx", FileFormat.Pptx2010);
}
```

### Batch Processing
```csharp
string[] files = Directory.GetFiles(folder, "*.pptx");
foreach (string file in files)
{
    using (Presentation presentation = new Presentation())
    {
        presentation.LoadFromFile(file);
        // Process...
    }
}
```

### Error Handling
```csharp
try
{
    presentation.LoadFromFile("input.pptx");
}
catch (FileNotFoundException ex)
{
    Console.WriteLine($"File not found: {ex.Message}");
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
}
```

## Additional Resources

- [API Documentation](https://www.e-iceblue.com/Introduce/presentation-for-net.html)
- [NuGet Package](https://www.nuget.org/packages/Spire.Presentation/)
