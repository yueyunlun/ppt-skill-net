---
name: batch-convert
description: Batch convert multiple presentations to PDF
---

# Batch Convert to PDF Example

```csharp
using Spire.Presentation;
using System.IO;
using System.Linq;

string inputFolder = "input_ppt";
string outputFolder = "output_pdf";

Directory.CreateDirectory(outputFolder);

// Get all PPT files
string[] pptFiles = Directory.GetFiles(inputFolder, "*.pptx")
    .Concat(Directory.GetFiles(inputFolder, "*.ppt"))
    .ToArray();

foreach (string pptFile in pptFiles)
{
    string fileName = Path.GetFileNameWithoutExtension(pptFile);
    string outputPath = Path.Combine(outputFolder, fileName + ".pdf");

    using (Presentation presentation = new Presentation())
    {
        presentation.LoadFromFile(pptFile);
        presentation.SaveToFile(outputPath, FileFormat.PDF);
        Console.WriteLine($"Converted: {fileName}");
    }
}

Console.WriteLine($"Complete! Converted {pptFiles.Length} files");
```
