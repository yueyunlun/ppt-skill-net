---
title: PPT翻译功能
category: spire-presentation
description: 使用AI服务自动翻译PPT中的文本内容，保持格式和结构完整
---

# PPT翻译功能

## 概述

PPT翻译功能可以自动提取演示文稿中的所有文本内容，使用AI服务进行翻译，然后将翻译结果替换回PPT中，同时保持原始的格式、结构和样式。

该功能支持：
- **多种元素类型** - 文本框、表格、SmartArt等
- **结构保持** - 完整保持PPT的原始结构和格式
- **智能识别** - 自动识别源语言和目标语言
- **批量处理** - 支持批量翻译多个PPT文件
- **灵活配置** - 支持不同的AI翻译服务和模型

## 翻译流程

PPT翻译的完整流程包括以下步骤：

```
原始PPT → 文本提取 → AI翻译 → 文本替换 → 翻译后的PPT
```

1. **文本提取** - 从PPT中提取所有文本内容，生成结构化数据
2. **AI翻译** - 调用AI服务对提取的文本进行翻译
3. **文本替换** - 将翻译后的文本按照结构替换回PPT
4. **保存输出** - 保存翻译后的PPT文件

## 文本提取

### 提取数据结构

```csharp
using System;
using System.Collections.Generic;

// PPT文本提取结果
public class PptTextExtraction
{
    public List<SlideTextData> Slides { get; set; }

    public PptTextExtraction()
    {
        Slides = new List<SlideTextData>();
    }
}

// 幻灯片文本数据
public class SlideTextData
{
    public int SlideIndex { get; set; }
    public List<ShapeTextData> Shapes { get; set; }

    public SlideTextData()
    {
        Shapes = new List<ShapeTextData>();
    }
}

// 形状文本数据
public class ShapeTextData
{
    public int ShapeIndex { get; set; }
    public List<int> ShapeIndices { get; set; }  // 用于GroupShape的嵌套索引
    public TextContent TextContent { get; set; }
}

// 文本内容
public class TextContent
{
    public List<ParagraphData> Paragraphs { get; set; }
    public List<CellData> TableCells { get; set; }
    public List<SmartArtNodeData> SmartArtNodes { get; set; }

    public TextContent()
    {
        Paragraphs = new List<ParagraphData>();
        TableCells = new List<CellData>();
        SmartArtNodes = new List<SmartArtNodeData>();
    }
}

// 段落数据
public class ParagraphData
{
    public int ParagraphIndex { get; set; }
    public string Text { get; set; }
}

// 表格单元格数据
public class CellData
{
    public int Row { get; set; }
    public int Column { get; set; }
    public string Text { get; set; }
}

// SmartArt节点数据
public class SmartArtNodeData
{
    public int NodeIndex { get; set; }
    public string Text { get; set; }
}
```

### 文本提取类

```csharp
using System;
using System.Collections.Generic;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using Spire.Presentation.Tables;

public class PptTextExtractor
{
    // 提取PPT中所有文本内容
    public static PptTextExtraction ExtractAllText(string pptxPath)
    {
        Presentation presentation = new Presentation();
        PptTextExtraction extraction = new PptTextExtraction();
        
        try
        {
            presentation.LoadFromFile(pptxPath);
            
            Console.WriteLine($"开始提取文本，共 {presentation.Slides.Count} 张幻灯片");
            
            // 遍历所有幻灯片
            for (int slideIdx = 0; slideIdx < presentation.Slides.Count; slideIdx++)
            {
                ISlide slide = presentation.Slides[slideIdx];
                SlideTextData slideData = new SlideTextData
                {
                    SlideIndex = slideIdx
                };
                
                Console.WriteLine($"  处理第 {slideIdx + 1} 张幻灯片，共 {slide.Shapes.Count} 个形状");
                
                // 遍历幻灯片中的所有形状
                for (int shapeIdx = 0; shapeIdx < slide.Shapes.Count; shapeIdx++)
                {
                    IShape shape = slide.Shapes[shapeIdx];
                    ShapeTextData shapeData = ExtractShapeText(shape, shapeIdx);
                    
                    // 只添加包含文本的形状
                    if (shapeData.TextContent != null)
                    {
                        slideData.Shapes.Add(shapeData);
                    }
                }
                
                extraction.Slides.Add(slideData);
            }
            
            Console.WriteLine($"文本提取完成，共提取了 {extraction.Slides.Count} 张幻灯片的文本");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"提取文本时发生错误: {ex.Message}");
            throw;
        }
        finally
        {
            presentation.Dispose();
        }
        
        return extraction;
    }
    
    // 提取形状中的文本
    private static ShapeTextData ExtractShapeText(IShape shape, int shapeIndex)
    {
        ShapeTextData shapeData = new ShapeTextData
        {
            ShapeIndex = shapeIndex,
            TextContent = null
        };
        
        // 处理IAutoShape
        if (shape is IAutoShape autoShape)
        {
            TextContent textContent = ExtractAutoShapeText(autoShape);
            if (textContent != null && textContent.Paragraphs.Count > 0)
            {
                shapeData.TextContent = textContent;
            }
        }
        // 处理ITable
        else if (shape is ITable table)
        {
            TextContent textContent = ExtractTableText(table);
            if (textContent != null && textContent.TableCells.Count > 0)
            {
                shapeData.TextContent = textContent;
            }
        }
        // 处理ISmartArt
        else if (shape is ISmartArt smartArt)
        {
            TextContent textContent = ExtractSmartArtText(smartArt);
            if (textContent != null && textContent.SmartArtNodes.Count > 0)
            {
                shapeData.TextContent = textContent;
            }
        }
        // 处理GroupShape
        else if (shape is GroupShape groupShape)
        {
            List<ShapeTextData> groupShapes = new List<ShapeTextData>();
            List<int> indices = new List<int>();
            ExtractGroupShapes(groupShape, groupShapes, indices, shapeIndex);
            
            // 如果组内有包含文本的形状，保存第一个有文本的形状信息
            foreach (var groupShapeData in groupShapes)
            {
                if (groupShapeData.TextContent != null)
                {
                    shapeData.ShapeIndices = groupShapeData.ShapeIndices;
                    shapeData.TextContent = groupShapeData.TextContent;
                    break;
                }
            }
        }
        
        return shapeData;
    }
    
    // 提取AutoShape中的文本
    private static TextContent ExtractAutoShapeText(IAutoShape autoShape)
    {
        if (autoShape.TextFrame == null || autoShape.TextFrame.Paragraphs.Count == 0)
        {
            return null;
        }
        
        TextContent textContent = new TextContent
        {
            Paragraphs = new List<ParagraphData>()
        };
        
        for (int paraIdx = 0; paraIdx < autoShape.TextFrame.Paragraphs.Count; paraIdx++)
        {
            TextParagraph paragraph = autoShape.TextFrame.Paragraphs[paraIdx];
            string text = paragraph.Text.Trim();
            
            if (!string.IsNullOrEmpty(text))
            {
                textContent.Paragraphs.Add(new ParagraphData
                {
                    ParagraphIndex = paraIdx,
                    Text = text
                });
            }
        }
        
        return textContent.Paragraphs.Count > 0 ? textContent : null;
    }
    
    // 提取Table中的文本
    private static TextContent ExtractTableText(ITable table)
    {
        TextContent textContent = new TextContent
        {
            TableCells = new List<CellData>()
        };
        
        try
        {
            for (int row = 0; row < table.TableRows.Count; row++)
            {
                for (int col = 0; col < table.ColumnsList.Count; col++)
                {
                    ICell cell = table[col, row];
                    string text = cell.TextFrame.Text.Trim();
                    
                    if (!string.IsNullOrEmpty(text))
                    {
                        textContent.TableCells.Add(new CellData
                        {
                            Row = row,
                            Column = col,
                            Text = text
                        });
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"提取表格文本时出错: {ex.Message}");
        }
        
        return textContent.TableCells.Count > 0 ? textContent : null;
    }
    
    // 提取SmartArt中的文本
    private static TextContent ExtractSmartArtText(ISmartArt smartArt)
    {
        TextContent textContent = new TextContent
        {
            SmartArtNodes = new List<SmartArtNodeData>()
        };
        
        int nodeIndex = 0;
        ExtractSmartArtNodes(smartArt.Nodes, textContent.SmartArtNodes, ref nodeIndex);
        
        return textContent.SmartArtNodes.Count > 0 ? textContent : null;
    }
    
    // 递归提取SmartArt节点
    private static void ExtractSmartArtNodes(ISmartArtNodeCollection nodes, 
                                            List<SmartArtNodeData> nodeList, 
                                            ref int index)
    {
        foreach (ISmartArtNode node in nodes)
        {
            if (node.TextFrame != null)
            {
                string text = node.TextFrame.Text.Trim();
                if (!string.IsNullOrEmpty(text))
                {
                    nodeList.Add(new SmartArtNodeData
                    {
                        NodeIndex = index++,
                        Text = text
                    });
                }
            }
            
            // 递归处理子节点
            if (node.ChildNodes.Count > 0)
            {
                ExtractSmartArtNodes(node.ChildNodes, nodeList, ref index);
            }
        }
    }
    
    // 提取GroupShape中的形状
    private static void ExtractGroupShapes(GroupShape groupShape, 
                                          List<ShapeTextData> shapes,
                                          List<int> indices,
                                          int currentIndex)
    {
        for (int subIdx = 0; subIdx < groupShape.Shapes.Count; subIdx++)
        {
            indices.Add(subIdx);
            ShapeTextData shapeData = ExtractShapeText(groupShape.Shapes[subIdx], subIdx);
            
            if (shapeData.TextContent != null)
            {
                shapeData.ShapeIndices = new List<int>(indices);
                shapes.Add(shapeData);
            }
            
            // 递归处理嵌套的GroupShape
            if (groupShape.Shapes[subIdx] is GroupShape subGroupShape)
            {
                ExtractGroupShapes(subGroupShape, shapes, indices, currentIndex);
            }
            
            indices.RemoveAt(indices.Count - 1);
        }
    }
    
    // 将提取结果保存为JSON
    public static void SaveExtractionAsJson(PptTextExtraction extraction, string jsonPath)
    {
        string json = System.Text.Json.JsonSerializer.Serialize(extraction, new System.Text.Json.JsonSerializerOptions
        {
            WriteIndented = true
        });
        System.IO.File.WriteAllText(jsonPath, json);
        Console.WriteLine($"提取结果已保存到: {jsonPath}");
    }
    
    // 从JSON加载提取结果
    public static PptTextExtraction LoadExtractionFromJson(string jsonPath)
    {
        string json = System.IO.File.ReadAllText(jsonPath);
        return System.Text.Json.JsonSerializer.Deserialize<PptTextExtraction>(json);
    }
}
```

## AI翻译集成

### AI翻译服务接口

```csharp
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

public interface IAiTranslationService
{
    /// <summary>
    /// 翻译单条文本
    /// </summary>
    Task<string> TranslateTextAsync(string text, string targetLanguage);
    
    /// <summary>
    /// 翻译PPT提取数据
    /// </summary>
    Task<List<SlideTextData>> TranslatePptExtractionAsync(PptTextExtraction extraction, string targetLanguage);
}
```

### 示例 1: 使用OpenAI翻译服务

```csharp
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

public class OpenAiTranslator : IAiTranslationService
{
    private string ApiKey { get; set; }
    private string Model { get; set; }
    private string Endpoint { get; set; }
    
    public OpenAiTranslator(string apiKey, string model = "gpt-4", string endpoint = "https://api.openai.com/v1/chat/completions")
    {
        ApiKey = apiKey;
        Model = model;
        Endpoint = endpoint;
    }
    
    public async Task<string> TranslateTextAsync(string text, string targetLanguage)
    {
        // 构建翻译prompt
        string systemPrompt = $@"你是一个专业的文档翻译专家。
请将以下文本翻译为{targetLanguage}，确保翻译的准确性和专业性。
只返回翻译结果，不要添加任何解释或额外内容。";
        
        // 调用OpenAI API
        var requestBody = new
        {
            model = Model,
            messages = new[]
            {
                new { role = "system", content = systemPrompt },
                new { role = "user", content = text }
            },
            temperature = 0.3,
            max_tokens = 1000
        };
        
        using (var httpClient = new HttpClient())
        {
            httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {ApiKey}");
            httpClient.DefaultRequestHeaders.Add("Content-Type", "application/json");
            
            var jsonContent = System.Text.Json.JsonSerializer.Serialize(requestBody);
            var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
            
            var response = await httpClient.PostAsync(Endpoint, content);
            var responseContent = await response.Content.ReadAsStringAsync();
            
            if (response.IsSuccessStatusCode)
            {
                var result = System.Text.Json.JsonSerializer.Deserialize<JsonElement>(responseContent);
                return result.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString();
            }
            else
            {
                throw new Exception($"翻译失败: {responseContent}");
            }
        }
    }
    
    public async Task<List<SlideTextData>> TranslatePptExtractionAsync(PptTextExtraction extraction, string targetLanguage)
    {
        Console.WriteLine($"开始翻译PPT内容，目标语言: {targetLanguage}");
        
        List<SlideTextData> translatedSlides = new List<SlideTextData>();
        
        foreach (SlideTextData slide in extraction.Slides)
        {
            Console.WriteLine($"  翻译第 {slide.SlideIndex + 1} 张幻灯片，共 {slide.Shapes.Count} 个形状");
            
            SlideTextData translatedSlide = new SlideTextData
            {
                SlideIndex = slide.SlideIndex,
                Shapes = new List<ShapeTextData>()
            };
            
            foreach (ShapeTextData shape in slide.Shapes)
            {
                ShapeTextData translatedShape = new ShapeTextData
                {
                    ShapeIndex = shape.ShapeIndex,
                    ShapeIndices = shape.ShapeIndices,
                    TextContent = await TranslateTextContentAsync(shape.TextContent, targetLanguage)
                };
                
                translatedSlide.Shapes.Add(translatedShape);
            }
            
            translatedSlides.Add(translatedSlide);
        }
        
        Console.WriteLine($"翻译完成，共翻译了 {translatedSlides.Count} 张幻灯片");
        
        return translatedSlides;
    }
    
    private async Task<TextContent> TranslateTextContentAsync(TextContent original, string targetLanguage)
    {
        if (original == null) return null;
        
        TextContent translated = new TextContent();
        
        // 翻译段落
        if (original.Paragraphs != null && original.Paragraphs.Count > 0)
        {
            translated.Paragraphs = new List<ParagraphData>();
            foreach (ParagraphData para in original.Paragraphs)
            {
                string translatedText = await TranslateTextAsync(para.Text, targetLanguage);
                translated.Paragraphs.Add(new ParagraphData
                {
                    ParagraphIndex = para.ParagraphIndex,
                    Text = translatedText
                });
            }
        }
        
        // 翻译表格单元格
        if (original.TableCells != null && original.TableCells.Count > 0)
        {
            translated.TableCells = new List<CellData>();
            foreach (CellData cell in original.TableCells)
            {
                string translatedText = await TranslateTextAsync(cell.Text, targetLanguage);
                translated.TableCells.Add(new CellData
                {
                    Row = cell.Row,
                    Column = cell.Column,
                    Text = translatedText
                });
            }
        }
        
        // 翻译SmartArt节点
        if (original.SmartArtNodes != null && original.SmartArtNodes.Count > 0)
        {
            translated.SmartArtNodes = new List<SmartArtNodeData>();
            foreach (SmartArtNodeData node in original.SmartArtNodes)
            {
                string translatedText = await TranslateTextAsync(node.Text, targetLanguage);
                translated.SmartArtNodes.Add(new SmartArtNodeData
                {
                    NodeIndex = node.NodeIndex,
                    Text = translatedText
                });
            }
        }
        
        return translated;
    }
}
```

### 示例 2: 使用Claude翻译服务

```csharp
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

public class ClaudeTranslator : IAiTranslationService
{
    private string ApiKey { get; set; }
    private string Model { get; set; }
    
    public ClaudeTranslator(string apiKey, string model = "claude-3-opus-20240229")
    {
        ApiKey = apiKey;
        Model = model;
    }
    
    public async Task<string> TranslateTextAsync(string text, string targetLanguage)
    {
        // 这里需要实现Claude API调用
        // 示例代码，用户需要根据实际的Claude API进行调整
        
        throw new NotImplementedException("请根据Claude API文档实现翻译功能");
    }
    
    public async Task<List<SlideTextData>> TranslatePptExtractionAsync(PptTextExtraction extraction, string targetLanguage)
    {
        // 使用与OpenAiTranslator相同的逻辑
        // 但调用Claude API进行翻译
        
        throw new NotImplementedException("请根据Claude API文档实现翻译功能");
    }
}
```

## 文本替换

### 文本替换类

```csharp
using System;
using System.Collections.Generic;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using Spire.Presentation.Tables;

public class PptTextReplacer
{
    // 基于翻译结果替换PPT中的文本
    public static void ReplaceTranslatedText(Presentation presentation, List<SlideTextData> translatedSlides)
    {
        Console.WriteLine("开始替换翻译后的文本");
        
        foreach (SlideTextData translatedSlide in translatedSlides)
        {
            int slideIndex = translatedSlide.SlideIndex;
            if (slideIndex >= presentation.Slides.Count)
            {
                Console.WriteLine($"  警告：幻灯片索引 {slideIndex} 超出范围");
                continue;
            }
            
            ISlide slide = presentation.Slides[slideIndex];
            
            Console.WriteLine($"  处理第 {slideIndex + 1} 张幻灯片，共 {translatedSlide.Shapes.Count} 个形状");
            
            foreach (ShapeTextData translatedShape in translatedSlide.Shapes)
            {
                List<int> shapeIndices = translatedShape.ShapeIndices ?? 
                    new List<int> { translatedShape.ShapeIndex };
                
                IShape parentShape = GetShapeByIndices(slide.Shapes, shapeIndices);
                
                if (parentShape == null)
                {
                    Console.WriteLine($"    警告：无法找到形状索引 [{string.Join(", ", shapeIndices)}]");
                    continue;
                }
                
                // 替换文本
                ReplaceShapeText(parentShape, translatedShape.TextContent);
                Console.WriteLine($"    已替换形状 {string.Join(", ", shapeIndices)} 的文本");
            }
        }
        
        Console.WriteLine("文本替换完成");
    }
    
    // 根据索引获取形状
    private static IShape GetShapeByIndices(ShapeCollection shapes, List<int> indices)
    {
        if (indices == null || indices.Count == 0)
            return null;
        
        IShape currentShape = shapes[indices[0]];
        
        // 处理GroupShape中的嵌套
        if (indices.Count > 1 && currentShape is GroupShape groupShape)
        {
            for (int i = 1; i < indices.Count; i++)
            {
                if (groupShape == null || groupShape.Shapes.Count <= 0)
                    return null;
                
                int nextIndex = indices[i];
                if (nextIndex >= groupShape.Shapes.Count)
                    return null;
                
                IShape nextShape = groupShape.Shapes[nextIndex];
                
                if (nextShape is GroupShape nextGroup)
                {
                    groupShape = nextGroup;
                }
                else
                {
                    currentShape = nextShape;
                    break;
                }
            }
        }
        
        return currentShape;
    }
    
    // 替换形状中的文本
    private static void ReplaceShapeText(IShape shape, TextContent translatedContent)
    {
        if (translatedContent == null) return;
        
        // 处理IAutoShape
        if (shape is IAutoShape autoShape && autoShape.TextFrame != null)
        {
            if (translatedContent.Paragraphs != null)
            {
                for (int i = 0; i < translatedContent.Paragraphs.Count && 
                     i < autoShape.TextFrame.Paragraphs.Count; i++)
                {
                    autoShape.TextFrame.Paragraphs[i].Text = translatedContent.Paragraphs[i].Text;
                }
            }
        }
        // 处理ITable
        else if (shape is ITable table)
        {
            if (translatedContent.TableCells != null)
            {
                foreach (CellData cell in translatedContent.TableCells)
                {
                    if (cell.Row < table.TableRows.Count && 
                        cell.Column < table.ColumnsList.Count)
                    {
                        table[cell.Column, cell.Row].TextFrame.Text = cell.Text;
                    }
                }
            }
        }
        // 处理ISmartArt
        else if (shape is ISmartArt smartArt)
        {
            if (translatedContent.SmartArtNodes != null)
            {
                int nodeIndex = 0;
                ReplaceSmartArtNodes(smartArt.Nodes, translatedContent.SmartArtNodes, ref nodeIndex);
            }
        }
    }
    
    // 递归替换SmartArt节点文本
    private static void ReplaceSmartArtNodes(ISmartArtNodeCollection nodes, 
                                            List<SmartArtNodeData> translatedNodes, 
                                            ref int index)
    {
        foreach (ISmartArtNode node in nodes)
        {
            if (index < translatedNodes.Count && node.TextFrame != null)
            {
                node.TextFrame.Text = translatedNodes[index].Text;
                index++;
            }
            
            // 递归处理子节点
            if (node.ChildNodes.Count > 0)
            {
                ReplaceSmartArtNodes(node.ChildNodes, translatedNodes, ref index);
            }
        }
    }
}
```

## 完整示例

### 示例 3: 端到端翻译流程

```csharp
using System;
using System.IO;
using Spire.Presentation;

public class PptTranslationService
{
    public static void TranslatePpt(string inputPath, string outputPath, string targetLanguage, IAiTranslationService translator)
    {
        Console.WriteLine("========================================");
        Console.WriteLine("PPT翻译服务");
        Console.WriteLine("========================================");
        Console.WriteLine();
        
        Console.WriteLine($"输入文件: {inputPath}");
        Console.WriteLine($"输出文件: {outputPath}");
        Console.WriteLine($"目标语言: {targetLanguage}");
        Console.WriteLine();
        
        try
        {
            // 1. 提取文本
            Console.WriteLine("步骤 1/4: 提取PPT文本");
            PptTextExtraction extraction = PptTextExtractor.ExtractAllText(inputPath);
            
            // 保存提取结果用于调试
            string extractionJson = Path.Combine(Path.GetDirectoryName(outputPath), "extraction.json");
            PptTextExtractor.SaveExtractionAsJson(extraction, extractionJson);
            Console.WriteLine();
            
            // 2. 翻译文本
            Console.WriteLine("步骤 2/4: 翻译文本内容");
            List<SlideTextData> translatedSlides = await translator.TranslatePptExtractionAsync(extraction, targetLanguage);
            Console.WriteLine();
            
            // 3. 替换文本
            Console.WriteLine("步骤 3/4: 替换翻译后的文本");
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(inputPath);
            
            PptTextReplacer.ReplaceTranslatedText(presentation, translatedSlides);
            Console.WriteLine();
            
            // 4. 保存结果
            Console.WriteLine("步骤 4/4: 保存翻译后的PPT");
            presentation.SaveToFile(outputPath, FileFormat.Pptx2010);
            presentation.Dispose();
            
            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine("翻译完成！");
            Console.WriteLine("========================================");
            Console.WriteLine($"输出文件: {outputPath}");
            Console.WriteLine();
            Console.WriteLine("提示:");
            Console.WriteLine("- 原始格式和结构已完整保持");
            Console.WriteLine("- 建议在PowerPoint中检查翻译结果");
            Console.WriteLine("- 如需要调整，可手动修改");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"翻译失败: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}

// 使用示例
public class Program
{
    public static async Task Main(string[] args)
    {
        // 配置
        string inputPath = "presentation.pptx";
        string outputPath = "presentation_en.pptx";
        string targetLanguage = "English";
        string apiKey = "your-api-key"; // 替换为实际的API密钥
        
        // 创建翻译服务
        IAiTranslationService translator = new OpenAiTranslator(apiKey);
        
        // 执行翻译
        await PptTranslationService.TranslatePpt(inputPath, outputPath, targetLanguage, translator);
    }
}
```

### 示例 4: 批量翻译PPT

```csharp
using System;
using System.IO;
using System.Threading.Tasks;

public class BatchPptTranslator
{
    // 批量翻译PPT文件
    public static async Task TranslateBatch(string inputFolder, string outputFolder, string targetLanguage, IAiTranslationService translator)
    {
        Console.WriteLine("========================================");
        Console.WriteLine("批量PPT翻译");
        Console.WriteLine("========================================");
        Console.WriteLine();
        
        // 确保输出目录存在
        Directory.CreateDirectory(outputFolder);
        
        // 获取所有PPT文件
        string[] pptFiles = Directory.GetFiles(inputFolder, "*.pptx");
        
        Console.WriteLine($"找到 {pptFiles.Length} 个PPT文件");
        Console.WriteLine();
        
        int successCount = 0;
        int failCount = 0;
        
        foreach (string inputFile in pptFiles)
        {
            string fileName = Path.GetFileNameWithoutExtension(inputFile);
            string outputFile = Path.Combine(outputFolder, $"{fileName}_translated.pptx");
            
            Console.WriteLine($"正在翻译: {fileName}");
            
            try
            {
                await PptTranslationService.TranslatePpt(inputFile, outputFile, targetLanguage, translator);
                successCount++;
                Console.WriteLine($"✓ 翻译成功: {fileName}");
            }
            catch (Exception ex)
            {
                failCount++;
                Console.WriteLine($"✗ 翻译失败: {fileName} - {ex.Message}");
            }
            
            Console.WriteLine();
        }
        
        Console.WriteLine("========================================");
        Console.WriteLine("批量翻译完成！");
        Console.WriteLine("========================================");
        Console.WriteLine($"成功: {successCount} 个");
        Console.WriteLine($"失败: {failCount} 个");
        Console.WriteLine($"总计: {pptFiles.Length} 个");
    }
}
```

### 示例 5: 智能语言识别

```csharp
using System;
using System.Linq;

public class LanguageDetector
{
    // 检测文本语言
    public static string DetectLanguage(string text)
    {
        // 简单的语言检测逻辑
        // 实际项目中可以使用更专业的语言检测库
        
        if (string.IsNullOrWhiteSpace(text))
            return "Unknown";
        
        // 检查是否包含中文字符
        int chineseCharCount = text.Count(c => c >= 0x4E00 && c <= 0x9FFF);
        int totalCharCount = text.Length;
        
        if (chineseCharCount > 0)
        {
            float chineseRatio = (float)chineseCharCount / totalCharCount;
            
            if (chineseRatio > 0.5)
            {
                return "Chinese";
            }
        }
        
        // 检查是否包含英文字母
        int englishCharCount = text.Count(c => 
            (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'));
        
        if (englishCharCount > 0)
        {
            return "English";
        }
        
        return "Unknown";
    }
    
    // 根据检测的语言确定目标语言
    public static string DetermineTargetLanguage(string detectedLanguage, string userRequirement = null)
    {
        if (!string.IsNullOrEmpty(userRequirement))
        {
            // 根据用户需求确定目标语言
            if (userRequirement.Contains("英文") || userRequirement.Contains("English"))
            {
                return "English";
            }
            else if (userRequirement.Contains("中文") || userRequirement.Contains("Chinese"))
            {
                return "Chinese";
            }
            else if (userRequirement.Contains("日文") || userRequirement.Contains("Japanese"))
            {
                return "Japanese";
            }
            else if (userRequirement.Contains("韩文") || userRequirement.Contains("Korean"))
            {
                return "Korean";
            }
        }
        
        // 根据检测的源语言确定目标语言
        switch (detectedLanguage)
        {
            case "Chinese":
                return "English";
            case "English":
                return "Chinese";
            default:
                return "English"; // 默认翻译为英文
        }
    }
    
    // 检测PPT的语言并确定翻译目标
    public static (string sourceLanguage, string targetLanguage) DetectAndDetermineLanguages(string pptxPath)
    {
        // 提取文本
        PptTextExtraction extraction = PptTextExtractor.ExtractAllText(pptxPath);
        
        // 收集所有文本
        string allText = "";
        foreach (SlideTextData slide in extraction.Slides)
        {
            foreach (ShapeTextData shape in slide.Shapes)
            {
                if (shape.TextContent.Paragraphs != null)
                {
                    allText += string.Join(" ", shape.TextContent.Paragraphs.Select(p => p.Text)) + " ";
                }
            }
        }
        
        // 检测语言
        string detectedLanguage = DetectLanguage(allText);
        
        // 确定目标语言
        string targetLanguage = DetermineTargetLanguage(detectedLanguage);
        
        return (detectedLanguage, targetLanguage);
    }
}
```

## 翻译规则

### 翻译规则配置

```csharp
public class TranslationRules
{
    // 不同PPT类型的翻译策略
    public static string GetTranslationPromptForType(string pptType, string targetLanguage)
    {
        switch (pptType)
        {
            case "学术":
                return $"你是一个专业的学术文档翻译专家。\n" +
                       $"请将以下内容翻译为{targetLanguage}，确保：\n" +
                       $"1. 术语翻译准确\n" +
                       $"2. 学术表达规范\n" +
                       $"3. 保持逻辑严谨性\n" +
                       $"只返回翻译结果，不要添加任何解释。";
                       
            case "商务":
                return $"你是一个专业的商务文档翻译专家。\n" +
                       $"请将以下内容翻译为{targetLanguage}，确保：\n" +
                       $"1. 专业术语准确\n" +
                       $"2. 商务表达得体\n" +
                       $"3. 保持商务礼仪\n" +
                       $"只返回翻译结果，不要添加任何解释。";
                       
            case "营销":
                return $"你是一个专业的营销文案翻译专家。\n" +
                       $"请将以下内容翻译为{targetLanguage}，确保：\n" +
                       $"1. 语言富有吸引力\n" +
                       $"2. 具有行动号召力\n" +
                       $"3. 符合营销风格\n" +
                       $"只返回翻译结果，不要添加任何解释。";
                       
            default:
                return $"你是一个专业的文档翻译专家。\n" +
                       $"请将以下内容翻译为{targetLanguage}，确保翻译的准确性和专业性。\n" +
                       $"只返回翻译结果，不要添加任何解释。";
        }
    }
    
    // 特殊术语处理
    public static Dictionary<string, string> GetSpecialTerms()
    {
        return new Dictionary<string, string>
        {
            // 英文到中文的特殊术语
            { "Artificial Intelligence", "人工智能" },
            { "Machine Learning", "机器学习" },
            { "Deep Learning", "深度学习" },
            { "Neural Network", "神经网络" },
            
            // 中文到英文的特殊术语
            { "人工智能", "Artificial Intelligence" },
            { "机器学习", "Machine Learning" },
            { "深度学习", "Deep Learning" },
            { "神经网络", "Neural Network" }
        };
    }
}
```

## 注意事项

1. **格式保持** - 翻译过程中需要保持原有的字体、颜色、大小等格式属性
2. **表格完整性** - 必须保持表格的行数、列数和结构，不能删除或添加单元格
3. **特殊字符** - 某些特殊字符在翻译过程中可能需要特殊处理
4. **性能考虑** - 大型PPT翻译可能需要较长时间，建议提供进度反馈
5. **错误处理** - AI调用失败时需要有合理的降级策略
6. **API配额** - 注意AI服务的API调用配额和费用
7. **文本长度** - 某些AI服务有单次文本长度限制，需要分批翻译

## 最佳实践

1. **预处理** - 在翻译前检查PPT文件是否可读
2. **备份原文件** - 翻译前备份原始PPT文件
3. **分批处理** - 大型PPT建议分批翻译以避免超时
4. **结果验证** - 翻译完成后人工检查关键内容
5. **术语管理** - 建立术语库确保专业术语翻译准确
6. **缓存机制** - 对相同文本使用缓存减少API调用
7. **日志记录** - 记录翻译过程以便问题排查

## 相关功能

- [文本处理](./03-text-content.md) - 文本内容和格式
- [形状处理](./04-shapes-images.md) - 形状和文本框
- [表格](./05-tables.md) - 表格数据处理
- [SmartArt](./07-smartart.md) - SmartArt文本处理
- [高级功能](./12-advanced-features.md) - 其他高级功能
