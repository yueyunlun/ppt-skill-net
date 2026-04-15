---
title: AI驱动的PPT创建
category: spire-presentation
description: 基于预置模板和结构化大纲，自动创建完整的PPT演示文稿
---

# AI驱动的PPT创建

## 概述

AI驱动的PPT创建功能可以基于预置的专业模板和结构化大纲，自动生成完整的演示文稿。该功能支持：

- **14种预置模板** - 涵盖创意、商务、工作总结等多种风格
- **结构化大纲** - 封面、目录、章节、内容页、感谢页的完整结构
- **自动图表创建** - 支持折线图、柱状图、饼图等数据图表
- **占位符替换** - 自动识别并替换模板中的占位符
- **灵活内容填充** - 支持文本、要点、图表等多种内容类型

通过模板和大纲的结合，可以快速生成专业美观的PPT，大幅提升工作效率。

## 预置模板库

### 模板列表

| 模板文件 | 风格 | 适用场景 |
|---------|------|---------|
| 创意手绘风活动策划PPT模板.pptx | 创意手绘 | 活动、创意、年轻化 |
| 大气简洁工作报告PPT模板.pptx | 大气简洁 | 工作报告、商务、正式 |
| 极简大气年终报告PPT模板.pptx | 极简大气 | 年终报告、总结汇报 |
| 清新活动策划方案汇报PPT模板.pptx | 清新活动 | 活动策划、方案汇报 |
| 清新淡雅工作总结计划PPT模板.pptx | 清新淡雅 | 工作总结、计划制定 |
| 清爽扁平化工作总结汇报PPT模板.pptx | 清爽扁平 | 总结汇报、数据展示 |
| 渐变圆圈泡泡工作总结PPT模板.pptx | 渐变泡泡 | 工作总结、轻松汇报 |
| 简洁大方工作总结PPT模板.pptx | 简洁大方 | 工作总结、通用汇报 |
| 简洁通用工作汇报总结PPT模板.pptx | 简洁通用 | 工作汇报、总结 |
| 简约大气通用总结计划PPT模板.pptx | 简约大气 | 总结计划、目标规划 |
| 简约彩色扁平化报告PPT模板.pptx | 简约彩色 | 报告演示、数据分析 |
| 简约淡雅工作汇报总结.pptx | 简约淡雅 | 工作汇报、总结 |
| 简约通用工作总结报告PPT模板.pptx | 简约通用 | 总结报告、月报 |
| 蓝色大气工作汇报PPT模板.pptx | 蓝色大气 | 工作汇报、商务演示 |

### 模板配置类

```csharp
using System;
using System.Collections.Generic;
using System.IO;

public class PptTemplate
{
    public string FileName { get; set; }
    public string Name { get; set; }
    public string Style { get; set; }
    public string[] UseCases { get; set; }
}

public static class PptTemplateLibrary
{
    public static readonly string TemplateDirectory = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "templates");

    public static List<PptTemplate> Templates { get; } = new List<PptTemplate>
    {
        new PptTemplate 
        { 
            FileName = "创意手绘风活动策划PPT模板.pptx",
            Name = "创意手绘风",
            Style = "创意手绘",
            UseCases = new[] { "活动", "创意", "年轻化" }
        },
        new PptTemplate 
        { 
            FileName = "大气简洁工作报告PPT模板.pptx",
            Name = "大气简洁",
            Style = "大气简洁",
            UseCases = new[] { "工作报告", "商务", "正式" }
        },
        new PptTemplate 
        { 
            FileName = "极简大气年终报告PPT模板.pptx",
            Name = "极简大气",
            Style = "极简大气",
            UseCases = new[] { "年终报告", "总结汇报", "年度总结" }
        },
        new PptTemplate 
        { 
            FileName = "清新活动策划方案汇报PPT模板.pptx",
            Name = "清新活动策划",
            Style = "清新活动",
            UseCases = new[] { "活动策划", "方案汇报", "活动方案" }
        },
        new PptTemplate 
        { 
            FileName = "清新淡雅工作总结计划PPT模板.pptx",
            Name = "清新淡雅",
            Style = "清新淡雅",
            UseCases = new[] { "工作总结", "计划制定", "工作计划" }
        },
        new PptTemplate 
        { 
            FileName = "清爽扁平化工作总结汇报PPT模板.pptx",
            Name = "清爽扁平化",
            Style = "清爽扁平",
            UseCases = new[] { "总结汇报", "数据展示", "工作汇报" }
        },
        new PptTemplate 
        { 
            FileName = "渐变圆圈泡泡工作总结PPT模板.pptx",
            Name = "渐变泡泡",
            Style = "渐变泡泡",
            UseCases = new[] { "工作总结", "轻松汇报", "轻松汇报" }
        },
        new PptTemplate 
        { 
            FileName = "简洁大方工作总结PPT模板.pptx",
            Name = "简洁大方",
            Style = "简洁大方",
            UseCases = new[] { "工作总结", "通用汇报", "总结" }
        },
        new PptTemplate 
        { 
            FileName = "简洁通用工作汇报总结PPT模板.pptx",
            Name = "简洁通用",
            Style = "简洁通用",
            UseCases = new[] { "工作汇报", "总结", "汇报" }
        },
        new PptTemplate 
        { 
            FileName = "简约大气通用总结计划PPT模板.pptx",
            Name = "简约大气",
            Style = "简约大气",
            UseCases = new[] { "总结计划", "目标规划", "计划" }
        },
        new PptTemplate 
        { 
            FileName = "简约彩色扁平化报告PPT模板.pptx",
            Name = "简约彩色",
            Style = "简约彩色",
            UseCases = new[] { "报告演示", "数据分析", "报告" }
        },
        new PptTemplate 
        { 
            FileName = "简约淡雅工作汇报总结.pptx",
            Name = "简约淡雅",
            Style = "简约淡雅",
            UseCases = new[] { "工作汇报", "总结", "汇报总结" }
        },
        new PptTemplate 
        { 
            FileName = "简约通用工作总结报告PPT模板.pptx",
            Name = "简约通用",
            Style = "简约通用",
            UseCases = new[] { "总结报告", "月报", "月度总结" }
        },
        new PptTemplate 
        { 
            FileName = "蓝色大气工作汇报PPT模板.pptx",
            Name = "蓝色大气",
            Style = "蓝色大气",
            UseCases = new[] { "工作汇报", "商务演示", "汇报" }
        }
    };

    // 获取模板完整路径
    public static string GetTemplatePath(string fileName)
    {
        return Path.Combine(TemplateDirectory, fileName);
    }

    // 根据名称获取模板
    public static PptTemplate GetTemplateByName(string name)
    {
        return Templates.Find(t => t.Name.Contains(name) || t.FileName.Contains(name));
    }

    // 根据用途获取推荐模板
    public static PptTemplate GetTemplateByUseCase(string useCase)
    {
        foreach (var template in Templates)
        {
            foreach (var caseKeyword in template.UseCases)
            {
                if (useCase.Contains(caseKeyword) || caseKeyword.Contains(useCase))
                {
                    return template;
                }
            }
        }
        // 默认返回第一个模板
        return Templates[0];
    }
}
```

## PPT大纲结构

### 大纲数据结构

```csharp
using System;
using System.Collections.Generic;

// PPT大纲根结构
public class PptOutline
{
    public CoverSlide CoverSlide { get; set; }
    public TocSlide TocSlide { get; set; }
    public List<Section> Sections { get; set; }

    public PptOutline()
    {
        Sections = new List<Section>();
    }
}

// 封面页
public class CoverSlide
{
    public string Title { get; set; }
    public string Subtitle { get; set; }
    public string Presenter { get; set; }
    public string Date { get; set; }
}

// 目录页
public class TocSlide
{
    public List<string> Sections { get; set; }

    public TocSlide()
    {
        Sections = new List<string>();
    }
}

// 章节
public class Section
{
    public string SectionTitle { get; set; }
    public List<ContentSlide> Slides { get; set; }

    public Section()
    {
        Slides = new List<ContentSlide>();
    }
}

// 内容页
public class ContentSlide
{
    public string SlideTitle { get; set; }
    public List<BulletPoint> BulletPoints { get; set; }
    public ChartData Chart { get; set; }
    public string Summary { get; set; }

    public ContentSlide()
    {
        BulletPoints = new List<BulletPoint>();
    }
}

// 要点
public class BulletPoint
{
    public string ContentTitle { get; set; }
    public string Content { get; set; }
}

// 图表数据
public class ChartData
{
    public string Type { get; set; }  // linechart, barchart, piechart
    public string Title { get; set; }
    public List<string> XAxis { get; set; }
    public List<double> YAxis { get; set; }

    public ChartData()
    {
        XAxis = new List<string>();
        YAxis = new List<double>();
    }
}
```

## 大纲生成示例

### 示例 1: 创建基本大纲

```csharp
public static class PptOutlineGenerator
{
    // 生成示例大纲
    public static PptOutline GenerateSampleOutline(string topic)
    {
        return new PptOutline
        {
            CoverSlide = new CoverSlide
            {
                Title = topic,
                Subtitle = "专业汇报 · 高效沟通",
                Presenter = "汇报人：您的姓名",
                Date = DateTime.Now.ToString("yyyy年MM月dd日")
            },
            TocSlide = new TocSlide
            {
                Sections = new List<string>
                {
                    "背景介绍",
                    "主要内容",
                    "实施计划",
                    "预期成果"
                }
            },
            Sections = new List<Section>
            {
                new Section
                {
                    SectionTitle = "背景介绍",
                    Slides = new List<ContentSlide>
                    {
                        new ContentSlide
                        {
                            SlideTitle = "项目背景",
                            BulletPoints = new List<BulletPoint>
                            {
                                new BulletPoint 
                                { 
                                    ContentTitle = "市场需求", 
                                    Content = "当前市场对该类产品的需求日益增长，用户规模不断扩大" 
                                },
                                new BulletPoint 
                                { 
                                    ContentTitle = "政策支持", 
                                    Content = "国家相关政策为项目发展提供了有力支持和保障" 
                                },
                                new BulletPoint 
                                { 
                                    ContentTitle = "技术基础", 
                                    Content = "成熟的技术方案为项目实施提供了坚实基础" 
                                }
                            }
                        }
                    }
                },
                new Section
                {
                    SectionTitle = "主要内容",
                    Slides = new List<ContentSlide>
                    {
                        new ContentSlide
                        {
                            SlideTitle = "市场规模分析",
                            BulletPoints = new List<BulletPoint>
                            {
                                new BulletPoint 
                                { 
                                    ContentTitle = "市场容量", 
                                    Content = "整体市场规模达到千亿级别，增长潜力巨大" 
                                },
                                new BulletPoint 
                                { 
                                    ContentTitle = "竞争格局", 
                                    Content = "市场竞争格局清晰，头部企业占据主要份额" 
                                }
                            },
                            Chart = new ChartData
                            {
                                Type = "barchart",
                                Title = "市场规模增长趋势（亿元）",
                                XAxis = new List<string> { "2020", "2021", "2022", "2023", "2024" },
                                YAxis = new List<double> { 150, 200, 280, 350, 450 }
                            },
                            Summary = "市场规模呈现稳定增长态势，年复合增长率超过20%"
                        }
                    }
                }
            }
        };
    }
}
```

### 示例 2: 生成销售业绩报告大纲

```csharp
// 生成销售业绩报告大纲
public static PptOutline GenerateSalesReportOutline(string quarter)
{
    return new PptOutline
    {
        CoverSlide = new CoverSlide
        {
            Title = $"{quarter}销售业绩报告",
            Subtitle = "数据驱动 · 持续增长",
            Presenter = "销售部",
            Date = DateTime.Now.ToString("yyyy年MM月dd日")
        },
        TocSlide = new TocSlide
        {
            Sections = new List<string>
            {
                "销售业绩概览",
                "各区域表现",
                "产品销售分析",
                "下季度规划"
            }
        },
        Sections = new List<Section>
        {
            new Section
            {
                SectionTitle = "销售业绩概览",
                Slides = new List<ContentSlide>
                {
                    new ContentSlide
                    {
                        SlideTitle = "总体业绩完成情况",
                        BulletPoints = new List<BulletPoint>
                        {
                            new BulletPoint 
                            { 
                                ContentTitle = "完成率", 
                                Content = "整体销售目标完成率达到115%，超额完成任务" 
                            },
                            new BulletPoint 
                            { 
                                ContentTitle = "同比增长", 
                                Content = "与去年同期相比增长28%，增长势头强劲" 
                            },
                            new BulletPoint 
                            { 
                                ContentTitle = "客户满意度", 
                                Content = "客户满意度评分达到4.8分（满分5分）" 
                            }
                        }
                    }
                }
            },
            new Section
            {
                SectionTitle = "各区域表现",
                Slides = new List<ContentSlide>
                {
                    new ContentSlide
                    {
                        SlideTitle = "区域销售对比",
                        BulletPoints = new List<BulletPoint>
                        {
                            new BulletPoint 
                            { 
                                ContentTitle = "华东区", 
                                Content = "销售额占比最高，达到35%" 
                            },
                            new BulletPoint 
                            { 
                                ContentTitle = "华南区", 
                                Content = "增长速度最快，同比增长40%" 
                            }
                        },
                        Chart = new ChartData
                        {
                            Type = "piechart",
                            Title = "各区域销售额占比",
                            XAxis = new List<string> { "华东区", "华南区", "华北区", "西南区", "西北区" },
                            YAxis = new List<double> { 35, 25, 20, 12, 8 }
                        },
                        Summary = "华东区继续保持领先地位，华南区增长迅速，市场潜力巨大"
                    }
                }
            }
        }
    };
}
```

## 基于模板创建PPT

### 示例 3: 核心创建逻辑

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using Spire.Presentation.Charts;

public class PptCreator
{
    // 基于模板和大纲创建PPT
    public static void CreatePpt(string templatePath, PptOutline outline, string outputPath)
    {
        Presentation ppt = new Presentation();
        
        try
        {
            // 1. 加载模板
            ppt.LoadFromFile(templatePath);
            
            Console.WriteLine("模板加载成功");
            
            // 2. 获取第一张幻灯片作为封面
            if (ppt.Slides.Count > 0)
            {
                ISlide coverSlide = ppt.Slides[0];
                FillCoverSlide(coverSlide, outline.CoverSlide);
                Console.WriteLine("封面页填充完成");
            }
            
            // 3. 创建目录页
            if (outline.TocSlide != null && outline.TocSlide.Sections.Count > 0)
            {
                ISlide tocSlide = CreateTocSlide(ppt, outline.TocSlide);
                Console.WriteLine("目录页创建完成");
            }
            
            // 4. 创建章节和内容页
            foreach (Section section in outline.Sections)
            {
                // 创建章节页
                ISlide sectionSlide = CreateSectionSlide(ppt, section.SectionTitle);
                Console.WriteLine($"章节页创建完成：{section.SectionTitle}");
                
                // 创建内容页
                foreach (ContentSlide contentSlide in section.Slides)
                {
                    ISlide slide = CreateContentSlide(ppt, contentSlide);
                    Console.WriteLine($"内容页创建完成：{contentSlide.SlideTitle}");
                }
            }
            
            // 5. 创建感谢页
            ISlide thankSlide = CreateThankSlide(ppt);
            Console.WriteLine("感谢页创建完成");
            
            // 6. 保存PPT
            ppt.SaveToFile(outputPath, FileFormat.Pptx2010);
            Console.WriteLine($"PPT已保存到：{outputPath}");
        }
        finally
        {
            ppt.Dispose();
        }
    }
    
    // 填充封面
    private static void FillCoverSlide(ISlide slide, CoverSlide cover)
    {
        foreach (IShape shape in slide.Shapes)
        {
            if (shape is IAutoShape autoShape && autoShape.TextFrame != null)
            {
                string text = autoShape.TextFrame.Text;
                
                // 替换标题占位符
                if (text.Contains("$title$") || text.Contains("标题") || text.Contains("Title"))
                {
                    autoShape.TextFrame.Text = cover.Title;
                    // 设置标题样式
                    autoShape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
                    autoShape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 48;
                    autoShape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
                    autoShape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(31, 41, 55);
                    autoShape.TextFrame.Paragraphs[0].TextRanges[0].IsBold = TriState.True;
                }
                // 替换副标题占位符
                else if (text.Contains("$subtitle$") || text.Contains("副标题") || text.Contains("Subtitle"))
                {
                    autoShape.TextFrame.Text = cover.Subtitle;
                    autoShape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
                    autoShape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 24;
                }
                // 替换汇报人占位符
                else if (text.Contains("$presenter$") || text.Contains("汇报人"))
                {
                    autoShape.TextFrame.Text = cover.Presenter;
                    autoShape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
                }
                // 替换日期占位符
                else if (text.Contains("$date$") || text.Contains("日期") || text.Contains("Date"))
                {
                    autoShape.TextFrame.Text = cover.Date;
                    autoShape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
                }
            }
        }
    }
    
    // 创建目录页
    private static ISlide CreateTocSlide(Presentation ppt, TocSlide toc)
    {
        ISlide slide = ppt.Slides.AppendEmptySlide();
        
        // 设置背景色
        slide.Background.Type = BackgroundType.Custom;
        slide.Background.FillFormat.FillType = FillFormatType.Solid;
        slide.Background.FillFormat.SolidFillColor.Color = Color.White;
        
        float slideWidth = ppt.SlideSize.Size.Width;
        float slideHeight = ppt.SlideSize.Size.Height;
        
        // 添加标题
        RectangleF titleRect = new RectangleF(50, 50, slideWidth - 100, 60);
        IAutoShape titleShape = slide.Shapes.AppendShape(ShapeType.Rectangle, titleRect);
        titleShape.Fill.FillType = FillFormatType.Solid;
        titleShape.Fill.SolidColor.Color = Color.FromArgb(30, 58, 138);
        titleShape.ShapeStyle.LineColor.Color = Color.Transparent;
        
        titleShape.AppendTextFrame("目录");
        titleShape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
        titleShape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 36;
        titleShape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.White;
        titleShape.TextFrame.Paragraphs[0].TextRanges[0].IsBold = TriState.True;
        
        // 添加目录项
        float startY = 130;
        int itemNumber = 1;
        
        foreach (string section in toc.Sections)
        {
            // 序号
            RectangleF numberRect = new RectangleF(100, startY + 5, 30, 30);
            IAutoShape numberShape = slide.Shapes.AppendShape(ShapeType.Ellipse, numberRect);
            numberShape.Fill.FillType = FillFormatType.Solid;
            numberShape.Fill.SolidColor.Color = Color.FromArgb(30, 58, 138);
            numberShape.ShapeStyle.LineColor.Color = Color.Transparent;
            
            numberShape.AppendTextFrame(itemNumber.ToString());
            numberShape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
            numberShape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 18;
            numberShape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.White;
            
            // 目录项
            RectangleF itemRect = new RectangleF(150, startY, slideWidth - 250, 40);
            IAutoShape itemShape = slide.Shapes.AppendShape(ShapeType.Rectangle, itemRect);
            itemShape.Fill.FillType = FillFormatType.Solid;
            itemShape.Fill.SolidColor.Color = Color.FromArgb(243, 244, 246);
            itemShape.ShapeStyle.LineColor.Color = Color.FromArgb(30, 58, 138);
            itemShape.LineWidth = 1;
            
            itemShape.AppendTextFrame(section);
            itemShape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
            itemShape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 18;
            itemShape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(31, 41, 55);
            
            startY += 50;
            itemNumber++;
        }
        
        return slide;
    }
    
    // 创建章节页
    private static ISlide CreateSectionSlide(Presentation ppt, string sectionTitle)
    {
        ISlide slide = ppt.Slides.AppendEmptySlide();
        
        float slideWidth = ppt.SlideSize.Size.Width;
        float slideHeight = ppt.SlideSize.Size.Height;
        
        // 设置背景色
        slide.Background.Type = BackgroundType.Custom;
        slide.Background.FillFormat.FillType = FillFormatType.Solid;
        slide.Background.FillFormat.SolidFillColor.Color = Color.FromArgb(30, 58, 138);
        
        // 创建装饰性元素
        RectangleF decorRect = new RectangleF(slideWidth - 200, slideHeight - 200, 180, 180);
        IAutoShape decorShape = slide.Shapes.AppendShape(ShapeType.Ellipse, decorRect);
        decorShape.Fill.FillType = FillFormatType.Solid;
        decorShape.Fill.SolidColor.Color = Color.FromArgb(59, 130, 246);
        decorShape.ShapeStyle.LineColor.Color = Color.Transparent;
        decorShape.Rotation = 45;
        
        // 创建大标题
        RectangleF titleRect = new RectangleF(50, slideHeight / 2 - 60, slideWidth - 100, 120);
        IAutoShape titleShape = slide.Shapes.AppendShape(ShapeType.Rectangle, titleRect);
        titleShape.Fill.FillType = FillFormatType.Solid;
        titleShape.Fill.SolidColor.Color = Color.Transparent;
        titleShape.ShapeStyle.LineColor.Color = Color.Transparent;
        
        titleShape.AppendTextFrame(sectionTitle);
        titleShape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
        titleShape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 48;
        titleShape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.White;
        titleShape.TextFrame.Paragraphs[0].TextRanges[0].IsBold = TriState.True;
        
        return slide;
    }
    
    // 创建内容页
    private static ISlide CreateContentSlide(Presentation ppt, ContentSlide content)
    {
        ISlide slide = ppt.Slides.AppendEmptySlide();
        
        float slideWidth = ppt.SlideSize.Size.Width;
        float slideHeight = ppt.SlideSize.Size.Height;
        
        // 设置背景色
        slide.Background.Type = BackgroundType.Custom;
        slide.Background.FillFormat.FillType = FillFormatType.Solid;
        slide.Background.FillFormat.SolidFillColor.Color = Color.White;
        
        // 添加标题栏
        RectangleF titleBarRect = new RectangleF(0, 0, slideWidth, 80);
        IAutoShape titleBarShape = slide.Shapes.AppendShape(ShapeType.Rectangle, titleBarRect);
        titleBarShape.Fill.FillType = FillFormatType.Solid;
        titleBarShape.Fill.SolidColor.Color = Color.FromArgb(30, 58, 138);
        titleBarShape.ShapeStyle.LineColor.Color = Color.Transparent;
        
        // 添加标题
        RectangleF titleRect = new RectangleF(50, 20, slideWidth - 100, 40);
        IAutoShape titleShape = slide.Shapes.AppendShape(ShapeType.Rectangle, titleRect);
        titleShape.Fill.FillType = FillFormatType.Solid;
        titleShape.Fill.SolidColor.Color = Color.Transparent;
        titleShape.ShapeStyle.LineColor.Color = Color.Transparent;
        
        titleShape.AppendTextFrame(content.SlideTitle);
        titleShape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
        titleShape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 32;
        titleShape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.White;
        titleShape.TextFrame.Paragraphs[0].TextRanges[0].IsBold = TriState.True;
        
        // 确定内容区域
        float contentAreaWidth = slideWidth - 100;
        float contentAreaHeight = slideHeight - 120;
        float contentAreaStartY = 100;
        
        // 如果有图表，创建图表
        if (content.Chart != null)
        {
            CreateChart(slide, content.Chart, 50, contentAreaStartY, slideWidth / 2 - 75, contentAreaHeight - 60);
        }
        
        // 添加要点
        float startY = content.Chart != null ? contentAreaStartY : contentAreaStartY;
        float contentX = content.Chart != null ? slideWidth / 2 : 50;
        float contentWidth = content.Chart != null ? slideWidth / 2 - 50 : contentAreaWidth;
        
        foreach (BulletPoint bullet in content.BulletPoints)
        {
            // 要点标题
            RectangleF bulletTitleRect = new RectangleF(contentX, startY, contentWidth, 25);
            IAutoShape bulletTitleShape = slide.Shapes.AppendShape(ShapeType.Rectangle, bulletTitleRect);
            bulletTitleShape.Fill.FillType = FillFormatType.Solid;
            bulletTitleShape.Fill.SolidColor.Color = Color.Transparent;
            bulletTitleShape.ShapeStyle.LineColor.Color = Color.Transparent;
            
            bulletTitleShape.AppendTextFrame(bullet.ContentTitle);
            bulletTitleShape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
            bulletTitleShape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 20;
            bulletTitleShape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(30, 58, 138);
            bulletTitleShape.TextFrame.Paragraphs[0].TextRanges[0].IsBold = TriState.True;
            
            startY += 30;
            
            // 要点内容
            RectangleF bulletContentRect = new RectangleF(contentX, startY, contentWidth - 10, 35);
            IAutoShape bulletContentShape = slide.Shapes.AppendShape(ShapeType.Rectangle, bulletContentRect);
            bulletContentShape.Fill.FillType = FillFormatType.Solid;
            bulletContentShape.Fill.SolidColor.Color = Color.Transparent;
            bulletContentShape.ShapeStyle.LineColor.Color = Color.Transparent;
            
            bulletContentShape.AppendTextFrame(bullet.Content);
            bulletContentShape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
            bulletContentShape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 14;
            bulletContentShape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(75, 85, 99);
            
            startY += 45;
        }
        
        // 添加总结（如果有）
        if (!string.IsNullOrEmpty(content.Summary))
        {
            RectangleF summaryRect = new RectangleF(contentX, startY, contentWidth, 40);
            IAutoShape summaryShape = slide.Shapes.AppendShape(ShapeType.Rectangle, summaryRect);
            summaryShape.Fill.FillType = FillFormatType.Solid;
            summaryShape.Fill.SolidColor.Color = Color.FromArgb(243, 244, 246);
            summaryShape.ShapeStyle.LineColor.Color = Color.FromArgb(30, 58, 138);
            summaryShape.LineWidth = 1;
            
            summaryShape.AppendTextFrame($"总结：{content.Summary}");
            summaryShape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
            summaryShape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 14;
            summaryShape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(31, 41, 55);
        }
        
        return slide;
    }
    
    // 创建图表
    private static void CreateChart(ISlide slide, ChartData chartData, float x, float y, float width, float height)
    {
        RectangleF chartRect = new RectangleF(x, y, width, height);
        
        IChart chart;
        ChartType chartType;
        
        // 根据类型创建图表
        switch (chartData.Type.ToLower())
        {
            case "linechart":
                chartType = ChartType.Line;
                break;
            case "barchart":
                chartType = ChartType.ColumnClustered;
                break;
            case "piechart":
                chartType = ChartType.Pie;
                break;
            default:
                chartType = ChartType.ColumnClustered;
                break;
        }
        
        chart = slide.Shapes.AppendChartInit(chartType, chartRect, false);
        
        // 设置图表标题
        chart.ChartTitle.TextProperties.Text = chartData.Title;
        chart.ChartTitle.TextProperties.IsCentered = true;
        chart.HasTitle = true;
        
        // 设置系列名
        chart.ChartData[0, 0].Text = "数据";
        
        // 设置横轴
        for (int i = 0; i < chartData.XAxis.Count; i++)
        {
            chart.ChartData[i + 1, 0].Text = chartData.XAxis[i];
        }
        
        // 设置纵轴数值
        for (int i = 0; i < chartData.YAxis.Count; i++)
        {
            chart.ChartData[i + 1, 1].NumberValue = chartData.YAxis[i];
        }
        
        // 创建系列
        IChartSeries series = chart.Series[0];
        series.Values = chart.ChartData["B2", $"B{chartData.YAxis.Count + 1}"];
        series.HasDataLabels = true;
        
        // 饼图特殊设置
        if (chartType == ChartType.Pie)
        {
            series.DataLabels.CategoryNameVisible = true;
            series.DataLabels.PercentValueVisible = true;
        }
    }
    
    // 创建感谢页
    private static ISlide CreateThankSlide(Presentation ppt)
    {
        ISlide slide = ppt.Slides.AppendEmptySlide();
        
        float slideWidth = ppt.SlideSize.Size.Width;
        float slideHeight = ppt.SlideSize.Size.Height;
        
        // 设置背景色
        slide.Background.Type = BackgroundType.Custom;
        slide.Background.FillFormat.FillType = FillFormatType.Solid;
        slide.Background.FillFormat.SolidFillColor.Color = Color.FromArgb(30, 58, 138);
        
        // 添加装饰性元素
        RectangleF decor1Rect = new RectangleF(50, slideHeight - 300, 100, 100);
        IAutoShape decor1 = slide.Shapes.AppendShape(ShapeType.Ellipse, decor1Rect);
        decor1.Fill.FillType = FillFormatType.Solid;
        decor1.Fill.SolidColor.Color = Color.FromArgb(59, 130, 246);
        decor1.ShapeStyle.LineColor.Color = Color.Transparent;
        
        RectangleF decor2Rect = new RectangleF(slideWidth - 150, 50, 100, 100);
        IAutoShape decor2 = slide.Shapes.AppendShape(ShapeType.Ellipse, decor2Rect);
        decor2.Fill.FillType = FillFormatType.Solid;
        decor2.Fill.SolidColor.Color = Color.FromArgb(59, 130, 246);
        decor2.ShapeStyle.LineColor.Color = Color.Transparent;
        
        // 创建主要文本
        RectangleF rect = new RectangleF(0, slideHeight / 2 - 50, slideWidth, 100);
        IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, rect);
        shape.Fill.FillType = FillFormatType.Solid;
        shape.Fill.SolidColor.Color = Color.Transparent;
        shape.ShapeStyle.LineColor.Color = Color.Transparent;
        
        shape.AppendTextFrame("谢谢观看！");
        shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
        shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 60;
        shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.White;
        shape.TextFrame.Paragraphs[0].TextRanges[0].IsBold = TriState.True;
        
        // 添加副文本
        RectangleF subRect = new RectangleF(0, slideHeight / 2 + 30, slideWidth, 50);
        IAutoShape subShape = slide.Shapes.AppendShape(ShapeType.Rectangle, subRect);
        subShape.Fill.FillType = FillFormatType.Solid;
        subShape.Fill.SolidColor.Color = Color.Transparent;
        subShape.ShapeStyle.LineColor.Color = Color.Transparent;
        
        subShape.AppendTextFrame("如有疑问，欢迎交流");
        subShape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
        subShape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 24;
        subShape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(147, 197, 253);
        
        return slide;
    }
}
```

### 示例 4: 使用示例

```csharp
// 基于模板创建PPT的完整示例
public class PptCreationExample
{
    public static void CreatePptFromTemplate()
    {
        // 1. 选择模板
        PptTemplate template = PptTemplateLibrary.GetTemplateByUseCase("工作报告");
        string templatePath = PptTemplateLibrary.GetTemplatePath(template.FileName);
        
        Console.WriteLine($"使用模板：{template.Name}");
        Console.WriteLine($"模板路径：{templatePath}");
        
        // 2. 生成大纲
        string topic = "2024年销售业绩报告";
        PptOutline outline = PptOutlineGenerator.GenerateSalesReportOutline("第一季度");
        
        // 3. 创建PPT
        string outputPath = "销售业绩报告.pptx";
        PptCreator.CreatePpt(templatePath, outline, outputPath);
        
        Console.WriteLine("PPT创建完成！");
    }
}
```

## 完整示例

### 示例 5: 端到端PPT创建

```csharp
using System;
using System.IO;

public class CompletePptCreationDemo
{
    public static void CreateCompletePpt(string topic, string useCase, string presenter)
    {
        Console.WriteLine("========================================");
        Console.WriteLine("AI驱动PPT创建");
        Console.WriteLine("========================================");
        Console.WriteLine();
        
        try
        {
            // 1. 选择合适的模板
            Console.WriteLine("步骤 1/5: 选择模板");
            PptTemplate template = PptTemplateLibrary.GetTemplateByUseCase(useCase);
            string templatePath = PptTemplateLibrary.GetTemplatePath(template.FileName);
            
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"警告：模板文件不存在：{templatePath}");
                Console.WriteLine("使用默认模板...");
                template = PptTemplateLibrary.Templates[0];
                templatePath = PptTemplateLibrary.GetTemplatePath(template.FileName);
            }
            
            Console.WriteLine($"  模板名称：{template.Name}");
            Console.WriteLine($"  模板风格：{template.Style}");
            Console.WriteLine($"  适用场景：{string.Join("、", template.UseCases)}");
            Console.WriteLine();
            
            // 2. 生成大纲
            Console.WriteLine("步骤 2/5: 生成PPT大纲");
            PptOutline outline = GenerateDynamicOutline(topic, presenter);
            Console.WriteLine($"  封面标题：{outline.CoverSlide.Title}");
            Console.WriteLine($"  目录项数：{outline.TocSlide.Sections.Count}");
            Console.WriteLine($"  章节数量：{outline.Sections.Count}");
            
            int totalSlides = outline.Sections.Sum(s => s.Slides.Count) + 3; // 封面、目录、感谢
            Console.WriteLine($"  总页数：{totalSlides}");
            Console.WriteLine();
            
            // 3. 创建PPT
            Console.WriteLine("步骤 3/5: 基于模板创建PPT");
            string outputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output");
            Directory.CreateDirectory(outputDir);
            
            string outputPath = Path.Combine(outputDir, $"{topic}.pptx");
            PptCreator.CreatePpt(templatePath, outline, outputPath);
            Console.WriteLine();
            
            // 4. 验证输出
            Console.WriteLine("步骤 4/5: 验证输出文件");
            if (File.Exists(outputPath))
            {
                FileInfo fileInfo = new FileInfo(outputPath);
                Console.WriteLine($"  文件大小：{fileInfo.Length / 1024} KB");
                Console.WriteLine($"  创建时间：{fileInfo.CreationTime:yyyy-MM-dd HH:mm:ss}");
            }
            Console.WriteLine();
            
            Console.WriteLine("========================================");
            Console.WriteLine("PPT创建完成！");
            Console.WriteLine("========================================");
            Console.WriteLine($"输出文件：{outputPath}");
            Console.WriteLine();
            Console.WriteLine("提示：");
            Console.WriteLine("- 可以在PowerPoint中打开查看效果");
            Console.WriteLine("- 可根据需要进一步调整内容和样式");
            Console.WriteLine("- 建议在不同PowerPoint版本中预览效果");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"创建失败：{ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
    
    // 动态生成大纲
    private static PptOutline GenerateDynamicOutline(string topic, string presenter)
    {
        return new PptOutline
        {
            CoverSlide = new CoverSlide
            {
                Title = topic,
                Subtitle = "专业汇报 · 高效沟通",
                Presenter = $"汇报人：{presenter}",
                Date = DateTime.Now.ToString("yyyy年MM月dd日")
            },
            TocSlide = new TocSlide
            {
                Sections = new List<string>
                {
                    "背景与目标",
                    "主要内容",
                    "实施计划",
                    "预期成果"
                }
            },
            Sections = new List<Section>
            {
                new Section
                {
                    SectionTitle = "背景与目标",
                    Slides = new List<ContentSlide>
                    {
                        new ContentSlide
                        {
                            SlideTitle = "项目背景",
                            BulletPoints = new List<BulletPoint>
                            {
                                new BulletPoint { ContentTitle = "市场需求", Content = "当前市场对该类产品的需求日益增长" },
                                new BulletPoint { ContentTitle = "发展机遇", Content = "行业发展趋势为项目带来了新机遇" },
                                new BulletPoint { ContentTitle = "战略意义", Content = "项目符合公司战略发展方向" }
                            }
                        },
                        new ContentSlide
                        {
                            SlideTitle = "项目目标",
                            BulletPoints = new List<BulletPoint>
                            {
                                new BulletPoint { ContentTitle = "短期目标", Content = "在3个月内完成产品原型开发" },
                                new BulletPoint { ContentTitle = "中期目标", Content = "在6个月内实现产品上线" },
                                new BulletPoint { ContentTitle = "长期目标", Content = "在1年内达到市场占有率目标" }
                            }
                        }
                    }
                },
                new Section
                {
                    SectionTitle = "主要内容",
                    Slides = new List<ContentSlide>
                    {
                        new ContentSlide
                        {
                            SlideTitle = "核心功能介绍",
                            BulletPoints = new List<BulletPoint>
                            {
                                new BulletPoint { ContentTitle = "功能一", Content = "提供智能化的数据分析能力" },
                                new BulletPoint { ContentTitle = "功能二", Content = "支持多种数据格式导入导出" },
                                new BulletPoint { ContentTitle = "功能三", Content = "实现跨平台协同工作" }
                            }
                        },
                        new ContentSlide
                        {
                            SlideTitle = "市场分析",
                            BulletPoints = new List<BulletPoint>
                            {
                                new BulletPoint { ContentTitle = "市场规模", Content = "目标市场规模持续扩大" },
                                new BulletPoint { ContentTitle = "竞争格局", Content = "市场竞争态势日趋激烈" }
                            },
                            Chart = new ChartData
                            {
                                Type = "barchart",
                                Title = "市场份额分布",
                                XAxis = new List<string> { "产品A", "产品B", "产品C", "产品D", "其他" },
                                YAxis = new List<double> { 35, 25, 20, 15, 5 }
                            },
                            Summary = "我们的产品在目标市场中占据领先地位"
                        }
                    }
                },
                new Section
                {
                    SectionTitle = "实施计划",
                    Slides = new List<ContentSlide>
                    {
                        new ContentSlide
                        {
                            SlideTitle = "时间安排",
                            BulletPoints = new List<BulletPoint>
                            {
                                new BulletPoint { ContentTitle = "第一阶段", Content = "需求分析和产品设计（1-2月）" },
                                new BulletPoint { ContentTitle = "第二阶段", Content = "核心功能开发（3-5月）" },
                                new BulletPoint { ContentTitle = "第三阶段", Content = "测试优化和上线（6月）" }
                            }
                        }
                    }
                },
                new Section
                {
                    SectionTitle = "预期成果",
                    Slides = new List<ContentSlide>
                    {
                        new ContentSlide
                        {
                            SlideTitle = "成果预期",
                            BulletPoints = new List<BulletPoint>
                            {
                                new BulletPoint { ContentTitle = "产品目标", Content = "按时完成产品开发并上线运行" },
                                new BulletPoint { ContentTitle = "市场目标", Content = "获得目标用户的认可和好评" },
                                new BulletPoint { ContentTitle = "收益目标", Content = "实现预期的商业价值和收益" }
                            }
                        },
                        new ContentSlide
                        {
                            SlideTitle = "风险评估与应对",
                            BulletPoints = new List<BulletPoint>
                            {
                                new BulletPoint { ContentTitle = "技术风险", Content = "通过技术方案评审降低风险" },
                                new BulletPoint { ContentTitle = "市场风险", Content = "通过市场调研和试点验证降低风险" },
                                new BulletPoint { ContentTitle = "资源风险", Content = "通过合理配置和备用方案降低风险" }
                            }
                        }
                    }
                }
            }
        };
    }
}

// 主程序入口
public class Program
{
    public static void Main(string[] args)
    {
        Console.WriteLine("AI驱动PPT创建系统");
        Console.WriteLine();
        
        // 获取用户输入
        Console.Write("请输入PPT主题：");
        string topic = Console.ReadLine();
        
        Console.Write("请输入使用场景（如：工作报告、活动策划、年终报告）：");
        string useCase = Console.ReadLine();
        
        Console.Write("请输入汇报人姓名：");
        string presenter = Console.ReadLine();
        
        // 创建PPT
        CompletePptCreationDemo.CreateCompletePpt(topic, useCase, presenter);
    }
}
```

## 注意事项

1. **模板依赖**：需要预置PPT模板文件在templates目录中
2. **大纲结构**：必须按照指定格式生成大纲结构
3. **图表数据**：确保图表数据的X轴和Y轴数量一致
4. **字体兼容**：建议使用常用字体以确保跨平台兼容性
5. **版本兼容**：保存为PPTX 2010格式以获得最佳兼容性

## 最佳实践

1. **模板选择**：根据主题和场景选择最合适的模板
2. **内容规划**：提前规划好大纲结构和内容要点
3. **图表设计**：确保图表数据准确且有说服力
4. **风格统一**：保持整个PPT的风格和配色统一
5. **预览检查**：创建完成后在不同设备上预览效果

## 相关功能

- [基本操作](./02-basic-operations.md) - 幻灯片创建和管理
- [文本处理](./03-text-content.md) - 文本内容和格式
- [形状处理](./04-shapes-images.md) - 形状和图片
- [图表](./06-charts.md) - 图表创建和样式
- [格式转换](./11-conversion.md) - PPT格式转换
