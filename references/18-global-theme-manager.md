---
title: 全局主题与色调接管
category: spire-presentation
description: 统一PPT的色调方案和字体样式，提供预设主题和自定义主题能力
---

# 全局主题与色调接管

## 概述

全局主题与色调接管功能提供完整的演示文稿样式统一能力，包括：
- **统一色调方案**：预设6种专业主题或自定义主题色板
- **统一字体方案**：一键切换无衬线/衬线字体体系
- **统一形状样式**：圆角、阴影、线条粗细标准化
- **智能颜色替换**：背景、填充、边框、文本色自动替换
- **批量应用主题**：一键为整个演示文稿应用主题

通过主题接管，您可以快速将风格各异的PPT统一为专业的视觉风格，提升品牌一致性和演示效果。

## 预设主题库

### 沉稳科技蓝

**适用场景**：科技、互联网、IT、数据分析、技术演讲

```csharp
public static readonly ThemeConfig TechBlueTheme = new ThemeConfig
{
    Name = "沉稳科技蓝",
    Description = "专业、稳重、科技感，适合技术和商务场景",
    Colors = new ColorScheme
    {
        Primary = Color.FromArgb(30, 58, 138),      // #1E3A8A 深蓝
        Secondary = Color.FromArgb(59, 130, 246),  // #3B82F6 中蓝
        Accent = Color.FromArgb(37, 99, 235),      // #2563EB 标准蓝
        Light = Color.FromArgb(96, 165, 250),       // #60A5FA 浅蓝
        Background = Color.FromArgb(243, 244, 246), // #F3F4F6 浅灰
        Text = Color.FromArgb(31, 41, 55),          // #1F2937 深灰
        White = Color.FromArgb(255, 255, 255),      // #FFFFFF 白色
        Success = Color.FromArgb(16, 185, 129),     // #10B981 绿色
        Warning = Color.FromArgb(245, 158, 11),     // #F59E0B 橙色
        Danger = Color.FromArgb(239, 68, 68)        // #EF4444 红色
    },
    Fonts = new FontScheme
    {
        MajorLatin = new TextFont("Segoe UI"),
        MinorLatin = new TextFont("Open Sans"),
        MajorEastAsian = new TextFont("微软雅黑"),
        MinorEastAsian = new TextFont("微软雅黑")
    },
    Shapes = new ShapeStyleScheme
    {
        CornerRadius = 8,
        ShadowEnabled = true,
        ShadowOpacity = 0.15f,
        LineWidth = 1,
        LineColor = Color.FromArgb(30, 58, 138)
    }
};
```

### 简约商务灰

**适用场景**：商务、金融、咨询、管理、培训

```csharp
public static readonly ThemeConfig BusinessGrayTheme = new ThemeConfig
{
    Name = "简约商务灰",
    Description = "专业、简洁、稳重，适合商务和金融场景",
    Colors = new ColorScheme
    {
        Primary = Color.FromArgb(75, 85, 99),       // #4B5563 中灰
        Secondary = Color.FromArgb(107, 114, 128), // #6B7280 浅灰
        Accent = Color.FromArgb(156, 163, 175),    // #9CA3AF 亮灰
        Light = Color.FromArgb(209, 213, 219),    // #D1D5DB 浅浅灰
        Background = Color.FromArgb(249, 250, 251), // #F9FAFB 极浅灰
        Text = Color.FromArgb(17, 24, 39),         // #111827 深黑
        White = Color.FromArgb(255, 255, 255),
        Success = Color.FromArgb(16, 185, 129),
        Warning = Color.FromArgb(245, 158, 11),
        Danger = Color.FromArgb(239, 68, 68)
    },
    Fonts = new FontScheme
    {
        MajorLatin = new TextFont("Helvetica Neue"),
        MinorLatin = new TextFont("Arial"),
        MajorEastAsian = new TextFont("思源黑体"),
        MinorEastAsian = new TextFont("思源黑体")
    },
    Shapes = new ShapeStyleScheme
    {
        CornerRadius = 4,
        ShadowEnabled = true,
        ShadowOpacity = 0.1f,
        LineWidth = 1,
        LineColor = Color.FromArgb(75, 85, 99)
    }
};
```

### 活力活力橙

**适用场景**：创意、活动、营销、产品发布、培训

```csharp
public static readonly ThemeConfig VibrantOrangeTheme = new ThemeConfig
{
    Name = "活力活力橙",
    Description = "热情、活力、创意，适合活动和营销场景",
    Colors = new ColorScheme
    {
        Primary = Color.FromArgb(249, 115, 22),    // #F97316 亮橙
        Secondary = Color.FromArgb(251, 146, 60),  // #FB923C 浅橙
        Accent = Color.FromArgb(253, 186, 116),   // #FDBA74 浅浅橙
        Light = Color.FromArgb(254, 215, 170),   // #FED7AA 浅浅浅橙
        Background = Color.FromArgb(255, 251, 235), // #FFFBEB 暖白
        Text = Color.FromArgb(29, 78, 216),       // #1D4ED8 深蓝（对比色）
        White = Color.FromArgb(255, 255, 255),
        Success = Color.FromArgb(16, 185, 129),
        Warning = Color.FromArgb(217, 119, 6),
        Danger = Color.FromArgb(220, 38, 38)
    },
    Fonts = new FontScheme
    {
        MajorLatin = new TextFont("Verdana"),
        MinorLatin = new TextFont("Tahoma"),
        MajorEastAsian = new TextFont("华文黑体"),
        MinorEastAsian = new TextFont("华文黑体")
    },
    Shapes = new ShapeStyleScheme
    {
        CornerRadius = 12,
        ShadowEnabled = true,
        ShadowOpacity = 0.2f,
        LineWidth = 2,
        LineColor = Color.FromArgb(249, 115, 22)
    }
};
```

### 自然生态绿

**适用场景**：环保、健康、教育、可持续发展、医疗

```csharp
public static readonly ThemeConfig EcoGreenTheme = new ThemeConfig
{
    Name = "自然生态绿",
    Description = "清新、健康、环保，适合环保和教育场景",
    Colors = new ColorScheme
    {
        Primary = Color.FromArgb(16, 185, 129),   // #10B981 翠绿
        Secondary = Color.FromArgb(52, 211, 153),  // #34D399 浅绿
        Accent = Color.FromArgb(110, 231, 183),   // #6EE7B7 浅浅绿
        Light = Color.FromArgb(187, 247, 208),    // #BBF7D0 浅浅浅绿
        Background = Color.FromArgb(240, 253, 244), // #F0FDF4 极浅绿
        Text = Color.FromArgb(6, 78, 59),          // #064E3A 深绿
        White = Color.FromArgb(255, 255, 255),
        Success = Color.FromArgb(5, 150, 105),
        Warning = Color.FromArgb(245, 158, 11),
        Danger = Color.FromArgb(220, 38, 38)
    },
    Fonts = new FontScheme
    {
        MajorLatin = new TextFont("Roboto"),
        MinorLatin = new TextFont("Segoe UI"),
        MajorEastAsian = new TextFont("苹方"),
        MinorEastAsian = new TextFont("苹方")
    },
    Shapes = new ShapeStyleScheme
    {
        CornerRadius = 10,
        ShadowEnabled = true,
        ShadowOpacity = 0.15f,
        LineWidth = 1,
        LineColor = Color.FromArgb(16, 185, 129)
    }
};
```

### 优雅紫罗兰

**适用场景**：奢侈、艺术、时尚、创意、高端品牌

```csharp
public static readonly ThemeConfig ElegantPurpleTheme = new ThemeConfig
{
    Name = "优雅紫罗兰",
    Description = "优雅、高端、艺术，适合奢侈和时尚场景",
    Colors = new ColorScheme
    {
        Primary = Color.FromArgb(124, 58, 237),   // #7C3AED 深紫
        Secondary = Color.FromArgb(139, 92, 246),  // #8B5CF6 中紫
        Accent = Color.FromArgb(167, 139, 250),   // #A78BFA 浅紫
        Light = Color.FromArgb(196, 181, 253),    // #C4B5FD 浅浅紫
        Background = Color.FromArgb(245, 243, 255), // #F5F3FF 极浅紫
        Text = Color.FromArgb(55, 48, 163),        // #3730A3 深深紫
        White = Color.FromArgb(255, 255, 255),
        Success = Color.FromArgb(16, 185, 129),
        Warning = Color.FromArgb(245, 158, 11),
        Danger = Color.FromArgb(220, 38, 38)
    },
    Fonts = new FontScheme
    {
        MajorLatin = new TextFont("Arial"),
        MinorLatin = new TextFont("Georgia"),
        MajorEastAsian = new TextFont("兰亭黑"),
        MinorEastAsian = new TextFont("兰亭黑")
    },
    Shapes = new ShapeStyleScheme
    {
        CornerRadius = 15,
        ShadowEnabled = true,
        ShadowOpacity = 0.2f,
        LineWidth = 1,
        LineColor = Color.FromArgb(124, 58, 237)
    }
};
```

### 经典深色模式

**适用场景**：技术、演示、夜间、游戏、开发者

```csharp
public static readonly ThemeConfig DarkModeTheme = new ThemeConfig
{
    Name = "经典深色模式",
    Description = "专业、护眼、现代，适合技术和演示场景",
    Colors = new ColorScheme
    {
        Primary = Color.FromArgb(59, 130, 246),   // #3B82F6 蓝色
        Secondary = Color.FromArgb(75, 85, 99),    // #4B5563 灰色
        Accent = Color.FromArgb(99, 102, 241),    // #6366F1 靛蓝
        Light = Color.FromArgb(156, 163, 175),    // #9CA3AF 浅灰
        Background = Color.FromArgb(17, 24, 39),  // #111827 深黑
        Text = Color.FromArgb(243, 244, 246),    // #F3F4F6 浅灰
        White = Color.FromArgb(255, 255, 255),
        Success = Color.FromArgb(16, 185, 129),
        Warning = Color.FromArgb(245, 158, 11),
        Danger = Color.FromArgb(239, 68, 68)
    },
    Fonts = new FontScheme
    {
        MajorLatin = new TextFont("Segoe UI"),
        MinorLatin = new TextFont("Consolas"),
        MajorEastAsian = new TextFont("微软雅黑"),
        MinorEastAsian = new TextFont("微软雅黑")
    },
    Shapes = new ShapeStyleScheme
    {
        CornerRadius = 6,
        ShadowEnabled = true,
        ShadowOpacity = 0.25f,
        LineWidth = 1,
        LineColor = Color.FromArgb(75, 85, 99)
    }
};
```

## 主题色调定义

### 色调数据结构

```csharp
public class ThemeConfig
{
    public string Name { get; set; }
    public string Description { get; set; }
    public ColorScheme Colors { get; set; }
    public FontScheme Fonts { get; set; }
    public ShapeStyleScheme Shapes { get; set; }
}

public class ColorScheme
{
    public Color Primary { get; set; }      // 主色
    public Color Secondary { get; set; }    // 辅色
    public Color Accent { get; set; }       // 强调色
    public Color Light { get; set; }        // 浅色
    public Color Background { get; set; }   // 背景色
    public Color Text { get; set; }         // 文本色
    public Color White { get; set; }        // 白色
    public Color Success { get; set; }      // 成功色
    public Color Warning { get; set; }      // 警告色
    public Color Danger { get; set; }       // 危险色
}

public class FontScheme
{
    public TextFont MajorLatin { get; set; }      // 西文主要字体
    public TextFont MinorLatin { get; set; }      // 西文次要字体
    public TextFont MajorEastAsian { get; set; }  // 东亚主要字体
    public TextFont MinorEastAsian { get; set; }  // 东亚次要字体
}

public class ShapeStyleScheme
{
    public int CornerRadius { get; set; }         // 圆角半径
    public bool ShadowEnabled { get; set; }       // 是否启用阴影
    public float ShadowOpacity { get; set; }      // 阴影透明度
    public float LineWidth { get; set; }          // 线条粗细
    public Color LineColor { get; set; }         // 线条颜色
}
```

### 智能配色生成

```csharp
public class ColorPaletteGenerator
{
    // 根据主色调生成完整色板
    public static ColorScheme GeneratePalette(Color primaryColor)
    {
        return new ColorScheme
        {
            Primary = primaryColor,
            Secondary = AdjustBrightness(primaryColor, 0.2f),
            Accent = AdjustBrightness(primaryColor, 0.4f),
            Light = AdjustBrightness(primaryColor, 0.6f),
            Background = GetComplementaryBackground(primaryColor),
            Text = GetContrastTextColor(primaryColor),
            White = Color.FromArgb(255, 255, 255),
            Success = Color.FromArgb(16, 185, 129),
            Warning = Color.FromArgb(245, 158, 11),
            Danger = Color.FromArgb(239, 68, 68)
        };
    }

    // 调整亮度
    private static Color AdjustBrightness(Color color, float factor)
    {
        int r = Math.Min(255, Math.Max(0, (int)(color.R * (1 + factor))));
        int g = Math.Min(255, Math.Max(0, (int)(color.G * (1 + factor))));
        int b = Math.Min(255, Math.Max(0, (int)(color.B * (1 + factor))));
        return Color.FromArgb(r, g, b);
    }

    // 获取互补背景色
    private static Color GetComplementaryBackground(Color primaryColor)
    {
        int brightness = (primaryColor.R * 299 + primaryColor.G * 587 + primaryColor.B * 114) / 1000;
        return brightness > 128
            ? Color.FromArgb(17, 24, 39)    // 深色背景
            : Color.FromArgb(249, 250, 251); // 浅色背景
    }

    // 获取对比文本色
    private static Color GetContrastTextColor(Color backgroundColor)
    {
        int brightness = (backgroundColor.R * 299 + backgroundColor.G * 587 + backgroundColor.B * 114) / 1000;
        return brightness > 128
            ? Color.FromArgb(17, 24, 39)    // 深色文本
            : Color.FromArgb(243, 244, 246); // 浅色文本
    }
}
```

## 字体统一方案

### 无衬线字体方案

```csharp
public class FontSchemePresets
{
    // 无衬线字体方案
    public static FontScheme SansSerifScheme = new FontScheme
    {
        MajorLatin = new TextFont("Segoe UI"),
        MinorLatin = new TextFont("Open Sans"),
        MajorEastAsian = new TextFont("微软雅黑"),
        MinorEastAsian = new TextFont("思源黑体")
    };

    // 衬线字体方案
    public static FontScheme SerifScheme = new FontScheme
    {
        MajorLatin = new TextFont("Georgia"),
        MinorLatin = new TextFont("Times New Roman"),
        MajorEastAsian = new TextFont("宋体"),
        MinorEastAsian = new TextFont("宋体")
    };

    // 等宽字体方案
    public static FontScheme MonospaceScheme = new FontScheme
    {
        MajorLatin = new TextFont("Consolas"),
        MinorLatin = new TextFont("Courier New"),
        MajorEastAsian = new TextFont("微软雅黑"),
        MinorEastAsian = new TextFont("思源黑体")
    };
}
```

### 字体应用逻辑

```csharp
public class FontApplier
{
    // 应用字体方案
    public static void ApplyFontScheme(Presentation presentation, FontScheme scheme)
    {
        // 应用到母版
        ApplyToMasters(presentation, scheme);

        // 应用到所有幻灯片
        foreach (ISlide slide in presentation.Slides)
        {
            ApplyToSlide(slide, scheme);
        }
    }

    // 应用到母版
    private static void ApplyToMasters(Presentation presentation, FontScheme scheme)
    {
        foreach (IMasterSlide master in presentation.Masters)
        {
            // 更新主题字体
            master.Theme.MajorFont.LatinFont = scheme.MajorLatin;
            master.Theme.MinorFont.LatinFont = scheme.MinorLatin;
            master.Theme.MajorFont.EastAsianFont = scheme.MajorEastAsian;
            master.Theme.MinorFont.EastAsianFont = scheme.MinorEastAsian;

            // 应用到母版中的形状
            ApplyToShapes(master.Shapes, scheme);
        }
    }

    // 应用到幻灯片
    private static void ApplyToSlide(ISlide slide, FontScheme scheme)
    {
        ApplyToShapes(slide.Shapes, scheme);
    }

    // 应用到形状
    private static void ApplyToShapes(ShapeCollection shapes, FontScheme scheme)
    {
        foreach (IShape shape in shapes)
        {
            if (shape is IAutoShape autoShape && autoShape.TextFrame != null)
            {
                ApplyToTextFrame(autoShape.TextFrame, scheme);
            }
            else if (shape is IChart chart)
            {
                ApplyToChart(chart, scheme);
            }
            else if (shape is ITable table)
            {
                ApplyToTable(table, scheme);
            }
            else if (shape is ISmartArt smartArt)
            {
                ApplyToSmartArt(smartArt, scheme);
            }
        }
    }

    // 应用到文本框
    private static void ApplyToTextFrame(ITextFrame textFrame, FontScheme scheme)
    {
        foreach (TextParagraph para in textFrame.Paragraphs)
        {
            foreach (TextRange range in para.TextRanges)
            {
                // 根据字号决定使用主要或次要字体
                if (range.FontHeight > 18)
                {
                    range.LatinFont = scheme.MajorLatin;
                    range.EastAsianFont = scheme.MajorEastAsian;
                }
                else
                {
                    range.LatinFont = scheme.MinorLatin;
                    range.EastAsianFont = scheme.MinorEastAsian;
                }
            }
        }
    }

    // 应用到图表
    private static void ApplyToChart(IChart chart, FontScheme scheme)
    {
        // 图表标题
        chart.ChartTitle.TextProperties.Paragraphs[0].TextRanges[0].LatinFont = scheme.MajorLatin;
        chart.ChartTitle.TextProperties.Paragraphs[0].TextRanges[0].EastAsianFont = scheme.MajorEastAsian;

        // 图例
        if (chart.HasLegend)
        {
            chart.Legend.TextProperties.Paragraphs[0].TextRanges[0].LatinFont = scheme.MinorLatin;
            chart.Legend.TextProperties.Paragraphs[0].TextRanges[0].EastAsianFont = scheme.MinorEastAsian;
        }

        // 坐标轴
        foreach (IChartAxis axis in chart.Axes)
        {
            foreach (TextRange range in axis.TextProperties.Paragraphs[0].TextRanges)
            {
                range.LatinFont = scheme.MinorLatin;
                range.EastAsianFont = scheme.MinorEastAsian;
            }
        }
    }

    // 应用到表格
    private static void ApplyToTable(ITable table, FontScheme scheme)
    {
        for (int row = 0; row < table.Rows.Count; row++)
        {
            for (int col = 0; col < table.Columns.Count; col++)
            {
                ICell cell = table[row, col];
                foreach (TextRange range in cell.TextFrame.Paragraphs[0].TextRanges)
                {
                    if (row == 0) // 表头使用主要字体
                    {
                        range.LatinFont = scheme.MajorLatin;
                        range.EastAsianFont = scheme.MajorEastAsian;
                    }
                    else // 正文使用次要字体
                    {
                        range.LatinFont = scheme.MinorLatin;
                        range.EastAsianFont = scheme.MinorEastAsian;
                    }
                }
            }
        }
    }

    // 应用到SmartArt
    private static void ApplyToSmartArt(ISmartArt smartArt, FontScheme scheme)
    {
        foreach (ISmartArtNode node in smartArt.Nodes)
        {
            foreach (TextRange range in node.TextFrame.Paragraphs[0].TextRanges)
            {
                range.LatinFont = scheme.MajorLatin;
                range.EastAsianFont = scheme.MajorEastAsian;
            }

            // 递归处理子节点
            ApplyToSmartArtNodes(node.ChildNodes, scheme);
        }
    }

    private static void ApplyToSmartArtNodes(ISmartArtNodeCollection nodes, FontScheme scheme)
    {
        foreach (ISmartArtNode node in nodes)
        {
            foreach (TextRange range in node.TextFrame.Paragraphs[0].TextRanges)
            {
                range.LatinFont = scheme.MajorLatin;
                range.EastAsianFont = scheme.MajorEastAsian;
            }

            if (node.ChildNodes.Count > 0)
            {
                ApplyToSmartArtNodes(node.ChildNodes, scheme);
            }
        }
    }
}
```

## 全局颜色应用

### 颜色替换核心类

```csharp
public class ColorApplier
{
    // 应用颜色方案
    public static void ApplyColorScheme(Presentation presentation, ColorScheme colors)
    {
        // 应用背景色
        ApplyBackground(presentation, colors.Background);

        // 应用到所有幻灯片
        foreach (ISlide slide in presentation.Slides)
        {
            ApplyToSlide(slide, colors);
        }

        // 应用到主题
        ApplyToTheme(presentation, colors);
    }

    // 应用背景色
    private static void ApplyBackground(Presentation presentation, Color backgroundColor)
    {
        foreach (ISlide slide in presentation.Slides)
        {
            slide.Background.Type = BackgroundType.Custom;
            slide.Background.FillFormat.FillType = FillFormatType.Solid;
            slide.Background.FillFormat.SolidFillColor.Color = backgroundColor;
        }
    }

    // 应用到幻灯片
    private static void ApplyToSlide(ISlide slide, ColorScheme colors)
    {
        ApplyToShapes(slide.Shapes, colors);
    }

    // 应用到形状
    private static void ApplyToShapes(ShapeCollection shapes, ColorScheme colors)
    {
        foreach (IShape shape in shapes)
        {
            if (shape is IAutoShape autoShape)
            {
                ApplyToAutoShape(autoShape, colors);
            }
            else if (shape is IChart chart)
            {
                ApplyToChart(chart, colors);
            }
            else if (shape is ITable table)
            {
                ApplyToTable(table, colors);
            }
            else if (shape is ISmartArt smartArt)
            {
                ApplyToSmartArt(smartArt, colors);
            }
        }
    }

    // 应用到自选图形
    private static void ApplyToAutoShape(IAutoShape shape, ColorScheme colors)
    {
        // 填充色
        if (shape.Fill.FillType == FillFormatType.Solid)
        {
            shape.Fill.SolidColor.Color = colors.Primary;
        }

        // 线条色
        shape.ShapeStyle.LineColor.Color = colors.LineColor ?? colors.Secondary;

        // 文本色
        if (shape.TextFrame != null)
        {
            foreach (TextParagraph para in shape.TextFrame.Paragraphs)
            {
                foreach (TextRange range in para.TextRanges)
                {
                    range.Fill.FillType = FillFormatType.Solid;
                    range.Fill.SolidColor.Color = colors.Text;
                }
            }
        }
    }

    // 应用到图表
    private static void ApplyToChart(IChart chart, ColorScheme colors)
    {
        // 图表区域填充
        chart.ChartArea.Fill.FillType = FillFormatType.Solid;
        chart.ChartArea.Fill.SolidColor.Color = colors.Background;

        // 图表系列颜色
        int seriesIndex = 0;
        Color[] seriesColors = new Color[] {
            colors.Primary, colors.Secondary, colors.Accent,
            colors.Light, colors.Success, colors.Warning
        };

        foreach (IChartSeries series in chart.Series)
        {
            Color seriesColor = seriesColors[seriesIndex % seriesColors.Length];
            series.Format.Fill.FillType = FillFormatType.Solid;
            series.Format.Fill.SolidColor.Color = seriesColor;
            seriesIndex++;
        }

        // 图表文本色
        chart.ChartTitle.TextProperties.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = colors.Text;

        if (chart.HasLegend)
        {
            chart.Legend.TextProperties.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = colors.Text;
        }

        foreach (IChartAxis axis in chart.Axes)
        {
            foreach (TextRange range in axis.TextProperties.Paragraphs[0].TextRanges)
            {
                range.Fill.SolidColor.Color = colors.Text;
            }
        }
    }

    // 应用到表格
    private static void ApplyToTable(ITable table, ColorScheme colors)
    {
        for (int row = 0; row < table.Rows.Count; row++)
        {
            for (int col = 0; col < table.Columns.Count; col++)
            {
                ICell cell = table[row, col];

                // 表头使用主色
                if (row == 0)
                {
                    cell.FillFormat.FillType = FillFormatType.Solid;
                    cell.FillFormat.SolidFillColor.Color = colors.Primary;
                    cell.BorderTop.FillFormat.FillType = FillFormatType.Solid;
                    cell.BorderTop.FillFormat.SolidFillColor.Color = colors.Light;
                    cell.BorderBottom.FillFormat.FillType = FillFormatType.Solid;
                    cell.BorderBottom.FillFormat.SolidFillColor.Color = colors.Light;
                    cell.BorderLeft.FillFormat.FillType = FillFormatType.Solid;
                    cell.BorderLeft.FillFormat.SolidFillColor.Color = colors.Light;
                    cell.BorderRight.FillFormat.FillType = FillFormatType.Solid;
                    cell.BorderRight.FillFormat.SolidFillColor.Color = colors.Light;

                    // 表头文本色为白色
                    foreach (TextRange range in cell.TextFrame.Paragraphs[0].TextRanges)
                    {
                        range.Fill.FillType = FillFormatType.Solid;
                        range.Fill.SolidColor.Color = colors.White;
                    }
                }
                else
                {
                    // 正文使用浅色背景
                    cell.FillFormat.FillType = FillFormatType.Solid;
                    cell.FillFormat.SolidFillColor.Color = (row % 2 == 0) ? colors.Background : colors.Light;

                    // 文本色
                    foreach (TextRange range in cell.TextFrame.Paragraphs[0].TextRanges)
                    {
                        range.Fill.FillType = FillFormatType.Solid;
                        range.Fill.SolidColor.Color = colors.Text;
                    }
                }
            }
        }
    }

    // 应用到SmartArt
    private static void ApplyToSmartArt(ISmartArt smartArt, ColorScheme colors)
    {
        smartArt.ColorStyle = SmartArtColorType.Accented1;

        // 应用主题色
        foreach (ISmartArtNode node in smartArt.Nodes)
        {
            node.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid;
            node.TextFrame.TextRange.Fill.SolidColor.Color = colors.White;

            if (node.ChildNodes.Count > 0)
            {
                ApplyToSmartArtNodes(node.ChildNodes, colors);
            }
        }
    }

    private static void ApplyToSmartArtNodes(ISmartArtNodeCollection nodes, ColorScheme colors)
    {
        foreach (ISmartArtNode node in nodes)
        {
            node.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid;
            node.TextFrame.TextRange.Fill.SolidColor.Color = colors.White;

            if (node.ChildNodes.Count > 0)
            {
                ApplyToSmartArtNodes(node.ChildNodes, colors);
            }
        }
    }

    // 应用到主题
    private static void ApplyToTheme(Presentation presentation, ColorScheme colors)
    {
        foreach (IMasterSlide master in presentation.Masters)
        {
            master.Theme.ColorScheme[SchemeColor.Dark1] = colors.Background;
            master.Theme.ColorScheme[SchemeColor.Light1] = colors.Text;
            master.Theme.ColorScheme[SchemeColor.Accent1] = colors.Primary;
            master.Theme.ColorScheme[SchemeColor.Accent2] = colors.Secondary;
            master.Theme.ColorScheme[SchemeColor.Accent3] = colors.Accent;
            master.Theme.ColorScheme[SchemeColor.Accent4] = colors.Light;
        }
    }
}
```

## 形状样式统一

### 形状样式应用

```csharp
public class ShapeStyleApplier
{
    // 应用形状样式方案
    public static void ApplyShapeStyleScheme(Presentation presentation, ShapeStyleScheme shapes)
    {
        foreach (ISlide slide in presentation.Slides)
        {
            ApplyToSlide(slide, shapes);
        }
    }

    // 应用到幻灯片
    private static void ApplyToSlide(ISlide slide, ShapeStyleScheme shapes)
    {
        foreach (IShape shape in slide.Shapes)
        {
            if (shape is IAutoShape autoShape)
            {
                ApplyToAutoShape(autoShape, shapes);
            }
        }
    }

    // 应用到自选图形
    private static void ApplyToAutoShape(IAutoShape shape, ShapeStyleScheme shapes)
    {
        // 应用圆角（仅限矩形和圆角矩形）
        if (shape.ShapeType == ShapeType.Rectangle ||
            shape.ShapeType == ShapeType.RoundedRectangle)
        {
            // 设置圆角
            // 注意：Spire.Presentation的圆角设置可能需要通过其他方式实现
            // 这里是示例代码
        }

        // 应用阴影
        if (shapes.ShadowEnabled)
        {
            shape.EffectDag.EnableOuterShadowEffect = true;
            // 设置阴影透明度
            // 注意：具体API可能需要调整
        }

        // 应用边框
        shape.LineWidth = shapes.LineWidth;
        if (shape.ShapeStyle.LineColor != null)
        {
            shape.ShapeStyle.LineColor.Color = shapes.LineColor;
        }
    }
}
```

## 批量应用主题

### 示例 1: 应用预设主题

```csharp
using Spire.Presentation;

public class ThemeManager
{
    // 应用预设主题
    public static void ApplyPresetTheme(string inputFile, string outputFile, ThemeType themeType)
    {
        Presentation presentation = new Presentation();
        presentation.LoadFromFile(inputFile);

        // 获取主题配置
        ThemeConfig theme = GetThemeByType(themeType);

        Console.WriteLine($"正在应用主题: {theme.Name}");
        Console.WriteLine($"描述: {theme.Description}");

        // 应用颜色方案
        Console.WriteLine("正在应用颜色方案...");
        ColorApplier.ApplyColorScheme(presentation, theme.Colors);

        // 应用字体方案
        Console.WriteLine("正在应用字体方案...");
        FontApplier.ApplyFontScheme(presentation, theme.Fonts);

        // 应用形状样式
        Console.WriteLine("正在应用形状样式...");
        ShapeStyleApplier.ApplyShapeStyleScheme(presentation, theme.Shapes);

        // 保存
        presentation.SaveToFile(outputFile, FileFormat.Pptx2010);
        presentation.Dispose();

        Console.WriteLine($"主题应用完成！输出文件: {outputFile}");
    }

    // 根据类型获取主题
    private static ThemeConfig GetThemeByType(ThemeType type)
    {
        switch (type)
        {
            case ThemeType.TechBlue:
                return TechBlueTheme;
            case ThemeType.BusinessGray:
                return BusinessGrayTheme;
            case ThemeType.VibrantOrange:
                return VibrantOrangeTheme;
            case ThemeType.EcoGreen:
                return EcoGreenTheme;
            case ThemeType.ElegantPurple:
                return ElegantPurpleTheme;
            case ThemeType.DarkMode:
                return DarkModeTheme;
            default:
                return TechBlueTheme;
        }
    }
}

public enum ThemeType
{
    TechBlue,
    BusinessGray,
    VibrantOrange,
    EcoGreen,
    ElegantPurple,
    DarkMode
}
```

### 示例 2: 使用示例

```csharp
// 应用沉稳科技蓝主题
ThemeManager.ApplyPresetTheme(
    "presentation.pptx",
    "presentation_with_theme.pptx",
    ThemeType.TechBlue
);

// 应用简约商务灰主题
ThemeManager.ApplyPresetTheme(
    "presentation.pptx",
    "presentation_with_theme.pptx",
    ThemeType.BusinessGray
);
```

### 示例 3: 批量应用主题

```csharp
using System.IO;
using Spire.Presentation;

public class BatchThemeApplier
{
    // 批量为多个PPT应用主题
    public static void ApplyThemeToBatch(
        string inputFolder,
        string outputFolder,
        ThemeType themeType)
    {
        // 确保输出目录存在
        Directory.CreateDirectory(outputFolder);

        // 获取所有PPT文件
        string[] files = Directory.GetFiles(inputFolder, "*.pptx");

        Console.WriteLine($"找到 {files.Length} 个PPT文件");

        foreach (string file in files)
        {
            string fileName = Path.GetFileName(file);
            string outputFile = Path.Combine(outputFolder, fileName);

            Console.WriteLine($"正在处理: {fileName}");

            try
            {
                ThemeManager.ApplyPresetTheme(file, outputFile, themeType);
                Console.WriteLine($"✓ 完成: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 失败: {fileName} - {ex.Message}");
            }
        }

        Console.WriteLine($"批量处理完成！输出目录: {outputFolder}");
    }
}

// 使用示例
BatchThemeApplier.ApplyThemeToBatch(
    @"C:\Presentations\Input",
    @"C:\Presentations\Output",
    ThemeType.TechBlue
);
```

## 自定义主题创建

### 示例 4: 创建自定义主题

```csharp
public class CustomThemeCreator
{
    // 创建自定义主题
    public static ThemeConfig CreateCustomTheme(
        string themeName,
        Color primaryColor,
        FontScheme fontScheme,
        ShapeStyleScheme shapeScheme = null)
    {
        ThemeConfig theme = new ThemeConfig
        {
            Name = themeName,
            Description = $"自定义主题 - {themeName}",
            Colors = ColorPaletteGenerator.GeneratePalette(primaryColor),
            Fonts = fontScheme,
            Shapes = shapeScheme ?? new ShapeStyleScheme
            {
                CornerRadius = 8,
                ShadowEnabled = true,
                ShadowOpacity = 0.15f,
                LineWidth = 1,
                LineColor = primaryColor
            }
        };

        return theme;
    }

    // 保存自定义主题
    public static void SaveTheme(ThemeConfig theme, string filePath)
    {
        string json = System.Text.Json.JsonSerializer.Serialize(theme, new System.Text.Json.JsonSerializerOptions
        {
            WriteIndented = true
        });
        File.WriteAllText(filePath, json);
    }

    // 加载自定义主题
    public static ThemeConfig LoadTheme(string filePath)
    {
        string json = File.ReadAllText(filePath);
        return System.Text.Json.JsonSerializer.Deserialize<ThemeConfig>(json);
    }
}

// 使用示例
// 创建自定义主题
ThemeConfig myTheme = CustomThemeCreator.CreateCustomTheme(
    "我的公司主题",
    Color.FromArgb(234, 88, 12), // 橙红色
    FontSchemePresets.SansSerifScheme
);

// 保存主题
CustomThemeCreator.SaveTheme(myTheme, "my_theme.json");

// 应用自定义主题
Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

ColorApplier.ApplyColorScheme(presentation, myTheme.Colors);
FontApplier.ApplyFontScheme(presentation, myTheme.Fonts);
ShapeStyleApplier.ApplyShapeStyleScheme(presentation, myTheme.Shapes);

presentation.SaveToFile("presentation_with_custom_theme.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 5: 基于公司品牌色创建主题

```csharp
public class BrandThemeCreator
{
    // 根据公司品牌色创建主题
    public static ThemeConfig CreateFromBrandColor(
        string companyName,
        Color brandColor)
    {
        ThemeConfig theme = new ThemeConfig
        {
            Name = $"{companyName}品牌主题",
            Description = $"基于{companyName}品牌色创建的专业主题",
            Colors = ColorPaletteGenerator.GeneratePalette(brandColor),
            Fonts = FontSchemePresets.SansSerifScheme,
            Shapes = new ShapeStyleScheme
            {
                CornerRadius = 8,
                ShadowEnabled = true,
                ShadowOpacity = 0.15f,
                LineWidth = 1,
                LineColor = brandColor
            }
        };

        return theme;
    }
}

// 使用示例
Color brandColor = Color.FromArgb(59, 130, 246); // 公司品牌色
ThemeConfig companyTheme = BrandThemeCreator.CreateFromBrandColor("我的公司", brandColor);
```

## 高级样式覆盖

### 示例 6: 保留特定样式

```csharp
public class SelectiveThemeApplier
{
    // 选择性应用主题
    public static void ApplyThemeWithExceptions(
        Presentation presentation,
        ThemeConfig theme,
        List<int> excludeSlides = null,
        List<string> excludeShapes = null)
    {
        excludeSlides = excludeSlides ?? new List<int>();
        excludeShapes = excludeShapes ?? new List<string>();

        for (int i = 0; i < presentation.Slides.Count; i++)
        {
            ISlide slide = presentation.Slides[i];

            // 跳过排除的幻灯片
            if (excludeSlides.Contains(i))
            {
                Console.WriteLine($"跳过幻灯片 {i + 1}");
                continue;
            }

            Console.WriteLine($"处理幻灯片 {i + 1}");

            foreach (IShape shape in slide.Shapes)
            {
                // 跳过排除的形状（通过名称）
                if (excludeShapes.Contains(shape.Name))
                {
                    Console.WriteLine($"  跳过形状: {shape.Name}");
                    continue;
                }

                // 应用样式
                if (shape is IAutoShape autoShape)
                {
                    ApplyToAutoShapeSelective(autoShape, theme);
                }
                else if (shape is IChart chart)
                {
                    ApplyToChartSelective(chart, theme);
                }
                // ... 其他形状类型
            }
        }
    }

    // 选择性应用到自选图形
    private static void ApplyToAutoShapeSelective(IAutoShape shape, ThemeConfig theme)
    {
        // 只修改颜色，不修改形状和效果
        if (shape.Fill.FillType == FillFormatType.Solid)
        {
            shape.Fill.SolidColor.Color = theme.Colors.Primary;
        }

        shape.ShapeStyle.LineColor.Color = theme.Colors.Secondary;
    }

    // 选择性应用到图表
    private static void ApplyToChartSelective(IChart chart, ThemeConfig theme)
    {
        // 只修改系列颜色，不修改图表类型和布局
        Color[] seriesColors = new Color[] {
            theme.Colors.Primary, theme.Colors.Secondary, theme.Colors.Accent
        };

        int seriesIndex = 0;
        foreach (IChartSeries series in chart.Series)
        {
            Color seriesColor = seriesColors[seriesIndex % seriesColors.Length];
            series.Format.Fill.FillType = FillFormatType.Solid;
            series.Format.Fill.SolidColor.Color = seriesColor;
            seriesIndex++;
        }
    }
}

// 使用示例
Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 应用主题，但排除第一张封面页和特定的logo形状
SelectiveThemeApplier.ApplyThemeWithExceptions(
    presentation,
    TechBlueTheme,
    new List<int> { 0 },  // 排除封面
    new List<string> { "LogoShape", "Watermark" }  // 排除特定形状
);

presentation.SaveToFile("presentation_selective_theme.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 完整示例

### 示例 7: 端到端主题应用

```csharp
using Spire.Presentation;
using System;

public class CompleteThemeManager
{
    public static void ApplyThemeComplete(
        string inputFile,
        string outputFile,
        ThemeType themeType,
        bool includeShapes = true,
        bool includeBackground = true)
    {
        Presentation presentation = new Presentation();
        presentation.LoadFromFile(inputFile);

        Console.WriteLine("========================================");
        Console.WriteLine("全局主题管理器");
        Console.WriteLine("========================================");
        Console.WriteLine();

        Console.WriteLine($"输入文件: {inputFile}");
        Console.WriteLine($"幻灯片总数: {presentation.Slides.Count}");
        Console.WriteLine();

        // 获取主题配置
        ThemeConfig theme = GetThemeByType(themeType);
        Console.WriteLine($"应用主题: {theme.Name}");
        Console.WriteLine($"主题描述: {theme.Description}");
        Console.WriteLine();

        // 1. 应用颜色方案
        Console.WriteLine("步骤 1/3: 应用颜色方案");
        Console.WriteLine($"  主色: RGB({theme.Colors.Primary.R}, {theme.Colors.Primary.G}, {theme.Colors.Primary.B})");
        Console.WriteLine($"  背景色: RGB({theme.Colors.Background.R}, {theme.Colors.Background.G}, {theme.Colors.Background.B})");
        Console.WriteLine($"  文本色: RGB({theme.Colors.Text.R}, {theme.Colors.Text.G}, {theme.Colors.Text.B})");

        ColorApplier.ApplyColorScheme(presentation, theme.Colors);
        Console.WriteLine("  ✓ 颜色方案应用完成");
        Console.WriteLine();

        // 2. 应用字体方案
        Console.WriteLine("步骤 2/3: 应用字体方案");
        Console.WriteLine($"  西文主要字体: {theme.Fonts.MajorLatin.FontName}");
        Console.WriteLine($"  西文次要字体: {theme.Fonts.MinorLatin.FontName}");
        Console.WriteLine($"  东亚主要字体: {theme.Fonts.MajorEastAsian.FontName}");
        Console.WriteLine($"  东亚次要字体: {theme.Fonts.MinorEastAsian.FontName}");

        FontApplier.ApplyFontScheme(presentation, theme.Fonts);
        Console.WriteLine("  ✓ 字体方案应用完成");
        Console.WriteLine();

        // 3. 应用形状样式（可选）
        if (includeShapes)
        {
            Console.WriteLine("步骤 3/3: 应用形状样式");
            Console.WriteLine($"  圆角半径: {theme.Shapes.CornerRadius}");
            Console.WriteLine($"  阴影: {(theme.Shapes.ShadowEnabled ? "启用" : "禁用")}");
            Console.WriteLine($"  线条粗细: {theme.Shapes.LineWidth}");

            ShapeStyleApplier.ApplyShapeStyleScheme(presentation, theme.Shapes);
            Console.WriteLine("  ✓ 形状样式应用完成");
        }

        // 保存
        presentation.SaveToFile(outputFile, FileFormat.Pptx2010);
        presentation.Dispose();

        Console.WriteLine();
        Console.WriteLine("========================================");
        Console.WriteLine("主题应用完成！");
        Console.WriteLine("========================================");
        Console.WriteLine($"输出文件: {outputFile}");
        Console.WriteLine($"处理时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        Console.WriteLine();
        Console.WriteLine("提示:");
        Console.WriteLine("- 如需进一步调整，请手动编辑PPT");
        Console.WriteLine("- 建议在不同版本的PowerPoint中预览效果");
        Console.WriteLine("- 主题效果可能因PPT原有样式而略有差异");
    }

    private static ThemeConfig GetThemeByType(ThemeType type)
    {
        switch (type)
        {
            case ThemeType.TechBlue:
                return TechBlueTheme;
            case ThemeType.BusinessGray:
                return BusinessGrayTheme;
            case ThemeType.VibrantOrange:
                return VibrantOrangeTheme;
            case ThemeType.EcoGreen:
                return EcoGreenTheme;
            case ThemeType.ElegantPurple:
                return ElegantPurpleTheme;
            case ThemeType.DarkMode:
                return DarkModeTheme;
            default:
                return TechBlueTheme;
        }
    }
}

// 主程序
public class Program
{
    public static void Main(string[] args)
    {
        string inputFile = "presentation.pptx";
        string outputFile = "presentation_with_theme.pptx";
        ThemeType themeType = ThemeType.TechBlue;

        Console.WriteLine("全局主题管理器");
        Console.WriteLine();

        // 选择主题
        Console.WriteLine("请选择主题:");
        Console.WriteLine("1. 沉稳科技蓝");
        Console.WriteLine("2. 简约商务灰");
        Console.WriteLine("3. 活力活力橙");
        Console.WriteLine("4. 自然生态绿");
        Console.WriteLine("5. 优雅紫罗兰");
        Console.WriteLine("6. 经典深色模式");
        Console.Write("请选择 (1-6): ");

        string choice = Console.ReadLine();
        if (int.TryParse(choice, out int themeChoice) && themeChoice >= 1 && themeChoice <= 6)
        {
            themeType = (ThemeType)(themeChoice - 1);
        }

        Console.WriteLine();
        Console.Write("是否应用形状样式? (Y/N): ");
        bool includeShapes = Console.ReadLine().Trim().ToUpper() == "Y";

        Console.WriteLine();
        CompleteThemeManager.ApplyThemeComplete(inputFile, outputFile, themeType, includeShapes);
    }
}
```

## 注意事项

1. **字体兼容性**: 确保目标系统安装了指定的字体，否则会自动回退
2. **颜色一致性**: 应用主题前建议备份原文件
3. **样式覆盖**: 主题会覆盖现有的颜色和字体设置
4. **性能考虑**: 大型PPT处理时间可能较长
5. **版本兼容**: 主题效果在不同PowerPoint版本中可能略有差异

## 最佳实践

1. **主题选择**: 根据演示场合和受众选择合适的主题
2. **品牌一致性**: 建议创建公司专属主题并复用
3. **样式保留**: 对特殊元素使用选择性应用功能
4. **预览测试**: 应用主题后在不同设备上预览效果
5. **逐步应用**: 对大型PPT可分批应用主题

## API 参考

### 核心类

| 类 | 描述 |
|----|------|
| `ThemeManager` | 主题管理器主类 |
| `ColorApplier` | 颜色方案应用器 |
| `FontApplier` | 字体方案应用器 |
| `ShapeStyleApplier` | 形状样式应用器 |
| `ColorPaletteGenerator` | 色板生成器 |
| `CustomThemeCreator` | 自定义主题创建器 |

### 数据结构

| 类 | 描述 |
|----|------|
| `ThemeConfig` | 主题配置数据结构 |
| `ColorScheme` | 颜色方案数据结构 |
| `FontScheme` | 字体方案数据结构 |
| `ShapeStyleScheme` | 形状样式方案数据结构 |

### 枚举类型

| 枚举 | 描述 |
|------|------|
| `ThemeType` | 主题类型 |

## 相关功能

- [文本处理](./03-text-content.md) - 文本字体设置
- [形状处理](./04-shapes-images.md) - 形状样式设置
- [高级功能](./12-advanced-features.md) - 主题基础操作
- [图表](./06-charts.md) - 图表颜色设置
- [SmartArt](./07-smartart.md) - SmartArt样式设置
