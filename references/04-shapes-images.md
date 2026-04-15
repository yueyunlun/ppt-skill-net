---
title: 形状和图像
category: spire-presentation
description: 使用 Spire.Presentation 添加和管理各种形状和图像
---

# 形状和图像

## 概述

Spire.Presentation 提供了丰富的形状和图像处理功能，包括：
- 20+ 种内置形状类型
- 自定义形状
- 图片插入和格式化
- 形状样式和效果
- 形状排列和组合

## 形状类型

### 基础形状
- `ShapeType.Rectangle` - 矩形
- `ShapeType.RoundedRectangle` - 圆角矩形
- `ShapeType.Ellipse` - 椭圆
- `ShapeType.Oval` - 椭圆
- `ShapeType.Line` - 线条
- `ShapeType.Diamond` - 菱形
- `ShapeType.Parallelogram` - 平行四边形
- `ShapeType.Trapezoid` - 梯形

### 箭头形状
- `ShapeType.RightArrow` - 右箭头
- `ShapeType.LeftArrow` - 左箭头
- `ShapeType.UpArrow` - 上箭头
- `ShapeType.DownArrow` - 下箭头
- `ShapeType.LeftRightArrow` - 左右箭头
- `ShapeType.UpDownArrow` - 上下箭头
- `ShapeType.CurvedRightArrow` - 弯曲右箭头
- `ShapeType.CurvedLeftArrow` - 弯曲左箭头

### 星形形状
- `ShapeType.FivePointStar` - 五角星
- `ShapeType.SixPointedStar` - 六角星
- `ShapeType.EightPointStar` - 八角星
- `ShapeType.TenPointStar` - 十角星
- `ShapeType.TwelvePointStar` - 十二角星
- `ShapeType.FourPointStar` - 四角星

### 流程图形状
- `ShapeType.FlowChartProcess` - 流程
- `ShapeType.FlowChartDecision` - 决策
- `ShapeType.FlowChartTerminator` - 终止符
- `ShapeType.FlowChartPreparation` - 准备
- `ShapeType.FlowChartData` - 数据
- `ShapeType.FlowChartDocument` - 文档

### 其他形状
- `ShapeType.Triangle` - 三角形
- `ShapeType.RightTriangle` - 直角三角形
- `ShapeType.Pentagon` - 五边形
- `ShapeType.Hexagon` - 六边形
- `ShapeType.Heptagon` - 七边形
- `ShapeType.Octagon` - 八边形
- `ShapeType.Decagon` - 十边形
- `ShapeType.Dodecagon` - 十二边形
- `ShapeType.Cloud` - 云形
- `ShapeType.LightningBolt` - 闪电
- `ShapeType.Heart` - 心形
- `ShapeType.Smile` - 笑脸
- `ShapeType.Sun` - 太阳
- `ShapeType.Moon` - 月亮

## 示例

### 示例 1: 添加基本形状

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

Presentation presentation = new Presentation();

// 添加矩形
RectangleF rect1 = new RectangleF(50, 50, 200, 100);
IAutoShape rectangle = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect1
);
rectangle.Fill.FillType = FillFormatType.Solid;
rectangle.Fill.SolidColor.Color = Color.Blue;
rectangle.ShapeStyle.LineColor.Color = Color.DarkBlue;

// 添加椭圆
RectangleF rect2 = new RectangleF(300, 50, 150, 150);
IAutoShape ellipse = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Ellipse,
    rect2
);
ellipse.Fill.FillType = FillFormatType.Solid;
ellipse.Fill.SolidColor.Color = Color.Red;
ellipse.ShapeStyle.LineColor.Color = Color.DarkRed;

presentation.SaveToFile("BasicShapes.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 2: 添加箭头形状

```csharp
// 添加右箭头
RectangleF arrowRect = new RectangleF(100, 100, 200, 80);
IAutoShape arrow = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.RightArrow,
    arrowRect
);
arrow.Fill.FillType = FillFormatType.Solid;
arrow.Fill.SolidColor.Color = Color.Orange;
arrow.ShapeStyle.LineColor.Color = Color.DarkOrange;
arrow.Line.FillFormat.FillType = FillFormatType.Solid;
arrow.Line.FillFormat.SolidColor.Color = Color.DarkOrange;
arrow.Line.Weight = 2f;
```

### 示例 3: 添加星形

```csharp
// 添加五角星
RectangleF starRect = new RectangleF(150, 100, 150, 150);
IAutoShape star = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.FivePointStar,
    starRect
);
star.Fill.FillType = FillFormatType.Solid;
star.Fill.SolidColor.Color = Color.Yellow;
star.ShapeStyle.LineColor.Color = Color.Gold;

// 设置发光效果
star.Effect.DistortionShadow.Type = DistortionEffectType.Glow;
star.Effect.DistortionShadow.Amount = 20f;
star.Effect.DistortionShadow.Color = Color.Yellow;
```

### 示例 4: 插入图片

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 插入嵌入图片
RectangleF imageRect = new RectangleF(50, 50, 400, 300);
IEmbedImage image = presentation.Slides[0].Shapes.AppendEmbedImage(
    ShapeType.Rectangle,
    "image.png",
    imageRect
);
image.Line.FillFormat.FillType = FillFormatType.None;

presentation.SaveToFile("WithImage.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 5: 插入链接图片

```csharp
// 插入链接图片（不嵌入到PPT中）
RectangleF imageRect = new RectangleF(50, 50, 400, 300);
ILinkImage linkImage = presentation.Slides[0].Shapes.AppendLinkImage(
    ShapeType.Rectangle,
    "http://example.com/image.jpg",
    imageRect
);
```

### 示例 6: 调整图片大小和裁剪

```csharp
Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 获取图片形状
IEmbedImage image = presentation.Slides[0].Shapes[0] as IEmbedImage;

if (image != null)
{
    // 调整大小
    image.Width = 500;
    image.Height = 400;

    // 设置图片裁剪
    image.Picture.Fill.PictureFillMode = PictureFillMode.Stretch;

    // 或设置图片裁剪边距
    image.Crop.F = 0.1f;  // 上
    image.Crop.T = 0.1f;  // 下
    image.Crop.L = 0.1f;  // 左
    image.Crop.R = 0.1f;  // 右
}

presentation.SaveToFile("ResizedImage.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 7: 设置图片透明度

```csharp
Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

IEmbedImage image = presentation.Slides[0].Shapes[0] as IEmbedImage;

if (image != null)
{
    // 设置透明度（0.0 = 完全透明，1.0 = 不透明）
    image.Picture.Fill.PictureTransparency = 0.5f;
}

presentation.SaveToFile("TransparentImage.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 8: 设置图片边框

```csharp
IEmbedImage image = presentation.Slides[0].Shapes[0] as IEmbedImage;

if (image != null)
{
    // 设置边框
    image.Line.FillFormat.FillType = FillFormatType.Solid;
    image.Line.FillFormat.SolidColor.Color = Color.Black;
    image.Line.Weight = 3f;
    image.Line.DashStyle = LineDashStyleType.Solid;
}
```

### 示例 9: 创建圆角矩形

```csharp
RectangleF roundedRect = new RectangleF(100, 100, 300, 150);
IAutoShape roundedShape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.RoundedRectangle,
    roundedRect
);
roundedShape.Fill.FillType = FillFormatType.Solid;
roundedShape.Fill.SolidColor.Color = Color.LightGreen;
roundedShape.ShapeStyle.LineColor.Color = Color.DarkGreen;

// 设置圆角半径（如果支持）
// roundedShape.CornerRadius = 20;
```

### 示例 10: 添加线条

```csharp
// 添加直线
PointF start = new PointF(50, 100);
PointF end = new PointF(400, 100);
IAutoShape line = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Line,
    new RectangleF(50, 100, 350, 0)
);

line.Line.FillFormat.FillType = FillFormatType.Solid;
line.Line.FillFormat.SolidColor.Color = Color.Blue;
line.Line.Weight = 2f;
```

### 示例 11: 添加带箭头的线条

```csharp
PointF start = new PointF(50, 150);
PointF end = new PointF(400, 150);
IAutoShape arrowLine = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Line,
    new RectangleF(50, 150, 350, 0)
);

arrowLine.Line.FillFormat.FillType = FillFormatType.Solid;
arrowLine.Line.FillFormat.SolidColor.Color = Color.Red;
arrowLine.Line.Weight = 3f;

// 设置箭头
arrowLine.Line.BeginArrowHeadStyle = LineArrowHeadStyleType.Arrow;
arrowLine.Line.BeginArrowHeadWidth = LineArrowHeadWidthType.Medium;
arrowLine.Line.BeginArrowHeadLength = LineArrowHeadLengthType.Medium;
```

### 示例 12: 形状样式设置

```csharp
IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;

// 设置填充
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.LightBlue;

// 或设置渐变填充
shape.Fill.FillType = FillFormatType.Gradient;
shape.Fill.Gradient.GradientStops.Append(0f, KnownColors.LightBlue);
shape.Fill.Gradient.GradientStops.Append(1f, KnownColors.DarkBlue);

// 或设置图片填充
shape.Fill.FillType = FillFormatType.Picture;
shape.Fill.Picture.Fill.PictureFillMode = PictureFillMode.Stretch;
shape.Fill.Picture.Fill.Url = "background.jpg";

// 设置边框
shape.ShapeStyle.LineColor.Color = Color.DarkBlue;
shape.Line.FillFormat.FillType = FillFormatType.Solid;
shape.Line.FillFormat.SolidColor.Color = Color.DarkBlue;
shape.Line.Weight = 2f;

// 设置阴影
shape.Effect.DistortionShadow.Type = DistortionEffectType.Shadow;
shape.Effect.DistortionShadow.Color = Color.Gray;
shape.Effect.DistortionShadow.Amount = 10f;
shape.Effect.DistortionShadow.Direction = 135f;
shape.Effect.DistortionShadow.Distance = 5f;
shape.Effect.DistortionShadow.Blur = 3f;
```

### 示例 13: 形状旋转和翻转

```csharp
IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;

// 旋转形状（角度）
shape.Rotation = 45f;

// 水平翻转
shape.FlipH = true;

// 垂直翻转
shape.FlipV = true;
```

### 示例 14: 形状排列

```csharp
// 形状对齐
presentation.Slides[0].Shapes.Align(ShapesAlignmentType.AlignLeft);
presentation.Slides[0].Shapes.Align(ShapesAlignmentType.AlignCenter);
presentation.Slides[0].Shapes.Align(ShapesAlignmentType.AlignRight);
presentation.Slides[0].Shapes.Align(ShapesAlignmentType.AlignTop);
presentation.Slides[0].Shapes.Align(ShapesAlignmentType.AlignMiddle);
presentation.Slides[0].Shapes.Align(ShapesAlignmentType.AlignBottom);

// 形状分布
presentation.Slides[0].Shapes.Distribute(ShapesDistributionType.DistributeHorizontally);
presentation.Slides[0].Shapes.Distribute(ShapesDistributionType.DistributeVertically);

// 形状叠放顺序
presentation.Slides[0].Shapes[0].ZOrder(ShapeZOrderType.BringForward);
presentation.Slides[0].Shapes[0].ZOrder(ShapeZOrderType.SendBackward);
presentation.Slides[0].Shapes[0].ZOrder(ShapeZOrderType.BringToFront);
presentation.Slides[0].Shapes[0].ZOrder(ShapeZOrderType.SendToBack);
```

### 示例 15: 复制形状

```csharp
// 在同一张幻灯片中复制
IAutoShape sourceShape = presentation.Slides[0].Shapes[0] as IAutoShape;
IAutoShape copiedShape = sourceShape.Clone() as IAutoShape;
presentation.Slides[0].Shapes.Append(copiedShape);

// 复制到另一张幻灯片
presentation.Slides[1].Shapes.Append(sourceShape.Clone() as IShape);
```

### 示例 16: 删除形状

```csharp
// 按索引删除
presentation.Slides[0].Shapes.RemoveAt(0);

// 按名称删除
for (int i = presentation.Slides[0].Shapes.Count - 1; i >= 0; i--)
{
    if (presentation.Slides[0].Shapes[i].Name == "MyShape")
    {
        presentation.Slides[0].Shapes.RemoveAt(i);
    }
}
```

### 示例 17: 提取图片

```csharp
using System.IO;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 提取所有图片
int imageIndex = 0;
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IEmbedImage image)
        {
            File.WriteAllBytes($"image_{imageIndex}.png", image.Picture.EmbedImage.Data);
            imageIndex++;
        }
    }
}

presentation.Dispose();
```

### 示例 18: 添加文本到形状

```csharp
IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;

// 添加文本
shape.AppendTextFrame("Hello World!");

// 或设置文本
shape.TextFrame.Text = "Hello World!";

// 添加多个段落
TextParagraph para1 = new TextParagraph();
para1.Text = "第一段";
shape.TextFrame.Paragraphs.Append(para1);

TextParagraph para2 = new TextParagraph();
para2.Text = "第二段";
shape.TextFrame.Paragraphs.Append(para2);

// 设置文本样式
shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 24;
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.White;
```

### 示例 19: 设置形状透明度

```csharp
IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;

// 设置整体透明度
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.FromArgb(128, Color.Blue); // 50% 透明

// 或使用设置透明度方法
// shape.Fill.SolidColor.Color.SetOpacity(0.5f);
```

### 示例 20: 形状组合

```csharp
// 获取要组合的形状
IShape shape1 = presentation.Slides[0].Shapes[0];
IShape shape2 = presentation.Slides[0].Shapes[1];
IShape shape3 = presentation.Slides[0].Shapes[2];

// 创建组（Spire.Presentation 对组合的支持有限）
// 通常需要使用其他方法或手动管理形状位置
```

## 形状属性参考

### IAutoShape 主要属性

| 属性 | 类型 | 描述 |
|------|------|------|
| `ShapeType` | ShapeType | 形状类型 |
| `Fill` | FillFormat | 填充格式 |
| `Line` | LineFormat | 线条格式 |
| `ShapeStyle` | ShapeStyle | 形状样式 |
| `TextFrame` | TextFrame | 文本框 |
| `Rotation` | float | 旋转角度 |
| `FlipH` | bool | 水平翻转 |
| `FlipV` | bool | 垂直翻转 |
| `Width` | float | 宽度 |
| `Height` | float | 高度 |
| `X` | float | X 坐标 |
| `Y` | float | Y 坐标 |
| `Effect` | EffectFormat | 效果格式 |

### IEmbedImage 主要属性

| 属性 | 类型 | 描述 |
|------|------|------|
| `Picture` | PictureFill | 图片填充 |
| `Crop` | PictureCrop | 裁剪设置 |
| `Line` | LineFormat | 边框格式 |

### LineArrowHeadStyleType

| 样式 | 描述 |
|------|------|
| `None` | 无 |
| `Arrow` | 箭头 |
| `StealthArrow` | 隐形箭头 |
| `DiamondArrow` | 菱形箭头 |
| `OvalArrow` | 椭圆箭头 |
| `OpenArrow` | 开放箭头 |
| `ChevronArrow` | V 形箭头 |

### ShapesAlignmentType

| 对齐方式 | 描述 |
|----------|------|
| `AlignLeft` | 左对齐 |
| `AlignCenter` | 居中对齐 |
| `AlignRight` | 右对齐 |
| `AlignTop` | 顶部对齐 |
| `AlignMiddle` | 中部对齐 |
| `AlignBottom` | 底部对齐 |

### ShapeZOrderType

| 顺序 | 描述 |
|------|------|
| `BringForward` | 上移一层 |
| `SendBackward` | 下移一层 |
| `BringToFront` | 置于顶层 |
| `SendToBack` | 置于底层 |

## 注意事项

1. **图片格式**: 支持常见图片格式（PNG, JPG, GIF, BMP等）
2. **图片大小**: 大图片会增加文件体积，建议优化图片尺寸
3. **形状组合**: Spire.Presentation 对形状组合的支持有限
4. **透明度**: 设置透明度时，确保使用 ARGB 颜色

## 最佳实践

1. **图片优化**: 使用适当大小的图片以减少文件体积
2. **形状命名**: 为形状设置有意义的名称以便后续操作
3. **样式一致**: 保持形状样式的一致性以获得专业外观
4. **资源管理**: 及时释放大型图片占用的资源

## 相关功能

- [文本处理](./03-text-content.md) - 形状中的文本
- [动画](./09-animations.md) - 形状动画
- [图表](./06-charts.md) - 图表作为特殊形状
- [SmartArt](./07-smartart.md) - SmartArt 图形
