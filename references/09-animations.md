---
title: 动画效果
category: spire-presentation
description: 使用 Spire.Presentation 为幻灯片、形状和文本添加动画效果
---

# 动画效果

## 概述

Spire.Presentation 提供了丰富的动画功能，包括：
- 幻灯片切换效果
- 形状进入、强调、退出动画
- 文本动画
- 自定义路径动画
- 动画时长和延迟设置

## 幻灯片切换效果

### 示例 1: 设置幻灯片切换效果

```csharp
using Spire.Presentation;
using Spire.Presentation.Drawing.Transition;

Presentation presentation = new Presentation();

// 设置第一张幻灯片的切换效果
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Circle;

// 设置切换速度
presentation.Slides[0].SlideShowTransition.Duration = 2.0f;

// 设置切换音效
presentation.Slides[0].SlideShowTransition.SoundEffect = AudioData.Embedded;

presentation.SaveToFile("WithTransition.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 2: 常用切换效果

```csharp
// 淡入淡出
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Fade;

// 推入
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Push;

// 擦除
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Wipe;

// 百叶窗
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Blinds;

// 旋转
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Rotate;

// 缩放
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Zoom;

// 棋盘
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Checkerboard;

// 盒状
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Box;
```

## 形状动画

### 示例 3: 为形状添加进入动画

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing.Animation;

Presentation presentation = new Presentation();

// 添加形状
RectangleF rect = new RectangleF(100, 100, 200, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.Blue;
shape.AppendTextFrame("动画形状");

// 添加进入动画
AnimationEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape,
    AnimationEffectType.FlyIn
);

// 设置动画方向
effect.Subtype = AnimationSubtype.FromLeft;

// 设置持续时间
effect.Timing.Duration = 1.0f;

presentation.SaveToFile("ShapeAnimation.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 4: 常用进入动画

```csharp
AnimationEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape,
    AnimationEffectType.FadeIn  // 淡入
);

// 其他进入动画类型
AnimationEffectType.FlyIn           // 飞入
AnimationEffectType.ZoomIn          // 缩放进入
AnimationEffectType.Wipe            // 擦除
AnimationEffectType.Split           // 分割
AnimationEffectType.Bounce          // 弹跳
AnimationEffectType.Boomerang       // 回旋
AnimationEffectType.FloatIn         // 浮入
AnimationEffectType.Spin            // 旋转
AnimationEffectType.Wheel           // 轮子
AnimationEffectType.Random          // 随机
```

### 示例 5: 强调动画

```csharp
AnimationEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape,
    AnimationEffectType.Emphasis
);

// 设置具体强调效果
effect.Subtype = AnimationSubtype.None;

// 常用强调动画
AnimationEffectType.Pulse           // 脉冲
AnimationEffectType.Teeter          // 摇摆
AnimationEffectType.SpinEmphasis    // 强调旋转
AnimationEffectType.GrowShrink      // 放大/缩小
AnimationEffectType.BoldFlash       // 加粗闪烁
AnimationEffectType.Blink           // 闪烁
AnimationEffectType.ColorPulse      // 颜色脉冲
```

### 示例 6: 退出动画

```csharp
AnimationEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape,
    AnimationEffectType.ExitFly
);

// 设置方向
effect.Subtype = AnimationSubtype.ToRight;

// 常用退出动画
AnimationEffectType.ExitFade       // 淡出
AnimationEffectType.ExitFly        // 飞出
AnimationEffectType.ExitZoom       // 缩放退出
AnimationEffectType.ExitWipe       // 擦除
AnimationEffectType.ExitSplit      // 分割
```

### 示例 7: 动画序列

```csharp
// 添加第一个动画
AnimationEffect effect1 = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape1,
    AnimationEffectType.FadeIn
);
effect1.Timing.Duration = 1.0f;

// 添加第二个动画（在前一个之后）
AnimationEffect effect2 = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape2,
    AnimationEffectType.FlyIn
);
effect2.Timing.TriggerType = AnimationTriggerType.AfterPrevious;
effect2.Timing.Duration = 1.5f;

// 添加第三个动画（与前一个同时）
AnimationEffect effect3 = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape3,
    AnimationEffectType.ZoomIn
);
effect3.Timing.TriggerType = AnimationTriggerType.WithPrevious;
effect3.Timing.Duration = 1.0f;
```

### 示例 8: 设置动画延迟和重复

```csharp
AnimationEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape,
    AnimationEffectType.FadeIn
);

// 设置延迟时间（秒）
effect.Timing.TriggerDelayTime = 0.5f;

// 设置持续时间
effect.Timing.Duration = 2.0f;

// 设置重复次数
effect.Timing.RepeatCount = 3;

// 设置重复类型
effect.Timing.RepeatType = AnimationRepeatType.Count;
// 其他选项: UntilNextClick, UntilEndOfSlide, UntilStop
```

### 示例 9: 自定义路径动画

```csharp
using Spire.Presentation.Drawing.Animation;

AnimationEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape,
    AnimationEffectType.PathUser
);

// 创建自定义路径
MotionPath motionPath = new MotionPath();

// 添加路径点
motionPath.Add(MotionCommandPathType.MoveTo, new PointF(100, 100), MotionPathPointsType.Auto, false);
motionPath.Add(MotionCommandPathType.LineTo, new PointF(200, 200), MotionPathPointsType.Auto, false);
motionPath.Add(MotionCommandPathType.LineTo, new PointF(300, 100), MotionPathPointsType.Auto, true);

// 应用路径
effect.MotionPath = motionPath;
```

### 示例 10: 预设路径动画

```csharp
// 使用预设路径
AnimationEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape,
    AnimationEffectType.PathCircle  // 圆形路径
);

// 其他预设路径
AnimationEffectType.Path4PointStar      // 四角星
AnimationEffectType.Path8PointStar      // 八角星
AnimationEffectType.PathLoop            // 循环
AnimationEffectType.PathCrescentMoon    // 新月
AnimationEffectType.PathHeart           // 心形
AnimationEffectType.PathHexagon         // 六边形
AnimationEffectType.PathOctagon         // 八边形
AnimationEffectType.PathPentagon        // 五边形
AnimationEffectType.PathSquare          // 正方形
AnimationEffectType.PathTeardrop        // 泪滴
AnimationEffectType.PathTurnDown        // 向下转弯
AnimationEffectType.PathTurnUp          // 向上转弯
AnimationEffectType.PathZigzag          // 锯齿形
```

## 文本动画

### 示例 11: 按字符动画

```csharp
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing.Animation;

Presentation presentation = new Presentation();

// 添加文本形状
RectangleF rect = new RectangleF(100, 100, 400, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle,
    rect
);
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.White;
shape.AppendTextFrame("这是文本动画示例");

// 添加动画
AnimationEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape,
    AnimationEffectType.FadeIn
);

// 设置按字符动画
effect.TextAnimation.Type = AnimateTextType.ByCharacter;
effect.TextAnimation.DelayBetween = 0.1f;  // 字符间延迟

presentation.SaveToFile("TextAnimation.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 12: 按段落动画

```csharp
AnimationEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape,
    AnimationEffectType.FadeIn
);

// 设置按段落动画
effect.TextAnimation.Type = AnimateTextType.ByParagraph;
effect.TextAnimation.DelayBetween = 0.2f;  // 段落间延迟
```

### 示例 13: 按词动画

```csharp
AnimationEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape,
    AnimationEffectType.FadeIn
);

// 设置按词动画
effect.TextAnimation.Type = AnimateTextType.ByWord;
effect.TextAnimation.DelayBetween = 0.05f;  // 词间延迟
```

### 示例 14: 一次性显示全部

```csharp
AnimationEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape,
    AnimationEffectType.FadeIn
);

// 设置一次性显示全部
effect.TextAnimation.Type = AnimateTextType.AllAtOnce;
```

## 动画设置

### 示例 15: 设置动画触发方式

```csharp
AnimationEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape,
    AnimationEffectType.FadeIn
);

// 点击触发
effect.Timing.TriggerType = AnimationTriggerType.OnClick;

// 自动触发（与上一个动画同时）
effect.Timing.TriggerType = AnimationTriggerType.WithPrevious;

// 自动触发（在上一个动画之后）
effect.Timing.TriggerType = AnimationTriggerType.AfterPrevious;
```

### 示例 16: 获取动画信息

```csharp
Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 获取幻灯片的所有动画
var animations = presentation.Slides[0].Timeline.MainSequence;

foreach (AnimationEffect effect in animations)
{
    Console.WriteLine($"动画类型: {effect.EffectType}");
    Console.WriteLine($"动画子类型: {effect.Subtype}");
    Console.WriteLine($"持续时间: {effect.Timing.Duration} 秒");
    Console.WriteLine($"延迟时间: {effect.Timing.TriggerDelayTime} 秒");
    Console.WriteLine($"触发类型: {effect.Timing.TriggerType}");
}

presentation.Dispose();
```

### 示例 17: 删除动画

```csharp
// 删除特定动画
Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

var animations = presentation.Slides[0].Timeline.MainSequence;
animations.RemoveAt(0);  // 删除第一个动画

// 删除所有动画
animations.Clear();

presentation.SaveToFile("NoAnimations.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 18: 设置动画声音

```csharp
AnimationEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(
    shape,
    AnimationEffectType.FadeIn
);

// 设置动画声音
effect.SoundEffect = AnimationSoundEffect.BuiltInSound;
effect.SoundEffectName = "Chime";

// 或使用自定义声音
effect.SoundEffect = AnimationSoundEffect.CustomSound;
effect.SoundEffectData = File.ReadAllBytes("sound.wav");
```

## 动画类型参考

### 进入动画类型

| 类型 | 描述 |
|------|------|
| `FadeIn` | 淡入 |
| `FlyIn` | 飞入 |
| `ZoomIn` | 缩放进入 |
| `Wipe` | 擦除 |
| `Split` | 分割 |
| `Bounce` | 弹跳 |
| `Boomerang` | 回旋 |
| `FloatIn` | 浮入 |
| `Spin` | 旋转 |
| `Wheel` | 轮子 |
| `Circle` | 圆形 |
| `Diamond` | 菱形 |
| `In` | 放大 |
| `Plus` | 加号 |
| `Random` | 随机 |

### 强调动画类型

| 类型 | 描述 |
|------|------|
| `Emphasis` | 强调 |
| `Pulse` | 脉冲 |
| `Teeter` | 摇摆 |
| `SpinEmphasis` | 强调旋转 |
| `GrowShrink` | 放大/缩小 |
| `BoldFlash` | 加粗闪烁 |
| `Blink` | 闪烁 |
| `ColorPulse` | 颜色脉冲 |
| `Darken` | 变暗 |
| `Desaturate` | 降低饱和度 |
| `Lighten` | 变亮 |

### 退出动画类型

| 类型 | 描述 |
|------|------|
| `ExitFade` | 淡出 |
| `ExitFly` | 飞出 |
| `ExitZoom` | 缩放退出 |
| `ExitWipe` | 擦除 |
| `ExitSplit` | 分割 |
| `Out` | 缩小 |
| `ExitSpin` | 退出旋转 |

### 动画方向 (AnimationSubtype)

| 方向 | 描述 |
|------|------|
| `None` | 无方向 |
| `FromLeft` | 从左 |
| `FromRight` | 从右 |
| `FromTop` | 从上 |
| `FromBottom` | 从下 |
| `FromLeftTop` | 从左上 |
| `FromRightBottom` | 从右下 |
| `ToRight` | 向右 |
| `ToLeft` | 向左 |
| `ToTop` | 向上 |
| `ToBottom` | 向下 |

## 注意事项

1. **性能**: 大量动画可能影响演示文稿的性能
2. **兼容性**: 某些动画效果在不同版本的 PowerPoint 中表现可能不同
3. **用户体验**: 避免使用过多动画，以免分散观众注意力
4. **导出**: 某些动画效果在转换为 PDF 或其他格式时可能丢失

## 最佳实践

1. **适度使用**: 每张幻灯片使用 1-3 个动画效果
2. **保持一致**: 在整个演示文稿中使用相似的动画风格
3. **测试播放**: 在不同设备上测试动画效果
4. **提供替代**: 为动画提供静态内容作为备选

## 相关功能

- [形状处理](./04-shapes-images.md) - 为形状添加动画
- [文本处理](./03-text-content.md) - 文本动画
- [多媒体](./08-multimedia.md) - 多媒体与动画结合
