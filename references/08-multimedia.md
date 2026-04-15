---
title: 音频和视频
category: spire-presentation
description: 使用 Spire.Presentation 在演示文稿中添加和管理音频、视频
---

# 音频和视频

## 概述

Spire.Presentation 提供了完整的音频和视频处理功能，包括插入、提取、设置播放模式等操作。

## 示例

### 示例 1: 插入音频

```csharp
using System;
using System.Drawing;
using Spire.Presentation;
using System.IO;

Presentation presentation = new Presentation();

// 插入音频到指定位置
RectangleF audioRect = new RectangleF(100, 100, 100, 100);
presentation.Slides[0].Shapes.AppendAudioMedia(
    Path.GetFullPath("music.wav"),
    audioRect
);

presentation.SaveToFile("WithAudio.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 2: 插入视频

```csharp
using System;
using System.Drawing;
using Spire.Presentation;
using System.IO;

Presentation presentation = new Presentation();

// 插入视频到指定位置
RectangleF videoRect = new RectangleF(100, 100, 300, 200);
presentation.Slides[0].Shapes.AppendVideoMedia(
    Path.GetFullPath("video.mp4"),
    videoRect
);

presentation.SaveToFile("WithVideo.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 3: 设置视频播放模式

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 插入视频
RectangleF videoRect = new RectangleF(100, 100, 400, 250);
IAudioVideo videoShape = presentation.Slides[0].Shapes.AppendVideoMedia(
    "video.mp4",
    videoRect
);

// 设置播放模式
videoShape.PlayMode = VideoPlayModeType.Auto; // 自动播放
// 其他选项: Click, AllSlides, LoopingUntilStopped

// 设置音量
videoShape.Volume = AudioVolumeMode.Medium;
// 其他选项: Muted, Low, Loud

presentation.SaveToFile("VideoWithSettings.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 4: 提取音频

```csharp
using System.IO;
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 遍历所有幻灯片的音频
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IAudio audioShape)
        {
            // 保存音频到文件
            File.WriteAllBytes($"audio_{shape.Name}.wav", audioShape.Data.Data);
        }
    }
}

presentation.Dispose();
```

### 示例 5: 提取视频

```csharp
using System.IO;
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 遍历所有幻灯片的视频
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IVideo videoShape)
        {
            // 保存视频到文件
            File.WriteAllBytes($"video_{shape.Name}.mp4", videoShape.BinaryData);
        }
    }
}

presentation.Dispose();
```

### 示例 6: 替换视频

```csharp
using System.Drawing;
using Spire.Presentation;
using System.IO;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 查找视频形状
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IVideo videoShape)
        {
            // 读取新视频数据
            byte[] newVideoData = File.ReadAllBytes("new_video.mp4");

            // 替换视频
            videoShape.BinaryData = newVideoData;
            break;
        }
}

presentation.SaveToFile("VideoReplaced.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 7: 设置音频播放模式

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 插入音频
RectangleF audioRect = new RectangleF(100, 100, 100, 100);
IAudio audioShape = presentation.Slides[0].Shapes.AppendAudioMedia(
    "background_music.wav",
    audioRect
);

// 设置播放模式
audioShape.PlayMode = AudioPlayModeType.Auto; // 自动播放
// 其他选项: Click, AcrossSlides

// 设置是否在幻灯片放映时隐藏
audioShape.HideAtShowing = false; // 播放时显示音频图标

// 设置音量
audioShape.Volume = AudioVolumeMode.Low;
// 其他选项: Muted, Medium, Loud

// 设置循环播放
audioShape.LoopSound = true;

presentation.SaveToFile("AudioWithSettings.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 8: 隐藏音频图标

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 插入音频
RectangleF audioRect = new RectangleF(100, 100, 100, 100);
IAudio audioShape = presentation.Slides[0].Shapes.AppendAudioMedia(
    "background.wav",
    audioRect
);

// 设置在放映时隐藏音频图标
audioShape.HideAtShowing = true;

presentation.SaveToFile("HiddenAudio.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 9: 获取音频效果

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 查找音频并获取信息
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IAudio audioShape)
        {
            Console.WriteLine($"音频名称: {shape.Name}");
            Console.WriteLine($"播放模式: {audioShape.PlayMode}");
            Console.WriteLine($"音量: {audioShape.Volume}");
            Console.WriteLine($"循环播放: {audioShape.LoopSound}");
            Console.WriteLine($"放映时隐藏: {audioShape.HideAtShowing}");
            Console.WriteLine($"数据大小: {audioShape.Data.Data.Length} 字节");
        }
    }
}

presentation.Dispose();
```

### 示例 10: 设置视频全屏播放

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 插入视频
RectangleF videoRect = new RectangleF(100, 100, 400, 250);
IAudioVideo videoShape = presentation.Slides[0].Shapes.AppendVideoMedia(
    "video.mp4",
    videoRect
);

// 设置全屏播放
videoShape.PlayFullScreen = true;

presentation.SaveToFile("FullScreenVideo.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 11: 裁剪视频

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 查找视频并裁剪
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IVideo videoShape)
        {
            // 设置裁剪（以秒为单位）
            videoShape.StartTime = 5.0; // 从第5秒开始
            videoShape.EndTime = 15.0;   // 到第15秒结束
        }
    }
}

presentation.SaveToFile("CroppedVideo.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 12: 设置视频裁剪显示

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 查找视频
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IVideo videoShape)
        {
            // 设置视频裁剪显示区域
            videoShape.Crop.F = 0.1f; // 上
            videoShape.Crop.T = 0.1f; // 下
            videoShape.Crop.L = 0.1f; // 左
            videoShape.Crop.R = 0.1f; // 右
        }
    }
}

presentation.SaveToFile("DisplayCropped.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 13: 在多个幻灯片中添加背景音乐

```csharp
using System.Drawing;
using Spire.Presentation;

Presentation presentation = new Presentation();

// 在每张幻灯片添加背景音乐
for (int i = 0; i < 5; i++)
{
    presentation.Slides.Append();

    RectangleF audioRect = new RectangleF(10, 10, 50, 50);
    IAudio audio = presentation.Slides[i].Shapes.AppendAudioMedia(
        "background.mp3",
        audioRect
    );

    // 设置跨幻灯片播放
    audio.PlayMode = AudioPlayModeType.AccrossSlides;
    audio.HideAtShowing = true; // 隐藏图标
}

presentation.SaveToFile("BackgroundMusic.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 14: 删除音频或视频

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("presentation.pptx");

// 删除所有音频
foreach (ISlide slide in presentation.Slides)
{
    for (int i = slide.Shapes.Count - 1; i >= 0; i--)
    {
        if (slide.Shapes[i] is IAudio)
        {
            slide.Shapes.RemoveAt(i);
        }
    }
}

// 删除所有视频
foreach (ISlide slide in presentation.Slides)
{
    for (int i = slide.Shapes.Count - 1; i >= 0; i--)
    {
        if (slide.Shapes[i] is IVideo)
        {
            slide.Shapes.RemoveAt(i);
        }
    }
}

presentation.SaveToFile("MediaRemoved.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 音频属性

### IAudio 主要属性

| 属性 | 类型 | 描述 |
|------|------|------|
| `PlayMode` | AudioPlayModeType | 播放模式（自动/点击/跨幻灯片） |
| `Volume` | AudioVolumeMode | 音量（静音/低/中/高） |
| `HideAtShowing` | bool | 播放时是否隐藏图标 |
| `LoopSound` | bool | 是否循环播放 |
| `Data` | AudioData | 音频数据 |

### 音频播放模式 (AudioPlayModeType)

| 模式 | 描述 |
|------|------|
| `Auto` | 自动播放 |
| `Click` | 点击播放 |
| `AcrossSlides` | 跨幻灯片播放 |

### 音量模式 (AudioVolumeMode)

| 模式 | 描述 |
|------|------|
| `Muted` | 静音 |
| `Low` | 低音量 |
| `Medium` | 中音量 |
| `Loud` | 高音量 |

## 视频属性

### IVideo 主要属性

| 属性 | 类型 | 描述 |
|------|------|------|
| `PlayMode` | VideoPlayModeType | 播放模式 |
| `Volume` | AudioVolumeMode | 音量 |
| `PlayFullScreen` | bool | 是否全屏播放 |
| `BinaryData` | byte[] | 视频数据 |
| `Crop` | PictureCrop | 视频裁剪设置 |
| `StartTime` | double | 开始时间（秒） |
| `EndTime` | double | 结束时间（秒） |

### 视频播放模式 (VideoPlayModeType)

| 模式 | 描述 |
|------|------|
| `Auto` | 自动播放 |
| `Click` | 点击播放 |
| `AllSlides` | 所有幻灯片 |
| `LoopingUntilStopped` | 循环直到停止 |

## 支持的格式

### 音频格式
- MP3
- WAV
- WMA
- M4A

### 视频格式
- MP4
- AVI
- WMV
- MOV
- MKV（部分支持）

## 注意事项

1. **文件大小**: 大型音频/视频文件会导致 PPT 文件体积增大
2. **编码兼容性**: 确保使用 PowerPoint 支持的编解码器
3. **播放顺序**: 多个音频的播放顺序可能需要特别设置
4. **跨版本兼容性**: 某些格式在不同 PowerPoint 版本中支持度不同

## 最佳实践

1. **使用压缩格式**: 优先使用压缩后的音频/视频以减小文件大小
2. **测试播放**: 在不同设备上测试多媒体内容的播放效果
3. **备份原始文件**: 保留原始音频/视频文件以便后续修改
4. **考虑用户体验**: 避免自动播放会打扰用户的音频

## 相关功能

- [形状处理](./04-shapes-images.md) - 音频/视频作为形状处理
- [动画](./09-animations.md) - 为多媒体添加动画效果
- [超链接](./10-hyperlinks.md) - 使用超链接链接到多媒体
