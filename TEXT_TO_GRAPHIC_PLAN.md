# 文本降维与图形化功能添加计划

## 功能需求

**核心功能**: 将幻灯片中指定的一大段枯燥文本自动转换为流程图或图形化表示，使用 SmartArt 或 Shape 完成。

## 实现方案

### 方案概述

该功能需要：
1. 从幻灯片中提取指定文本
2. 分析文本结构（识别步骤、列表、因果关系等）
3. 根据文本特征选择合适的图形类型（流程图/关系图/列表等）
4. 自动生成 SmartArt 或 Shape 布局
5. 替换或附加原文本位置

### 技术实现点

#### 1. 文本分析
- 识别序号（1. 2. 3. 或 第一、第二、第三）
- 识别关键词（开始、结束、如果、那么、因为、所以）
- 识别标点符号（分号、逗号、句号等分隔符）
- 识别列表格式（项目符号、编号）

#### 2. 图形类型选择策略

| 文本特征 | 推荐图形 | SmartArt 布局 |
|---------|---------|---------------|
| 顺序步骤 | 流程图 | BasicProcess, ChevronProcess |
| 循环/迭代 | 循环图 | BasicCycle, Cycle |
| 层次结构 | 层次图 | Hierarchy, OrganizationalChart |
| 并列关系 | 列表 | BasicBendingProcess |
| 因果关系 | 关系图 | Balance, ConvergingRadial |
| 对比/选择 | 分支图 | Decision (用 Shape 模拟) |

#### 3. 实现功能点

**基本功能**:
- 文本提取和解析
- 节点生成
- 连接线生成（如果使用 Shape）
- 样式应用

**高级功能**:
- 自动布局计算
- 智能缩放适配
- 颜色主题应用
- 添加动画效果
- 保留原文本作为注释/备注

## 需要添加的文档内容

### 1. 新增参考文档

创建 `references/16-text-to-graphic.md`，包含：
- 文本分析算法说明
- 图形类型选择逻辑
- 转换流程图
- API 使用示例

### 2. 新增示例文件

创建 `examples/advanced/text-to-flowchart.cs`，包含：
- 简单文本转流程图示例
- 复杂文本转层次图示例
- 智能选择图形类型示例
- 批量转换示例

### 3. 更新 SKILL.md

在 "When to Use" 中添加：
- "Convert text to flowcharts or diagrams"
- "Visualize complex text as SmartArt graphics"

在 "Quick Reference" 表格中添加：
- Text to Graphic Conversion | [Text to Graphic](./references/16-text-to-graphic.md)

### 4. 更新 evals/trigger_eval.json

添加正触发测试：
- "Convert this text to a flowchart"
- "Visualize these steps as a diagram"
- "Turn this paragraph into SmartArt"

添加负触发测试：
- "Convert image to text" (这是 OCR，不是我们要的功能)
- "Extract text from diagram"

### 5. 创建实用工具类（可选）

在 `examples/scripts/` 中创建：
- `TextAnalyzer.cs` - 文本分析工具
- `GraphicGenerator.cs` - 图形生成工具
- `LayoutOptimizer.cs` - 布局优化工具

## 详细实现步骤

### 步骤 1: 创建核心参考文档

**文件**: `references/16-text-to-graphic.md`

内容结构：
1. 概述 - 功能说明和使用场景
2. 文本分析 - 如何解析文本结构
3. 图形选择 - 根据文本特征选择图形
4. 基础转换 - 简单文本到 SmartArt
5. 高级转换 - 复杂文本到自定义 Shape
6. 布局优化 - 自动调整布局
7. 样式定制 - 颜色、字体、大小
8. 完整示例 - 端到端实现
9. API 参考 - 相关类和方法

### 步骤 2: 创建示例代码

**文件**: `examples/advanced/text-to-flowchart.cs`

示例 1: 简单步骤文本转流程图
- 输入: "1. 需求分析 2. 设计方案 3. 开发实现 4. 测试验证"
- 输出: BasicProcess SmartArt

示例 2: 复杂文本转层次图
- 输入: 带缩进的层次文本
- 输出: Hierarchy SmartArt

示例 3: 智能图形选择
- 根据文本关键词自动选择图形类型

示例 4: 批量转换
- 处理多个文本框

### 步骤 3: 更新现有文档

更新 `SKILL.md`:
- 添加新功能的触发条件
- 更新快速参考表

更新 `references/07-smartart.md`:
- 添加"文本到 SmartArt 转换"章节的链接

### 步骤 4: 更新评估文件

在 `evals/trigger_eval.json` 中添加相关测试用例。

## 代码示例模板

```csharp
// 文本转流程图核心逻辑
public static ISmartArt ConvertTextToFlowchart(
    Presentation presentation,
    ISlide slide,
    string text,
    RectangleF position)
{
    // 1. 分析文本结构
    List<string> steps = AnalyzeSteps(text);

    // 2. 选择合适的布局
    SmartArtLayoutType layout = SelectLayout(steps);

    // 3. 创建 SmartArt
    ISmartArt smartArt = slide.Shapes.AppendSmartArt(
        position,
        layout
    );

    // 4. 添加节点
    foreach (string step in steps)
    {
        ISmartArtNode node = smartArt.Nodes.AddNode();
        node.TextFrame.Text = step;
    }

    // 5. 应用样式
    smartArt.ColorStyle = SmartArtColorType.Colorful;
    smartArt.SmartArtStyle = SmartArtStyleType.WhiteOutline;

    return smartArt;
}
```

## 注意事项

1. **文本长度限制**: 过长文本需要截断或分页
2. **节点数量限制**: SmartArt 对节点数量有限制
3. **布局自动调整**: 需要根据内容自动调整大小
4. **保留原文本**: 考虑是否保留原文本到备注
5. **错误处理**: 无法分析的文本应给出提示

## 预期效果

用户输入:
```
将这段文本转换为流程图：
1. 接收用户请求
2. 验证数据格式
3. 处理业务逻辑
4. 返回响应结果
```

输出:
自动生成一个包含4个步骤的流程图 SmartArt。

## 时间估算

- 文档编写: 1-2小时
- 示例代码: 1-2小时
- 更新现有文件: 30分钟
- 测试验证: 1小时

总计: 约3.5-5.5小时
