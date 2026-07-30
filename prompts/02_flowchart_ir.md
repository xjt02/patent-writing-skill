# 02 — Flowchart IR（流程图中间表示）

> **阶段**：阶段 2
> **输入**：abstract_page.md + intake JSON
> **输出**：flowchart_ir.json（visio-scientific-flowchart 兼容）

## 目标

从摘要中提炼流程图，生成 visio-scientific-flowchart 可消费的 JSON IR。

## IR Schema（与 visio-scientific-flowchart 对齐）

```json
{
  "meta": {
    "title": "<流程图标题，≤20 字>",
    "diagram_type": "flowchart",
    "direction": "vertical",
    "with_arrows": false
  },
  "nodes": [
    {
      "text": "<节点文字，≤10 字>",
      "kind": "process | decision | io_input | terminal_start | terminal_end",
      "position": {"x": 1.0, "y": <float>, "w": 0.9, "h": 0.35}
    }
  ],
  "edges": [
    {
      "src": "<源节点 text>",
      "dst": "<目标节点 text>",
      "label": "<可选标签，如'是'/'否'>",
      "style": "solid"
    }
  ]
}
```

## 节点 kind 规范

| kind | 用途 | 视觉特征 |
|---|---|---|
| `terminal_start` | 流程开始 | 圆角矩形 |
| `io_input` | 数据采集 / 输入 | 平行四边形 |
| `process` | 处理步骤 | 矩形 |
| `decision` | 判断 / 决策 | 菱形 |
| `terminal_end` | 流程结束 | 圆角矩形 |

## 硬约束

1. **节点数 ≤ 12**：多于 12 节点必须合并或分层
2. **每节点文字 ≤ 10 字**：超长用简称（如"图像采集" → "采图"）
3. **起始节点唯一**：用 `terminal_start`，位置最高（y 最大）
4. **结束节点唯一**：用 `terminal_end`，位置最低（y 最小）
5. **decision 节点必须有 2 条出边**：分别标 "是" / "否"（或等价词）
6. **垂直方向** `direction=vertical`，节点 y 单调递减
7. **位置参数**：
   - x ∈ [1.0, 2.5]（页面宽度内）
   - y 从 8.0 开始，每节点递减 0.6
   - w=0.9, h=0.35（标准尺寸）

## Few-shot 示例

**输入摘要**：
> 基于仪表数据集，通过 HRNet 网络训练获得仪表识别模型；采集仪表实时图像，输入至仪表识别模型中，获取仪表数值。

**输出 IR**（参考 `examples/sample_001_pointer_meter.md`）：

```json
{
  "meta": {"title": "指针式仪表识别流程", "diagram_type": "flowchart", "direction": "vertical", "with_arrows": false},
  "nodes": [
    {"text": "采集仪表图像", "kind": "io_input", "position": {"x": 1.0, "y": 8.0, "w": 0.9, "h": 0.35}},
    {"text": "ROI 裁剪", "kind": "process", "position": {"x": 1.0, "y": 7.4, "w": 0.9, "h": 0.35}},
    {"text": "HRNet 推理", "kind": "process", "position": {"x": 1.0, "y": 6.8, "w": 0.9, "h": 0.35}},
    {"text": "输出关键点", "kind": "process", "position": {"x": 1.0, "y": 6.2, "w": 0.9, "h": 0.35}},
    {"text": "关键点合规？", "kind": "decision", "position": {"x": 1.0, "y": 5.6, "w": 0.9, "h": 0.35}},
    {"text": "计算刻度索引", "kind": "process", "position": {"x": 1.0, "y": 5.0, "w": 0.9, "h": 0.35}},
    {"text": "计算仪表数值", "kind": "process", "position": {"x": 1.0, "y": 4.4, "w": 0.9, "h": 0.35}},
    {"text": "输出读数", "kind": "terminal_end", "position": {"x": 1.0, "y": 3.8, "w": 0.9, "h": 0.35}},
    {"text": "低置信度报警", "kind": "process", "position": {"x": 2.2, "y": 5.6, "w": 0.9, "h": 0.35}}
  ],
  "edges": [
    {"src": "采集仪表图像", "dst": "ROI 裁剪"},
    {"src": "ROI 裁剪", "dst": "HRNet 推理"},
    {"src": "HRNet 推理", "dst": "输出关键点"},
    {"src": "输出关键点", "dst": "关键点合规？"},
    {"src": "关键点合规？", "dst": "计算刻度索引", "label": "是"},
    {"src": "计算刻度索引", "dst": "计算仪表数值"},
    {"src": "计算仪表数值", "dst": "输出读数"},
    {"src": "关键点合规？", "dst": "低置信度报警", "label": "否"}
  ]
}
```

## 反例

- ❌ 节点文字超 10 字（如"采集高清摄像机视频流图像"）
- ❌ decision 节点只有 1 条出边
- ❌ 多于 12 节点（复杂流程应分层）
- ❌ 起始节点不是 `terminal_start`

## 收尾

- 写到 `assets/cases/{case_id}/flowchart_ir.json`
- 用 `tools/insert_vsdx.py` 生成 vsdx + PNG 双产物
- 进入 `prompts/03_claims.md`