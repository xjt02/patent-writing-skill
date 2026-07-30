# sample_001 — 指针仪表识别（端到端 P2 验收样本）

## 概况

| 项 | 值 |
|---|---|
| 案件 ID | sample_001 |
| 节点数 | 9（1 起始 + 7 过程 + 1 终止，含 1 决策） |
| 出边数 | 9 |
| 公式数 | 0 |
| 节 1 扉页摘要 | 空（端到端只验证骨架） |
| 节 3 权利要求书 | 空 |
| 节 4 说明书 | 空 |
| 节 5 说明书附图 | 空（仅占位段） |

## 产物清单

| 文件 | 大小 | 用途 |
|---|---|---|
| `flowchart_ir.json` | 1.6 KB | visio-scientific-flowchart 输入 IR |
| `sample_001.vsdx` | 25 KB | Visio 流程图源文件（可手工微调） |
| `sample_001.png` | 17.8 KB | Visio COM 导出的 PNG（嵌入节 2） |
| `sample_001.docx` | 44 KB | 最终专利文档（5 节 + 嵌入图） |

## 验证项

- [x] 5 节齐全（扉页/附图/权利要求/说明书/说明书附图）
- [x] 节 1 页眉：`说    明    书    摘    要`
- [x] 节 2 页眉：`摘    要    附    图`
- [x] 节 3 页眉：`权    利    要    求    书 `（含国标尾随空格）
- [x] 节 4 页眉：`说    明    书`（**修正后**）
- [x] 节 5 页眉：`说    明    书    附    图`
- [x] 节 2 含 1 个 `<w:drawing>` 嵌入图
- [x] `word/media/image2.png` 落地

## 复现命令

```bash
cd ${SKILL_ROOT}/.claude/skills/patent-writing-skill
python tools/insert_vsdx.py \
  examples/sample_001/flowchart_ir.json \
  assets/templates/专利空白模版_fixed.docx \
  /tmp/output \
  --case-id sample_001
```

## 限制（端到端骨架测试）

- 节 1 扉页摘要、节 3 权利要求书、节 4 说明书仍是模板空白段
- 公式段未注入（sample_001 无公式）
- 节 5 仅保留模板占位段

完整文字生成需配合 `prompts/06_assemble.md` + `md_to_docx.py`（P2 下一步）。