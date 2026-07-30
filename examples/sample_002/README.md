# sample_002 — 指针仪表识别（md_to_docx.py 端到端验收）

## 概况

| 项 | 值 |
|---|---|
| 案件 ID | sample_002 |
| 节 1 扉页摘要 | 133 字（≤300 字 ✓） |
| 节 2 摘要附图 | 无（流程图非必需） |
| 节 3 权利要求书 | 1 独立 + 2 从属，共 3 条 |
| 节 4 说明书 | 空（未提供） |
| 节 5 说明书附图 | 2 个占位段（图 2 / 图 3） |
| 公式 | 无 |

## 产物清单

| 文件 | 大小 | 用途 |
|---|---|---|
| `claims.json` | 425 B | 权利要求列表 |
| `sample_002.docx` | 31.4 KB | 最终专利文档 |

## 验证项

- [x] 5 节齐全
- [x] 5 个页眉全部正确（含节 4 修正后的 `说    明    书`）
- [x] 节 1 摘要 133 字（≤ 300）
- [x] 节 3 包含 `1、` `2、` `3、` 编号
- [x] 节 5 占位段 `图2  待贴附：XXX` `图3  待贴附：XXX`
- [x] 文档属性：title=`一种指针仪表识别方法` author=`[发明人]`
- [x] 模板继承：未触 `docDefaults`，页眉/页脚/sectPr 完全保留

## 复现命令

```bash
cd ${SKILL_ROOT}/.claude/skills/patent-writing-skill
python tools/md_to_docx.py \
  --template assets/templates/专利空白模版_fixed.docx \
  --case-id sample_002 \
  --abstract-text "本发明公开了一种指针仪表识别方法及其系统，属于工业仪表智能识别技术领域..." \
  --claims examples/sample_002/claims.json \
  --placeholders "检测流程示意图" "系统架构框图" \
  --doc-title "一种指针仪表识别方法" \
  --doc-author "[发明人]" \
  --output /tmp/sample_002.docx
```

## 与 sample_001 的对比

| 维度 | sample_001 | sample_002 |
|---|---|---|
| 主工具 | `insert_vsdx.py`（PNG 嵌入） | `md_to_docx.py`（5 节组装） |
| 测试目标 | Visio COM 出 PNG + docx 嵌入 | 模板继承 + 内容注入 |
| 内容 | 仅 IR + 占位段 | 摘要 + 权利要求 + 占位段 |
| 端到端 | ✓ | ✓ |

## 已知限制

- 说明书（节 4）未注入内容（待与 prompts/06_assemble.md 完整联动）
- 没有公式段（待 P1 完成 `extract_formulas.py` 后验证）