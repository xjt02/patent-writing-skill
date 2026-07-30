# 06 — Assemble（5 节结构组装）

> **阶段**：阶段 2
> **输入**：所有上游产物（intake / abstract / IR / claims / description / formulas / vsdx + PNG）
> **输出**：最终 .docx

## 目标

基于 `assets/templates/专利空白模版_fixed.docx` 的 5 节骨架，把所有内容注入对应节，输出最终 Word 文档。

## 输入清单

| 文件 | 内容 | 注入节 |
|---|---|---|
| `intake.json` | 案件元数据 | 文档属性（不写入正文） |
| `abstract_page.md` | 扉页摘要正文 | 节 1 |
| `flowchart_ir.json` | 流程图 IR | 节 2 |
| `assets/flowcharts/{case_id}.vsdx` | vsdx 源文件 | 附带交付 |
| `assets/flowcharts/{case_id}.png` | PNG 嵌入图 | 节 2 |
| `claims.json` | 权利要求列表 | 节 3 |
| `description.md` | 说明书正文 | 节 4 |
| `formulas.json` | OMML 公式列表 | 节 4（替换 `<<<FORMULA:...>>>`） |

## 组装流程

```
1. 复制 assets/templates/专利空白模版_fixed.docx → assets/cases/{case_id}/{case_id}.docx
2. 节 1：注入扉页摘要正文（占位段）
3. 节 2：嵌入 PNG（中心对齐，宽度 ≤ 5.5 inch）
4. 节 3：注入权利要求 5 条（自动编号 1、2、3、4、5）
5. 节 4：注入说明书 5 小节（替换 FORMULA 占位为 OMML）
6. 节 5：保留为占位段（"图 N  待贴附：..."）
7. 附带交付：assets/flowcharts/{case_id}.vsdx
8. 写文档属性：title / author / created
```

## 硬约束

### 节 1 — 扉页摘要

- 单段，首行缩进 2 字符
- 正文 14 pt + 宋体
- 字数 ≤ 300

### 节 2 — 摘要附图

- 嵌入 PNG，宽度 ≤ 5.5 inch
- 居中
- 段前段后各空 1 行

### 节 3 — 权利要求书

- 独立权利要求（id=1）：拆段 = 前导段（顶格，不缩进）+ 步骤段（每段缩进 2 字符）
- 从属权利要求（id=2-5）：每条 1 段，不缩进
- 自动编号 `1、` `2、` `3、`...
- 从属权利要求引用前一条 id

### 节 4 — 说明书

- 5 个小节标题居中或左对齐（沿用模板）
- 5 个小节标题段不缩进
- 5 个小节正文段首行缩进 2 字符
- 公式段前后各加 1 个空段
- 公式段落继承 `line=360 lineRule=auto`

### 节 5 — 说明书附图

- 仅占位段：`图 N  待贴附：XXX`
- N 从 2 开始（因为图 1 在节 2 用了）
- 占位段首行不缩进

## 字号与字体继承

所有正文段必须继承 `docDefaults`：
- ascii = Times New Roman
- eastAsia = 宋体
- sz = 14 pt

**不得**给 run 显式设置不同字体/字号（除非是页眉页脚）。

## 验证清单（生成后必跑）

> 注：以下为人工检查项，无自动化脚本。生成后用 Word/WPS 打开确认。

- [ ] 5 节齐全，页眉文字正确
- [ ] 节 2 含 1 张流程图（vsdx 转 PNG）
- [ ] 节 4 含 OMML 可编辑公式（双击验证）
- [ ] 节 3 权利要求 1-5 条，独立权利要求含步骤编号

## 收尾

- 最终文件：`assets/cases/{case_id}/{case_id}.docx`
- 附带：`assets/cases/{case_id}/{case_id}.vsdx`
- 进入 `prompts/99_self_check.md`