---
name: patent-writing-skill
description: Use when the user provides a Chinese invention patent abstract and wants a Word document conforming to the 5-section Chinese invention patent layout (扉页摘要 / 摘要附图 / 权利要求书 / 说明书 / 说明书附图), with editable Office Math formulas (OMML), Visio vsdx flowcharts embedded in section 2, prescribed fonts (宋体 + Times New Roman), 14pt body, 1.5× line spacing, 2-character first-line indent, and independent headers/footers per section. Triggers on phrases like "发明专利"、"专利文档"、"权利要求书"、"说明书"、"摘要附图"、"OMML 公式"、"vsdx 流程图"、"专利写作"。Do not use for patent mining, prior-art search, or non-Chinese patent jurisdictions.
---

# patent-writing-skill

## Overview

把一段中文发明专利摘要 → 输出符合国标的 5 节结构 Word 文档。**所有 prompts/ 与 tools/ 已就位**——本 skill 是入口与编排层，**不重新发明工具**。

**核心原则**：
1. **模板继承**：用 `assets/templates/专利空白模版_fixed.docx` 作底，不重置 docDefaults / sectPr / 页眉页脚
2. **OMML 可编辑**：公式走 `mathml2omml` → `<m:oMath>`，Word 双击可改
3. **vsdx 强制**：流程图必须 .vsdx + .png 双产物（vsdx 手工微调，PNG 嵌入节 2）
4. **5 节顺序固定**：扉页摘要 → 摘要附图 → 权利要求书 → 说明书 → 说明书附图

## When to Use

- 用户输入摘要 + 要求"生成发明专利 Word"
- 用户输入 IR + 要求"画流程图并嵌入专利"
- 用户输入权利要求 JSON + 要求"写到 docx"
- 用户输入含 MathML 公式的说明书 + 要求"可编辑 OMML 公式"

**不要用**于：专利挖掘、查新、交底书、PCT 国际申请、外观专利、实用新型。

## Pipeline（6 步）

每步 `Read` 对应 prompt 即可，详见 `prompts/`：

| # | Prompt | 输入 | 输出 |
|---|---|---|---|
| 0 | `prompts/00_intake.md` | 用户摘要 + 案件元数据 | `intake.json` |
| 1 | `prompts/01_abstract_page.md` | 摘要 | `abstract_page.md`（≤300 字，4 要素） |
| 2 | `prompts/02_flowchart_ir.md` | 摘要 | `flowchart_ir.json`（visio-scientific-flowchart 兼容） |
| 3 | `prompts/03_claims.md` | 摘要 + intake | `claims.json`（1 独立 + 3-5 从属） |
| 4 | `prompts/04_description.md` | 摘要 + claims | `description.md`（5 小节） |
| 5 | `prompts/05_formula_omml.md` | description.md | `formulas.json`（`{name: omml_xml}`） |
| 6 | `prompts/06_assemble.md` | 全部上游 | 调 `tools/md_to_docx.py` 出最终 .docx |

每步完成后跑 `prompts/99_self_check.md`（不入正文）。

## Tools

8 个工具，全部在 `tools/`。**首选脚本路径**——不要用 LLM 直接生成 docx 字符串。

| 工具 | 用途 | 何时用 |
|---|---|---|
| `extract_style.py` | 模板 → `style.json` | 阶段 1（已就位） |
| `fix_template.py` | 修正模板（节 4 页眉、残留清理） | 一次性（已跑过） |
| `style_applier.py` | `check` 校验 / `apply` 应用 style.json | 自检阶段 |
| `mathml_to_omml.py` | MathML → OMML 转换器 | `extract_formulas.py` 内部调 |
| `extract_formulas.py` | `description.md` → `formulas.json` | 阶段 5 |
| `claim_formatter.py` | `claims.json` 校验 / 编号预览 | 阶段 3 自检 |
| `insert_vsdx.py` | IR → vsdx + PNG + 嵌入节 2 | 阶段 2 → 6 之间 |
| `md_to_docx.py` | 5 节结构组装主入口 | 阶段 6（终步） |

调用约定：
- 全用 `python tools/<name>.py ...`（路径相对 skill 根）
- 入参/出参 schema 见各 prompt 末尾「收尾」节
- 失败时**第一手看 stderr**，不要重试 3 次就放弃

## 端到端示例

```bash
cd ${SKILL_ROOT}/.claude/skills/patent-writing-skill

# 阶段 0：模板就位
ls assets/templates/  # 专利空白模版_fixed.docx 应在

# 阶段 2+6：流程图嵌入
python tools/insert_vsdx.py \
  examples/sample_001/flowchart_ir.json \
  assets/templates/专利空白模版_fixed.docx \
  /tmp/out --case-id sample_001

# 阶段 5：公式提取
python tools/extract_formulas.py /tmp/desc.md -o /tmp/formulas.json

# 阶段 3：权利要求校验
python tools/claim_formatter.py validate examples/sample_002/claims.json

# 阶段 6：终步组装
python tools/md_to_docx.py \
  --template assets/templates/专利空白模版_fixed.docx \
  --case-id mycase \
  --abstract-text "本发明..." \
  --claims mycase_claims.json \
  --flowchart-png mycase.png \
  --description mycase_desc.md \
  --formulas mycase_formulas.json \
  --placeholders "检测流程示意图" \
  --output mycase.docx
```

完整端到端样例：`examples/sample_001/`（vsdx 嵌入）、`examples/sample_002/`（5 节组装）。

## 硬约束（不可破坏）

### 文档结构

- 恰好 **5 节**，节序：扉页摘要 / 摘要附图 / 权利要求书 / 说明书 / 说明书附图
- 节 4 页眉 = `说    明    书`（已修正模板）
- 节 3 页眉 = `权    利    要    求    书 `（含尾随空格，国标惯例）
- 节 5 页眉 = `说    明    书    附    图`

### 排版

- 正文：`docDefaults/rPrDefault/rPr/rFonts/ascii=Times New Roman`、`eastAsia=宋体`、`sz=28`（14 pt）
- 行距：`line=360 lineRule=auto`（1.5×）
- 首行缩进：`firstLine=200 firstLineChars=200`（2 字符）
- 页眉：`sz=32`（16 pt）
- 页脚：`sz=24`（12 pt），`line=200 lineRule=exact`

### 公式

- 必须用 `<m:oMath>`（**OMML**），不要 PNG
- 公式段前后各加 1 个空段
- 注入用 `tools/mathml_to_omml.py:inject_into_paragraph()`

### 流程图

- 必须 vsdx + PNG 双产物
- PNG 嵌入节 2，宽度 ≤ 5.5 inch，居中
- vsdx 通过 `visio-scientific-flowchart.to_vsdx()` 生成

完整规范见 `references/docx_style_guide.md`（14 节硬约束清单）。

## 已知坑

- **Visio 进程残留**：`Documents.Open` 可能因 VISIO.EXE 残留报文件共享冲突 → `insert_vsdx.py` 内 `cleanup_visio()` 自动杀
- **mathml2omml 导出名**：包导出的函数是 `convert`，本 skill 的 `mathml_to_omml.py` 包装成 `mathml_to_omml()`
- **节 5 sectPr 直接挂 body**：`_find_sect_end` 双模式（段内 sectPr / body 直挂 sectPr）
- **claims.json schema 兼容**：裸数组 `[{...}]` 与 `{"claims":[...]}` 都接受
- **公式占位 `<<FORMULA:NAME>>`**：单独占一行，MathML 块以 `</math>` 闭合

## 反例（不要做）

- ❌ 用 `document-skills:docx` 套件（**不支持 OMML**）
- ❌ 用 `python-docx` 手动拼 `<m:oMath>`（已封装在 `mathml_to_omml.py`，别重写）
- ❌ 把公式转 PNG 嵌入（用户要求"可编辑"）
- ❌ 重置 docDefaults / sectPr / 页眉页脚（破坏模板继承）
- ❌ LLM 直接拼 docx 字符串（用 `md_to_docx.py`）
- ❌ 跳过 `prompts/99_self_check.md`（27 项检查很重要）

## 进阶

- 端到端流水线尚未完全脚本化，目前需要 6 步手动调（阶段 2-6）
- vsdx 端到端需要 Windows + Visio COM + Python 3.10+ + pywin32
- OMML 在 WPS / LibreOffice 可能渲染差异（Word 最稳）
