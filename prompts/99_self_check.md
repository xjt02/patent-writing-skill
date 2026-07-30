# 99 — Self Check（自检清单）

> **阶段**：阶段 2（最末）
> **目的**：在交付前自动校验产物；**不写入文档**

## 自检项

### 结构（必须全过）

- [ ] 文档恰好 5 节
- [ ] 节 1 页眉 = `说    明    书    摘    要`
- [ ] 节 2 页眉 = `摘    要    附    图`
- [ ] 节 3 页眉 = `权    利    要    求    书 `（尾随空格）
- [ ] 节 4 页眉 = `说    明    书`（**已修正**）
- [ ] 节 5 页眉 = `说    明    书    附    图`
- [ ] 每节页脚含 PAGE 域（页码），从 1 重启

### 字体（继承自 docDefaults）

- [ ] `docDefaults/rPrDefault/rPr/rFonts/ascii` = `Times New Roman`
- [ ] `docDefaults/rPrDefault/rPr/rFonts/eastAsia` = `宋体`
- [ ] `docDefaults/rPrDefault/rPr/sz` = `28`（14 pt）

### 排版

- [ ] 正文段 `firstLine=200 firstLineChars=200`
- [ ] 正文段 `line=360 lineRule=auto`
- [ ] 页脚段 `line=200 lineRule=exact`
- [ ] 页眉段 `sz=32 szCs=28`（16 pt）
- [ ] 页脚页码 `sz=24 szCs=24`（12 pt）

### 内容

- [ ] 节 1 扉页字数 ≤ 300
- [ ] 节 1 段数 = 1
- [ ] 节 2 含至少 1 个嵌入图片（drawing/inline）
- [ ] 节 3 权利要求条数 ∈ [4, 6]
- [ ] 节 3 含 1 条独立权利要求
- [ ] 节 4 含 5 个小节标题（"技术领域""背景技术""发明内容""附图说明""具体实施方式"）
- [ ] 节 4 总字数 ∈ [2000, 3000]
- [ ] 节 4 含至少 1 个 OMML 公式（如说明书中出现公式占位）
- [ ] 节 5 含至少 2 个占位段

### 公式（如适用）

- [ ] 每个 `<<<FORMULA:...>>>` 都被 OMML 替换（无残留）
- [ ] OMML 含 `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"`
- [ ] 公式段前后各有 1 个空段
- [ ] 用 Word/WPS 打开 → 双击公式 → 可编辑（人工验收）

### 流程图（如适用）

- [ ] vsdx 文件已附带（`assets/cases/{case_id}/{case_id}.vsdx`）
- [ ] PNG 嵌入节 2，宽度 ≤ 5.5 inch
- [ ] 流程图 IR 节点数 ≤ 12
- [ ] 决策节点有 2 条出边

## 自检输出

```json
{
  "case_id": "<YYYYMMDDHHmmss>",
  "passed": true | false,
  "checks": [
    {"name": "structure.section_count", "passed": true, "expected": 5, "actual": 5},
    {"name": "structure.header.section_4", "passed": true, "expected": "说    明    书", "actual": "说    明    书"},
    ...
  ],
  "errors": [
    "节 4 缺少 OMML 公式，但 description.md 含 3 个公式占位"
  ],
  "warnings": [
    "说明书总字数偏少（1800 字，目标 2000-3000）"
  ]
}
```

## 失败处理

若 `passed=false`，输出错误清单，由调用方决定：
- 重跑对应 prompt（如 06_assemble）
- 人工干预修改后重新 assemble
- 放弃该案件

## 不入正文

> **重要**：本自检清单的输出**不得**写入最终 .docx，仅用于调试和验收。