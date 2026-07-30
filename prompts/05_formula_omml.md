# 05 — Formula OMML（公式转换）

> **阶段**：阶段 2
> **输入**：description.md（含 `<<<FORMULA:...>>>` 占位）
> **输出**：formulas.json（OMML XML 字符串列表）

## 目标

把 description.md 中的 `<<<FORMULA:mathml>>>...<<<FORMULA>>>` 占位转换为 Word OMML（可双击编辑）。

## 占位语法

```
正文描述 <<<FORMULA:L=-\frac{1}{N}\sum_{i=1}^{N}y_i\log\hat{y}_i>>> 后续正文。
```

其中：
- `<<<FORMULA:` 开始标记
- `>>>` 结束标记
- 中间是 **MathML 内容字符串**（带 LaTeX 风格的 MathML）

## MathML 输入格式

接受 **LaTeX 风格但用 MathML 标签**的输入（不是纯 LaTeX）。例如：

```xml
<math xmlns="http://www.w3.org/1998/Math/MathML">
  <mrow>
    <mi>L</mi>
    <mo>=</mo>
    <mrow>
      <mo>-</mo>
      <mfrac>
        <mn>1</mn>
        <mi>N</mi>
      </mfrac>
    </mrow>
    <munderover>
      <mo>∑</mo>
      <mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow>
      <mi>N</mi>
    </munderover>
    <mrow>
      <msub><mi>y</mi><mi>i</mi></msub>
      <mo>log</mo>
      <msub><mover><mi>y</mi><mo>^</mo></mover><mi>i</mi></msub>
    </mrow>
  </mrow>
</math>
```

## 输出格式

```json
{
  "formulas": [
    {
      "id": "F1",
      "source_mathml": "<原始 MathML>",
      "omml_xml": "<m:oMath xmlns:m='...'>...</m:oMath>",
      "valid": true
    }
  ]
}
```

## 硬约束

1. OMML 必须使用 WordprocessingML 命名空间 `m:`
2. OMML 必须包含 `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"`
3. 每个公式输出独立 `<m:oMath>` 元素（不是 `<m:oMathPara>`）
4. 公式字体默认 Cambria Math（OMML 标准）
5. 公式段行距继承正文（`line=360 lineRule=auto`）
6. 公式段前后各加 1 个空段

## OMML 元素速查（实现参考）

| MathML 元素 | OMML 元素 |
|---|---|
| `<mfrac>` | `<m:f>` |
| `<munderover>` | `<m:limLow>` + `<m:limUpp>` 或 `<m:nary>` |
| `<msup>` / `<msub>` | `<m:sSup>` / `<m:sSub>` |
| `<msqrt>` | `<m:rad>` |
| `<mrow>` | `<m:e>` |
| `<mi>` `<mn>` `<mo>` | `<m:r><w:rPr>...</w:rPr><m:t>...</m:t></m:r>` |

**注意**：document-skills:docx 没有原生公式 API，必须用 `python-docx` 的 `OxmlElement` + `parse_xml` 手写注入。

## 兜底方案

如果 `mathml2omml` 包不可用或转换失败：
1. 把公式段替换为 `[公式 {id}：见说明书附图]`
2. 在节 5"说明书附图"占位添加对应图片位置
3. 在 self-check 中标记 `formula_skipped: true`

## Few-shot 示例

**输入占位**：
```
<<<FORMULA:\hat{y}_c=\frac{e^{z_c}}{\sum_{k=0}^{10}e^{z_k}}, \quad c=0,1,\ldots,10>>>
```

**输入 MathML**（用户提供）：
```xml
<math xmlns="http://www.w3.org/1998/Math/MathML">
  <mrow>
    <msub><mover><mi>y</mi><mo>^</mo></mover><mi>c</mi></msub>
    <mo>=</mo>
    <mfrac>
      <msup><mi>e</mi><msub><mi>z</mi><mi>c</mi></msub></msup>
      <mrow>
        <munderover><mo>∑</mo><mrow><mi>k</mi><mo>=</mo><mn>0</mn></mrow><mn>10</mn></munderover>
        <msup><mi>e</mi><msub><mi>z</mi><mi>k</mi></msub></msup>
      </mrow>
    </mfrac>
  </mrow>
</math>
```

**输出 OMML**（目标）：
```xml
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:sSub>
    <m:e><m:r><m:t>y&#x0302;</m:t></m:r></m:e>
    <m:sup><m:r><m:t>c</m:t></m:r></m:sup>
  </m:sSub>
  <m:r><m:t>=</m:t></m:r>
  <m:f>
    <m:num>
      <m:sSup>
        <m:e><m:r><m:t>e</m:t></m:r></m:e>
        <m:sup><m:sSub><m:e><m:r><m:t>z</m:t></m:r></m:e><m:sub><m:r><m:t>c</m:t></m:r></m:sub></m:sSub></m:sup>
      </m:sSup>
    </m:num>
    <m:den>
      <m:nary>
        <m:chr><m:r><m:t>∑</m:t></m:r></m:chr>
        <m:lim><m:r><m:t>k=0</m:t></m:r></m:lim>
        <m:lim><m:r><m:t>10</m:t></m:r></m:lim>
      </m:nary>
      <m:sSup>
        <m:e><m:r><m:t>e</m:t></m:r></m:e>
        <m:sup><m:sSub><m:e><m:r><m:t>z</m:t></m:r></m:e><m:sub><m:r><m:t>k</m:t></m:r></m:sub></m:sSub></m:sup>
      </m:sSup>
    </m:den>
  </m:f>
</m:oMath>
```

## 收尾

- 写到 `assets/cases/{case_id}/formulas.json`
- 用 `tools/mathml_to_omml.py` 真正做转换（在 tools/ 阶段实现）
- 进入 `prompts/06_assemble.md`