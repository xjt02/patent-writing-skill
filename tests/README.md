# 阶段 5 — 压力场景测试报告

## 概况

| # | 场景 | 输入 | 期望 | 状态 |
|---|---|---|---|---|
| T1 | 无公式摘要 | 200 字指针仪表摘要 + 流程图 IR + 3 条权利要求 | 5 节齐全 + 流程图嵌入 | ✓ |
| T2 | 含公式摘要 | 61 行说明书（含 2 个 OMML 公式）+ IR + claims | OMML 嵌入节 4 + 流程图嵌入 | ✓ |
| T3 | 模糊摘要 | 缺技术领域的 50 字摘要 | 主动追问，不强行生成 | ✓（协议已写） |

## T1 — 无公式摘要

**输入**：
- 摘要（145 字，< 300）：见 `tests/output/T1/T1.docx` 第 1 段
- 流程图：`examples/sample_001/sample_001.png`（vsdx 同源）
- 权利要求：`examples/sample_002/claims.json`（1 独立 + 2 从属）
- 附图占位："检测流程示意图" + "系统架构框图"

**输出**：`tests/output/T1/T1.docx`（45 KB）

**验证**：
- [x] 5 节齐全
- [x] 5 个页眉全部正确（含节 4 修正后的 `说    明    书`）
- [x] 节 2 含 1 个 `<w:drawing>`
- [x] 节 3 3 条权利要求
- [x] 摘要 145 字（< 300）

## T2 — 含公式摘要

**输入**：
- 摘要（136 字）：与 T1 同主题但精简
- 流程图：sample_001.png
- 权利要求：sample_002 claims
- 说明书：`${SKILL_ROOT}/AppData/Local/Temp/desc_test.md`（61 行）
- 公式：LOSS（交叉熵）+ SOFTMAX（10 类输出归一化）

**输出**：`tests/output/T2/T2.docx`（约 50 KB）
**中间产物**：`tests/output/T2/formulas.json`（2 条 OMML 字符串）

**验证**：
- [x] OMML 元素数：2（`oMath` 2 个，`</m:oMath>` 2 个）
- [x] 数学元素：m:f 2 个 / m:nary 4 个 / m:sSub 5 个 / m:limUpp 2 个
- [x] 流程图嵌入：1 个 `<w:drawing>`
- [x] 公式段前后各有空段（`extract_formulas.py` + `md_to_docx.py` 共同保证）
- [x] OMML 含 `xmlns:m` 命名空间（mathml2omml 包装层注入）

**Word 打开验证**（人工）：
- 公式段落双击 → Word 公式编辑器激活 → 可编辑 ✓
- 节 2 流程图 → 1080×613 px PNG 居中 ✓

## T3 — 模糊摘要追问协议

**协议位置**：`prompts/00_intake.md` 末尾「摘要 4 要素检查」节

**4 要素 + 判定关键词**：

| # | 要素 | 关键词 |
|---|---|---|
| 1 | 技术领域 | 属于、涉及、应用、应用于、本发明公开 |
| 2 | 解决问题 | 解决、针对、改善、提高、降低、克服 |
| 3 | 技术方案 | 包括、采用、通过、基于、步骤、特征在于 |
| 4 | 有益效果 | 可用于、适用于、因此、有助于 |

**追问规则**：
- 任一要素缺 → 触发 `AskUserQuestion`，每次只问 1 项
- header 字段直接用要素名（"技术问题" / "技术方案" 等）
- 提供 2-4 个常见选项 + "其他" 让用户填
- ❌ 不编造、不合并问题、不跳过

**测试模拟**：

输入：「一种基于 Transformer 的图像分类方法，准确率高。」

缺 3 项：技术领域、解决问题、有益效果。**第 1 轮**追问技术领域 → 用户填「计算机视觉」→ **第 2 轮**追问解决问题 → 用户填「准确率不足」→ **第 3 轮**追问有益效果 → 用户填「可应用于医学影像」→ 进入 prompts/01。

## 关键发现

1. **md_to_docx.py 段/节切换**已稳：模板里节 5 sectPr 直接挂 body，`_find_sect_end` 双模式正常
2. **paragraph + 图片**：先 `doc.add_paragraph` 落到末尾，再 `body.remove` + `sect2_end_p.addprevious` 移动。**不能直接 `Paragraph(lxml_elem, None)`**——需要 doc instance
3. **公式提取抓取条件**：`</math>` 闭合优先（最稳健），`<<END>>` / `<<FORMULA:*>>` 兜底
4. **claim_formatter 兼容**：`[{...}]` 与 `{"claims":[...]}` 都接受

## 已知限制

- 端到端流水线仍需手动调 6 个工具（未串成单条命令）
- 公式 OMML 在 WPS / LibreOffice 渲染可能与 Word 略有差异（Word 双击可编辑是核心要求）
- T3 追问协议未跑真实对话（依赖 AskUserQuestion 在主对话中触发）
