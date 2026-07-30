# 专利写作 Skill

将"摘要"输入 → 自动产出符合中国发明专利格式规范的 Word 文档+流程图vsdx格式原件。

## 5 节结构

| 节 | 页眉 | 内容 |
|---|---|---|
| 1 | 说    明    书    摘    要 | 扉页摘要 |
| 2 | 摘    要    附    图 | 摘要附图（vsdx 流程图） |
| 3 | 权    利    要    求    书 | 权利要求书（独立 + 从属） |
| 4 | 说    明    书 | 说明书（技术领域 / 背景技术 / 发明内容 / 附图说明 / 具体实施方式） |
| 5 | 说    明    书    附    图 | 说明书附图（人工贴图占位） |

## 排版规范

- 字体：宋体（东亚） + Times New Roman（ASCII）
- 字号：正文 14pt / 页眉 16pt / 页脚 12pt
- 行距：1.5 倍
- 首行缩进：2 字符
- 页眉字段分隔：4 个半角空格
- 公式：MathML → OMML（可编辑）
- 流程图：vsdx + PNG 双产物


## 目录

```
patent-writing-skill/
├── assets/         # 模板、style.json、流程图 IR
├── prompts/        # 8 个分步指令
├── references/     # 排版规范、专利法要点、句式库
├── tools/          # Python 工具链
├── tests/          # pytest 测试
└── examples/       # 测试用样本摘要，其中demo文件夹中展示了输出示例
```

## 依赖

### 必须

- Python ≥ 3.10
- Microsoft Windows（Visio COM 调用仅支持 Windows）
- Microsoft Visio（已安装，用于流程图生成）
- `pip install -r requirements.txt`

### 可选（流程图功能）

- `visio-scientific-flowchart` skill：需单独克隆到 `~/.claude/skills/visio-scientific-flowchart/`
  - 或者设置环境变量 `VISIO_SKILL_PATH` 指向该 skill 目录
  - 如果未安装，流程图生成步骤会跳过（其他功能不受影响）

### 环境变量

| 变量 | 用途 | 默认值 |
|---|---|---|
| `VISIO_SKILL_PATH` | 指向 `visio-scientific-flowchart` 目录 | `~/.claude/skills/visio-scientific-flowchart` |

## 开发状态

阶段 9.1 完成，v7 终稿验证通过。详见 plan 文件。

## Project Architecture

```
摘要文本
  ↓ [LLM + prompts/00-05]
intake.json / abstract_page.md / flowchart_ir.json / claims.json / description.md / formulas.json
  ↓ [tools/md_to_docx.py + tools/insert_vsdx.py + tools/extract_formulas.py]
符合 GB/T 的 Word 文档（5 节 + OMML 公式 + vsdx/PNG 流程图）
```

## Project Structure

```
patent-writing-skill/
├── SKILL.md               # Skill 主入口（供 Claude Code 加载）
├── README.md              # 本文件
├── requirements.txt      # Python 依赖
├── assets/
│   ├── style.json        # 排版属性（从模板反推）
│   ├── templates/
│   │   ├── 专利空白模版.docx          # 用户原模板
│   │   └── 专利空白模版_fixed.docx   # 修正后模板
│   └── cases/           # 测试用例输入/输出
├── prompts/             # 8 个分步 prompt（00-06, 99）
├── references/          # 5 个参考文档
├── tools/               # 8 个 Python 工具
└── examples/            # 2 个样本摘要
```

## Roadmap

| 版本 | 阶段 | 状态 |
|---|---|---|
| v1.0 | 阶段 0-6 | MVP：5 节骨架 + 模板修正 + prompts + tools |
| v2.0 | 阶段 7 | LLM 生成层端到端实战 |
| v3.0 | 阶段 8 | 说明书下划线标题 + MathML 过滤 |
| v4.0 | 阶段 8.1 | 小节段尾空行 + 节标题加粗 + 公式紧贴 |
| v5.0 | 阶段 9 | 内容结构重排 + 步骤命名统一 |
| v6.0 | 阶段 9.1 | 独立权利要求拆段 + 缩进 |
| **打包版** | 阶段 10 | 清理敏感信息 + 依赖文档化 |

## Limitations

| 限制项 | 说明 |
|---|---|
| **必须：Windows + Visio** | 流程图生成依赖 `visio-scientific-flowchart` skill + Microsoft Visio COM，无法在 Linux/macOS 运行 |
| **必须：python-docx** | 仅支持 `.docx` 格式，不支持 `.doc`（Word 97-2003） |
| **可选：Visio 流程图** | 不装 `visio-scientific-flowchart` 时跳过，其他功能正常 |
| **公式兼容性** | OMML 公式在 Microsoft Word 中可编辑；WPS/LibreOffice 兼容性取决于版本 |
| **权利要求上限** | 建议总数 ≤ 6 条（超过时 claim_formatter 会警告） |
| **中文专利格式** | 仅适配中国国家知识产权局发明专利格式，不支持实用新型/外观设计 |

## License

本 skill 仅供个人使用与内部研究。打包分发时请自行确保：

- 不包含真实案件摘要数据
- 示例摘要已做脱敏处理（发明人姓名 → "示例发明人"）
- 模板文件 `专利空白模版.docx` 版权归原模板提供者所有