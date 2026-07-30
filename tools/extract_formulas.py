"""
extract_formulas.py
===================

从 description.md 抓 `<<FORMULA:NAME>>` 占位，把其中嵌入的 MathML 转 OMML，
输出 formulas.json（dict by name）。

占位语法（在 description.md 文本中）：

    损失函数 <<FORMULA:LOSS>> 用于评估模型性能。

其中 <<FORMULA:LOSS>>> 占位单独占一行时（推荐）：

    <<FORMULA:LOSS>>

占位内必须是合法 MathML（不是 LaTeX 文本）。mathml2omml 仅做 MathML→OMML 转换。

用法：
    python tools/extract_formulas.py <description.md> -o <formulas.json>

输出 schema：
    {
      "F1": "<m:oMath xmlns:m='...'>...</m:oMath>",
      "F2": "<m:oMath ...>...</m:oMath>",
      ...
    }

兜底：转换失败时写 "valid": false 的占位提示，由 md_to_docx.py 跳过并打印警告。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

try:
    # 优先用同目录的本地包装（统一 mathml_to_omml 入口）
    sys.path.insert(0, str(Path(__file__).parent))
    from mathml_to_omml import mathml_to_omml
except ImportError:
    try:
        from mathml2omml import convert as mathml_to_omml
    except ImportError:
        print("缺少 mathml2omml，请先运行: pip install mathml2omml", file=sys.stderr)
        raise


# `<<FORMULA:NAME>>` 行（推荐，整行占位）或 `<<FORMULA:NAME>>` 嵌入文本
_FORMULA_LINE_RE = re.compile(r"^<<FORMULA:(\w+)>>\s*$")
_FORMULA_INLINE_RE = re.compile(r"<<FORMULA:(\w+)>>")

# 公式内容格式：占位后下一行到下一个 <<END>> / <<FORMULA:...>> / 文件尾
# 但更稳健的做法：让 <<FORMULA:NAME>>> 占位自闭合，里面直接放 MathML：
#   <<FORMULA:LOSS>><math>...</math><<END>>
# 不过为了和 prompt 05 的 few-shot 兼容，采纳最简形式：
#   <<FORMULA:NAME>>        ← 单独一行
#   <math>...</math>        ← 下一行起为 MathML 内容
#   <<END>>                  ← 结束（也允许 <<END FORMULA>> 之类）
# 为了实现简单，本工具只支持「整行占位 + 下一行起读到下一个 <<FORMULA:*>> 或文件尾」

_END_MARKERS = re.compile(r"^<<(END|FORMULA:\w+)>>")


def _extract_mathml_blocks(md_text: str) -> dict[str, str]:
    """
    抓出 (name, mathml_text) 对。

    格式（推荐）：
        <<FORMULA:F1>>
        <math xmlns="...">...</math>

        下一段文本 ...
        <<FORMULA:F2>>
        <math ...>...</math>
    """
    lines = md_text.splitlines()
    result: dict[str, str] = {}
    i = 0
    while i < len(lines):
        m = _FORMULA_LINE_RE.match(lines[i])
        if not m:
            i += 1
            continue
        name = m.group(1)
        i += 1
        # 抓 MathML 块：直到遇到 `</math>` 闭合标签，或下一个 `<<FORMULA:*>>`/`<<END>>` 标记
        # （`</math>` 优先：这是最稳健的终止条件）
        buf: list[str] = []
        while i < len(lines):
            line = lines[i]
            if _END_MARKERS.match(line.strip()):
                break
            buf.append(line)
            i += 1
            if "</math>" in line:
                break
        mathml = "\n".join(buf).strip()
        if mathml:
            result[name] = mathml
    return result


def convert_all(md_text: str) -> dict[str, dict]:
    """
    把 description.md 里的全部公式占位转换为 OMML。

    返回：{name: {"mathml": str, "omml": str, "valid": bool, "error": str|None}}
    """
    blocks = _extract_mathml_blocks(md_text)
    out: dict[str, dict] = {}
    for name, mathml in blocks.items():
        try:
            omml = mathml_to_omml(mathml)
            out[name] = {
                "mathml": mathml,
                "omml": omml,
                "valid": True,
                "error": None,
            }
        except Exception as e:
            out[name] = {
                "mathml": mathml,
                "omml": "",
                "valid": False,
                "error": f"{type(e).__name__}: {e}",
            }
    return out


def main(argv=None):
    parser = argparse.ArgumentParser(description="从 description.md 提取公式并转 OMML")
    parser.add_argument("input", type=Path, help="description.md 路径")
    parser.add_argument("-o", "--output", type=Path, required=True,
                        help="输出 formulas.json 路径")
    parser.add_argument("--strict", action="store_true",
                        help="遇失败公式就 exit 1")
    args = parser.parse_args(argv)

    if not args.input.is_file():
        sys.exit(f"输入不存在: {args.input}")

    md_text = args.input.read_text(encoding="utf-8")
    converted = convert_all(md_text)

    # 简化 schema：只保留 {name: omml} 便于 md_to_docx.py 直接用
    # 失败项保留 omml="" 并打 stderr 警告
    out: dict[str, str] = {}
    failures = []
    for name, info in converted.items():
        if info["valid"]:
            out[name] = info["omml"]
        else:
            out[name] = ""
            failures.append((name, info["error"]))
            print(f"  ⚠ 公式 {name} 转换失败: {info['error']}", file=sys.stderr)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(out, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"已写入: {args.output}")
    print(f"  公式数: {len(converted)} (成功 {len(converted) - len(failures)}, 失败 {len(failures)})")

    if args.strict and failures:
        sys.exit(1)


if __name__ == "__main__":
    main()