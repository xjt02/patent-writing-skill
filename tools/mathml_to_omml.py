"""
mathml_to_omml.py
=================

MathML → OMML 转换器（包装 mathml2omml）+ python-docx 注入辅助。

用法：
    # 命令行：把 MathML 文件转 OMML
    python tools/mathml_to_omml.py convert <input.mathml> [-o output.xml]

    # 命令行：把公式注入到 docx（替换指定段）
    python tools/mathml_to_omml.py inject <docx> <para_index> <mathml> -o <output>

    # Python API：
    from mathml_to_omml import mathml_to_omml, inject_into_paragraph
    omml_str = mathml_to_omml("<math>...</math>")
    inject_into_paragraph(doc, paragraph, omml_str)

依赖：
    pip install mathml2omml
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

from lxml import etree

try:
    import mathml2omml
except ImportError:
    print("缺少 mathml2omml，请先运行: pip install mathml2omml", file=sys.stderr)
    raise


# --- 命名空间 ---
NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def mathml_to_omml(mathml_str: str) -> str:
    """
    把 MathML 字符串转 OMML 字符串。

    输入：
        <math xmlns="..."><mrow><mi>L</mi><mo>=</mo><mn>1</mn></mrow></math>

    输出：
        <m:oMath>...</m:oMath>

    异常：
        mathml2omml.MathMLHandler 异常
    """
    if not mathml_str.strip():
        raise ValueError("MathML 输入为空")
    return mathml2omml.convert(mathml_str)


def parse_omml(omml_str: str) -> etree._Element:
    """把 OMML 字符串解析为 lxml 元素。

    注意：mathml2omml 输出的 OMML 不含 xmlns:m 声明；
    这里显式注入 nsmap 以便 lxml 能正确解析 m: 前缀。
    """
    # 注入 xmlns:m 声明
    if "xmlns:m" not in omml_str:
        omml_str = omml_str.replace(
            "<m:oMath",
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
            1,
        )
    return etree.fromstring(omml_str)


def inject_into_paragraph(paragraph, omml_str: str) -> None:
    """
    把 OMML 公式追加到指定 paragraph 的末尾。

    paragraph: python-docx Paragraph 对象
    omml_str: OMML 字符串（<m:oMath>...</m:oMath>）
    """
    omml_elem = parse_omml(omml_str)
    # python-docx 暴露 paragraph._p（CT_P lxml 元素）
    # 把 OMML 直接附加为子节点，但需要包在 <m:r> 内（Word 要求 oMath 出现在 run-level）
    # 实际上 oMath 可以作为 paragraph 的直接子节点（属于 math run），不需要包 <w:r>
    # 但 document.xml 中既有 run-level oMath 也有 paragraph-level oMathPara
    # 我们这里用 paragraph-level：把 oMath 直接加到 <w:p> 下
    paragraph._p.append(omml_elem)


def inject_into_paragraph_via_run(paragraph, omml_str: str) -> None:
    """
    备用方案：把 OMML 包在 <m:r> 内再附加。
    适用于某些 Word 版本对 oMath 在 paragraph 下的渲染差异。
    """
    omml_elem = parse_omml(omml_str)
    # 创建 <m:r> 包装
    new_r = etree.SubElement(paragraph._p, f"{M}r")
    new_r.append(omml_elem)


def main(argv=None):
    parser = argparse.ArgumentParser(description="MathML → OMML 转换与注入")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_conv = sub.add_parser("convert", help="MathML 文件转 OMML 字符串")
    p_conv.add_argument("input", type=Path)
    p_conv.add_argument("-o", "--output", type=Path)

    p_inj = sub.add_parser("inject", help="把公式注入到 docx 的指定段")
    p_inj.add_argument("docx", type=Path)
    p_inj.add_argument("para_index", type=int,
                       help="0-based 正文段索引")
    p_inj.add_argument("mathml_file", type=Path,
                       help="包含 MathML 的文件")
    p_inj.add_argument("-o", "--output", type=Path, required=True)

    args = parser.parse_args(argv)

    if args.cmd == "convert":
        if not args.input.is_file():
            sys.exit(f"输入不存在: {args.input}")
        mathml = args.input.read_text(encoding="utf-8")
        omml = mathml_to_omml(mathml)
        if args.output:
            args.output.write_text(omml, encoding="utf-8")
            print(f"已写入 {args.output}")
        else:
            print(omml)

    elif args.cmd == "inject":
        if not args.docx.is_file():
            sys.exit(f"输入 docx 不存在: {args.docx}")
        if not args.mathml_file.is_file():
            sys.exit(f"MathML 文件不存在: {args.mathml_file}")

        from docx import Document
        doc = Document(str(args.docx))
        mathml = args.mathml_file.read_text(encoding="utf-8")
        omml = mathml_to_omml(mathml)

        paragraphs = doc.paragraphs
        if args.para_index >= len(paragraphs):
            sys.exit(f"para_index {args.para_index} 超出范围（共 {len(paragraphs)} 段）")

        para = paragraphs[args.para_index]
        inject_into_paragraph(para, omml)
        doc.save(str(args.output))
        print(f"已注入 OMML 到段 {args.para_index}，输出: {args.output}")


if __name__ == "__main__":
    main()