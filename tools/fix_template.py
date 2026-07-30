"""
fix_template.py
===============

修正"专利空白模版.docx"的已知问题：
1. 第四节（说明书）页眉 `权    利    说    明    书 ` → `说    明    书`
2. 第一节首页页眉删除 `MP1208682` 残留
3. 第一节首页页脚删除 `10002` 与 `2002.8` 残留

用法：
    python tools/fix_template.py <input.docx> -o <output.docx>

实现：
- docx 是 ZIP 容器，本脚本直接操作 word/header*.xml 与 word/footer*.xml
- 不依赖 python-docx 内部机制
- 保持其他 part 不变
"""
from __future__ import annotations

import argparse
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


def _text_of(elem) -> str:
    """w:t 的文本。"""
    return "".join(t.text or "" for t in elem.iter(f"{W}t"))


def fix_header6_section4(xml_bytes: bytes) -> tuple[bytes, int]:
    """
    修正 header6.xml 的页眉：
    - 旧文本："权 利 说 明 书"（分散多 run）
    - 新文本："说    明    书"

    返回 (新 XML, 是否修改)
    """
    root = etree.fromstring(xml_bytes)
    changed = False
    for p in root.findall(f".//{W}p"):
        runs = p.findall(f"{W}r")
        if not runs:
            continue
        full = "".join(_text_of(r) for r in runs)
        # 节 4 页眉特征：含"权 利 说 明 书"，且不含"要求"
        if ("权" in full and "利" in full and "说" in full
            and "明" in full and "书" in full
            and "要求" not in full):
            # 把首个 run 的所有 w:t 清空，重新写入"说    明    书"
            first_run = runs[0]
            for t in first_run.findall(f"{W}t"):
                first_run.remove(t)
            new_t = etree.SubElement(first_run, f"{W}t")
            new_t.set(
                "{http://www.w3.org/XML/1998/namespace}space", "preserve"
            )
            new_t.text = "说    明    书"
            # 删除后续 run
            for r in runs[1:]:
                r.getparent().remove(r)
            changed = True
            break  # 只改第一个匹配的段落
    new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return new_xml, int(changed)


def fix_header3_residue(xml_bytes: bytes) -> tuple[bytes, int]:
    """删除 header3.xml 中含 'MP120' 的段落。"""
    root = etree.fromstring(xml_bytes)
    removed = 0
    for p in list(root.findall(f".//{W}p")):
        full = _text_of(p)
        if "MP120" in full or "8682" in full:
            p.getparent().remove(p)
            removed += 1
    new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return new_xml, removed


def fix_footer3_residue(xml_bytes: bytes) -> tuple[bytes, int, int]:
    """
    清理 footer3.xml：
    - 删除含 '10002' 的 run
    - 删除含 '2002.8' 的整个段落
    """
    root = etree.fromstring(xml_bytes)
    removed_runs = 0
    removed_paras = 0
    for r in list(root.findall(f".//{W}r")):
        full = _text_of(r)
        if "10002" in full:
            r.getparent().remove(r)
            removed_runs += 1
    for p in list(root.findall(f".//{W}p")):
        full = _text_of(p)
        if "2002.8" in full:
            p.getparent().remove(p)
            removed_paras += 1
    new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return new_xml, removed_runs, removed_paras


def main(argv=None):
    parser = argparse.ArgumentParser(description="修正专利空白模版的已知错误")
    parser.add_argument("input", type=Path, help="输入 docx")
    parser.add_argument("-o", "--output", type=Path, required=True,
                        help="输出 docx 路径")
    args = parser.parse_args(argv)

    if not args.input.is_file():
        sys.exit(f"输入文件不存在: {args.input}")

    args.output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(args.input, args.output)

    # 读取所有 entry
    with zipfile.ZipFile(args.output, "r") as zin:
        info_list = zin.infolist()
        entries = {info.filename: (zin.read(info.filename), info) for info in info_list}

    # 处理目标 entry
    fixed_header6 = 0
    removed_h3 = 0
    removed_f3_runs = 0
    removed_f3_paras = 0

    new_entries = {}
    for name, (data, info) in entries.items():
        if name == "word/header6.xml":
            data, fixed_header6 = fix_header6_section4(data)
        elif name == "word/header3.xml":
            data, removed_h3 = fix_header3_residue(data)
        elif name == "word/footer3.xml":
            data, removed_f3_runs, removed_f3_paras = fix_footer3_residue(data)
        new_entries[name] = (data, info)

    # 写出新 zip
    with zipfile.ZipFile(args.output, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, (data, info) in new_entries.items():
            new_info = zipfile.ZipInfo(filename=name, date_time=info.date_time)
            new_info.compress_type = info.compress_type
            new_info.external_attr = info.external_attr
            zout.writestr(new_info, data)

    print(f"输出: {args.output}")
    print(f"  修正页眉: {fixed_header6} 处")
    print(f"  删除残留: header3 段 {removed_h3} 个, "
          f"footer3 run {removed_f3_runs} 个, "
          f"footer3 段 {removed_f3_paras} 个")


if __name__ == "__main__":
    main()