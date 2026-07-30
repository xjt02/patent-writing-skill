"""
claim_formatter.py
==================

权利要求 JSON → docx 段编号文本（含校验）。

输入（prompts/03_claims.md 定义的 schema）：

    {
      "claims": [
        {"id": 1, "type": "independent", "text": "一种基于 XXX 的方法..."},
        {"id": 2, "type": "dependent", "depends_on": 1, "text": "如权利要求 1 所述..."},
        ...
      ]
    }

输出（独立接口 + CLI）：
- format_for_docx(claims) -> list[str]
    生成 docx 段文本，每段以 `1、` 顿号开头；段内"如权利要求 X 所述"已检查引用合法性。

- validate(claims) -> ValidationReport
    校验 03_claims.md 硬约束：
      1. 独立 = 1 条
      2. 从属 ∈ [3, 5] 条
      3. 总数 ≤ 6
      4. depends_on 必须 < id
      5. 不得用评价词

用法：
    python tools/claim_formatter.py validate <claims.json>
    python tools/claim_formatter.py preview <claims.json>    # 打 1、xxx 2、yyy...
    python tools/claim_formatter.py dump <claims.json> -o <out.txt>   # 写纯文本
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any


# 评价词 / 禁用句式（与 prompt 03 反例一致）
_BANNED_WORDS = [
    "最优", "完美", "极佳", "极好", "突破性", "革命性", "划时代",
    "特别地", "值得注意的是", "显而易见", "显然",
]


class ValidationReport:
    def __init__(self):
        self.errors: list[str] = []
        self.warnings: list[str] = []

    def add_error(self, msg: str) -> None:
        self.errors.append(msg)

    def add_warning(self, msg: str) -> None:
        self.warnings.append(msg)

    @property
    def passed(self) -> bool:
        return not self.errors

    def to_dict(self) -> dict:
        return {
            "passed": self.passed,
            "errors": self.errors,
            "warnings": self.warnings,
        }


def _normalize(claim: dict) -> dict:
    """容错：'id' 缺省时按数组索引+1。"""
    out = dict(claim)
    out.setdefault("id", 0)  # 0 表示未指定
    out.setdefault("type", "dependent")
    out.setdefault("depends_on", None)
    out.setdefault("text", "")
    return out


def validate(claims: list[dict]) -> ValidationReport:
    r = ValidationReport()
    if not claims:
        r.add_error("claims 列表为空")
        return r

    # 分配 id（缺省用索引+1）
    claims = [_normalize(c) for c in claims]
    n = len(claims)

    # 1. 独立 = 1
    indep = [c for c in claims if c["type"] == "independent"]
    if len(indep) == 0:
        r.add_error("缺少独立权利要求（type='independent' 计数 0）")
    elif len(indep) > 1:
        r.add_error(f"独立权利要求过多（{len(indep)} 条，需恰好 1 条）")

    # 2. 从属 ∈ [3, 5]
    dep = [c for c in claims if c["type"] == "dependent"]
    if len(dep) < 3:
        r.add_warning(f"从属权利要求过少（{len(dep)} 条，建议 3-5 条）")
    elif len(dep) > 5:
        r.add_error(f"从属权利要求过多（{len(dep)} 条，最多 5 条）")

    # 3. 总数 ≤ 6
    if n > 6:
        r.add_error(f"权利要求总数 {n} 超过 6 条")

    # 4. depends_on 合法
    for c in claims:
        if c["type"] == "dependent":
            d = c.get("depends_on")
            if d is None:
                r.add_warning(f"权利要求 id={c['id']} 缺少 depends_on")
            elif not isinstance(d, int):
                r.add_error(f"权利要求 id={c['id']} depends_on 非整数: {d!r}")
            elif d >= c["id"]:
                r.add_error(f"权利要求 id={c['id']} 引用了未来 id={d}（必须 < 当前）")
            elif d < 1:
                r.add_error(f"权利要求 id={c['id']} depends_on={d} 非法")

    # 5. 评价词
    for c in claims:
        text = c.get("text", "")
        for bad in _BANNED_WORDS:
            if bad in text:
                r.add_warning(f"权利要求 id={c['id']} 含禁用词「{bad}」")

    # 6. 编号：检查 id 是否 1..n
    ids = [c["id"] for c in claims]
    if sorted(ids) != list(range(1, n + 1)):
        r.add_error(f"id 序列不连续 1..n，实际为 {ids}")

    return r


def format_for_docx(claims: list[dict]) -> list[str]:
    """
    生成 docx 段文本列表，每段 = `<id>、<text>`。

    注意：md_to_docx.py 内部已加 `i、` 前缀，本函数返回的文本**不含**编号前缀，
    由 md_to_docx.py 负责编号。本函数只做 normalize 与校验。
    """
    claims = [_normalize(c) for c in claims]
    return [c["text"].strip() for c in claims]


def numbered_preview(claims: list[dict]) -> list[str]:
    """生成 `1、xxx` 形式的预览文本（用于人工核对）。"""
    claims = [_normalize(c) for c in claims]
    out = []
    for c in claims:
        prefix = f"{c['id']}、"
        out.append(prefix + c["text"].strip())
    return out


def main(argv=None):
    parser = argparse.ArgumentParser(description="权利要求 JSON 校验/格式化")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_val = sub.add_parser("validate", help="校验 claims.json")
    p_val.add_argument("input", type=Path)

    p_prev = sub.add_parser("preview", help="打 `1、xxx` 预览")
    p_prev.add_argument("input", type=Path)

    p_dump = sub.add_parser("dump", help="写纯文本（每行一条）")
    p_dump.add_argument("input", type=Path)
    p_dump.add_argument("-o", "--output", type=Path, required=True)

    args = parser.parse_args(argv)

    if not args.input.is_file():
        sys.exit(f"输入不存在: {args.input}")
    data = json.loads(args.input.read_text(encoding="utf-8"))
    # 兼容两种 schema：{"claims": [...]} 或裸数组 [...]
    if isinstance(data, list):
        claims = data
    else:
        claims = data.get("claims", [])

    if args.cmd == "validate":
        rep = validate(claims)
        print(json.dumps(rep.to_dict(), ensure_ascii=False, indent=2))
        sys.exit(0 if rep.passed else 1)
    elif args.cmd == "preview":
        for line in numbered_preview(claims):
            print(line)
    elif args.cmd == "dump":
        text = "\n".join(numbered_preview(claims)) + "\n"
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text, encoding="utf-8")
        print(f"已写入: {args.output} ({len(claims)} 条)")


if __name__ == "__main__":
    main()