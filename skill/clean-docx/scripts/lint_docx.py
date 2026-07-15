# -*- coding: utf-8 -*-
"""docx 合规体检 CLI —— 对任意 .docx 按《字体排版规范 v1》的 11 条硬规则打分。

用法：
    python lint_docx.py 某文档.docx [更多.docx ...]
    python lint_docx.py --en 纯英文文档.docx     # 强制按纯英文（跳过中文字体项）

退出码：所有文件均满分(无 FAIL) 时为 0，否则为 1，方便接进自动验收。
体检逻辑在 lib/docx_style.py 的 lint()，与生成时调用的是同一套规则。
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "lib"))
import docx_style as ds          # noqa: E402
from docx import Document        # noqa: E402

_MARK = {"PASS": "✓", "FAIL": "✗", "SKIP": "–"}


def lint_file(path, english=None):
    doc = Document(path)
    rep = ds.lint(doc, english=english)
    print(f"\n=== {os.path.basename(path)} ===")
    print("  文档类型：", "纯英文（中文字体项已跳过）" if rep["english"] else "中文/中英混排")
    for it in rep["items"]:
        line = f"  {_MARK[it['status']]} [{it['id']:>2}] {it['name']}"
        if it["status"] == "FAIL":
            line += f"  期望 {it['expected']}，实测 {it['actual']}"
        elif it["status"] == "SKIP":
            line += f"  （{it['actual']}）"
        print(line)
    fails = [it for it in rep["items"] if it["status"] == "FAIL"]
    verdict = "合规" if not fails else f"不合规（{len(fails)} 项未过）"
    print(f"  ── 得分 {rep['score']}/{rep['total']}"
          f"（另跳过 {rep['skipped']} 项）→ {verdict}")
    return not fails


def main(argv):
    english = None
    paths = []
    for a in argv:
        if a in ("--en", "--english"):
            english = True
        else:
            paths.append(a)
    if not paths:
        print("用法：python lint_docx.py 文档.docx [...] [--en]")
        return 2
    all_ok = True
    for p in paths:
        if not os.path.isfile(p):
            print(f"\n!! 找不到文件：{p}")
            all_ok = False
            continue
        try:
            all_ok &= lint_file(p, english=english)
        except Exception as e:  # noqa: BLE001
            print(f"\n!! 读取失败 {p}：{e}")
            all_ok = False
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
