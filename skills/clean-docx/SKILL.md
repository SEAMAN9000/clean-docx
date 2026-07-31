---
name: clean-docx
description: >-
  生成中文或中英混排的 Word 文档（.docx）时，套用统一的字体排版规范，避免每次格式都不一样。
  规则要点：中文用宋体、英文和数字用 Times New Roman 且两者绝不混用；标题用黑体；正文小四 12pt、
  1.5 倍行距；加粗只给标题和关键数字；表格一律三线表；固定加页码，页眉和目录按文档篇幅自动决定。
  只要用户要写报告、论文、综述、方案、教材、备忘、通知等任何中文 Word 文档，或提到「排版」「字体」
  「格式统一」「生成 docx/Word 文档」，就主动使用本技能——即使用户没明说要套规范。同样覆盖这些近义场景：
  「把内容另存为/导出为 Word 或 docx」「把这些（中文）内容单独存成一个 docx/Word 文件」「md / Markdown
  转 Word / 转 docx」「用 python-docx 生成/拼一个 Word 文档」「顺手再附带出一份 docx / 附属文档」——
  这些只要产出物是中文 .docx 就该走本技能。英文说法同样触发：write a Chinese report / paper /
  review / proposal / teaching material / memo as Word；typesetting / fonts / consistent
  formatting；generate a .docx / Word document；save / export the content as Word；convert
  Markdown to Word；build a Word document with python-docx。硬规则：**凡是准备动手写 python-docx 代码、或调用
  anthropic-skills:docx 生成中文/中英混排文档，先停下来走本技能**，别图快手搓脚本（那正是排版会乱的根源）。
  它建立在 anthropic-skills:docx 之上，负责其中「排版好看且统一」这一环。不处理纯英文论文（用期刊自带模板）、
  PDF、Excel、PPT。
---

# 中文文档统一排版（学术规范版）

## 什么时候必须用本技能（触发硬线）

只要最终产物是一份中文/中英混排的 .docx，就走本技能——不管用户怎么措辞。除了「写报告/论文/方案」
这种明说，下面这些**容易被忽略的场景同样必须触发**：

- 「把这些内容另存/导出为 Word / docx」「单独出一个 docx 文件」「附带再生成一份 docx」
- 「md / Markdown 转 Word / 转 docx」
- **你自己准备动手写 python-docx 代码**去拼一个中文文档时——把「开始写 python-docx」这个动作
  本身当成触发信号。为图快临时手搓 python-docx 脚本、绕过本技能，正是排版乱掉的典型根源。

一句话：**别在没走本技能的情况下手写 python-docx 生成中文文档。** 真要手写或改已有文档，
也至少用下面的 `enforce_fonts` 兜底、用 `lint` 验收。

## 这个技能解决什么

Claude 生成 Word 文档时，常出现排版不统一：英文数字被套上中文字体、用代码字体当正文、
中文没设字体走默认、一篇里混用好几种西文字体。本技能用一套固定规则消除这些毛病，
让任何中文文档排版一致、耐看。规则定稿见 `references/spec.md`。

## 核心做法：用现成的工具函数库，别手搓 XML

设中文 eastAsia 字体、画三线表、插页码这些都要操作底层 XML，容易出错。
本技能在 `scripts/docx_style.py` 里把它们全封装好了。**生成文档时优先 import 这个库**，
而不是重新写 python-docx 细节。它的函数都已内置规范里的字体、字号、行距。

```python
import sys
sys.path.insert(0, "<本技能目录>/scripts")  # 换成 SKILL.md 所在目录下的 scripts
import docx_style as ds

doc = ds.new_document()                 # A4 + 2.5cm 页边距
ds.add_page_number(doc)                 # 页码固定加

# 按预估页数和一级标题数，自动决定要不要页眉/目录
layout = ds.decide_layout(page_estimate=8, num_h1=4)
if layout["header"]:
    ds.add_header(doc, "唐诗里的四季")

ds.add_title(doc, "唐诗里的四季",
             "The Four Seasons in Tang Poetry")
if layout["toc"]:
    ds.add_toc(doc)                     # 目录（打开后按 F9 更新）

ds.add_heading(doc, "一、为什么按四季读唐诗", 1)
ds.add_body(doc, [("《全唐诗》收诗近 ", False), ("5 万", True), (" 首。", False)])
ds.add_heading(doc, "1.1 主要参考", 2)
ds.add_bullets(doc, ["《全唐诗》", "《唐诗三百首》"])
ds.add_table_title(doc, "表1 四季题材篇数（示例数据）")
ds.add_three_line_table(doc, [["季节", "篇数", "占比"],
                              ["春", "62", "31%"],
                              ["秋", "71", "35.5%"]])
doc.save("输出.docx")
```

各函数说明见库文件顶部注释，都很短、可直接读。

## 规范铁律（写正文时也要守）

即使不用函数库、需要手写或改已有文档，也要遵守这几条——它们是「乱」与「不乱」的分界：

1. **中英文字体分开**：中文（含标点）走宋体/黑体，英文字母和阿拉伯数字走 Times New Roman。
   绝不把英文数字套上中文字体，也绝不把中文套上英文字体。
2. **不拿代码字体当正文**：Consolas、Courier、等宽字体是给代码用的，正文一律不用。
3. **加粗克制**：只加粗标题，以及正文里真正要强调的关键词或数字（如「2.9%」）。不给整句整段加粗。
4. **表格用三线表**：只留顶、底、表头下三条横线，没有竖线；表内字比正文小一号；表头加粗、内容居中。
5. **图表标注位置**：表标题在表上方，图标题在图下方。

## 字号速查

| 部位 | 中文字体 | 西文/数字 | 字号 |
|---|---|---|---|
| 大标题 | 黑体 | Times New Roman | 18pt |
| 一级标题 | 黑体 | Times New Roman | 15pt |
| 二级标题 | 黑体 | Times New Roman | 13pt |
| 正文 | 宋体 | Times New Roman | 12pt（小四） |
| 表内文字 | 宋体 | Times New Roman | 11pt |
| 页码 | 宋体 | Times New Roman | 10.5pt（五号） |

正文 1.5 倍行距、首行缩进 2 字符；页面 A4、四边 2.5cm。

## 页眉 / 目录：按文档大小决定

用 `ds.decide_layout(page_estimate, num_h1)` 自动判断，规则：

- 短文档（约 ≤5 页，如周报、备忘、单篇）：页眉、目录都不加。
- 中等（约 6–15 页，多个一级标题，如报告、综述、方案）：加页眉；一级标题 ≥3 个就加目录。
- 长文档（约 >15 页 或 分章节，如教材、书稿）：页眉、目录都加。

页码任何文档都加。目录和页码都是 Word「域」，文档打开后按 F9（或确认更新域）会刷新页码。
若用脚本转 PDF，先用 Word 更新域再导出。

## 修已有文档 / 手写了 run：用 enforce_fonts 兜底

`add_*` 系列函数生成的内容字体一定是对的。但如果你**改一份已有的 docx**，或在别处手写了 run，
就绕过了这些函数，中文容易回退到默认字体（这是最常见的「乱」）。这时在保存前调一次：

```python
n = ds.enforce_fonts(doc)   # 遍历每个文字片段，强制西文 Times、中文宋体/黑体（标题黑体）
print(f"已规整 {n} 个文字片段的字体")
```

它只改字体、不动字号和加粗，专治「中文回退」。

## 保存前先体检：结构(review) + 格式(lint)

两个检查互补，建议都跑：`review` 查**结构**、`lint` 查**格式**是否合规。

### 格式合规打分 lint

`ds.lint(doc)` 按《字体排版规范 v1》的 11 条硬规则给文档打分，返回 `{"score","total","items",...}`，
每条 PASS / FAIL / SKIP（不适用项如纯英文跳过中文字体、短文档不要求目录）。**FAIL 不必硬拦，
软提示即可**——告诉用户哪条没过、再决定是否修：

```python
rep = ds.lint(doc)
print(f"合规 {rep['score']}/{rep['total']}")
for it in rep["items"]:
    if it["status"] == "FAIL":
        print(f"  ✗ {it['name']}：期望 {it['expected']}，实测 {it['actual']}")
doc.save("输出.docx")          # 软提示：照常保存，把不合规项一并告诉用户
```

也可对任意现成 .docx 在命令行打分：`python scripts/lint_docx.py 某文档.docx`（纯英文加 `--en`）。
检查项即字体/字号/行距/三线表/页码/A4/缩进/页眉目录/中文 run 是否回退等 11 条。

### 结构审查 review

在 `doc.save()` 之前调用 `ds.review(doc)`，它返回一份警告列表，专查这几类容易出错的结构问题：

- **重复子标题**（同级同名标题出现多次）
- **标题层级跳级**或未从一级开始（如开头直接是二级、从一级跳到三级）
- **序号不连续**：一级『一、三、』缺号、二级『1.1 1.3』缺号、序号重复或倒退，
  以及二级标题章节号与所属一级标题不符（如第二章下出现『3.1』）
- **图/表编号不连续**：图1、图2…表1、表2… 缺号、重复、不从 1 开始
- **图表与标题数量不匹配**：有图却没图题、有表题却没对应表格
- **交叉引用失效**：正文写"见图5"却没有图5；图表定义了却从未在正文被引用
- **公式编号不连续 / 公式引用失效**：式(1)(3) 缺号、"由式(5)"却无 (5) 号公式
- **公式未渲染**：正文残留 `\frac`、`$...$`、`^{ }` 等 LaTeX 源码，没做成可编辑公式
- **空标题**、**把整句话当标题**（以句末标点结尾）
- **目录会为空**：插了目录(TOC)却没有任何带层级的标题

```python
for w in ds.review(doc):
    print("⚠", w)        # 有问题就提示用户或修正后再保存
doc.save("输出.docx")
```

注意：标题必须用 `ds.add_heading()` 生成，它会写入"大纲级别"，自动目录才收得到——
直接用 `add_body` 或手写加粗当标题，目录会是空的。

## 完成后

生成的 .docx 可用 anthropic-skills:docx 的方法校验。若需要给用户看效果图，
用 Word 转 PDF（记得先更新域）再转图片。
