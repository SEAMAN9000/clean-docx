# clean-docx

> 让 AI 生成的中文 / 中英混排 Word 文档（.docx）自动套上一套统一的学术排版——中文宋体、
> 西文 Times New Roman、标题黑体、三线表、自动页码，生成即合规，不用事后手修。

📖 English: [README.en.md](./README.en.md)

快速安装（全平台通用，需 Node.js；手动方式见[安装](#安装)）：

```bash
npx skills add SEAMAN9000/clean-docx -g
```

## 什么时候用

AI 直接生成的 Word 文档排版常常失控：英文数字被套上中文字体、代码字体混进正文、
同一篇文档里几种字体打架——每次生成都得手工返工。本 skill 把一套定稿的排版规范做成
函数库和检查器：生成时自动套用，保存前自动体检。

✅ 适合：

- 让 AI 写中文报告、论文、综述、方案、教材、备忘等，直接产出 .docx
- 把 Markdown / 散稿转成排版统一的 Word 文档
- 手上有一份排版乱掉的 .docx，想打分找问题、或一键把字体规整回来
- 自己写 python-docx 脚本，想直接复用一套排好的版式函数

❌ 不适合：

- 纯英文论文——用目标期刊自带的模板
- PDF / Excel / PPT——不归本 skill 管，用各自的专用工具
- 想要自定义主题或品牌视觉——本 skill 有意钉死一套规范（见「设计取舍」）

## 用起来长什么样

阅读约定：`>` 开头的是你打的字，其余是屏幕上的回显。

```text
> 把《唐诗里的四季》这篇稿子整理成 Word 文档

（Claude 自动调起 clean-docx，用其函数库生成文档，保存前自动体检）

已生成 唐诗里的四季.docx（3 页）：
  结构审查 review：0 条警告
  格式体检 lint：10/10 合规（另跳过 1 项）
```

也可以在命令行给任何现成的 .docx 打分（以下为真实输出）：

```text
$ python skills/clean-docx/scripts/lint_docx.py 唐诗里的四季_案例.docx
=== 唐诗里的四季_案例.docx ===
  文档类型： 中文/中英混排
  ✓ [ 1] 正文字号 12pt
  ✓ [ 2] 行距 1.5 倍
  ✓ [ 3] 标题中文字体 黑体
  ✓ [ 4] 一级标题字号 15pt
  ✓ [ 5] 表格三线表
  ✓ [ 6] 页脚页码
  ✓ [ 7] 页面 A4 + 2.5cm 边距
  ✓ [ 8] 正文首行缩进 2 字符
  ✓ [ 9] 一级标题段前后距
  – [10] 篇幅对应的页眉/目录  （约 1 页）
  ✓ [11] 中文 run 均设字体(不回退)
  ── 得分 10/10（另跳过 1 项）→ 合规
```

你做过的事：说一句要什么，其余全自动。文档内容和检查明细会因稿件而异，
稳定不变的是流程（生成 → 结构审查 → 格式体检 → 保存）和这套排版规则。

排版效果（三页演示文档，内容仅作排版展示，图表数字为虚构示例）：

| 第 1 页 | 第 2 页 | 第 3 页 |
|---|---|---|
| ![示例第1页](demo/page1.png) | ![示例第2页](demo/page2.png) | ![示例第3页](demo/page3.png) |

## Runtime 支持

### ✅ Claude Code

- 安装路径：`~/.claude/skills/clean-docx/`
- 项目记忆文件：`CLAUDE.md`
- 触发：语义自动触发（见「触发」），无斜杠命令

### ✅ Codex CLI

- 本 skill 是标准 `SKILL.md` + 脚本结构，没有 Claude 专用功能，可直接用
- 安装路径：`~/.agents/skills/clean-docx/`（用户级）或项目内 `.agents/skills/clean-docx/`
  （保持文件名 `SKILL.md` 不变）；项目记忆文件是 `AGENTS.md`

### ✅ OpenCode

- OpenCode 兼容 `~/.claude/skills/` 路径，按 Claude Code 的装法装即可
- 项目记忆文件是 `AGENTS.md`

## 安装

通用依赖（所有平台都要）：

```bash
pip install python-docx
```

### 一键安装（推荐，全平台通用）

需要 Node.js（`npx` 随它附带，作用是临时下载并运行一个小工具）。一行命令自动检测
你电脑上装了哪些 AI 编程工具（Claude Code / Codex / OpenCode / Cursor 等 70+ 种），
把 skill 装到各工具对应的位置：

```bash
npx skills add SEAMAN9000/clean-docx
```

- 默认装到当前项目；加 `-g` 装到用户全局（推荐，所有项目可用）
- 只想装给某个工具：加 `-a`，如 `-a claude-code -a opencode`
- 安装器来自 [vercel-labs/skills](https://github.com/vercel-labs/skills)

装完验证：新开会话，说「用 clean-docx 生成一页测试文档」——能产出 .docx 并报告体检得分即装好。

### 手动安装：Claude Code

```bash
git clone https://github.com/SEAMAN9000/clean-docx.git
mkdir -p ~/.claude/skills
cp -r clean-docx/skills/clean-docx ~/.claude/skills/
```

Windows（PowerShell）：

```powershell
git clone https://github.com/SEAMAN9000/clean-docx.git
New-Item -ItemType Directory -Force "$env:USERPROFILE/.claude/skills"
Copy-Item -Recurse clean-docx/skills/clean-docx "$env:USERPROFILE/.claude/skills/"
```

### 手动安装：Codex CLI

```bash
git clone https://github.com/SEAMAN9000/clean-docx.git
mkdir -p ~/.agents/skills
cp -r clean-docx/skills/clean-docx ~/.agents/skills/
```

只想装给某一个项目，就把最后两行的 `~/.agents/skills` 换成该项目根目录下的
`.agents/skills`——Codex 会从当前目录一路往仓库根找这个文件夹。

装完验证：同上一句测试话术。若调不起，`scripts/` 下两个脚本可脱离 skill 独立使用
（见「当函数库直接用」）。

### 手动安装：OpenCode

与 Claude Code 完全相同——OpenCode 兼容 `~/.claude/skills/` 路径。

## 触发

无斜杠命令，靠语义自动触发。你的话里出现下列意思（对应 `SKILL.md` 的 description）即触发：

| 你说的话里包含 | English equivalent |
|---|---|
| 写报告 / 论文 / 综述 / 方案 / 教材 / 备忘 / 通知 等任何中文 Word 文档 | write a Chinese report / paper / review / proposal / teaching material / memo as Word |
| 排版 / 字体 / 格式统一 | typesetting / fonts / consistent formatting |
| 生成 docx / Word 文档 | generate a .docx / Word document |
| 把内容另存为 / 导出为 Word 或 docx | save / export the content as Word |
| md / Markdown 转 Word / 转 docx | convert Markdown to Word |
| 用 python-docx 生成 / 拼一个 Word 文档 | build a Word document with python-docx |

不会触发：纯英文论文、PDF、Excel、PPT，以及不落到 .docx 的泛泛排版讨论。
想强制调起：明说「用 clean-docx 技能……」。

## 跑完之后

流程（全自动，中途不需要你点头）：

1. 触发后，AI 导入 `scripts/docx_style.py` 函数库生成文档，不手搓 python-docx 细节；
2. 保存前自动跑两道互补检查：`review` 查结构（标题层级 / 序号连续 / 图表交叉引用 / 公式），
   `lint` 按 11 条硬规则查格式；
3. 检查不过不拦截——文档照常保存，同时告诉你哪条没过，改不改由你定。

产物：只有你要的那份 `输出.docx`（要求转 PDF 时多一份 PDF）。

改动边界：只写你指定的输出文件；不改你项目里的其他文件，不动 skill 自身目录。

失败语义：生成中途报错就没有产物（至多留下可直接删掉的半个文件），重说一次即可重来；
lint 不合规不算失败，见上。

注意：页码和目录是 Word「域」（打开文档后按 F9 或右键「更新域」刷新）；
用脚本转 PDF 前，先让 Word 更新域。

## 当函数库直接用

不经 AI，自己的 Python 脚本里也能用：

```python
import sys; sys.path.insert(0, "skills/clean-docx/scripts")
import docx_style as ds

doc = ds.new_document()                  # A4 + 2.5cm 页边距
ds.add_page_number(doc)
ds.add_title(doc, "报告标题", "English Subtitle")
ds.add_heading(doc, "一、研究背景", 1)
ds.add_body(doc, [("关键数字 ", False), ("2.9%", True), (" 会加粗。", False)])
ds.add_table_title(doc, "表1 示例")
ds.add_three_line_table(doc, [["指标", "2023", "2024"], ["A", "1", "2"]])
print(ds.review(doc))                    # 结构审查：警告列表
print(ds.lint(doc)["score"])             # 格式体检：得分
doc.save("输出.docx")
```

改已有文档、或在别处手写了文字片段之后，保存前调一次 `ds.enforce_fonts(doc)`
可强制规整全篇字体（专治中文回退到默认字体）。全部函数见库文件顶部注释，
排版规则定稿见 [skills/clean-docx/references/spec.md](skills/clean-docx/references/spec.md)。

## 仓库结构

- `skills/clean-docx/` —— skill 包本体（安装就复制这个）
- `demo/` —— 排版效果示例图

## 设计取舍

- **钉死一套规范，而不是可配置主题。** 换来零决策、每次生成都一致；代价是不合口味时，
  要自己改 `scripts/docx_style.py` 顶部的常量（字体字号都集中在那里）。
- **函数库事前保证，`enforce_fonts` 只是兜底。** 排版对错在生成那一刻就定下来，而不是事后修补；
  代价是绕过函数库手写内容就可能乱——所以才留了兜底和体检两道防线。
- **lint 不合规不硬拦。** 文档总能先拿到手；代价是不合规的文档也存得出来，修不修由你定。
- **页码、目录用 Word 域而不写死。** 内容再改页码也不会错；代价是打开文档要按一次 F9。
- **已知局限：** 只管中文 / 中英混排，纯英文论文请用期刊模板；页眉 / 目录「按篇幅决定」
  基于页数估算，边界情况可能判偏，生成时明说要或不要即可覆盖。

## License

MIT — 详见 [LICENSE](./LICENSE)
