---
name: tyust-thesis-formatter
description: >-
  Format existing Word (.docx) thesis drafts into compliant Taiyuan University of
  Science and Technology (太原科技大学) doctoral or master's degree theses based on
  the 2026 TYUST Graduate Thesis Writing Specification (GB/T 7713.1-2025 &
  GB/T 7714-2025) and official Word skeleton templates. Use whenever the user asks
  to format, typeset, convert, or audit a TYUST doctoral/master's thesis (理工类 /
  人文社科类 / 学术学位 / 专业学位 / 英文撰写), including cover, title pages,
  odd-page chapter breaks, headers/footers, three-line tables, equations, figures,
  footnotes, GB/T 7714-2025 references, TOC refresh, and spine generation.
---

# 太原科技大学研究生学位论文自动排版 Skill（2026 规范版）

将用户已有的旧论文 `.docx` 稿件无损解析并按《太原科技大学研究生学位论文撰写规范（2026 年 3 月）》与 6 套官方母版模板，输出格式完全合规、目录与页码已自动刷新的 Word 学位论文（及可选书脊文件）。

> **核心原则**：LLM 仅负责定型、补全元数据、复核结构映射与规范化参考文献；正文文字、图片、公式（OMML / MathType）、表格与脚注一律由确定性 Python 脚本无损搬运与精确排版，最后经 Word/WPS COM 自动刷新目录并导出 PDF 视觉核验。

---

## 模板选型矩阵

| 学位层次 (`degree_level`) | 学位类型 (`degree_type`) | 学科体例 (`discipline`) | 选用母版 (`resources/templates/`) |
| :--- | :--- | :--- | :--- |
| 博士 (`doctor`) | 学术学位 (`academic`) | 理工类 (`science`) | `doc_sci_aca.docx` |
| 博士 (`doctor`) | 专业学位 (`professional`) | 理工类 (`science`) | `doc_sci_pro.docx` |
| 硕士 (`master`) | 学术学位 (`academic`) | 理工类 (`science`) | `mas_sci_aca.docx` |
| 硕士 (`master`) | 专业学位 (`professional`) | 理工类 (`science`) | `mas_sci_pro.docx` |
| 硕士 (`master`) | 学术学位 (`academic`) | 人文社科类 (`humanities`) | `mas_hum_aca.docx` |
| 硕士 (`master`) | 专业学位 (`professional`) | 人文社科类 (`humanities`) | `mas_hum_pro.docx` |
| 博士 (`doctor`) | 学术/专业 | 人文社科类 (`humanities`) | 载入 `doc_sci_*.docx` + 切换人文社科章节编号与脚注规则 |
| 博士/硕士 | 学术/专业 | 英文撰写 (`language="en"`) | 载入对应学位母版 + 切换英文标题/页眉/前置顺序规则 |

---

## 五步执行工作流

### Step 1：定型与元数据确认

1. 确定 5 项配置：
   - `degree_level`: `master` | `doctor`
   - `degree_type`: `academic`（学术学位）| `professional`（专业学位）
   - `discipline`: `science`（理工类 `1 / 1.1 / 1.1.1`）| `humanities`（人文社科类 `一、/（一）/1.`）
   - `language`: `zh`（默认）| `en`
   - `bilingual_captions`: `false`（默认仅中文图题/表题）| `true`（同时输出英文图题/表题）
2. 先运行 Step 2 的提取脚本自动抓取旧论文中已有的题名、学科专业、摘要等字段，再参照 [examples/metadata.example.json](./examples/metadata.example.json) 生成 `metadata.json`。
3. 若缺少封面/扉页必填项（姓名、导师及职称、培养单位、中图分类号、UDC、提交/答辩日期、答辩委员会主席、英文扉页字段等），明确向用户询问或先填入醒目占位符并提醒用户核对。详见 [references/cover-and-frontmatter.md](./references/cover-and-frontmatter.md)。

### Step 2：无损解析旧论文（`extract_thesis.py`）

```powershell
python "<SKILL_DIR>\scripts\extract_thesis.py" `
  --input "<旧论文.docx>" `
  --out-dir "<work_dir>\extracted" `
  --discipline science
```

输出：
- `<work_dir>\extracted\thesis_ir.json`：结构化中间表示（封面元数据候选、中英文摘要、关键词、符号说明、章节树、段落富文本/OMML公式/脚注、三线表数据、图及图题、参考文献、附录、学术成果、致谢）
- `<work_dir>\extracted\media\*`：原论文提取的高清图片与嵌入资产
- `<work_dir>\extracted\extract_report.json`：结构识别摘要与低置信度提示（`review_hints`）

**Agent 语义复核要点（必做，禁止篡改原正文文字）**：
1. 检查 `extract_report.json` 中的 `review_hints`：确认各级标题切分是否准确（尤其是旧稿中未套用 `Heading` 样式的章节标题、章末小结、大纲行内冒号段落、方程表格等）。
2. 检查公式表示：若旧稿含纯文本公式（`review_hints` 会明确提示数量），在 `thesis_ir.json` 对应 `equation` 块中补写标准 `"latex"` 或多行 `"latex_lines"` 字段；原稿已带 Word 原生 OMML 或 MathType OLE 对象（`<w:object>`）的公式直接保留即可。
3. 检查英文摘要：若旧稿缺少 `ABSTRACT` 与 `Keywords`，按规范要求（英文标点后空 1 个半角空格、关键词间半角分号 `; ` 分隔、末尾无标点）补写或提示用户确认。
4. 检查参考文献：按 [references/gbt7714-2025.md](./references/gbt7714-2025.md) 核对文献标识代码（`[M]/[J]/[D]/[C]/[R]/[S]/[P]/[EB/OL]` 等）、序号 `[n]` 后 1 个半角空格以及缺项标注。
5. 如需微调结构，直接修改 `thesis_ir.json`（保留 `omml`、`ole_object`、`image`、`footnote_ref` 等无损段字段）。

### Step 3：确定性排版生成（`build_thesis.py`）

```powershell
python "<SKILL_DIR>\scripts\build_thesis.py" `
  --ir "<work_dir>\extracted\thesis_ir.json" `
  --meta "<work_dir>\metadata.json" `
  --output "<输出论文.docx>" `
  --spine-output "<输出书脊.docx>"
```

脚本严格按照 [references/spec-matrix.md](./references/spec-matrix.md) 与 [references/template-quirks.md](./references/template-quirks.md) 装配：
- 以官方清洗母版为底座（完整保留校徽、封面/中英文扉页排版及原创性与授权声明）；
- 摘要、ABSTRACT、目录、符号说明（如有）、每一章、参考文献、附录（如有）、学术成果、致谢全部使用真正的 Word `oddPage` 分节符；
- 奇数页页眉为本章标题，偶数页页眉为 `太原科技大学博士/硕士学位论文`（宋体/TNR 五号 10.5pt，单实线下框线）；
- 页码严格采用规范口径：**底端居中、Times New Roman 五号（10.5pt）、无修饰线**；前置（摘要～目录）用大写罗马数字 `I, II, III...`，正文第一章起用阿拉伯数字 `1, 2, 3...`；
- 标题采用官方模板居中一级标题（可通过 `--h1-align both` 切换为规范文字的两端对齐），`outlineLvl` 与 Word `TOC \o "1-3" \h \u` 完全联动；
- 三线表（顶底 1.5pt、表头下 0.75pt、无竖线、跨页重复表头）、图表按章规范编号与防分离（`keep_with_next`）、脚注每页重编 `①②③`；
- 公式采用 **无边框三列定位表**（总宽 16.0 cm：`1.85 / 12.30 / 1.85 cm`，中间居中放公式、右列右对齐放 `（章-序）`，相邻公式表自动隔断防合并）：`latex`/`latex_lines` 经 `latex2mathml` + 本机 Office `MML2OMML.XSL` 自动编译为 Word 原生可编辑 OMML（Cambria Math 12pt），原稿自带 OMML 与 MathType OLE 无损保留，正文行内变量默认自动转换为规范斜体与真上下标（可用 `--no-inline-math` 关闭）。

### Step 4：Word / WPS COM 刷新目录并导出预览（`finalize_word.py`）

```powershell
python "<SKILL_DIR>\scripts\finalize_word.py" `
  --docx "<输出论文.docx>" `
  --pdf "<输出论文.pdf>" `
  --preview-dir "<work_dir>\preview"
```

- 自动优先调用本机 Microsoft Word COM（失败时自动降级 WPS `Kwps.Application`）；
- 强制刷新所有域与自动目录（`TablesOfContents.Update()`），回存 `.docx`；
- 导出真实分页后的 `.pdf`，并使用 `pymupdf` 渲染封面、扉页、声明、摘要、目录、章节首页、图表公式页与参考文献页的 PNG 预览图。

### Step 5：自动化质检 + 视觉核验（`audit_thesis.py`）

```powershell
python "<SKILL_DIR>\scripts\audit_thesis.py" `
  --docx "<输出论文.docx>" `
  --pdf "<输出论文.pdf>" `
  --degree master `
  --discipline science `
  --report "<work_dir>\audit_report.json"
```

1. 先看 `audit_report.json` 中的全部硬检项（页边距、页眉页脚距、奇偶页分节、页眉与章标题一致性、页码格式与字号、字体/字号/行距合规率、三线表检查、参考文献数量与正文 `[n]` 闭环、目录页码是否刷新）。
2. 再使用 `read_file` 查看 `<work_dir>\preview\` 中关键页面截图，重点检查：
   - 封面与中英文扉页是否各严丝合缝占满单页（未撑到下一页）；
   - 目录前导符是否右对齐撑满版心、罗马与阿拉伯页码正确；
   - 图表与图题/表题是否同页、无表格文字挤爆、公式编号是否右齐；
   - 奇数页起始章节页眉正确、无意外孤行。
3. 发现任何瑕疵立即调整参数或 `thesis_ir.json` 并重跑 Step 3–5，直至全部通过。

---

## 按需查阅的参考文档

- 全量字号/行距/缩进/页边距参数表：[references/spec-matrix.md](./references/spec-matrix.md)
- 封面、中英文扉页、声明与书脊规则：[references/cover-and-frontmatter.md](./references/cover-and-frontmatter.md)
- 理工类 / 人文社科类 / 英文类差异：[references/discipline-variants.md](./references/discipline-variants.md)
- GB/T 7714-2025 参考文献著录细则：[references/gbt7714-2025.md](./references/gbt7714-2025.md)
- 官方模板与规范冲突及默认裁决说明：[references/template-quirks.md](./references/template-quirks.md)
