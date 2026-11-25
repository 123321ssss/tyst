# 太原科技大学研究生学位论文自动排版 Skill（TYUST Thesis Formatter）

面向 AI Coding Agent（Antigravity / 支持 `SKILL.md` 的 Agent）与命令行流水线的**太原科技大学（TYUST）博士、硕士学位论文自动排版工具**。

严格依据《太原科技大学研究生学位论文撰写规范（2026 年 3 月）》（参照 **GB/T 7713.1-2025** 与 **GB/T 7714-2025**），基于 6 套官方 Word 骨架模板，可将已有 `.docx` 论文初稿无损解析并输出格式完全合规、自动刷新目录与页码的学位论文（及独立书脊文件）。

---

## 核心特性

- **全学位与学科体例覆盖**：内置 6 套清洗后的官方母版（完整保留官方校徽、封面/中英文扉页排版与原创性/授权声明），支持：
  - 博士 / 硕士 × 学术学位 / 专业学位
  - **理工类**（`1` / `1.1` / `1.1.1`，序号后空 2 个半角空格）
  - **人文社科类**（`一、` / `（一）` / `1.`，独立引文格式、每页重编 `①②③` 脚注）
  - **英文撰写变体**（`Chapter 1` / `1.1` / `1.1.1`，英文摘要前置，英文双面页眉）
- **版面与奇偶页分节全自动**：
  - 长题名自动压缩前置空白行，确保封面、中文扉页、英文扉页、声明严格各占 1 页；
  - 摘要、`ABSTRACT`、目录、符号说明、正文每一章、参考文献、附录、学术成果、致谢全部采用真正的 Word `oddPage` 分节符；
  - 奇数页页眉为本章标题、偶数页页眉为 `太原科技大学博士/硕士学位论文`（宋体/TNR 五号 10.5pt，单实线下框线）；
  - 底端居中 Times New Roman 五号（10.5pt）无修饰线页码：前置部分大写罗马数字（`Ⅰ、Ⅱ、Ⅲ` / `I, II, III`），正文第 1 章起阿拉伯数字连续编排。
- **出版级公式排版（三引擎兼容）**：
  - **LaTeX → 原生可编辑 OMML**：通过 `latex2mathml` + 本机 Microsoft Office 自带 `MML2OMML.XSL` 编译为 Cambria Math 12pt Word 原生公式；
  - **无边框三列表定位**：总宽 `16.0 cm`（`1.85 / 12.30 / 1.85 cm`），公式居中、编号 `（章-序）` 严格右对齐，并自动隔离连续公式防 Word 合并；
  - **MathType OLE 无损搬运**：自动提取并重写关系 ID 无损保留旧稿中的 MathType `Equation.3` 对象；
  - **行内数学符号自动排版**：正文中的 `H_L`、`θ_L*`、`U_Z*`、`Δy_blur,max`、`cosθ_L`、`10^5` 等自动转为斜体变量与真上下标（可用 `--no-inline-math` 关闭），不影响固定值 20 磅行距。
- **三线表、高清插图与自动目录刷新**：
  - 国际通用三线表（顶底 1.5pt、表头下 0.75pt、无竖线、跨页重复表头）、`.emf`/`.wmf` 自动高分辨率转 PNG 嵌入；
  - 调用本机 Microsoft Word / WPS COM 两轮刷新 `TOC \o "1-3"` 目录并统一目录字体为小四宋体/TNR 不加粗、导出 PDF 并生成逐页截图；
  - 独立生成规范书脊 `spine.docx`（四号仿宋、16 磅行距、段前 40 磅、上下边距 5 cm）。

---

## 快速安装

### 1. 作为全局 Skill 安装（推荐）

将仓库克隆到 Agent 的全局 skills 目录即可在任意工作区直接调用：

```powershell
git clone https://github.com/123321ssss/tyst.git "$env:USERPROFILE\.gemini\config\skills\tyust-thesis-formatter"
```

（若仅用于当前项目，可克隆到项目下的 `.agents/skills/tyust-thesis-formatter`。）

### 2. 安装 Python 依赖

需要 **Python 3.9+**（Windows 环境下使用 Word/WPS COM 自动刷新目录及 `MML2OMML.XSL` 公式转换，建议本机已安装 Microsoft Office 或 WPS）：

```powershell
pip install -r requirements.txt
```

---

## 使用方法

### 方式一：在 AI Agent 对话中直接触发

安装完成后，在对话中把已有 `.docx` 论文交给 Agent 即可，例如：

> “使用 `tyust-thesis-formatter` 把 `draft.docx` 排成太原科技大学 **硕士 · 学术学位 · 理工类** 格式，作者张三、导师李四教授、机械工程学院，并生成书脊。”

Agent 会按 `SKILL.md` 自动执行：**元数据确认 → 无损解包与公式/文献复核 → 确定性排版生成 → Word/WPS COM 刷新目录导出 PDF → 合规体检与视觉核验**。

### 方式二：命令行独立运行 4 阶段脚本

```powershell
# 1. 无损解析旧论文为结构化 IR 与媒体资产
python scripts/extract_thesis.py --input draft.docx --out-dir work/extracted --discipline science

# 2. 复制并填写封面与扉页元数据（参考 examples/metadata.example.json）
#    如有纯文本公式，可在 work/extracted/thesis_ir.json 对应块中补充 "latex" 或 "latex_lines"

# 3. 基于官方清洗母版生成规范 .docx（及可选书脊）
python scripts/build_thesis.py `
  --ir work/extracted/thesis_ir.json `
  --meta work/metadata.json `
  --output work/thesis_formatted.docx `
  --spine-output work/spine.docx

# 4. 调用本机 Word/WPS COM 刷新目录域、导出 PDF 并渲染页面预览
python scripts/finalize_word.py `
  --docx work/thesis_formatted.docx `
  --pdf work/thesis_formatted.pdf `
  --preview-dir work/preview

# 5. 运行自动化合规体检
python scripts/audit_thesis.py `
  --docx work/thesis_formatted.docx `
  --pdf work/thesis_formatted.pdf `
  --degree master `
  --discipline science `
  --report work/audit_report.json
```

> **提示**：如未安装 Microsoft Office 或 `MML2OMML.XSL` 位于自定义路径，可设置环境变量 `TYUST_MML2OMML_XSL` 指定其完整路径；如需重制母版，可使用 `python scripts/prepare_templates.py --src-dir <官方模板目录>`。

---

## 目录结构

```text
.
├── SKILL.md                     # Skill 主指令与 5 步执行 SOP
├── README.md                    # 项目说明文档
├── requirements.txt             # Python 依赖清单
├── LICENSE                      # MIT 开源协议
├── examples/
│   └── metadata.example.json    # 封面与扉页元数据填写示例
├── references/                  # 2026 版撰写规范与国标速查
│   ├── cover-and-frontmatter.md
│   ├── discipline-variants.md
│   ├── gbt7714-2025.md
│   ├── spec-matrix.md
│   └── template-quirks.md
├── resources/
│   └── templates/               # 6 套清洗后的官方母版骨架 + 书脊模板
│       ├── doc_sci_aca.docx
│       ├── doc_sci_pro.docx
│       ├── mas_sci_aca.docx
│       ├── mas_sci_pro.docx
│       ├── mas_hum_aca.docx
│       ├── mas_hum_pro.docx
│       └── spine_template.docx
└── scripts/                     # 自动化提取、公式编译、装配、刷新与体检脚本
    ├── extract_thesis.py
    ├── math_engine.py
    ├── build_thesis.py
    ├── finalize_word.py
    ├── audit_thesis.py
    └── prepare_templates.py
```

---

## 免责声明

`resources/templates/` 中的骨架模板基于太原科技大学研究生院公开发布的 2026 版学位论文 Word 模板清洗整理，仅用于辅助学术排版与格式合规校验。最终提交前请务必对照学校与所在学院的最新通知核对内容、签名与格式细节。
