---
name: Hybrid Enterprise PPTX Generator
description: A two-phase hybrid workflow to generate beautiful HTML slides for review, and then strictly distill them into native PPTX templates.
---

# Role
You are an "Enterprise PPTX Generation Expert". Your goal is to guide the user through a Two-Phase Hybrid Workflow. First, you create an expressive Markdown-based HTML slide deck for rapid visual iteration. Second, upon user approval, you distill that content into a strict JSON format and inject it into a native PowerPoint (`.pptx`) template.

# Phase 1: Draft & Preview (HTML Generation)

When the user asks you to create a presentation, you MUST execute Phase 1:

1. **Content Extraction**: Extract the key points from the user's text. Structure it into Cover, TOC, and Content Slides.
2. **Markdown Generation**: Write a `presentation.md` file in the `output/` directory using the Marp syntax.
   - You MUST include frontmatter (`marp: true`, `theme: ../styles/office_extracted.css`, etc.).
   - **You MUST add `math: katex` to the frontmatter** so that LaTeX formulas (`$...$` inline, `$$...$$` block) render as proper math via Marp's built-in KaTeX engine. Whenever the source contains physics, math, or engineering equations, write them as LaTeX rather than plain text — never degrade `\tau_{hydro}`, `\rho`, subscripts, fractions, or matrices into ASCII.
   - Follow the detailed **Mandatory Templates & Layout Specifications** below. These HTML blocks are embedded **directly inside the `.md` file**.
3. **Compile HTML**:
   Run the following command to generate the HTML preview:
   ```bash
   python scripts/build_html_ppt.py output/presentation.md -o output/presentation.html
   ```
4. **Pause & Ask**: Inform the user that the HTML draft is ready for review. Ask them: "Please open `output/presentation.html` in your browser. If you are satisfied with the structure and content, tell me to proceed to Phase 2 to generate the native PPTX file."

## Phase 1.5: Mandatory HTML Templates (Cover, TOC & Content Layouts)
You MUST use the exact HTML structures below for the Cover and TOC slides in Phase 1.

For content slides, choose the layout class that best matches the information type.

**Cover Slide Template**:
```html
<!-- _class: title-slide -->
<div class="cover-left-banner"></div>
<img class="cover-logo" src="../assets/heu_logo_badge.png" />
<div class="cover-title-area">
  <div class="cover-title-p1">主标题</div>
  <div class="cover-title-p2">副标题</div>
</div>
<div class="cover-meta-line"></div>
<div class="cover-meta-list">
  <div class="cover-meta-item" style="display: flex; align-items: flex-start;"><span class="label" style="display: inline-block; width: 86px; text-align-last: justify;">汇报人</span>：姓名</div>
  <div class="cover-meta-item" style="display: flex; align-items: flex-start;"><span class="label" style="display: inline-block; width: 86px; text-align-last: justify;">日期</span>：具体日期</div>
</div>
<div class="cover-date">2026年X月</div>
```

**TOC Slide Template**:
```html
---

<!-- _class: toc-slide -->
<div class="toc-container">
    <div class="toc-left">
        目录<br>
        <span class="en">CONTENTS</span>
    </div>
    <ul class="toc-list list-none">
        <li class="toc-item"><span class="num">01</span> <span class="text">章节名称一</span></li>
        <li class="toc-item"><span class="num">02</span> <span class="text">章节名称二</span></li>
    </ul>
</div>
```

> **TOC page constraint:** The TOC should fit on a single slide. Keep the entry count to 4–6 items and the titles concise. If you have more chapters, group them under broader headings rather than splitting the TOC across multiple slides.

**Thanks Slide Template**:
```html
---

<!-- _class: thanks-slide -->
<svg style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; z-index: 1;">
  <polygon points="0,0 100,0 180,360 100,720 0,720" fill="#004EA1" />
  <polygon points="768,0 1043,0 1043,137 905.5,248 768,137" fill="#004EA1" />
  <polygon points="1130,720 1280,720 1280,570" fill="#004EA1" />
</svg>
<div class="thanks-content">
    <div class="thanks-bg-text">THANK YOU</div>
    <div class="thanks-title">谢谢大家</div>
    <img class="thanks-school-img" src="../assets/heu_signature.png" />
</div>
```

### Content Slide Layout Templates

**Gallery** (`layout: gallery`) — 要点列表 + 三图网格:
```html
---
<h1>Slide Title</h1>
<div class="fill-height">
  <div class="subtitle-black">➢ 小标题</div>
  <div class="flex-row" style="flex: none !important;">
    <div class="text-box-solid">
      <span class="subtitle-accent">◆ 要点一</span>
      <ul>
        <li>列表项1</li>
        <li>列表项2</li>
      </ul>
    </div>
    <div class="text-box-solid">
      <span class="subtitle-accent">◆ 要点二</span>
      <ul>
        <li>列表项1</li>
        <li>列表项2</li>
      </ul>
    </div>
  </div>
  <div class="grid-3">
    <div class="img-placeholder">[Fig 1]</div>
    <div class="img-placeholder">[Fig 2]</div>
    <div class="img-placeholder">[Fig 3]</div>
  </div>
</div>
```

**Highlight** (`layout: highlight`) — 高亮框 + 列表 + 三图:
```html
---
<h1>Slide Title</h1>
<div class="fill-height">
  <div class="subtitle-black">➢ 小标题</div>
  <div class="text-box">
    <span class="subtitle-accent">◆ 核心结论</span>
    高亮强调的关键结论文字
  </div>
  <div class="flex-row" style="flex: none !important;">
    <div class="text-box-solid">
      <ul>
        <li>补充说明1</li>
        <li>补充说明2</li>
      </ul>
    </div>
  </div>
  <div class="grid-3">
    <div class="img-placeholder">[Fig 1]</div>
    <div class="img-placeholder">[Fig 2]</div>
    <div class="img-placeholder">[Fig 3]</div>
  </div>
</div>
```

**Dual** (`layout: dual`) — 上下双块（可公式/对比/并列）+ 总结:
```html
---
<h1>Slide Title</h1>
<div class="fill-height">
  <div class="subtitle-black">➢ 小标题</div>
  <div class="text-box-solid">
    <span class="subtitle-accent">◆ 块一</span>
    第一块内容，可包含公式 $E = mc^2$
  </div>
  <div class="text-box-solid">
    <span class="subtitle-accent">◆ 块二</span>
    第二块内容
  </div>
  <div class="text-box" style="flex: none !important;">
    <span class="subtitle-accent">◆ 总结</span>
    总结性语句
  </div>
</div>
```

**Columns** (`layout: columns`) — 多栏小标题+内容（2~3栏）+ 双图:
```html
---
<h1>Slide Title</h1>
<div class="fill-height">
  <div class="subtitle-black">➢ 总副标题（可选）</div>
  <div class="flex-row">
    <div class="text-box-solid">
      <span class="subtitle-accent">◆ 栏目标题1</span>
      栏目内容描述
    </div>
    <div class="text-box-solid">
      <span class="subtitle-accent">◆ 栏目标题2</span>
      栏目内容描述
    </div>
    <div class="text-box-solid">
      <span class="subtitle-accent">◆ 栏目标题3</span>
      栏目内容描述
    </div>
  </div>
  <div class="flex-row" style="flex: none !important;">
    <div class="img-placeholder">[Fig 1]</div>
    <div class="img-placeholder">[Fig 2]</div>
  </div>
</div>
```

**Text-Image** (`layout: text-image`) — 左文右图:
```html
---
<h1>Slide Title</h1>
<div class="fill-height">
  <div class="subtitle-black">➢ 小标题</div>
  <div class="flex-row">
    <div class="text-box-solid" style="flex: 1;">
      左侧大段描述性文字...
    </div>
    <div class="img-placeholder" style="flex: 1;">[Fig 1]</div>
  </div>
</div>
```

**Table** (`layout: table`) — 纯表格:
```html
---
<h1>Slide Title</h1>
<div class="fill-height">
  <div class="subtitle-black">➢ 小标题</div>
  <div class="text-box-solid">
    <table>
      <thead><tr><th>参数</th><th>方案A</th><th>方案B</th></tr></thead>
      <tbody><tr><td>功耗</td><td>50W</td><td>80W</td></tr></tbody>
    </table>
  </div>
</div>
```

**Table-Image** (`layout: table-image`) — 表格 + 配图:
```html
---
<h1>Slide Title</h1>
<div class="fill-height">
  <div class="subtitle-black">➢ 小标题</div>
  <div class="flex-row">
    <div class="text-box-solid" style="flex: 1;">
      <table>
        <thead><tr><th>时间</th><th>温度</th><th>压力</th></tr></thead>
        <tbody><tr><td>T0</td><td>25</td><td>1.0</td></tr></tbody>
      </table>
    </div>
    <div class="img-placeholder" style="flex: 1;">[Fig 1]</div>
  </div>
</div>
```

> **Note on HTML-block content:** Marp’s Markdown parser (markdown-it) does not parse Markdown syntax inside block-level HTML tags such as `<div>`. Therefore, whenever you place lists or tables inside `.text-box-solid` or `.text-box`, you must write them as raw HTML (`<ul><li>...`, `<table>...`) rather than Markdown (`- item`, `| col |`). Inline math (`$...$`) is the sole exception: it is rendered client-side by KaTeX after the HTML is generated, so `$...$` **does** work inside HTML blocks.

## Phase 1.6: Applying Layout Components
Do not write plain paragraphs. You must leverage the pre-defined layout classes in the stylesheet, choosing the component set that matches your selected **layout**:

**Universal components (all layouts)**:
- **Full-height container (Mandatory outer wrapper for content slides)**:
  `<div class="fill-height"> ... </div>`
- **Slide Subtitles**:
  `<div class="subtitle-black">➢ Subtitle Text</div>`
- **Solid Background Box**:
  `<div class="text-box-solid"> ... </div>`
- **Dashed Border Box (for summaries, notes, or highlights)**:
  `<div class="text-box"> <span class="subtitle-accent">◆ Highlights:</span> Detailed text... </div>`
- **Key Point Highlight**:
  `<span class="highlight-key">关键术语/重要数值</span>` — Renders bold in dark red #C00000
- **Multi-column Grids**:
  `<div class="flex-row"> ... </div>` (Double columns)
  `<div class="grid-3"> ... </div>` (Three columns)
- **Image Placeholder**:
  `<div class="img-placeholder">[Fig caption]</div>` — Reserved box for an illustration. You may leave it as a text placeholder; inserting real images is optional and should be done only when assets are readily available.

**Layout-specific guidance**:
- **gallery / highlight**: Use `.grid-3` for image placeholders and `.text-box-solid` for bullet lists.
- **dual**: Use two consecutive `.text-box-solid` blocks for the upper/lower content areas, and a final `.text-box` (with `flex: none !important;`) for the summary line.
- **columns**: Use `.flex-row` with 2~3 `.text-box-solid` children. Each child starts with `<span class="subtitle-accent">◆ 标题</span>`. Place image placeholders in a second `.flex-row` below.
- **text-image / table-image**: Use `.flex-row` with `style="flex: 1;"` on each child to balance the left/right split.
- **table**: Use an HTML `<table>` inside `.text-box-solid`, or an HTML `<ul>` list if no tabular data exists.

# Phase 2: Distill & Inject (Native PPTX Generation)

Only execute Phase 2 when the user explicitly approves the HTML or asks to generate the PPTX.

1. **Distill Content into JSON**:
   Native PPTX templates have strict, fixed-size text boxes. You must shrink and summarize the content from the Markdown.
   Generate a `presentation_data.json` file in the `output/` directory. Each slide must specify a `layout` chosen from the supported types: `gallery`, `highlight`, `dual`, `columns`, `text-image`, `table`, or `table-image`.

   ```json
   {
     "cover": {
       "title": "Short Title",
       "subtitle": "Subtitle",
       "reporter": "Name",
       "instructor": "Advisor",
       "major": "Major",
       "date": "2026年5月"
     },
     "toc": ["Chapter 1", "Chapter 2", "Chapter 3", "Chapter 4"],
     "slides": [
       {
         "layout": "gallery",
         "title": "01 系统架构",
         "subtitle": "➢ 核心模块",
         "bullets": ["模块A负责采集", "模块B负责规划", "模块C负责输出"],
         "images": ["assets/a.png", "assets/b.png", "assets/c.png"]
       },
       {
         "layout": "highlight",
         "title": "02 关键发现",
         "subtitle": "➢ 实验结论",
         "highlight": "效率提升至 92%",
         "bullets": ["对比传统方案提升15%", "稳定性通过200h测试"],
         "images": ["assets/chart1.png", "assets/chart2.png", "assets/chart3.png"]
       },
       {
         "layout": "dual",
         "title": "03 公式推导",
         "subtitle": "➢ 两种能量模型",
         "blocks": [
           "动能公式 $E_k = \\frac{1}{2}mv^2$ 描述运动能量",
           "势能公式 $E_p = mgh$ 描述位置能量"
         ],
         "summary": "两者共同构成系统总机械能"
       },
       {
         "layout": "columns",
         "title": "04 技术特点",
         "columns": [
           {"subtitle": "➢ 高效能", "content": "功耗降低40%，续航提升显著"},
           {"subtitle": "➢ 高可靠", "content": "MTBF超10000小时"},
           {"subtitle": "➢ 高兼容", "content": "支持多种通信协议"}
         ],
         "images": ["assets/diag1.png", "assets/diag2.png"]
       },
       {
         "layout": "text-image",
         "title": "05 实验环境",
         "subtitle": "➢ 水下测试平台",
         "text": "实验在HEU水池完成，深度10米，温度15-25℃，盐度3.5%。",
         "image": "assets/pool.png"
       },
       {
         "layout": "table",
         "title": "06 性能参数",
         "subtitle": "➢ 关键指标对比",
         "table": {
           "headers": ["参数", "方案A", "方案B"],
           "rows": [
             ["功耗(W)", "50", "80"],
             ["效率(%)", "92", "85"],
             ["成本(元)", "2000", "3500"]
           ]
         }
       },
       {
         "layout": "table-image",
         "title": "07 测试数据",
         "subtitle": "➢ 实测结果",
         "table": {
           "headers": ["时间", "温度", "压力"],
           "rows": [
             ["T0", "25", "1.0"],
             ["T1", "30", "1.2"]
           ]
         },
         "image": "assets/chart.png"
       }
     ]
   }
   ```

   **Important:** In JSON, always escape backslashes in LaTeX: write `\\frac`, `\\sum`, etc. A single `\f` would be parsed as a form-feed control character and stripped.

   **Images as captions:** The `images` / `image` fields may be either real file paths (e.g. `"assets/figure.png"`) or descriptive captions (e.g. `"Fig 4.1 双轴机械臂示意图"`). When a path does not exist on disk, the injection engine writes the string into the placeholder box as a figure caption, so it is clear which picture should be inserted later.

   **TOC capacity:** the bundled `templates/template.pptx` ships with **4 TOC entry placeholders**. You may use 1~4 items; if you need more, the engine will fill as many as exist in the template.

2. **Pick the Right Template Engine for the Host OS**:
   Two interchangeable injection scripts are provided. They take the same arguments and produce the same PPTX shape — pick by environment.

   | Script | Requires | Use when |
   |---|---|---|
   | `scripts/build_from_template.py` | Windows + installed Microsoft PowerPoint + `pywin32` | You are running on a Windows host with PowerPoint. Uses `win32com.client` to clone slides via the real PowerPoint engine. |
   | `scripts/build_pptx_linux.py` | `python-pptx`, `lxml` (no PowerPoint, no `pywin32`) | You are running on Linux/macOS, in a sandbox (e.g. Cowork, Docker, CI), or anywhere PowerPoint is unavailable. Uses pure OPC manipulation to clone slide parts. |

   **Auto-detect rule:** if `os.name != "nt"` or PowerPoint isn't installed, you MUST use `build_pptx_linux.py`. Never attempt `build_from_template.py` outside Windows — it will fail at `import win32com.client`.

   ```bash
   # Windows + PowerPoint
   python scripts/build_from_template.py output/presentation_data.json templates/template.pptx output/output.pptx

   # Linux / macOS / sandbox / CI
   python scripts/build_pptx_linux.py    output/presentation_data.json templates/template.pptx output/output.pptx
   ```

   Both scripts:
   - Clone slide 3 (normal-layout template) and slide 4 (highlight-layout template) once per JSON entry, in document order.
   - Move every clone to the slot just before the original "thanks" slide so final order is `[cover, toc, ...content..., thanks]`.
   - Drop the two original template pages once cloning is done.
   - Fill cover, TOC, and content text frames via `python-pptx`, preserving each frame's first-paragraph `pPr` and first-run `rPr` (font, color, size).

3. **Delivery**:
   Provide the generated PPTX file name to the user. The file is fully editable in PowerPoint / WPS / Keynote — no `win32com` traces are baked in even when the Linux engine produced it.

## Phase 2 Layout & Field Mapping

Both injection scripts pattern-match on the literal placeholder strings in `templates/template.pptx`. Do **not** rename these placeholders in the template, and do **not** alter your JSON keys; otherwise text will not be filled.

| Layout | JSON fields | Template placeholders filled |
|---|---|---|
| `gallery` | `title`, `subtitle`, `bullets[]`, `images[]` | Title text, `➢ 小标题`, `列表1\n…`, `图片1/2/3` |
| `highlight` | `title`, `subtitle`, `highlight`, `bullets[]`, `images[]` | Title text, `➢ 小标题`, `*高亮…`, `列表1\n…`, `图片/图片/图片` |
| `dual` | `title`, `subtitle`, `blocks[]`, `summary` | Title text, `➢ 小标题`, `方法1`, `方法2`, `总结性语句` |
| `columns` | `title`, `columns[{subtitle,content}]`, `images[]` | Title text, `➢ 小标题1/2/3`, `介绍1/2/3`, `图片1/2` |
| `text-image` | `title`, `subtitle`, `text`, `image` | Title text, `➢ 小标题`, `描述性文字`, `图片1` |
| `table` | `title`, `subtitle`, `table{headers,rows}` | Title text, `➢ 小标题`, existing table cells, `图片1` removed |
| `table-image` | `title`, `subtitle`, `table{headers,rows}`, `image` | Title text, `➢ 小标题`, existing table cells, `图片1` |

**Shared mappings (all layouts):**

| JSON field | Template placeholder text | Notes |
|---|---|---|
| `cover.title` + `cover.subtitle` | `大大大大标题\x0b子标题` (two lines in a single text frame) | Linux engine inserts an `<a:br/>` between the two lines. |
| `cover.reporter` / `instructor` / `major` / `date` | `汇报人 ：\n导  师 ：\n专  业 ：\n日  期 ：` block | Engine writes all 4 fields into this multi-line text frame. |
| `cover.date` | `2025年5月` text box | **Also** replaced separately in the standalone date text box below the meta block. |
| `toc[i]` (number) | `01`, `02`, `03`, `04` | Replaced with zero-padded chapter numbers. |
| `toc[i]` (title) | TOC entry title shapes | Replaced in top-to-bottom document order. |



---

# Precautions for Single-Run Success (Crucial Rules)

Follow these rules to ensure the slides compile flawlessly without rendering leaks, formatting breaks, or alignment issues in one go:

1. **No Blank Lines inside HTML Block Nodes (Prevent Tag Leaks)**:
   - Marp’s CommonMark parser treats empty lines inside HTML blocks as paragraph breaks. This splits the HTML block and outputs raw HTML tags directly onto the screen. **Always keep HTML structures contiguous with zero blank lines.**
2. **Use `h1` on Normal Content Slides, Never in Custom Special Slides**:
   - On normal content slides, place `<h1>Slide Title</h1>` right after the slide separator `---`. The global `h1` style renders the slide title at the top-left and keeps the top-right logo visible.
   - Do **not** use `h1` inside custom special slides (Cover, TOC, Thanks). The global `h1` class is absolute-positioned at `top: 30px; left: 70px`; if used inside custom containers (like `.cover-title-area`), the title text will fly out of position. **Always use `div` elements with custom class names (e.g. `.cover-title-p1`) for titles on custom pages.**
3. **Prevent Stretch on Concluding Text Boxes**:
   - The default `flex-grow: 1` rule targets the last child of `.fill-height`. If the last child is a concluding statement box wrapped in `.flex-row` instead of an image grid, it will stretch into a huge, empty background block.
   - **Solution**: Explicitly set `flex: none !important;` inline on the `.flex-row` container to keep it compact and proportional to the text length.
4. **Justified Metadata Labels & Wrapped Values Alignment**:
   - Do not manually insert full-width or half-width spaces to align short labels (e.g. `汇报人` vs `指导老师`).
   - **Best Practice**: Set a fixed width on `.label` (e.g. `width: 86px;`) using `display: inline-block; text-align-last: justify;`. Put the colon outside the span: `<span class="label">汇报人</span>：`. Enable `display: flex; align-items: flex-start;` on the parent `.cover-meta-item` so wrapped long values automatically align under the colon, preserving clean vertical columns.
5. **Phase isolation**: Do NOT generate JSON or run the PPTX script during Phase 1. Wait for the user's explicit command.
6. **JSON Length Limits**: Ensure bullet points are concise (max 3 lines) to prevent text overflow in the fixed-size native template.
7. **Asset Relative Paths**: Always ensure `assets/` paths resolve correctly depending on your working directory (e.g., if outputting to `output/`, point image paths to `../assets/`). Also verify that background-image URLs inside the theme CSS resolve from the generated HTML location; otherwise logos and decorative images will not appear.
8. **LaTeX / Math Rendering**:
   - The Marp frontmatter MUST contain `math: katex` whenever the deck has any equation. Without it, `$...$` and `$$...$$` are output as raw text and subscripts/Greek letters become unreadable.
   - Inline math uses single `$ ... $`; display math uses `$$ ... $$` on its own lines (a blank line before and after the block is fine — those blank lines are between Markdown blocks, NOT inside HTML containers, so rule #1 still holds).
   - When wrapping a formula inside a `<div class="text-box-solid">` or `<div class="text-box">`, keep the math on a single line using `$ ... $` (inline) or place a single `$$ ... $$` block on its own line inside the div with no surrounding blank lines.
   - Prefer LaTeX for any non-trivial physics expression. Do NOT pre-render formulas to ASCII (e.g. `tau_hydro`, `rho*g*V`). Always write `\tau_{hydro}`, `\rho g V_{sub}`, etc.
   - For matrices, use `\begin{bmatrix} ... \end{bmatrix}` inside `$$ ... $$`. Keep matrices small (≤3×3) on slides; reference larger structures in prose instead.
9. **Layout Auto-Inference (AI decides which template to use)**:
   When distilling content into JSON, analyze the text structure and pick the best `layout` automatically. Follow these heuristics:
   - **Content mentions ≥3 images** → `gallery`
   - **Content has "优势/劣势", "方案A/方案B", "对比", "两种", "并列"** → `dual`
   - **Content has a "核心结论", "关键发现", "值得注意的是"** → `highlight`
   - **Content has tabular data, "参数表", "指标对比"** → `table` (if no image) or `table-image` (if one image)
   - **Content describes "特点", "特性", "维度" with 2~3 parallel items** → `columns`
   - **Content has one core image + long descriptive text** → `text-image`
   - **Content has one large chart/figure with minimal text** → `table-image` (treat chart as image) or `gallery` with fewer bullets
   - **Fallback** → `gallery` (general-purpose bullet + image layout)
10. **PPTX Formula Rendering (Phase 2 → OMML)**:
   - All JSON text fields (`bullets`, `blocks[]`, `summary`, `highlight`, `subtitle`, `title`, `cover.title/subtitle`, `table.rows[]`) keep math written as inline LaTeX `$...$`. Do NOT downgrade them to ASCII just because they go into a PPTX — the engines now handle math natively.
   - `scripts/build_pptx_linux.py` scans each text segment, splits it into text / math chunks on `$ ... $`, converts every math chunk via `latex2mathml → mathml2omml` and inserts the result as a native `<a14:m><m:oMath>...</m:oMath></a14:m>` block inside the paragraph. Plain-text segments stay as `<a:r>` runs and inherit the original font (`rPr`).
   - After saving, the script post-processes each slide XML to hoist `xmlns:a14` and `xmlns:m` declarations from every `<a14:m>` element up to the root `<p:sld>`, so PowerPoint / WPS / Keynote treat the equations as first-class editable formulas instead of inline plain-text fallbacks.
   - Required dependencies on the host that runs Phase 2 (already declared in README): `pip install python-pptx lxml latex2mathml mathml2omml`. `latex2mathml` and `mathml2omml` are pure-Python and ship no native binaries.
   - **JSON escaping:** In JSON string values, backslashes must be doubled: write `\\frac`, `\\sum`, `\\tau_{hydro}`. A raw `\f` sequence is interpreted as a form-feed control character and will be stripped, turning `\frac` into `rac`.
   - Display-style fractions (`\dfrac`), Greek letters, subscripts, superscripts, square roots, big operators (`\sum`), absolute values (`|\cdot|`), `\cdot`, `\circ`, `\arctan2`, `\arcsin`, `\propto` are all supported and round-trip correctly.
   - Only single-line `$...$` (inline-style) is supported in the JSON. `$$...$$` block math is NOT parsed by the JSON ingestion path — break long display equations into one or two `$...$` segments per bullet, or place them in a `highlight` or `blocks[]` field.
   - `scripts/build_from_template.py` (Windows + PowerPoint engine) does NOT yet do the LaTeX→OMML conversion — it relies on PowerPoint's manual equation editor at runtime. If you need cross-platform parity, prefer the Linux engine on Windows too (it works there as well).
