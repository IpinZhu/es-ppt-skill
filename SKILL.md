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
   - Follow the detailed **Mandatory Templates & Layout Specifications** below.
3. **Compile HTML**:
   Run the following command to generate the HTML preview:
   ```bash
   python scripts/build_html_ppt.py output/presentation.md -o output/presentation.html
   ```
4. **Pause & Ask**: Inform the user that the HTML draft is ready for review. Ask them: "Please open `output/presentation.html` in your browser. If you are satisfied with the structure and content, tell me to proceed to Phase 2 to generate the native PPTX file."

## Phase 1.5: Mandatory HTML Templates (Cover & TOC)
You MUST use the exact HTML structures below for the Cover and TOC slides in Phase 1.

**Cover Slide Template**:
```html
<!-- _class: title-slide -->
<div class="cover-left-banner"></div>
<img class="cover-logo" src="../assets/heu_logo_badge.png" />
<div class="cover-title-area">
  <div class="cover-title-p1">主标题</div>
  <div class="cover-title-p1">副标题</div>
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

## Phase 1.6: Applying Layout Components
Do not write plain paragraphs. You must leverage the pre-defined layout classes in the stylesheet:
- **Full-height container (Mandatory outer wrapper for content slides)**:
  `<div class="fill-height"> ... </div>`
- **Slide Subtitles**:
  `<div class="subtitle-black">➢ Subtitle Text</div>`
- **Solid Background Box (for simple text/paragraphs)**:
  `<div class="text-box-solid"> ... </div>`
- **Dashed Border Box (for summaries, notes, or highlights)**:
  `<div class="text-box"> <span class="subtitle-accent">◆ Highlights:</span> Detailed text... </div>`
- **Key Point Highlight (for emphasizing key terms/values inside any text box)**:
  `<span class="highlight-key">关键术语/重要数值</span>` — Renders bold in dark red #C00000
- **Multi-column Grids (for images & split layouts)**:
  `<div class="flex-row"> ... </div>` (Double columns)
  `<div class="grid-3"> <div class="img-placeholder">[Fig 1]</div> ... </div>` (Three columns)

# Phase 2: Distill & Inject (Native PPTX Generation)

Only execute Phase 2 when the user explicitly approves the HTML or asks to generate the PPTX.

1. **Distill Content into JSON**:
   Native PPTX templates have strict, fixed-size text boxes. You must shrink and summarize the content from the Markdown.
   Generate a `presentation_data.json` file in the `output/` directory. Each slide can specify a `layout`: `"normal"` (standard text layout) or `"highlight"` (includes a dashed highlight box).

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
     "toc": ["Chapter 1", "Chapter 2", "Chapter 3"],
     "slides": [
       {
         "layout": "normal",
         "title": "Chapter 1",
         "subtitle": "➢ Key Point",
         "bullets": ["Very short bullet 1", "Very short bullet 2"]
       },
       {
         "layout": "highlight",
         "title": "Chapter 2",
         "subtitle": "➢ Highlight Box",
         "bullets": ["Highlighted detail 1"],
         "highlight": "One-line emphasis text",
         "image": "assets/some_image.png"
       }
     ]
   }
   ```

   **TOC capacity:** the bundled `templates/template.pptx` ships with **3 TOC entry placeholders** (texts `"01"` and `"项目背景介绍"`). Keep `toc` to 3 items; if more chapters are needed, group them or extend the template separately.

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

## Phase 2 Bullet & Title Mapping (so the engines can find your text)

Both injection scripts pattern-match on the literal placeholder strings in `templates/template.pptx`. Do **not** rename these placeholders in the template, and do **not** alter your JSON keys; otherwise text will not be filled.

| JSON field | Template placeholder text | Notes |
|---|---|---|
| `cover.title` + `cover.subtitle` | `大大大大标题\x0b子标题` (two lines in a single text frame) | Linux engine inserts an `<a:br/>` between the two lines. |
| `cover.reporter` / `instructor` / `major` | `申 请 人 ：…` block | Linux engine writes 4 lines: `汇 报 人 / 指导老师 / 专 业 / 日 期`. |
| `cover.date` | `2025年5月` text box | Replaced as a single line. |
| `toc[i]` (number) | `01` (3 occurrences) | Replaced with `01`, `02`, `03`. |
| `toc[i]` (title) | `项目背景介绍` (3 occurrences) | Replaced in document order. |
| `slides[i].title` | `01 项目背景介绍` (text frame heading) | The chapter title at top-left of each content slide. |
| `slides[i].subtitle` | `➢ 核心痛点分析` | The arrow-prefixed subtitle below the title. |
| `slides[i].bullets` | `现有方案的局限性较高\n…` | One paragraph per bullet, joined by `\n`. |
| `slides[i].highlight` | `*HIGHLIGHT…` (highlight layout only) | Single-line emphasis inside the dashed box. |



---

# Precautions for Single-Run Success (Crucial Rules)

Follow these rules to ensure the slides compile flawlessly without rendering leaks, formatting breaks, or alignment issues in one go:

1. **No Blank Lines inside HTML Block Nodes (Prevent Tag Leaks)**:
   - Marp’s CommonMark parser treats empty lines inside HTML blocks as paragraph breaks. This splits the HTML block and outputs raw HTML tags directly onto the screen. **Always keep HTML structures contiguous with zero blank lines.**
2. **Never Use `h1` in Custom Special Slides (Avoid Global Style Pollution)**:
   - The global `h1` class is absolute-positioned at `top: 30px; left: 70px;` for normal content pages. If you use `h1` inside custom containers (like `.cover-title-area`), the title text will fly out of position. **Always use `div` elements with custom class names (e.g. `.cover-title-p1`) for titles on custom pages.**
3. **Prevent Stretch on Concluding Text Boxes**:
   - The default `flex-grow: 1` rule targets the last child of `.fill-height`. If the last child is a concluding statement box wrapped in `.flex-row` instead of an image grid, it will stretch into a huge, empty background block.
   - **Solution**: Explicitly set `flex: none !important;` inline on the `.flex-row` container to keep it compact and proportional to the text length.
4. **Justified Metadata Labels & Wrapped Values Alignment**:
   - Do not manually insert full-width or half-width spaces to align short labels (e.g. `汇报人` vs `指导老师`).
   - **Best Practice**: Set a fixed width on `.label` (e.g. `width: 86px;`) using `display: inline-block; text-align-last: justify;`. Put the colon outside the span: `<span class="label">汇报人</span>：`. Enable `display: flex; align-items: flex-start;` on the parent `.cover-meta-item` so wrapped long values automatically align under the colon, preserving clean vertical columns.
5. **Phase isolation**: Do NOT generate JSON or run the PPTX script during Phase 1. Wait for the user's explicit command.
6. **JSON Length Limits**: Ensure bullet points are concise (max 3 lines) to prevent text overflow in the fixed-size native template.
7. **Asset Relative Paths**: Always ensure `assets/` paths resolve correctly depending on your working directory (e.g., if outputting to `output/`, point image paths to `../assets/`).
8. **LaTeX / Math Rendering**:
   - The Marp frontmatter MUST contain `math: katex` whenever the deck has any equation. Without it, `$...$` and `$$...$$` are output as raw text and subscripts/Greek letters become unreadable.
   - Inline math uses single `$ ... $`; display math uses `$$ ... $$` on its own lines (a blank line before and after the block is fine — those blank lines are between Markdown blocks, NOT inside HTML containers, so rule #1 still holds).
   - When wrapping a formula inside a `<div class="text-box-solid">` or `<div class="text-box">`, keep the math on a single line using `$ ... $` (inline) or place a single `$$ ... $$` block on its own line inside the div with no surrounding blank lines.
   - Prefer LaTeX for any non-trivial physics expression. Do NOT pre-render formulas to ASCII (e.g. `tau_hydro`, `rho*g*V`). Always write `\tau_{hydro}`, `\rho g V_{sub}`, etc.
   - For matrices, use `\begin{bmatrix} ... \end{bmatrix}` inside `$$ ... $$`. Keep matrices small (≤3×3) on slides; reference larger structures in prose instead.
9. **PPTX Formula Rendering (Phase 2 → OMML)**:
   - Both `presentation_data.json` text fields (`bullets`, `highlight`, `subtitle`, `title`, `cover.title/subtitle`) keep math written as inline LaTeX `$...$`. Do NOT downgrade them to ASCII just because they go into a PPTX — the engines now handle math natively.
   - `scripts/build_pptx_linux.py` scans each text segment, splits it into text / math chunks on `$ ... $`, converts every math chunk via `latex2mathml → mathml2omml` and inserts the result as a native `<a14:m><m:oMath>...</m:oMath></a14:m>` block inside the paragraph. Plain-text segments stay as `<a:r>` runs and inherit the original font (`rPr`).
   - After saving, the script post-processes each slide XML to hoist `xmlns:a14` and `xmlns:m` declarations from every `<a14:m>` element up to the root `<p:sld>`, so PowerPoint / WPS / Keynote treat the equations as first-class editable formulas instead of inline plain-text fallbacks.
   - Required dependencies on the host that runs Phase 2 (already declared in README): `pip install python-pptx lxml latex2mathml mathml2omml`. `latex2mathml` and `mathml2omml` are pure-Python and ship no native binaries.
   - Display-style fractions (`\dfrac`), Greek letters, subscripts, superscripts, square roots, big operators (`\sum`), absolute values (`|\cdot|`), `\cdot`, `\circ`, `\arctan2`, `\arcsin`, `\propto` are all supported and round-trip correctly.
   - Only single-line `$...$` (inline-style) is supported in the JSON. `$$...$$` block math is NOT parsed by the JSON ingestion path — break long display equations into one or two `$...$` segments per bullet, or place them in a `highlight` field.
   - `scripts/build_from_template.py` (Windows + PowerPoint engine) does NOT yet do the LaTeX→OMML conversion — it relies on PowerPoint's manual equation editor at runtime. If you need cross-platform parity, prefer the Linux engine on Windows too (it works there as well).
