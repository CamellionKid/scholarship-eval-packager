---
name: scholarship-eval-packager
description: >-
  Parse university scholarship evaluation rules (web URLs, embedded VSB PDFs/images,
  or local documents), audit applicant evidence files (papers, conference programs,
  cadre certificates, transcripts), and generate a complete submission folder with a
  formula-driven Excel workbook, filled Word (.doc/.docx) form, and numbered proof
  attachments. Use when the user asks to fill out a scholarship self-evaluation form,
  calculate scholarship points from rules and certificates, or package academic
  application materials.
---

# Scholarship Eval Packager

Pipeline: **RULES** (extract weights, tiers, caps) → **EVIDENCE** (audit proofs, resolve links, verify transcript math) → **BUILD** (`scripts/build_scholarship_pack.py` → formula `.xlsx` + filled `.doc`/`.docx` + numbered attachments) → **VERIFY** (render via LibreOffice + `pdftoppm` and inspect layout).

## Workflow

### 1. Parse Evaluation Rules

1. **Fetch rule source**:
   - If the user provides a university CMS URL (e.g. Visual SiteBuilder `showVsbpdfIframe`), inspect the HTML for `vsb_pdf_image_data` or embedded PDF paths and download all pages. See `references/schema-and-rules.md` for extraction commands.
   - Save the complete rule pages as a combined PDF inside `附件_支撑材料/04_评审依据文件/`.
2. **Extract scoring model**:
   - Record module weights (e.g., Research 55%, Academic Grades 30%, Non-research Awards 10%, Social Practice 5%).
   - Record exact tier points, caps, and deduction conditions (e.g., newspaper tiers, conference presentation levels, student cadre grades, attendance thresholds).
   - When a rule revision introduces a new sub-tier (such as regional vs. provincial newspapers), score conservatively using the explicit tier and document the alternative in the Excel `note` column so the user can adjust the yellow input cell if needed.

### 2. Audit & Reconcile Evidence

1. **Inspect all files** in the user's material directory:
   - **Papers**: verify author order, publication date, newspaper/journal tier, page/section, and CNKI/DOI indexing status.
   - **Academic conferences**: verify conference level (national vs. provincial/municipal) and confirm oral/keynote presentation from the official program agenda (`会议手册`) plus acceptance/receipt proofs.
   - **Social practice / cadre service**: verify role, assessment grade (`优秀` / `良好` / `合格`), service duration (must meet minimum duration, typically one academic year), and class activity ledger deductions.
   - **Transcripts (`成绩单`)**: extract every course code, name, credit, and score; cross-check against any college-published average score sheet (`平均成绩公示表`) for rounding conventions (e.g., `ROUND(AVERAGE(...), 3)`).
2. **Resolve non-portable pointers**:
   - Replace macOS `.textClipping` files or Markdown `![[wikilink]]` references by locating the actual image/PDF file on disk and copying the real file into the attachment package.

### 3. Build Formula-Driven Excel, Word Form & Attachment Folder

1. Prepare a JSON specification following `references/schema-and-rules.md` (see `assets/example_spec.json`).
2. Run the packager script:
   ```bash
   python3 scripts/build_scholarship_pack.py --spec /tmp/spec.json
   ```
   The script automatically:
   - Backs up the blank `.doc` / `.docx` template into `附件_支撑材料/00_原始空白模板备份/` without losing original files.
   - Copies and renames all proof files into numbered folders (`01_科研情况...`, `02_社会实践...`, `03_学业成绩...`, `04_评审依据文件`).
   - Generates a 3-sheet `.xlsx` workbook:
     - `1_分项自动计分明细表`: yellow editable raw-score cells, `=F*G` weighted score formulas, module subtotals, grand total `=SUM(...)`, and clickable relative hyperlinks to each proof file.
     - `2_学院标准自评分表(可打印)`: matches the exact columns of the college's Word form, with both a descriptive text row and a pure formula verification row linked to Sheet 1.
     - `3_课程成绩明细表`: lists all courses and computes `=SUM()`, `=AVERAGE()`, `=ROUND(AVERAGE(), 3)`, and weighted grade score automatically.
   - Fills the `.docx` form and converts it back to `.doc` via `soffice` so both formats stay in sync.

### 4. Visual Render & Verification

1. Pass `--render-dir /tmp/render_check` to `scripts/build_scholarship_pack.py` (requires `soffice` and `pdftoppm`).
2. Visually inspect every rendered PNG page of both the Word table and the 3 Excel sheets:
   - Confirm student IDs and total scores do not wrap awkwardly across lines.
   - Confirm each Excel worksheet fits onto a single A4 landscape page (`fitToWidth = 1`, `fitToHeight = 1`).
   - Clean up temporary files in `/tmp` after verification.
