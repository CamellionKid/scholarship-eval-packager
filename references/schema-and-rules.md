# Specification Schema & Extraction Patterns

## 1. Extracting Visual SiteBuilder (VSB) Embedded Rule PDFs/Images

Many Chinese university websites (Visual SiteBuilder 9) embed policy documents via `showVsbpdfIframe` and `vsb_pdf_image_data` instead of plain HTML text. Extract them with:

```bash
curl -s "<RULE_URL>" | grep -E "vsb_pdf_image_data|showVsbpdfIframe"
```

Download the page images in parallel and combine them into a local reference PDF using Pillow (`Image.save(..., save_all=True, append_images=...)`).

## 2. Resolving macOS `.textClipping` & Obsidian Wikilinks

When an evidence folder contains a `.textClipping` file or Markdown `![[screenshot.png]]`:

```bash
plutil -p "path/to/file.textClipping"
mdfind -name "screenshot.png"
```

Always copy the resolved `.png` / `.pdf` binary into `附件_支撑材料/` so reviewers on Windows can open every attachment.

## 3. JSON Specification Schema for `scripts/build_scholarship_pack.py`

```json
{
  "base_dir": "/abs/path/to/output_folder",
  "excel_filename": "学业奖学金自评分表_姓名_学号.xlsx",
  "word_template": "/abs/path/to/template.doc",
  "title": "2025—2026 学年硕士研究生学业奖学金自评自动核算表",
  "form_title": "2025—2026 学年硕士研究生学业奖学金自评分表",
  "applicant": {
    "id": "20240000001",
    "name": "张三",
    "major": "2024级 哲学（伦理学）",
    "department": "马克思主义学院",
    "advisor": "李四",
    "rule_doc": "《研究生学业奖学金评审细则》"
  },
  "copies": [
    {
      "src": "/abs/path/to/source_proof.pdf",
      "dst_rel": "附件_支撑材料/01_科研情况_55%/01_论文证明.pdf"
    }
  ],
  "items": [
    {
      "no": 1,
      "cat": "一、科研情况\n(55%)",
      "col_key": "papers",
      "col_name": "发表论文情况(自评分)",
      "detail": "论文或成果详细描述",
      "rule": "细则条款说明",
      "raw_score": 5,
      "weight": 0.55,
      "link_text": "01_论文证明.pdf",
      "link_path": "附件_支撑材料/01_科研情况_55%/01_论文证明.pdf",
      "note": "备注说明"
    }
  ],
  "courses": [
    {"no": 1, "code": "G1001", "name": "课程A", "credit": 2.0, "score": 90.0}
  ],
  "word_columns": [
    {"key": "id", "header": "学号", "width_cm": 2.2, "align": "center"},
    {"key": "name", "header": "学生姓名", "width_cm": 1.6, "align": "center"},
    {"key": "papers", "header": "发表论文情况(自评分)", "width_cm": 4.3, "align": "left"},
    {"key": "awards", "header": "获奖情况(自评分)", "width_cm": 4.8, "align": "left"},
    {"key": "projects", "header": "主持项目(自评分)", "width_cm": 1.8, "align": "center"},
    {"key": "non_research", "header": "非科研获奖（自评分）", "width_cm": 1.8, "align": "center"},
    {"key": "practice", "header": "引领校风院风学风建设（自评分）", "width_cm": 3.2, "align": "left"},
    {"key": "grades", "header": "课程成绩", "width_cm": 3.3, "align": "left"},
    {"key": "total", "header": "总分", "width_cm": 1.6, "align": "center"},
    {"key": "sign", "header": "是否赞同签名", "width_cm": 1.6, "align": "center"}
  ],
  "word_row_texts": {
    "id": "20240000001",
    "name": "张三",
    "papers": "1. 论文明细...\n\n总计：5×0.55 = 2.75分",
    "awards": "1. 会议发言...\n\n总计：(6+3)×0.55 = 4.95分",
    "projects": "无",
    "non_research": "无",
    "practice": "1. 校辩论队队员（良好，2分）\n\n总计：2×0.05 = 0.1分",
    "grades": "85+90+... = 1271\n1271÷15 = 84.733\n\n总计：84.733×0.3 = 25.4199分",
    "total": "33.2199",
    "sign": ""
  }
}
```
