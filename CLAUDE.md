# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

公文格式核稿系统 — a Chinese government document (公文) format checking and auto-correction web app. Users upload `.docx` files, the system detects format violations against GB/T 9704 standards, then generates corrected `.docx` output with one click.

## Commands

```bash
# Start the server (single command)
python app.py

# Or use the batch file (installs deps if needed, opens browser)
start.bat

# Install dependencies
pip install -r requirements.txt
```

Server runs at `http://localhost:8000`.

## Architecture

```
app.py          FastAPI entry — 7 endpoints (upload, format, download, preview, fonts list/download, health)
parser.py       Document structure analyzer — 4-pass detection pipeline
formatter.py    Format application engine — generates corrected .docx
config.py       All format rule constants (margins, fonts, spacing, element types)
utils.py        Regex patterns for structure classification + CJK/ASCII character splitting
static/
  index.html    Single-page frontend (vanilla JS, no framework)
```

### Data Flow

1. **Upload** (`POST /api/upload`) → saves to `uploads/{uuid}.docx` → `parse_document()` extracts structure → returns paragraphs + issues to frontend
2. **Format** (`POST /api/format`) → `format_document()` creates corrected copy in `outputs/{uuid}.docx` → frontend gets download URL
3. **Download** (`GET /api/download/{id}`) → streams the corrected file

In-memory `file_store` dict maps file IDs to metadata. No database.

### Parser Pipeline (4 passes)

1. **Per-paragraph classification** — `utils.classify_paragraph()` uses regex to tag each paragraph: `title`, `heading_1/2/3`, `figure_caption`, `speaker`, `body`, `attachment_title`, `empty`
2. **Context refinement** (`_refine_context`) — backward scan from document end to identify `signature_date`, `cc`, `note`, `attachment_title`; forward propagation marks attachment body paragraphs
3. **Format compliance check** (`_check_format`) — compares current vs. expected font/size/bold/alignment/indent/line-spacing for each paragraph
4. **Page-level check** (`_check_page_setup`) — verifies margins match standard (top 37/bottom 35/left 28/right 26 mm)

### Formatter Pipeline

1. `_reset_style_defaults()` — modify `styles.xml`: set `docDefaults` and Normal style spacing to 0
2. `_apply_page_setup()` — A4 paper, standard margins
3. `_reset_all_paragraph_spacing()` — XML-level clear of all 6 spacing attrs per paragraph
4. Per-paragraph: `_apply_paragraph_format()` — alignment, indent, fixed 28pt line spacing, CJK/ASCII font splitting
5. `_insert_signature_spacing()` — ensure 2 empty lines before signature block
6. `_apply_subtitle_bold()` — detect consecutive equal-length first sentences (≥3) and bold them

### Key Format Rules (config.py)

| Element | Font | Size | Indent |
|---------|------|------|--------|
| title | 方正小标宋_GBK | 22pt (2号) | None |
| heading_1 | 黑体 | 16pt (3号) | 1.13cm |
| heading_2 | 楷体_GB2312 | 16pt (3号) | 1.13cm |
| heading_3/body | 仿宋_GB2312 | 16pt (3号) | 1.13cm |
| signature/date | 仿宋_GB2312 | 16pt (3号) | None (right-aligned) |

All paragraphs: fixed 28pt line spacing, 0pt before/after spacing.

### Spacing Attributes

When working with paragraph spacing, **all 6 OOXML attributes** must be cleared — not just `w:before` and `w:after`:

- `w:before` / `w:after` — absolute spacing in twips
- `w:beforeLines` / `w:afterLines` — line-based spacing (100 = 1行, 200 = 2行). Word displays these as "X行"
- `w:beforeAutospacing` / `w:afterAutospacing` — auto spacing flags

Use `_clear_spacing_attrs(sp_element)` from `formatter.py` to zero all six.

### Font Requirements

Windows system must have: 方正小标宋_GBK, 仿宋_GB2312, 楷体_GB2312, 黑体. Font files in `fonts/` directory are available for download via `/api/fonts`.
