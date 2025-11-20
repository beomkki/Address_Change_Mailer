# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is a Korean patent law firm mail merge automation tool that generates Outlook `.msg` draft files from Excel data and Word templates. It automates the creation of trademark application and search notification emails to overseas agents, eliminating manual copy-pasting of recipients, CC addresses, and dynamic subject lines.

## Core Architecture

The system has two entry points that both call the same underlying engine:

1. **`generate_mail_merge.py`**: CLI tool with command-line arguments for batch processing
2. **`mail_merge_gui.py`**: Tkinter GUI with preset profiles (Filing/Search) for easier manual use

Both use the `run_mail_merge()` function in `generate_mail_merge.py` as the mail merge engine.

### Key Technical Components

**Placeholder System**: The engine accepts two styles of field markers interchangeably:
- `«필드명»` (guillemet style, used in Word mail merge)
- `<<필드명>>` (angle bracket style)

Both are normalized internally to Jinja2 templates (`{{ 필드명 }}`) for rendering via `docx2msg`.

**Template Processing Flow**:
1. Parse original Word template to collect all field names (`collect_fields`)
2. Convert `«field»` or `<<field>>` markers to `{{ field }}` Jinja2 syntax (`convert_paragraph_placeholders`)
3. Inject YAML header into document header section for subject line rendering
4. Load Excel data rows and match column headers to field names
5. Render each row through `docx2msg` to generate `.msg` files with proper To/CC addresses
6. Attach files from `첨부파일` column (semicolon-separated paths) via `mail.Attachments.Add()`

**Dependencies**: Requires Windows with Outlook installed because `docx2msg` uses `pywin32` to create `.msg` files via COM automation.

## Development Commands

**Environment setup**:
```bash
python -m venv .venv
.venv\Scripts\activate
pip install docx2msg docxtpl openpyxl python-docx pywin32
```

**Running the CLI (Filing profile)**:
```bash
python generate_mail_merge.py
```

**Running the CLI (Search profile)**:
```bash
python generate_mail_merge.py --excel Search_Merge.xlsx --template Search.docx --subject-template "(<<관리번호>><<국가코드>>-S) Trademark search(es) in <<국가명칭>>"
```

**Running the GUI**:
```bash
python mail_merge_gui.py
```

**Testing changes**:
- Run the GUI and generate at least one `.msg` file per profile
- Open generated `.msg` files in Outlook to verify subject, To/CC fields, and HTML body
- Check that filename matches the rendered subject line (sanitized)

## File Structure

**Core scripts**:
- `generate_mail_merge.py:201` - `run_mail_merge()` function is the main entry point
- `generate_mail_merge.py:97` - `convert_paragraph_placeholders()` handles split runs for placeholders
- `generate_mail_merge.py:44` - `collect_fields()` extracts all placeholders from template
- `mail_merge_gui.py:15` - `MERGE_PROFILES` defines preset configurations for Filing/Search

**Data files**:
- `Filing.docx` / `Search.docx`: Word templates with `«필드»` or `<<필드>>` placeholders
- `Filing_Merge.xlsx` / `Search_Merge.xlsx`: Excel files with headers matching template field names
- First column headers in Excel must match placeholder names exactly

**Default columns**:
- `수신` (To): Recipient email addresses
- `참조` (CC): Carbon copy email addresses
- `첨부파일` (Attachments): Semicolon-separated file paths to attach
- Other columns are merged into template body

**Output**:
- `output-msg/`: Generated `.msg` files with sanitized filenames based on subject lines

## Important Implementation Details

**Placeholder Handling Edge Cases**: Word sometimes splits `«field»` markers across multiple runs. The `convert_paragraph_placeholders()` function at `generate_mail_merge.py:97` handles three cases:
1. Entire marker in one run: `«field»` → direct replacement
2. Split across runs: `«` in one run, `field` in next, `»` in another → collect and replace
3. Partial markers: `text«field` or `field»text` → replace within run

**Subject Line Rendering**: The subject template is embedded in the document header as YAML front matter (`generate_mail_merge.py:29`), then rendered by `docx2msg` along with the body. The rendered subject is also used for filename generation.

**Filename Sanitization**: `sanitize_filename()` at `generate_mail_merge.py:162` strips non-ASCII and special characters, replacing them with underscores. Falls back to `message_{index}` if sanitization produces empty string.

**XML Escaping**: Field values are XML-escaped (`escape_docx_text` at `generate_mail_merge.py:155`) to prevent malformed documents when data contains `<`, `>`, or `&`.

**Attachment Processing**: Files are attached after rendering the mail body but before saving (around `generate_mail_merge.py:283`). The attachment field value is split by semicolon, and each path is:
1. Converted to absolute path if relative (relative to project root)
2. Checked for existence (warning printed if missing, but generation continues)
3. Added to mail via `mail.Attachments.Add(str(absolute_path))`

## Business Context (from helpme.txt)

The tool was created because:
- Word's mail merge cannot insert merge fields into email subject lines
- Manual workflow required copying subject from body, pasting into subject, and filling To/CC fields
- With 20+ recipients, manual process became too time-consuming
- Approval workflow requires saving `.msg` files before sending

## Troubleshooting

**"docx2msg 패키지를 찾을 수 없습니다"**: Install `docx2msg` via pip

**Missing fields warning**: If Excel columns don't match all template placeholders, the script prints a warning but continues (empty strings are substituted)

**COM/Outlook errors**: Requires Windows with Outlook installed; `docx2msg` uses `pywin32` COM automation

**Encoding issues**: All files use UTF-8; Korean text in templates and Excel must be saved as UTF-8
