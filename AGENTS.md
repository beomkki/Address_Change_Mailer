# Repository Guidelines

## Project Structure & Module Organization
- `generate_mail_merge.py`: CLI automation script; reads Excel data, renders Word template, and creates Outlook `.msg` drafts via `docx2msg`.
- `mail_merge_gui.py`: GUI version with file picker dialogs for easier use by non-technical users.
- `Filing.docx`, `Search.docx`: Word template files with merge fields (`«필드명»` format).
- `Filing_Merge.xlsx`, `Search_Merge.xlsx`: Excel data files with merge data.
- `output-msg/`: Auto-created workspace for generated `.msg` drafts. Safe to delete and regenerate.
- `README.MD`: Comprehensive user documentation with setup, usage, and troubleshooting guides.
- `helpme.txt`: User notes that explain business context and requirements.

## Build, Test, and Development Commands

### Development Setup
- `pip install docx2msg openpyxl python-docx pywin32` - Installs required runtime packages. Run once per environment.

### Running CLI Version
- `python generate_mail_merge.py` - Runs the full merge pipeline with default settings (`Filing_Merge.xlsx` + `Filing.docx`).
- `python generate_mail_merge.py --excel "Search_Merge.xlsx" --template "Search.docx"` - Override defaults for Search templates.
- `python generate_mail_merge.py --subject-template "«관리번호» - «국가명칭»"` - Use custom subject template.

### Running GUI Version
- `python mail_merge_gui.py` - Launches GUI application with file selection dialogs.

### Building EXE (Windows)
- `pip install pyinstaller` - Install PyInstaller for building executables.
- `pyinstaller --onefile --windowed mail_merge_gui.py` - Build standalone GUI executable.
- `pyinstaller --onefile --console generate_mail_merge.py` - Build CLI executable (shows console window).
- Output: `dist/mail_merge_gui.exe` or `dist/generate_mail_merge.exe`

### Testing
- Generate sample `.msg` files and open in Outlook to verify:
  - Subject line is correctly populated with merge fields
  - Recipients (To, CC) are properly set
  - Email body renders correctly with all merge fields replaced
  - HTML formatting is preserved

## Coding Style & Naming Conventions
- Use 4-space indentation, UTF-8 encoding, and snake_case identifiers (`normalize_field_markers`, `run_mail_merge`).
- Keep console output action-oriented (`Saved MSG: ...`). Avoid verbose logging inside the GUI.
- When adding placeholders to Word templates, type them as plain text (`«필드»` or `<<필드>>`); the engine normalises both forms.

## Testing Guidelines
- Manual validation is required: after each run, open a sample `.msg` in Outlook to confirm:
  - Subject, recipients (To, CC), and email body are correctly populated
  - All merge fields (`«필드명»`) are replaced with actual data
  - No merge field symbols (`«»`) remain in the output
  - HTML layout is preserved
  - Korean text displays correctly without encoding issues
- When modifying Excel/Word templates:
  - Cross-check at least one sample entry
  - Verify column names match merge field names exactly (case-sensitive)
  - Ensure `«»` symbols are used, not `<<>>` (see README.MD for details)
- For GUI changes:
  - Test file selection dialogs with various file paths
  - Verify progress updates and error messages display correctly
  - Test with missing/invalid files to ensure graceful error handling

## Commit & Pull Request Guidelines
- Follow imperative commit messages (`Add GUI file picker`, `Fix merge field encoding`). Keep commits scoped to a single logical change.
- Pull requests should include:
  - Purpose summary (what and why)
  - Testing notes (commands executed, sample `.msg` files reviewed)
  - List of dependencies installed (if any)
  - Screenshots for GUI changes
- Before committing:
  - Test with both Filing and Search templates
  - Verify README.MD is updated if user-facing features changed
  - Run at least one full merge test and verify output `.msg` files

## Distribution Guidelines
- EXE files are for end-users who don't have Python installed
- Include README.MD with the EXE distribution
- EXE file size will be ~20-30MB due to bundled Python interpreter
- Test EXE on clean Windows machine without Python to verify standalone functionality
- For team distribution: provide both Python scripts (for developers) and EXE (for non-technical users)

