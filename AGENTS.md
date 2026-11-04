# Excel Macro Toolkit Development Rules

## Workspace

- Use `D:\OneDrive\Program\data-editing-macro-toolkit` as the project path. Do not publish to the former C-drive OneDrive path.
- Preserve unrelated user changes. Stage only files that belong to the current task.

## Korean Encoding

- Keep editable source files (`.bas`, `.ps1`, `.md`) in UTF-8.
- PowerShell 5.1 scripts containing Korean must use UTF-8 with BOM. UTF-8 without BOM is decoded as the active ANSI code page and can corrupt string literals before the script starts.
- `src/MacroToolkit.bas` is the UTF-8 source of truth. Do not edit a CP949 export and copy it back over the source.
- The release script may create a CP949 VBA import copy because Excel VBA imports Korean modules through the active Windows code page.
- Never repair Korean text by guessing mojibake replacements. Recover the intended Korean from the UTF-8 source, documentation, or a known-good revision.
- After changing Korean UI text, run an explicit UTF-8 read and scan changed files for common mojibake fragments such as `?ㅏ`, `吏`, `怨듭떇`, and `寃`.
- Build the XLAM/XLSM and confirm the release manifest contains `"vbaImportEncoding": "CP949"`.

## Release

- Versioned releases must not retain API credentials. Only `release/latest` may retain local workbook-scoped credentials.
- Publish protected XLAM files through `scripts/publish_release.ps1`.
- Read the VBA password only from the `DATA_EDITING_TOOLKIT_VBA_PASSWORD` user environment variable. Never put the password in source, arguments, logs, or manifests.
- VBA protection must be non-interactive. Do not open the VBE project-properties dialog or use mouse/keyboard automation during publish.
- Update `docs/USER_MANUAL.md` whenever user-facing behavior changes.
- Run focused tests, `git diff --check`, encoding scans, and manifest checks before committing a release.
