# Data Editing Macro Toolkit

Excel data editing macro toolkit.

## Project Structure

- `src/`: VBA source files. Edit macro code here first.
- `dev/`: development workbook used while building and testing.
- `release/`: published version snapshots.
- `scripts/`: build and publish automation.

## Current Version

`0.1.0`

The source of truth is `TOOLKIT_VERSION` in `src/MacroToolkit.bas`.

## Current Features

- Convert text-formatted numbers to real numeric values.
- Normalize mixed Korean phone number formats into a consistent hyphenated format.

## Build Development Workbook

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_macro_workbook.ps1
```

This creates or replaces `dev/DataEditingMacroToolkit_dev.xlsm`.

## Publish Release

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish_release.ps1
```

This reads `TOOLKIT_VERSION` and creates:

- `release/v0.1.0/DataEditingMacroToolkit_v0.1.0.xlsm`
- `release/v0.1.0/MacroToolkit.bas`
- `release/v0.1.0/manifest.json`
- `release/latest/DataEditingMacroToolkit_latest.xlsm`

## Versioning Policy

Use semantic versioning:

- Patch: small fixes with no workflow change, such as `0.1.1`.
- Minor: new reviewed features, such as `0.2.0`.
- Major: breaking changes to macro names, behavior, or workbook structure, such as `1.0.0`.

Recommended workflow:

1. Edit `src/MacroToolkit.bas`.
2. Build and test `dev/DataEditingMacroToolkit_dev.xlsm`.
3. Update `TOOLKIT_VERSION` when ready.
4. Run `scripts/publish_release.ps1`.
5. Commit the source and release snapshot.

## Excel VBA Access Note

If Excel blocks automatic module import, enable:

File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model.

Alternatively, open the VBA editor with `Alt+F11`, choose File > Import File, and import `src/MacroToolkit.bas`.
