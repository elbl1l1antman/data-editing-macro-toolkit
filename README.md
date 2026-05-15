# Data Editing Macro Toolkit v0.1.0

This folder contains a macro-enabled Excel workbook shell and the VBA module for the toolkit.

## Files

- `DataEditingMacroToolkit_v0.1.0.xlsm`: macro-enabled workbook shell.
- `MacroToolkit.bas`: VBA module containing the toolkit macros.
- `build_macro_workbook.ps1`: rebuild script that imports the VBA module when Excel allows programmatic VBA access.

## Current Features

- Convert text-formatted numbers to real numeric values.
- Normalize mixed Korean phone number formats into a consistent hyphenated format.

## Macros

- `Toolkit_About`
- `Toolkit_SetupWorkbook`
- `ConvertTextNumbersInSelection`
- `ConvertTextNumbersInUsedRange`
- `ConvertTextNumbersInColumns`
- `NormalizePhonesInSelection`
- `NormalizePhonesInColumns`

## Installation Note

The local Excel security setting currently blocks programmatic VBA project access. To import the module automatically:

1. Open Excel.
2. Go to File > Options > Trust Center > Trust Center Settings > Macro Settings.
3. Enable "Trust access to the VBA project object model".
4. Run `build_macro_workbook.ps1` again.

Alternatively, open the VBA editor with `Alt+F11`, choose File > Import File, and import `MacroToolkit.bas`.

## Versioning

The first version is `0.1.0`. Future reviewed feature additions should increment the version and update `TOOLKIT_VERSION` in `MacroToolkit.bas`.
