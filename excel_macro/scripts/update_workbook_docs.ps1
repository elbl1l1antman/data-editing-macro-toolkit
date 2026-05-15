param(
    [string]$WorkbookPath = "C:\Users\COM\OneDrive\Program\excel_macro\dev\DataEditingMacroToolkit_dev.xlsm",
    [string]$ModulePath = "C:\Users\COM\OneDrive\Program\excel_macro\src\MacroToolkit.bas"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
if (-not (Test-Path -LiteralPath $ModulePath)) { throw "Module not found: $ModulePath" }

function Remove-SheetIfExists($workbook, [string]$name) {
    foreach ($sheet in @($workbook.Worksheets)) {
        if ($sheet.Name -eq $name) {
            $sheet.Delete()
            return
        }
    }
}

function Add-SheetAfter($workbook, [string]$name) {
    $sheet = $workbook.Worksheets.Add([System.Type]::Missing, $workbook.Worksheets.Item($workbook.Worksheets.Count))
    $sheet.Name = $name
    return $sheet
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null

try {
    $workbook = $excel.Workbooks.Open($WorkbookPath)

    Remove-SheetIfExists $workbook "README"
    Remove-SheetIfExists $workbook "Macro List"
    Remove-SheetIfExists $workbook "VBA Source Preview"

    $readme = Add-SheetAfter $workbook "README"
    $readme.Range("A1").Value = "Data Editing Macro Toolkit"
    $readme.Range("A2").Value = "Development workbook"
    $readme.Range("A4").Value = "Where the implementation lives"
    $readme.Range("B4").Value = "src\MacroToolkit.bas"
    $readme.Range("A5").Value = "In this workbook"
    $readme.Range("B5").Value = "See the 'VBA Source Preview' sheet for a readable copy of the VBA source."
    $readme.Range("A6").Value = "To run macros"
    $readme.Range("B6").Value = "Import src\MacroToolkit.bas in the VBA editor, or enable Trust access to the VBA project object model and rebuild."
    $readme.Range("A8").Value = "Current version"
    $readme.Range("B8").Value = "0.1.0"
    $readme.Range("A10").Value = "Important note"
    $readme.Range("B10").Value = "This workbook may not contain an executable VBA module yet if Excel security blocks automatic import. The source preview is included so you can inspect the implementation directly."
    $readme.Range("A1:B1").Font.Bold = $true
    $readme.Range("A4:A10").Font.Bold = $true
    $readme.Columns("A:B").AutoFit()

    $list = Add-SheetAfter $workbook "Macro List"
    $headers = @("Macro", "What it does", "Input")
    for ($i = 0; $i -lt $headers.Count; $i++) { $list.Cells.Item(1, $i + 1).Value = $headers[$i] }
    $rows = @(
        @("Toolkit_About", "Shows toolkit version and feature summary", "None"),
        @("Toolkit_SetupWorkbook", "Creates/updates the hidden toolkit metadata sheet", "None"),
        @("ConvertTextNumbersInSelection", "Converts text-formatted numbers in selected cells", "Current selection"),
        @("ConvertTextNumbersInUsedRange", "Converts text-formatted numbers in the active sheet used range", "Active sheet"),
        @("ConvertTextNumbersInColumns", "Converts text-formatted numbers in typed column letters", "InputBox columns, e.g. A,C,F:H"),
        @("NormalizePhonesInSelection", "Normalizes Korean phone numbers in selected cells", "Current selection"),
        @("NormalizePhonesInColumns", "Normalizes Korean phone numbers in typed column letters", "InputBox columns, e.g. B,D,F:G")
    )
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $rows[$r].Count; $c++) {
            $list.Cells.Item($r + 2, $c + 1).Value = $rows[$r][$c]
        }
    }
    $list.Range("A1:C1").Font.Bold = $true
    $list.Columns("A:C").AutoFit()

    $source = Add-SheetAfter $workbook "VBA Source Preview"
    $source.Range("A1").Value = "Line"
    $source.Range("B1").Value = "Code from src\MacroToolkit.bas"
    $source.Range("A1:B1").Font.Bold = $true
    $lines = Get-Content -Path $ModulePath
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $source.Cells.Item($i + 2, 1).Value2 = [string]($i + 1)
        $source.Cells.Item($i + 2, 2).Value2 = [string]$lines[$i]
    }
    $source.Columns("A").ColumnWidth = 8
    $source.Columns("B").ColumnWidth = 120
    $source.Columns("B").Font.Name = "Consolas"
    $source.Columns("B").Font.Size = 9
    $source.Range("A2").Select() | Out-Null
    $excel.ActiveWindow.FreezePanes = $true

    $readme.Activate() | Out-Null
    $workbook.Save()
    Write-Output "Updated workbook documentation sheets: $WorkbookPath"
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($true) | Out-Null } catch {}
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

