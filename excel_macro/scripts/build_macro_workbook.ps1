param(
    [string]$OutputPath,
    [string]$ModulePath
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir

if (-not $OutputPath) {
    $OutputPath = Join-Path $ProjectRoot "dev\DataEditingMacroToolkit_dev.xlsm"
}

if (-not $ModulePath) {
    $ModulePath = Join-Path $ProjectRoot "src\MacroToolkit.bas"
}

if (-not (Test-Path -LiteralPath $ModulePath)) {
    throw "VBA module not found: $ModulePath"
}

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputPath) | Out-Null

$ToolkitVersion = "unknown"
$ModuleText = Get-Content -Raw -Encoding UTF8 -Path $ModulePath
if ($ModuleText -match 'TOOLKIT_VERSION\s+As\s+String\s*=\s*"([^"]+)"') {
    $ToolkitVersion = $Matches[1]
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null

try {
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "MacroToolkit"
    $sheet.Range("A1").Value = "Data Editing Macro Toolkit"
    $sheet.Range("A2").Value = "Version"
    $sheet.Range("B2").Value = $ToolkitVersion
    $sheet.Range("A4").Value = "Install status"
    $sheet.Range("B4").Value = "Pending VBA import"
    $sheet.Range("A6").Value = "How to finish installation"
    $sheet.Range("A7").Value = "1. In Excel, enable Trust access to the VBA project object model."
    $sheet.Range("A8").Value = "2. Run this script again, or import src\MacroToolkit.bas in the VBA editor."
    $sheet.Range("A10").Value = "Included macros"
    $sheet.Range("A11").Value = "ConvertTextNumbersInSelection"
    $sheet.Range("A12").Value = "ConvertTextNumbersInUsedRange"
    $sheet.Range("A13").Value = "ConvertTextNumbersInColumns"
    $sheet.Range("A14").Value = "NormalizePhonesInSelection"
    $sheet.Range("A15").Value = "NormalizePhonesInColumns"
    $sheet.Range("A16").Value = "OutlierDetect_Setup"
    $sheet.Range("A17").Value = "OutlierDetect_Run"
    $sheet.Columns("A:B").AutoFit()

    if ($null -ne $workbook.VBProject) {
        $workbook.VBProject.VBComponents.Import($ModulePath) | Out-Null
        $excel.Run("'" + $workbook.Name + "'!Toolkit_SetupWorkbook")
        $workbook.Worksheets.Item("_MacroToolkit").Visible = 2
        $sheet.Range("B4").Value = "VBA imported"
    }

    $workbook.SaveAs($OutputPath, 52)
    $workbook.Close($true)
    Write-Output "Created: $OutputPath"

    $DocScript = Join-Path $ScriptDir "update_workbook_docs.ps1"
    if (Test-Path -LiteralPath $DocScript) {
        & $DocScript -WorkbookPath $OutputPath -ModulePath $ModulePath
    }
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) | Out-Null } catch {}
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}



