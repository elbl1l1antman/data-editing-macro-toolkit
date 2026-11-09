param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$publisher = Join-Path $ProjectRoot 'scripts\publish_release.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($publisher, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Publisher parse failed' }
$fn = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Add-ToolkitReleaseNotes' }, $true)
if (-not $fn) { throw 'Release note builder missing' }
Invoke-Expression $fn.Extent.Text
$source = Get-Content -Raw -Encoding UTF8 (Join-Path $ProjectRoot 'src\MacroToolkit.bas')
$version = [regex]::Match($source, 'TOOLKIT_VERSION As String = "([^"]+)"').Groups[1].Value
$helper = [regex]::Match($source, '(?ms)^Public Function Toolkit_GetLatestUpdates\(\) As String.*?^End Function').Value
if (-not $version -or -not $helper) { throw 'Version or release note helper missing' }
$excel = New-Object -ComObject Excel.Application
$workbook = $null
try {
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    Add-ToolkitReleaseNotes -Workbook $workbook -Path (Join-Path $ProjectRoot "docs\releases\v$version.md") -ReleaseVersion $version
    $notes = $workbook.Worksheets.Item('_ToolkitReleaseNotes')
    if ($notes.Visible -ne 2 -or $notes.Range('A1').Value2 -ne $version) { throw 'Release notes sheet version or visibility incorrect' }
    $first = [string]$notes.Range('A2').Value2
    if (-not $first) { throw 'Release summary is empty' }
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString("Option Explicit`r`nPublic Const TOOLKIT_VERSION As String = `"$version`"`r`n" + $helper)
    $summary = [string]$excel.Run("'$($workbook.Name)'!Toolkit_GetLatestUpdates")
    if (-not $summary.Contains($first)) { throw 'About summary does not include embedded release notes' }
    $notes.Range('A1').Value2 = 'old-version'
    $fallback = [string]$excel.Run("'$($workbook.Name)'!Toolkit_GetLatestUpdates")
    if ($fallback.Contains($first) -or -not $fallback) { throw 'Stale release notes were not rejected' }
    Write-Host "Embedded release summary verified for $version."
}
finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
