param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)

. (Join-Path $PSScriptRoot "test_helpers.ps1")

$ErrorActionPreference = 'Stop'
$modulePath = Join-Path $ProjectRoot 'src\MacroToolkit.bas'
$source = Get-ToolkitVbaSource -ProjectRoot $ProjectRoot

function Assert-Contains([string]$Actual, [string]$Expected) {
    if ($Actual.IndexOf($Expected, [StringComparison]::Ordinal) -lt 0) { throw "Missing: $Expected" }
}

foreach ($name in @(
    'SampleExtract_BuildStratumKey',
    'SampleExtract_NormalizeMatchValue',
    'SampleExtract_SystematicIndexes',
    'SampleExtract_AppendUniqueRound',
    'SampleExtract_TestSystematicPositions',
    'SampleExtract_TestErrorStats'
)) { Assert-Contains $source $name }

foreach ($token in @(
    'Public Sub SampleExtract_Run()',
    'Private Function SampleExtract_LoadAllocationTargets',
    'Private Function SampleExtract_BuildReviewText',
    'Private Sub SampleExtract_WriteResultSheet',
    'Private Sub SampleExtract_WriteReviewSheet',
    '__SAMPLE_SOURCE_ROW',
    '__SAMPLE_REPLACEMENT_ROUND',
    '__SAMPLE_WEIGHT',
    '표본오차는 단순임의추출 근사치'
)) { Assert-Contains $source $token }

foreach ($token in @(
    'minimumRows = IIf(wizard.HasHeader, 2, 1)',
    'Private Function SampleExtract_CopyRows',
    '"실제 원표본"',
    '"실제 대체표본"'
)) { Assert-Contains $source $token }

function Get-VbaBlock([string]$Name) {
    $start = [regex]::Match($source, "(?m)^(Public|Private)\s+(Function|Sub)\s+$([regex]::Escape($Name))\b.*$")
    if (-not $start.Success) { throw "Procedure not found: $Name" }
    $kind = $start.Groups[2].Value
    $tail = $source.Substring($start.Index)
    $finish = [regex]::Match($tail, "(?m)^End $kind\s*$")
    if (-not $finish.Success) { throw "Procedure end not found: $Name" }
    $tail.Substring(0, $finish.Index + $finish.Length)
}

$names = @(
    'SampleExtract_SystematicIndexes',
    'SampleExtract_TestSystematicPositions',
    'SampleExtract_TestErrorStats'
)
$code = "Option Explicit`r`n" + (($names | ForEach-Object { Get-VbaBlock $_ }) -join "`r`n`r`n")

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Add()
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($code)

    $positions = [string]$excel.Run('SampleExtract_TestSystematicPositions', 10, 4, 0.5)
    if ($positions -ne '1,4,6,9') { throw "Unexpected positions: $positions" }

    $census = [string]$excel.Run('SampleExtract_TestErrorStats', 5, 5, 95, 0.5)
    Assert-Contains $census 'SE=0'
    Assert-Contains $census 'MOE=0'
    Assert-Contains $census 'WEIGHT=1'

    $partial = [string]$excel.Run('SampleExtract_TestErrorStats', 100, 20, 95, 0.5)
    Assert-Contains $partial 'WEIGHT=5'
    Assert-Contains $partial 'PI=0.2'
    Write-Host 'Sample extraction logic tests passed.'
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) | Out-Null }
    if ($null -ne $excel) {
        $excel.Quit() | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
