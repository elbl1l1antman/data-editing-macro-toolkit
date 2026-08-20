param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"

$modulePath = Join-Path $ProjectRoot "src\MacroToolkit.bas"
$publishPath = Join-Path $ProjectRoot "scripts\publish_release.ps1"

if (-not (Test-Path -LiteralPath $modulePath)) {
    throw "MacroToolkit.bas not found: $modulePath"
}
if (-not (Test-Path -LiteralPath $publishPath)) {
    throw "publish_release.ps1 not found: $publishPath"
}

$moduleSource = Get-Content -LiteralPath $modulePath -Raw -Encoding UTF8
$publishSource = Get-Content -LiteralPath $publishPath -Raw -Encoding UTF8

$failures = New-Object System.Collections.Generic.List[string]

function Assert-Contains {
    param(
        [string]$Name,
        [string]$Source,
        [string]$Needle
    )

    if (-not $Source.Contains($Needle)) {
        $script:failures.Add("$Name missing: $Needle")
    }
}

function Assert-Matches {
    param(
        [string]$Name,
        [string]$Source,
        [string]$Pattern
    )

    if (-not [regex]::IsMatch($Source, $Pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $script:failures.Add("$Name pattern not found: $Pattern")
    }
}

Assert-Contains "module cancel flag" $moduleSource "Private mToolkitCancelRequested As Boolean"
Assert-Contains "module cancel function" $moduleSource "Private Function Toolkit_ProgressCancelRequested() As Boolean"
Assert-Contains "module cancel raise" $moduleSource "Private Sub Toolkit_RaiseIfCancelled()"
Assert-Contains "module cancel error number" $moduleSource "TOOLKIT_CANCEL_ERROR_NUMBER"
Assert-Matches "progress update checks cancel after events" $moduleSource "Private Sub Toolkit_ProgressUpdate\b.*DoEvents\s+Toolkit_RaiseIfCancelled"
Assert-Matches "row progress starts for every positive row count" $moduleSource "Private Function Toolkit_ShouldShowRowProgress\b.*Toolkit_ShouldShowRowProgress\s*=\s*\(totalRows\s*>\s*0\)"
Assert-Matches "row progress updates every call" $moduleSource "Private Sub Toolkit_RowProgressUpdate\b(?!.*TOOLKIT_PROGRESS_BATCH_ROWS).*Toolkit_ProgressUpdate processedRows, totalRows, message"
Assert-Contains "pair loop checks cancel" $moduleSource "Toolkit_RaiseIfCancelled"

Assert-Contains "progress form cancel button" $publishSource '$cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)'
Assert-Contains "progress form cancel caption assignment" $publishSource '$cmdCancel.Caption = '
Assert-Contains "progress form cancel flag" $publishSource "Public CancelRequested As Boolean"
Assert-Contains "progress form cancel property" $publishSource "Public Property Get IsCancelled() As Boolean"
Assert-Contains "progress form cancel click" $publishSource "Private Sub cmdCancel_Click()"
Assert-Contains "progress form close requests cancel" $publishSource "CancelRequested = True"

if ($failures.Count -gt 0) {
    throw ($failures -join [Environment]::NewLine)
}

Write-Host "Progress cancellation static tests passed."
