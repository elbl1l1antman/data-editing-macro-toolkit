param(
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot)
)
. (Join-Path $PSScriptRoot "test_helpers.ps1")

$ErrorActionPreference = "Stop"
$source = Get-ToolkitVbaSource -ProjectRoot $ProjectRoot

function Assert-Contains([string]$Name, [string]$Expected) {
    if (-not $source.Contains($Expected)) { throw "$Name missing: $Expected" }
}

Assert-Contains "progress throttle constant" "TOOLKIT_PROGRESS_REFRESH_SECONDS"
Assert-Contains "progress throttle helper" "Private Function Toolkit_ShouldRefreshProgress"
Assert-Contains "database batch size" "DBMERGE_OUTPUT_BATCH_ROWS"
Assert-Contains "database output batch" "Private Function DBMerge_BuildOutputBatch"
Assert-Contains "database bulk write" "Resize(batchCount, UBound(outputBatch, 2)).Value = outputBatch"

$dbStart = $source.IndexOf("Private Function DBMerge_WriteMergedSheet")
$dbEnd = $source.IndexOf("Private Function DBMerge_DataHeaders", $dbStart)
$dbBody = $source.Substring($dbStart, $dbEnd - $dbStart)
if ($dbBody -match 'outWs\.Cells\(outRow') {
    throw "DB merge still writes result cells row by row."
}

Write-Output "Performance path source checks passed."
