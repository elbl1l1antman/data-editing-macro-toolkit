param(
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot)
)
. (Join-Path $PSScriptRoot "test_helpers.ps1")

$ErrorActionPreference = "Stop"

$modulePath = Join-Path $ProjectRoot "src\MacroToolkit.bas"
$source = Get-ToolkitVbaSource -ProjectRoot $ProjectRoot

function Assert-Contains([string]$Name, [string]$Expected) {
    if (-not $source.Contains($Expected)) {
        throw "$Name missing: $Expected"
    }
}

Assert-Contains "rollback snapshot type" "Function Toolkit_CreateRangeSnapshot"
Assert-Contains "rollback restore type" "Sub Toolkit_RestoreRangeSnapshot"
Assert-Contains "formula2 snapshot" 'item.Add "formula2", area.Formula2'
Assert-Contains "text number rollback" "Toolkit_RestoreRangeSnapshot snapshot"
Assert-Contains "residual blank rollback" "Toolkit_RestoreRangeSnapshot snapshot"
Assert-Contains "database temporary output" "DBMerge_PrepareTemporaryOutputSheet"
Assert-Contains "database output commit" "DBMerge_CommitTemporaryOutputSheet"
Assert-Contains "database commit name rollback" 'temporarySheet.Name = DBMerge_UniqueSheetName(temporarySheet.Parent, "__DBMERGE_FAILED")'
Assert-Contains "source ID insertion marked after success" 'rollbackItem("insertedColumn") = True'
Assert-Contains "source ID cell state" 'changedCells.Add cellState'

Write-Output "Operation rollback source checks passed."
