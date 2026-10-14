param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)
. (Join-Path $PSScriptRoot "test_helpers.ps1")

$ErrorActionPreference = "Stop"
$moduleText = Get-ToolkitVbaSource -ProjectRoot $ProjectRoot
$version = [regex]::Match($moduleText, 'TOOLKIT_VERSION As String = "([^"]+)"').Groups[1].Value
if (-not $version) { throw "Could not read TOOLKIT_VERSION." }

$artifacts = @(
    (Join-Path $ProjectRoot "release\v$version\DataEditingMacroToolkit_v$version.xlsm"),
    (Join-Path $ProjectRoot "release\v$version\DataEditingMacroToolkit_v$version.xlam"),
    (Join-Path $ProjectRoot "release\latest\DataEditingMacroToolkit_latest.xlsm"),
    (Join-Path $ProjectRoot "release\latest\DataEditingMacroToolkit_latest.xlam")
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    foreach ($artifact in $artifacts) {
        if (-not (Test-Path -LiteralPath $artifact)) { throw "Missing release artifact: $artifact" }
        $workbook = $excel.Workbooks.Open($artifact, 0, $true)
        try {
            $macro = "'{0}'!Toolkit_CompileSmokeTest" -f $workbook.Name.Replace("'", "''")
            $result = [string]$excel.Run($macro)
            if (-not $result.StartsWith($version + "|forms=14")) {
                throw "Unexpected runtime version for $artifact`: $result"
            }
            if ($workbook.VBProject.Protection -ne 1) {
                throw "VBA project is not protected: $artifact"
            }
        }
        finally {
            $workbook.Close($false)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }
    }
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

Write-Output "Release artifact runtime tests passed for v$version."
