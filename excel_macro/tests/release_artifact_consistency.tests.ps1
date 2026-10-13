param(
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = "Stop"
$publish = Get-Content -LiteralPath (Join-Path $ProjectRoot "scripts\publish_release.ps1") -Raw -Encoding UTF8

function Assert-Contains([string]$Name, [string]$Expected) {
    if (-not $publish.Contains($Expected)) { throw "$Name missing: $Expected" }
}

function Assert-NotContains([string]$Name, [string]$Unexpected) {
    if ($publish.Contains($Unexpected)) { throw "$Name still contains: $Unexpected" }
}

Assert-Contains "assembled workbook save" '$workbook.SaveAs($tempReleaseWorkbook, 52)'
Assert-Contains "latest workbook copied from assembled workbook" 'Copy-Item -Path $tempReleaseWorkbook -Destination $LatestWorkbook -Force'
Assert-Contains "sanitize before artifact saves" '$publishStage = "sanitize all release API credentials"'
Assert-Contains "latest credentials disabled" '$latestMayContainLocalCredentials = $false'
Assert-Contains "all artifacts protected" 'foreach ($artifact in @($ReleaseWorkbook, $LatestWorkbook, $ReleaseAddin, $LatestAddin))'
Assert-NotContains "raw development workbook release copy" 'Copy-Item -Path $DevWorkbook -Destination $ReleaseWorkbook -Force'

Write-Output "Release artifact consistency source checks passed."
