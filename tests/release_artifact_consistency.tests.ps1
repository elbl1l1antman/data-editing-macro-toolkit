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

Assert-Contains "assembled add-in save" '$workbook.SaveAs($tempReleaseAddin, 55)'
Assert-Contains "latest add-in copied from assembled add-in" 'Copy-Item -Path $tempReleaseAddin -Destination $LatestAddin -Force'
Assert-Contains "stale release files removed" '$StaleReleaseExtensions = @(".bas", ".xlsm")'
Assert-Contains "sanitize before artifact saves" '$publishStage = "sanitize all release API credentials"'
Assert-Contains "latest credentials disabled" '$latestMayContainLocalCredentials = $false'
Assert-Contains "all artifacts verified" 'foreach ($artifact in @($ReleaseAddin, $LatestAddin))'
Assert-NotContains "raw development workbook release copy" 'Copy-Item -Path $DevWorkbook -Destination'
Assert-NotContains "release workbook output" '$ReleaseWorkbook'
Assert-NotContains "latest workbook output" '$LatestWorkbook'
Assert-NotContains "release source module copy" '-Destination (Join-Path $LatestDir $moduleFileName)'

Write-Output "Release artifact consistency source checks passed."
