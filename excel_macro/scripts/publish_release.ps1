param(
    [string]$Version,
    [string]$ProjectRoot
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ProjectRoot) {
    $ProjectRoot = Split-Path -Parent $ScriptDir
}

$ModulePath = Join-Path $ProjectRoot "src\MacroToolkit.bas"
$DevWorkbook = Join-Path $ProjectRoot "dev\DataEditingMacroToolkit_dev.xlsm"
$ReleaseRoot = Join-Path $ProjectRoot "release"

if (-not (Test-Path -LiteralPath $ModulePath)) {
    throw "VBA module not found: $ModulePath"
}

if (-not (Test-Path -LiteralPath $DevWorkbook)) {
    throw "Development workbook not found: $DevWorkbook"
}

if (-not $Version) {
    $moduleText = Get-Content -Raw -Path $ModulePath
    if ($moduleText -match 'TOOLKIT_VERSION\s+As\s+String\s*=\s*"([^"]+)"') {
        $Version = $Matches[1]
    }
    else {
        throw "Could not read TOOLKIT_VERSION from $ModulePath"
    }
}

$VersionTag = "v$Version"
$VersionDir = Join-Path $ReleaseRoot $VersionTag
$LatestDir = Join-Path $ReleaseRoot "latest"
$ReleaseWorkbookName = "DataEditingMacroToolkit_$VersionTag.xlsm"

New-Item -ItemType Directory -Force -Path $VersionDir,$LatestDir | Out-Null

$ReleaseWorkbook = Join-Path $VersionDir $ReleaseWorkbookName
$LatestWorkbook = Join-Path $LatestDir "DataEditingMacroToolkit_latest.xlsm"

Copy-Item -Path $DevWorkbook -Destination $ReleaseWorkbook -Force
Copy-Item -Path $DevWorkbook -Destination $LatestWorkbook -Force
Copy-Item -Path $ModulePath -Destination (Join-Path $VersionDir "MacroToolkit.bas") -Force
Copy-Item -Path $ModulePath -Destination (Join-Path $LatestDir "MacroToolkit.bas") -Force

$commit = "unknown"
try {
    $commit = (git -C $ProjectRoot rev-parse --short HEAD).Trim()
} catch {}

$manifest = [ordered]@{
    name = "Data Editing Macro Toolkit"
    version = $Version
    versionTag = $VersionTag
    publishedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss K")
    gitCommit = $commit
    workbook = $ReleaseWorkbookName
    sourceModule = "MacroToolkit.bas"
}

$manifest | ConvertTo-Json | Set-Content -Path (Join-Path $VersionDir "manifest.json")
$manifest | ConvertTo-Json | Set-Content -Path (Join-Path $LatestDir "manifest.json")

Write-Output "Published: $ReleaseWorkbook"
Write-Output "Updated latest: $LatestWorkbook"
