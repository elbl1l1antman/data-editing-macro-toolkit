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
    $moduleText = Get-Content -Raw -Encoding UTF8 -Path $ModulePath
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
$ReleaseAddinName = "DataEditingMacroToolkit_$VersionTag.xlam"
$LatestWorkbookName = "DataEditingMacroToolkit_latest.xlsm"
$LatestAddinName = "DataEditingMacroToolkit_latest.xlam"

New-Item -ItemType Directory -Force -Path $VersionDir,$LatestDir | Out-Null

$ReleaseWorkbook = Join-Path $VersionDir $ReleaseWorkbookName
$ReleaseAddin = Join-Path $VersionDir $ReleaseAddinName
$LatestWorkbook = Join-Path $LatestDir $LatestWorkbookName
$LatestAddin = Join-Path $LatestDir $LatestAddinName

Copy-Item -Path $DevWorkbook -Destination $ReleaseWorkbook -Force
Copy-Item -Path $DevWorkbook -Destination $LatestWorkbook -Force
Copy-Item -Path $ModulePath -Destination (Join-Path $VersionDir "MacroToolkit.bas") -Force
Copy-Item -Path $ModulePath -Destination (Join-Path $LatestDir "MacroToolkit.bas") -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open($DevWorkbook)

    try {
        foreach ($component in @($workbook.VBProject.VBComponents)) {
            if ($component.Type -eq 1 -and ($component.Name -eq "MacroToolkit" -or $component.Name -like "Module*")) {
                $workbook.VBProject.VBComponents.Remove($component)
            }
        }
        $imported = $workbook.VBProject.VBComponents.Import($ModulePath)
        $imported.Name = "MacroToolkit"
    }
    catch {
        throw "Could not import VBA module into .xlam. Enable Excel Trust Center > Macro Settings > Trust access to the VBA project object model. Details: $($_.Exception.Message)"
    }

    $workbook.IsAddin = $true

    $tempAddin = Join-Path $VersionDir ("_tmp_" + $ReleaseAddinName)
    if (Test-Path -LiteralPath $tempAddin) { Remove-Item -LiteralPath $tempAddin -Force }

    $workbook.SaveAs($tempAddin, 55)
    $workbook.Close($false)
    $workbook = $null

    Copy-Item -Path $tempAddin -Destination $ReleaseAddin -Force
    Copy-Item -Path $tempAddin -Destination $LatestAddin -Force
    Remove-Item -LiteralPath $tempAddin -Force
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) | Out-Null } catch {}
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

$commit = "unknown"
try {
    $commit = (git -C $ProjectRoot rev-parse --short HEAD).Trim()
    $dirty = (git -C $ProjectRoot status --short)
    if ($dirty) { $commit = "$commit-dirty" }
} catch {}

$manifest = [ordered]@{
    name = "Data Editing Macro Toolkit"
    version = $Version
    versionTag = $VersionTag
    publishedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss K")
    gitCommit = $commit
    addin = $ReleaseAddinName
    workbook = $ReleaseWorkbookName
    sourceModule = "MacroToolkit.bas"
    installFile = $ReleaseAddinName
    containsVbaModule = $true
    notes = "Install the .xlam file as an Excel add-in. The .xlsm file is kept for workbook inspection and development review."
}

$manifest | ConvertTo-Json | Set-Content -Encoding UTF8 -Path (Join-Path $VersionDir "manifest.json")
$manifest | ConvertTo-Json | Set-Content -Encoding UTF8 -Path (Join-Path $LatestDir "manifest.json")

Write-Output "Published add-in: $ReleaseAddin"
Write-Output "Published workbook: $ReleaseWorkbook"
Write-Output "Updated latest add-in: $LatestAddin"
