param(
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = "Stop"
$sourceDir = Join-Path $ProjectRoot "src"
$modules = @(Get-ChildItem -LiteralPath $sourceDir -Filter "*.bas" -File)

if ($modules.Count -lt 8) {
    throw "Expected functional VBA modules, found $($modules.Count)."
}

$names = New-Object System.Collections.Generic.List[string]
foreach ($module in $modules) {
    $lineCount = (Get-Content -LiteralPath $module.FullName).Count
    if ($lineCount -gt 7000) {
        throw "VBA module is still too large: $($module.Name) ($lineCount lines)."
    }

    $text = Get-Content -LiteralPath $module.FullName -Raw -Encoding UTF8
    foreach ($match in [regex]::Matches($text, '(?im)^\s*(?:Public|Private)\s+(?:Sub|Function)\s+([A-Za-z][A-Za-z0-9_]*)\b')) {
        $names.Add($match.Groups[1].Value)
    }
}

$duplicates = @($names | Group-Object | Where-Object Count -gt 1)
if ($duplicates.Count -gt 0) {
    throw "Duplicate VBA procedures: $($duplicates.Name -join ', ')"
}

$publish = Get-Content -LiteralPath (Join-Path $ProjectRoot "scripts\publish_release.ps1") -Raw -Encoding UTF8
if (-not $publish.Contains('Get-ToolkitModulePaths')) {
    throw "Publisher does not discover all VBA modules."
}
if (-not $publish.Contains('sourceModules')) {
    throw "Release manifest does not list all source modules."
}

Write-Output "VBA module structure checks passed."
