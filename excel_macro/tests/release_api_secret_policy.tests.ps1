param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"

$publishPath = Join-Path $ProjectRoot "scripts\publish_release.ps1"
if (-not (Test-Path -LiteralPath $publishPath)) {
    throw "publish_release.ps1 not found: $publishPath"
}

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

Assert-Contains "workbook sanitizer" $publishSource "function Clear-WorkbookStoredApiSecrets"
Assert-Contains "file sanitizer" $publishSource "function Clear-WorkbookFileStoredApiSecrets"
Assert-Contains "latest credential mode" $publishSource 'latestMayContainLocalCredentials = $true'
Assert-Contains "versioned credential mode" $publishSource 'versionedArtifactsSanitized = $true'
Assert-Contains "manifest policy" $publishSource "apiKeyReleasePolicy"

Assert-Matches "versioned workbook copy is sanitized" $publishSource 'Copy-Item\s+-Path\s+\$DevWorkbook\s+-Destination\s+\$ReleaseWorkbook\s+-Force.*Clear-WorkbookFileStoredApiSecrets\s+-Excel\s+\$excel\s+-Path\s+\$ReleaseWorkbook'
Assert-Matches "latest addin is saved before workbook sanitizer" $publishSource '\$workbook\.SaveAs\(\$tempLatestAddin,\s*55\).*Clear-WorkbookStoredApiSecrets\s+-Workbook\s+\$workbook.*\$workbook\.SaveAs\(\$tempReleaseAddin,\s*55\)'
Assert-Matches "latest addin copy uses latest temp" $publishSource 'Copy-Item\s+-Path\s+\$tempLatestAddin\s+-Destination\s+\$LatestAddin\s+-Force'
Assert-Matches "release addin copy uses sanitized release temp" $publishSource 'Copy-Item\s+-Path\s+\$tempReleaseAddin\s+-Destination\s+\$ReleaseAddin\s+-Force'

if ($failures.Count -gt 0) {
    throw ($failures -join [Environment]::NewLine)
}

Write-Host "Release API secret policy static tests passed."
