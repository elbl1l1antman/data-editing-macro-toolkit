param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$source = Get-Content -LiteralPath (Join-Path $ProjectRoot "scripts\publish_release.ps1") -Raw -Encoding UTF8
foreach ($token in @("RefreshOrderNumbers", "ClearOrderNumbers", 'ctl.Caption = CStr(orderNo) & ". " & ctl.Caption')) {
    if (-not $source.Contains($token)) { throw "Company name order UI token missing: $token" }
}

Write-Host "Company name order UI tests passed."
