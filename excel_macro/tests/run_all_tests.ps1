param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"
$tests = Get-ChildItem -LiteralPath (Join-Path $ProjectRoot "tests") -Filter "*.tests.ps1" | Sort-Object Name

foreach ($test in $tests) {
    Write-Host ("== " + $test.Name + " ==")
    & powershell -NoProfile -ExecutionPolicy Bypass -File $test.FullName -ProjectRoot $ProjectRoot
    if ($LASTEXITCODE -ne 0) {
        throw ($test.Name + " failed with exit code " + $LASTEXITCODE)
    }
}

Write-Host ("ALL_TESTS_PASSED=" + $tests.Count)
