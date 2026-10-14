function Get-ToolkitVbaSource {
    param([Parameter(Mandatory = $true)][string]$ProjectRoot)

    $modules = Get-ChildItem -LiteralPath (Join-Path $ProjectRoot "src") -Filter "*.bas" -File |
        Sort-Object @{ Expression = { if ($_.Name -eq "MacroToolkit.bas") { 0 } else { 1 } } }, Name
    return (($modules | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8 }) -join "`r`n")
}
