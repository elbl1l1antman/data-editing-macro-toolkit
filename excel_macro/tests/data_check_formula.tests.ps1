$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$modulePath = Join-Path $projectRoot "src\MacroToolkit.bas"
$module = Get-Content -Raw -Encoding UTF8 $modulePath

if ($module -notmatch 'Public Sub DataCheck_CreateCharacterCheckFormula\(\)') { throw "Character check macro is missing." }
if ($module -notmatch 'outputCell\.Formula2 = formulaText') { throw "Formula2 generation is missing." }
if ($module -notmatch 'outputCell\.Formula = formulaText') { throw "Formula fallback is missing." }
if (-not $module.Contains('Replace(result, "+",')) { throw "Plus pronunciation rule is missing." }
if (-not $module.Contains('Replace(result, "&",')) { throw "Ampersand pronunciation rule is missing." }

Write-Output "Data check formula static tests passed."
