param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)

$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\MacroToolkit.bas') -Raw -Encoding UTF8
$publisher = Get-Content -LiteralPath (Join-Path $ProjectRoot 'scripts\publish_release.ps1') -Raw -Encoding UTF8
$about = [regex]::Match($source, '(?ms)^Public Sub Toolkit_About\(\).*?^End Sub').Value
$email = 'elbr1ll1antant@gmail.com'
if (-not $about.Contains($email)) { throw 'Contact email missing from version popup.' }
if (-not $publisher.Contains('contact = "' + $email + '"')) { throw 'Manifest contact differs from version popup.' }
Write-Host 'Version popup and manifest contact tests passed.'
