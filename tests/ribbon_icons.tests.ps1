param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path)

$ErrorActionPreference = "Stop"
$source = Get-Content -Raw -Encoding UTF8 (Join-Path $ProjectRoot "scripts\publish_release.ps1")
$match = [regex]::Match($source, '(?s)<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">.*?</customUI>')
if (-not $match.Success) { throw "Ribbon XML not found" }
[xml]$xml = $match.Value
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace("r", "http://schemas.microsoft.com/office/2006/01/customui")
$controls = @($xml.SelectNodes('//r:group|//r:button', $ns))
if ($controls.Count -eq 0) { throw "No ribbon controls found" }
foreach ($control in $controls) {
    if (-not $control.imageMso) { throw "Missing image: $($control.id)" }
    if ($control.imageMso -in @('NumberFormatGeneral', 'CellsMerge', 'InsertFunction', 'GoToTop', 'ConditionalFormattingHighlightCellsRules')) {
        throw "Unsupported image: $($control.imageMso)"
    }
}
Write-Host "Ribbon icons configured for $($controls.Count) controls."
