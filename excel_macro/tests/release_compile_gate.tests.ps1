param(
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot)
)
. (Join-Path $PSScriptRoot "test_helpers.ps1")

$ErrorActionPreference = "Stop"
$module = Get-ToolkitVbaSource -ProjectRoot $ProjectRoot
$publish = Get-Content -LiteralPath (Join-Path $ProjectRoot "scripts\publish_release.ps1") -Raw -Encoding UTF8

function Assert-Contains([string]$Name, [string]$Text, [string]$Expected) {
    if (-not $Text.Contains($Expected)) { throw "$Name missing: $Expected" }
}

Assert-Contains "VBA smoke entry point" $module "Public Function Toolkit_CompileSmokeTest() As String"
Assert-Contains "publisher compile stage" $publish '$publishStage = "compile smoke test"'
Assert-Contains "publisher runs smoke test" $publish 'Toolkit_CompileSmokeTest'
Assert-Contains "final artifacts run smoke test" $publish '$verifySmokeResult = [string]$verifyExcel.Run($verifySmokeMacro)'
Assert-Contains "publisher encoding guard" $publish "function Assert-NoBrokenKoreanText"
Assert-Contains "publisher disables template events" $publish '$excel.EnableEvents = $false'
Assert-Contains "publisher opens editable template without links" $publish '$excel.Workbooks.Open($DevWorkbook, 0, $false)'

$smokeStart = $module.IndexOf("Public Function Toolkit_CompileSmokeTest() As String")
$smokeEnd = $module.IndexOf("End Function", $smokeStart)
$smokeBody = $module.Substring($smokeStart, $smokeEnd - $smokeStart)
if ($smokeBody.Contains("VBA.UserForms.Add")) {
    throw "Compile smoke test must not initialize UserForms while Excel is hidden."
}

Write-Output "Release compile gate source checks passed."
