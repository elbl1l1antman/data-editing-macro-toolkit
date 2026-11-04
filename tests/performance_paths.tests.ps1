param(
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot)
)
. (Join-Path $PSScriptRoot "test_helpers.ps1")

$ErrorActionPreference = "Stop"
$source = Get-ToolkitVbaSource -ProjectRoot $ProjectRoot

function Assert-Contains([string]$Name, [string]$Expected) {
    if (-not $source.Contains($Expected)) { throw "$Name missing: $Expected" }
}

Assert-Contains "progress throttle constant" "TOOLKIT_PROGRESS_REFRESH_SECONDS"
Assert-Contains "progress throttle helper" "Private Function Toolkit_ShouldRefreshProgress"
Assert-Contains "database batch size" "DBMERGE_OUTPUT_BATCH_ROWS"
Assert-Contains "database output batch" "Private Function DBMerge_BuildOutputBatch"
Assert-Contains "database bulk write" "Resize(batchCount, UBound(outputBatch, 2)).Value = outputBatch"
Assert-Contains "phone chunk size" "PHONE_NORMALIZE_CHUNK_ROWS"
Assert-Contains "phone chunk processor" "Private Sub Phone_NormalizeChunk"
Assert-Contains "phone chunk bulk read" "chunkRange.Value2"
Assert-Contains "phone chunk bulk write" "chunkRange.Value2 = values"
Assert-Contains "company chunk size" "COMPANY_NAME_CHUNK_ROWS"
Assert-Contains "company chunk processor" "Private Sub CompanyName_CleanChunk"
Assert-Contains "company bulk write" "outputChunk.Value2 = outputValues"
Assert-Contains "company regex cache" "Private Function CompanyName_GetRegex"
Assert-Contains "text number chunk processor" "Private Sub DataCleanup_ConvertTextNumberChunk"
Assert-Contains "digit strip chunk processor" "Private Sub DataCleanup_StripDigitsChunk"
Assert-Contains "representative split bulk read" "sourceValues = DataCleanup_ReadRangeMatrix(sourceRange, False)"
Assert-Contains "representative split bulk write" "secondRange.Value2 = secondValues"

$dbStart = $source.IndexOf("Private Function DBMerge_WriteMergedSheet")
$dbEnd = $source.IndexOf("Private Function DBMerge_DataHeaders", $dbStart)
$dbBody = $source.Substring($dbStart, $dbEnd - $dbStart)
if ($dbBody -match 'outWs\.Cells\(outRow') {
    throw "DB merge still writes result cells row by row."
}

$phoneStart = $source.IndexOf("Private Sub NormalizePhonesInRange")
$phoneEnd = $source.IndexOf("Private Function NormalizeKoreanPhone", $phoneStart)
$phoneBody = $source.Substring($phoneStart, $phoneEnd - $phoneStart)
if ($phoneBody -match 'For Each cell In target\.Cells') {
    throw "Phone normalization still processes the target cell by cell."
}

$splitStart = $source.IndexOf("Private Function Phone_CreateRepresentativeSplitColumns")
$splitEnd = $source.IndexOf("Private Sub NormalizePhonesInRange", $splitStart)
$splitBody = $source.Substring($splitStart, $splitEnd - $splitStart)
if ($splitBody -match 'sourceRange\.Cells\(') {
    throw "Representative-number splitting still reads source cells individually."
}

$numberStart = $source.IndexOf("Private Sub ConvertTextNumbersInRange")
$numberEnd = $source.IndexOf("Public Sub DataCleanup_ClearResidualBlankValues", $numberStart)
$numberBody = $source.Substring($numberStart, $numberEnd - $numberStart)
if ($numberBody -match 'For Each cell In target\.Cells') {
    throw "Text-number conversion still processes the target cell by cell."
}

$stripStart = $source.IndexOf("Private Sub StripDigitsToCopiedColumns")
$stripEnd = $source.IndexOf("Private Function TryNumberFromText", $stripStart)
$stripBody = $source.Substring($stripStart, $stripEnd - $stripStart)
if ($stripBody -match 'For Each cell In outputRange\.Cells') {
    throw "Digit stripping still processes the output cell by cell."
}

$companyPath = Join-Path $ProjectRoot "src\CompanyNameTools.bas"
$companySource = Get-Content -LiteralPath $companyPath -Raw -Encoding UTF8
$regexCreateCount = ([regex]::Matches($companySource, 'CreateObject\("VBScript\.RegExp"\)')).Count
if ($regexCreateCount -ne 1) {
    throw "Company-name cleanup should create RegExp objects only inside its shared cache. Found: $regexCreateCount"
}

Write-Output "Performance path source checks passed."
