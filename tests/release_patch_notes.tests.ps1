param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)
. (Join-Path $PSScriptRoot "test_helpers.ps1")

$ErrorActionPreference = "Stop"

$modulePath = Join-Path $ProjectRoot "src\MacroToolkit.bas"
$publishPath = Join-Path $ProjectRoot "scripts\publish_release.ps1"
$moduleSource = Get-ToolkitVbaSource -ProjectRoot $ProjectRoot
$publishSource = Get-Content -LiteralPath $publishPath -Raw -Encoding UTF8

if ($moduleSource -notmatch 'TOOLKIT_VERSION\s+As\s+String\s*=\s*"([^"]+)"') {
    throw "TOOLKIT_VERSION을 찾을 수 없습니다."
}

$version = $Matches[1]
$notesPath = Join-Path $ProjectRoot ("docs\releases\v{0}.md" -f $version)
$changeLogPath = Join-Path $ProjectRoot "CHANGELOG.md"
if (-not (Test-Path -LiteralPath $notesPath)) {
    throw "현재 버전 패치노트가 없습니다: $notesPath"
}

$notes = Get-Content -LiteralPath $notesPath -Raw -Encoding UTF8
foreach ($heading in @("## 주요 변경", "## 사용 방법", "## 수정된 문제", "## 주의사항")) {
    if (-not $notes.Contains($heading)) {
        throw "패치노트 필수 항목이 없습니다: $heading"
    }
}

if (Test-Path -LiteralPath $changeLogPath) {
    $changeLog = Get-Content -LiteralPath $changeLogPath -Raw -Encoding UTF8
    if (-not $changeLog.Contains("## $version")) {
        throw "통합 변경 이력에 현재 버전이 없습니다: $version"
    }
}

foreach ($needle in @(
    'PatchNotesSource',
    'PATCH_NOTES.md',
    'patchNotesFile = "PATCH_NOTES.md"'
)) {
    if (-not $publishSource.Contains($needle)) {
        throw "배포 스크립트에 패치노트 처리가 없습니다: $needle"
    }
}

Write-Host "Release patch notes static tests passed for v$version."
