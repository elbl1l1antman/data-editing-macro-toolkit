param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$workbook = $null; $modulePath = $null
$cachePath = Join-Path $env:TEMP ('DartCacheTest_' + [guid]::NewGuid().ToString('N') + '.xml')
try {
    $publisher = Get-Content -Raw -Encoding UTF8 (Join-Path $ProjectRoot 'scripts\publish_release.ps1')
    if (-not $publisher.Contains('onAction="Ribbon_DartRefreshCompanyMaster"')) { throw 'DART refresh ribbon button missing' }
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml('<result><list><corp_code>00123456</corp_code><corp_name>ACME</corp_name><stock_code>005930</stock_code></list><list><corp_code>00123457</corp_code><corp_name>SAME</corp_name></list><list><corp_code>00123458</corp_code><corp_name>SAME</corp_name></list></result>')
    $doc.Save($cachePath)
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Const BIZ_APP_NAME As String = "DataEditingMacroToolkit"
Public Const TOOLKIT_CANCEL_ERROR_NUMBER As Long = vbObjectError + 9130
Public Function Toolkit_GetRangeFromUserForm(a As String, b As String, ByRef c As Range) As Boolean
End Function
Public Sub Toolkit_ProgressShow(a As String, b As String, c As Long)
End Sub
Public Sub Toolkit_ProgressUpdate(a As Long, b As Long, c As String)
End Sub
Public Sub Toolkit_ProgressClose()
End Sub
Public Sub Toolkit_RaiseIfCancelled()
End Sub
Public Function CheckLocalCache(ByVal path As String) As String
    On Error GoTo Failed
    Dim names As Object, codes As Object, stocks As Object, issue As String
    Set names = CreateObject("Scripting.Dictionary")
    Set codes = CreateObject("Scripting.Dictionary")
    Set stocks = CreateObject("Scripting.Dictionary")
    Dart_LoadCompanyMaster String$(40, "X"), names, codes, stocks, False, path
    If Dart_FindCompany("ACME", names, codes, stocks, issue) <> "00123456" Then Err.Raise 513, , "Cached lookup failed"
    issue = ""
    If Dart_FindCompany("SAME", names, codes, stocks, issue) <> "" Or Len(issue) = 0 Then Err.Raise 513, , "Duplicate not flagged"
    CheckLocalCache = "PASS"
    Exit Function
Failed:
    CheckLocalCache = CStr(Err.Number) & ": " & Err.Description
End Function
Public Function TryReplaceCache(ByVal path As String) As String
    On Error GoTo Failed
    Dim doc As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.LoadXML "<result><list><corp_code>00999999</corp_code><corp_name>NEW</corp_name></list></result>"
    Dart_SaveMasterCache doc, path
    TryReplaceCache = "PASS"
    Exit Function
Failed:
    TryReplaceCache = "ERROR"
End Function
'@)
    $modulePath = Join-Path $env:TEMP ('DartTools_' + [guid]::NewGuid().ToString('N') + '.bas')
    $source = Get-Content -Raw -Encoding UTF8 (Join-Path $ProjectRoot 'src\DartTools.bas')
    if (-not $source.Contains('Dart_MasterNeedsWarning(cachePath)') -or -not $source.Contains('forceRefresh, cachePath')) { throw 'Stale warning or refresh flow missing' }
    $enc = [Text.Encoding]::GetEncoding(949, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
    [IO.File]::WriteAllText($modulePath, $source, $enc)
    $workbook.VBProject.VBComponents.Import($modulePath) | Out-Null
    $prefix = "'$($workbook.Name)'!"
    if ($excel.Run($prefix + 'CheckLocalCache', $cachePath) -ne 'PASS') { throw 'Cached master lookup failed' }
    if ($excel.Run($prefix + 'Dart_MasterNeedsWarning', $cachePath)) { throw 'Fresh master marked stale' }
    (Get-Item -LiteralPath $cachePath).LastWriteTime = (Get-Date).AddMonths(-7)
    if (-not $excel.Run($prefix + 'Dart_MasterNeedsWarning', $cachePath)) { throw 'Old master not marked stale' }
    $oldHash = (Get-FileHash -LiteralPath $cachePath).Hash
    $locked = [IO.File]::Open($cachePath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    try { if ($excel.Run($prefix + 'TryReplaceCache', $cachePath) -ne 'ERROR') { throw 'Locked cache replacement unexpectedly succeeded' } }
    finally { $locked.Dispose() }
    if ((Get-FileHash -LiteralPath $cachePath).Hash -ne $oldHash) { throw 'Failed replacement modified the old cache' }
    if ($excel.Run($prefix + 'TryReplaceCache', $cachePath) -ne 'PASS') { throw 'Cache replacement failed' }
    if ($excel.Run($prefix + 'Dart_MasterNeedsWarning', $cachePath)) { throw 'Refreshed cache marked stale' }
    $updated = New-Object System.Xml.XmlDocument
    $updated.Load($cachePath)
    if ($updated.SelectSingleNode('/result/list/corp_code').InnerText -ne '00999999') { throw 'Updated cache contents incorrect' }
    [IO.File]::WriteAllText($cachePath, '<result><invalid>', [Text.UTF8Encoding]::new($false))
    $damagedHash = (Get-FileHash -LiteralPath $cachePath).Hash
    if ($excel.Run($prefix + 'CheckLocalCache', $cachePath) -eq 'PASS') { throw 'Damaged cache accepted' }
    if ((Get-FileHash -LiteralPath $cachePath).Hash -ne $damagedHash) { throw 'Damaged cache unexpectedly replaced' }
    Write-Host 'DART local cache, stale warning and atomic replacement tests passed.'
}
finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    foreach ($p in @($cachePath, $modulePath)) { if ($p -and (Test-Path -LiteralPath $p)) { Remove-Item -LiteralPath $p -Force } }
}
