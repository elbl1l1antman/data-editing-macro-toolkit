param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\DataCleanupTools.bas') -Raw -Encoding UTF8
$match = [regex]::Match($source, '(?ms)^Public Function Phone_CombineParts\b.*?^End Function\s*$')
if (-not $match.Success) { throw 'Phone_CombineParts is missing.' }
$test = @'
Option Explicit
Public Function TestPhoneCombine() As String
    If Phone_CombineParts("010", "1234", "5678") <> "010-1234-5678" Then Err.Raise 5
    If Phone_CombineParts("02", "123", "4567") <> "02-123-4567" Then Err.Raise 5
    If Phone_CombineParts("0507", "1234", "5678") <> "0507-1234-5678" Then Err.Raise 5
    If Phone_CombineParts("2", "123", "4567") <> "02-123-4567" Then Err.Raise 5
    If Phone_CombineParts("31", "1234", "5678") <> "031-1234-5678" Then Err.Raise 5
    If Phone_CombineParts("10", "1234", "5678") <> "010-1234-5678" Then Err.Raise 5
    If Phone_CombineParts("507", "1234", "5678") <> "0507-1234-5678" Then Err.Raise 5
    If Phone_CombineParts("99", "1234", "5678") <> "" Then Err.Raise 5
    If Phone_CombineParts("31", "12", "5678") <> "" Then Err.Raise 5
    If Phone_CombineParts("010", "", "5678") <> "" Then Err.Raise 5
    If Phone_CombineParts("010", "abc", "5678") <> "" Then Err.Raise 5
    TestPhoneCombine = "PASS"
End Function
'@
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $module = $book.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($test + "`r`n" + $match.Value)
    if ([string]$excel.Run("'$($book.Name)'!TestPhoneCombine") -ne 'PASS') { throw 'Phone combine test failed.' }
    Write-Host 'Phone combine tests passed.'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
