param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)
. (Join-Path $PSScriptRoot "test_helpers.ps1")

$ErrorActionPreference = "Stop"
$moduleText = Get-ToolkitVbaSource -ProjectRoot $ProjectRoot
$version = [regex]::Match($moduleText, 'TOOLKIT_VERSION As String = "([^"]+)"').Groups[1].Value
if (-not $version) { throw "Could not read TOOLKIT_VERSION." }

$artifacts = @(
    (Join-Path $ProjectRoot "release\v$version\DataEditingMacroToolkit_v$version.xlsm"),
    (Join-Path $ProjectRoot "release\v$version\DataEditingMacroToolkit_v$version.xlam"),
    (Join-Path $ProjectRoot "release\latest\DataEditingMacroToolkit_latest.xlsm"),
    (Join-Path $ProjectRoot "release\latest\DataEditingMacroToolkit_latest.xlam")
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $driver = $excel.Workbooks.Add()
    $module = $driver.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Function CheckPackagedCompanyRules(ByVal workbookName As String) As String
    On Error GoTo Failed
    Dim prefix As String, options As Object, unitText As Variant, actual As String
    prefix = "'" & Replace(workbookName, "'", "''") & "'!"
    Set options = Application.Run(prefix & "CompanyName_DefaultOptions")
    options("removeLocationUnits") = True
    For Each unitText In Array("판교지점", "판교본사", "판교공장", "대구지점", "청주사업소", "수원영업소", "세종대로지점", "세종대로영업소")
        actual = Application.Run(prefix & "CompanyName_CleanOne", "한빛" & CStr(unitText), options)
        If actual <> "한빛" Then Err.Raise 513, , CStr(unitText) & ": " & actual
    Next unitText
    For Each unitText In Array("주)", "( 주 )", "( 주)", "(주 )", "（　주　）")
        actual = Application.Run(prefix & "CompanyName_CleanOne", CStr(unitText) & "한빛", options)
        If actual <> "한빛" Then Err.Raise 513, , "Corporate prefix: " & actual
    Next unitText
    Set options = Application.Run(prefix & "CompanyName_DefaultOptions")
    options("removeBranchMarkers") = True
    Dim examples As Variant, expected As Variant, i As Long
    examples = Array("신한은행본점", "국민은행본사", "삼성증권제1지점")
    expected = Array("신한은행", "국민은행", "삼성증권")
    For i = LBound(examples) To UBound(examples)
        actual = Application.Run(prefix & "CompanyName_CleanOne", CStr(examples(i)), options)
        If actual <> CStr(expected(i)) Then Err.Raise 513, , CStr(examples(i)) & ": " & actual
    Next i
    CheckPackagedCompanyRules = "PASS"
    Exit Function
Failed:
    CheckPackagedCompanyRules = CStr(Err.Number) & ": " & Err.Description
End Function
'@)
    foreach ($artifact in $artifacts) {
        if (-not (Test-Path -LiteralPath $artifact)) { throw "Missing release artifact: $artifact" }
        $workbook = $excel.Workbooks.Open($artifact, 0, $true)
        try {
            $macro = "'{0}'!Toolkit_CompileSmokeTest" -f $workbook.Name.Replace("'", "''")
            $result = [string]$excel.Run($macro)
            if (-not $result.StartsWith($version + "|forms=14")) {
                throw "Unexpected runtime version for $artifact`: $result"
            }
            if ($workbook.VBProject.Protection -ne 1) {
                throw "VBA project is not protected: $artifact"
            }
            $companyResult = $excel.Run("'$($driver.Name)'!CheckPackagedCompanyRules", $workbook.Name)
            if ($companyResult -ne 'PASS') { throw "Packaged company rules failed for $artifact`: $companyResult" }
        }
        finally {
            $workbook.Close($false)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }
    }
}
finally {
    if ($driver) { $driver.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

Write-Output "Release artifact runtime tests passed for v$version."
