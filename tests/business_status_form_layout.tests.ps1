param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path)
$ErrorActionPreference = 'Stop'
$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $ProjectRoot 'scripts\publish_release.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Publisher parse failed' }
$fn = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Add-BizStatusFieldsOptionsUserForm' }, $true)
if (-not $fn) { throw 'Business status form builder missing' }
Invoke-Expression $fn.Extent.Text
$excel = New-Object -ComObject Excel.Application
$driver = $null; $target = $null
try {
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $driver = $excel.Workbooks.Add()
    Add-BizStatusFieldsOptionsUserForm -Workbook $driver
    $module = $driver.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Function CheckFormBounds() As String
    Dim f As Object, c As Object
    Set f = VBA.UserForms.Add("frmBizStatusFieldsOptions")
    For Each c In f.Controls
        If TypeName(c) = "CommandButton" Then
            If c.Top + c.Height > f.InsideHeight Or c.Left + c.Width > f.InsideWidth Then
                CheckFormBounds = c.Name
                Unload f
                Exit Function
            End If
        End If
    Next c
    Unload f
    CheckFormBounds = "PASS"
End Function
'@)
    $result = $excel.Run("'$($driver.Name)'!CheckFormBounds")
    if ($result -ne 'PASS') { throw "Business status form button clipped: $result" }
    $source = Get-Content -Raw -Encoding UTF8 (Join-Path $ProjectRoot 'src\MacroToolkit.bas')
    $reset = [regex]::Match($source, '(?ms)^Public Sub Toolkit_ResetAllSheetsToTop\(\).*?^End Sub').Value
    if (-not $reset) { throw 'Immediate sheet reset procedure missing' }
    $reset = [regex]::Replace($reset, '(?m)^\s*MsgBox [^\r\n]+', "`r`n    ' Suppress dialogs during test")
    $module.CodeModule.AddFromString($reset)
    $target = $excel.Workbooks.Add()
    $first = $target.Worksheets.Item(1)
    $hidden = $target.Worksheets.Add()
    $hidden.Visible = 0
    $first.Activate()
    $first.Range('D8').Select() | Out-Null
    $excel.ActiveWindow.ScrollRow = 5
    $excel.ScreenUpdating = $false
    $excel.Run("'$($driver.Name)'!Toolkit_ResetAllSheetsToTop")
    if ($excel.Selection.Address() -ne '$A$1' -or $excel.ActiveWindow.ScrollRow -ne 1 -or $excel.ScreenUpdating) { throw 'Immediate A1 reset failed or did not restore screen updating' }
    if ($hidden.Visible -ne 0) { throw 'Hidden sheet visibility changed' }
    Write-Host 'Business status form buttons fit within the visible form.'
    Write-Host 'Immediate A1 reset skips hidden sheets and restores screen updating.'
}
finally {
    if ($target) { $target.Close($false) }
    if ($driver) { $driver.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
