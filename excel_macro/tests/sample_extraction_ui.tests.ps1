param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)

$ErrorActionPreference = 'Stop'
$publishPath = Join-Path $ProjectRoot 'scripts\publish_release.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($publishPath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
$fn = $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Add-SampleExtractionWizardUserForm' }, $true) | Select-Object -First 1
if ($null -eq $fn) { throw 'Add-SampleExtractionWizardUserForm not found' }
$source = $fn.Extent.Text
foreach ($token in @('Forms.MultiPage.1','1. 모집단과 층 선택','2. 추출 방법 설정','3. 실행 전 검토','모집단 범위 선택','ID가 없으면 결과용 ID 자동 생성','대체표본 배수','현재 행 순서 유지','표본오차는 단순임의추출 근사치입니다.','Public Property Get Action()','Public Sub ConfigureReview','Private Sub cmdBack_Click()','Private Sub cmdNext_Click()')) {
    if ($source.IndexOf($token, [StringComparison]::Ordinal) -lt 0) { throw "Missing UI token: $token" }
}

Invoke-Expression $source
$excel = $null; $workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Add()
    Add-SampleExtractionWizardUserForm -Workbook $workbook
    $component = $workbook.VBProject.VBComponents.Item('frmSampleExtractionWizard')
    $designer = $component.Designer
    $pages = $designer.Controls.Item('mpSteps')
    if ($pages.Pages.Count -ne 3) { throw 'Wizard must have three pages' }
    $expected = @('1. 모집단과 층 선택','2. 추출 방법 설정','3. 실행 전 검토')
    for ($i=0; $i -lt 3; $i++) {
        if ([string]$pages.Pages.Item($i).Caption -ne $expected[$i]) { throw "Unexpected page caption $i" }
        foreach ($control in $pages.Pages.Item($i).Controls) {
            if (($control.Left + $control.Width) -gt ($pages.Width - 10)) { throw "Control clipped horizontally: $($control.Name)" }
            if (($control.Top + $control.Height) -gt ($pages.Height - 18)) { throw "Control clipped vertically: $($control.Name)" }
        }
    }
    $formWidth = [double]$component.Properties.Item('Width').Value
    foreach ($name in @('cmdBack','cmdNext','cmdRun','cmdCancel')) {
        $control = $designer.Controls.Item($name)
        if (($control.Left + $control.Width) -gt ($formWidth - 8)) { throw "Button clipped: $name" }
    }
    Write-Host 'Sample extraction UI tests passed.'
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) | Out-Null }
    if ($null -ne $excel) { $excel.Quit() | Out-Null; [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
