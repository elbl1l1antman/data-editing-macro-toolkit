param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)

$ErrorActionPreference = 'Stop'
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$output = Join-Path $ProjectRoot 'dev\sample_extraction_test.xlsx'
$excel = $null; $workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $population = $workbook.Worksheets.Item(1); $population.Name = '모집단'
    $headers = @('ID','산업','지역','규모','정렬값')
    for ($c=0; $c -lt $headers.Count; $c++) { $population.Cells.Item(1,$c+1).Value2 = $headers[$c] }
    for ($r=2; $r -le 25; $r++) {
        $population.Cells.Item($r,1).Value2 = if ($r -eq 5) { '' } elseif ($r -eq 6) { 'ID4' } else { "ID$($r-1)" }
        $population.Cells.Item($r,2).Value2 = if ($r % 2) { 'SW' } else { 'Non-SW' }
        $population.Cells.Item($r,3).Value2 = if ($r % 3) { '수도권' } else { '' }
        $population.Cells.Item($r,4).Value2 = if ($r % 4) { '소형' } else { '대형' }
        $population.Cells.Item($r,5).Value2 = [double](100 - $r)
    }
    $allocation = $workbook.Worksheets.Add(); $allocation.Name = '표본설계_할당결과'
    @('산업','지역','최종할당') | ForEach-Object -Begin {$c=1} -Process { $allocation.Cells.Item(1,$c++).Value2 = $_ }
    $rows = @(@('SW','수도권',2),@('SW','수도권',1),@('Non-SW','수도권',3))
    for ($r=0; $r -lt $rows.Count; $r++) {
        for ($c=0; $c -lt 3; $c++) {
            $value = $rows[$r][$c]
            $allocation.Cells.Item($r+2,$c+1).Value2 = [string]$value
        }
    }
    $workbook.SaveAs($output, 51)
    Write-Host "Created: $output"
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) | Out-Null }
    if ($null -ne $excel) { $excel.Quit() | Out-Null; [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
