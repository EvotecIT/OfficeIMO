$path = '.\Budget-Dashboard.xlsx'
$budget = @(
    [pscustomobject]@{ Department = 'Engineering'; Budget = 240000; Actual = 218500; Forecast = 236000 }
    [pscustomobject]@{ Department = 'Operations'; Budget = 160000; Actual = 151200; Forecast = 164000 }
    [pscustomobject]@{ Department = 'Sales'; Budget = 190000; Actual = 177800; Forecast = 188500 }
    [pscustomobject]@{ Department = 'Support'; Budget = 110000; Actual = 104300; Forecast = 109000 }
)
$budget = $budget | Select-Object *, @{ Name = 'Variance'; Expression = { $_.Forecast - $_.Budget } }
$forecastVariance = ($budget | Measure-Object -Property Variance -Sum).Sum
$overspend = @($budget | Where-Object Variance -gt 0 | Select-Object -ExpandProperty Department)
$decision = if ($overspend.Count) { 'Decision: review overspend in ' + ($overspend -join ', ') + '.' } else { 'Decision: confirm the forecast remains within budget.' }

ExcelNew -Path $path -EvaluateFormulas {
    ExcelSheet 'Dashboard' {
        ExcelGridlines -Hide
        ExcelColumn -ColumnName A -Width 38
        ExcelColumn -ColumnName B -Width 24
        ExcelCell -Address A1 -Value 'Department budget review'
        ExcelRow -Row 1 -Bold $true -Height 32
        ExcelCell -Address A2 -Value 'Illustrative annual forecast | USD'
        ExcelCell -Address A3 -Value 'Total budget'
        ExcelCell -Address B3 -Formula 'SUM(Detail!B2:B5)' -NumberFormat '$#,##0'
        ExcelCell -Address A4 -Value 'Actual spend'
        ExcelCell -Address B4 -Formula 'SUM(Detail!C2:C5)' -NumberFormat '$#,##0'
        ExcelCell -Address A5 -Value 'Forecast variance'
        ExcelCell -Address B5 -Formula 'SUM(Detail!E2:E5)' -NumberFormat '[Red]$#,##0;[Green]-$#,##0'
        ExcelRow -Row 5 -Bold $true -BackgroundColor $(if ($forecastVariance -gt 0) { '#FEE2E2' } else { '#DCFCE7' }) -FirstColumn 1 -LastColumn 2
        ExcelCell -Address A7 -Value 'Negative variance means the forecast is below budget.'
        ExcelRow -Row 7 -WrapText $true -Height 34
        ExcelCell -Address A9 -Value $decision
        ExcelRow -Row 9 -WrapText $true -Height 34
        ExcelOrientation -Orientation Landscape
        ExcelPageSetup -FitToWidth 1 -FitToHeight 0
        ExcelPrintArea -Range 'A1:B9'
    }

    ExcelSheet 'Detail' {
        ExcelTable -Data $budget -TableName 'DepartmentBudget' -StartRow 1 -StartColumn 1 -TableStyle 'TableStyleMedium4' -AutoFit
        ExcelFreeze -TopRows 1
        ExcelConditionalDataBar -Range 'C2:C5' -Color '#5B9BD5'
        ExcelConditionalIconSet -Range 'D2:D5' -IconSet ThreeTrafficLights1
        foreach ($header in 'Budget', 'Actual', 'Forecast', 'Variance') {
            ExcelColumnStyleByHeader -Header $header -NumberFormat '$#,##0' -AutoFit
        }
        for ($row = 2; $row -le $budget.Count + 1; $row++) {
            ExcelCell -Address "E$row" -Formula "D$row-B$row" -NumberFormat '[Red]$#,##0;[Green]-$#,##0'
        }
        ExcelChart -Range 'A1:D5' -Row 7 -Column 1 -Type ColumnClustered -Title 'Budget, actual, and forecast' -WidthPixels 780 -HeightPixels 340
    }

    ExcelTableOfContents -SheetName 'Index' -AddBackLinks -BackLinkText 'Back to Index'
}
