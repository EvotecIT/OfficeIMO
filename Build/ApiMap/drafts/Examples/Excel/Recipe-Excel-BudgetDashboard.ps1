$path = '.\Budget-Dashboard.xlsx'
$budget = @(
    [pscustomobject]@{ Department = 'Engineering'; Budget = 240000; Actual = 218500; Forecast = 236000 }
    [pscustomobject]@{ Department = 'Operations'; Budget = 160000; Actual = 151200; Forecast = 164000 }
    [pscustomobject]@{ Department = 'Sales'; Budget = 190000; Actual = 177800; Forecast = 188500 }
    [pscustomobject]@{ Department = 'Support'; Budget = 110000; Actual = 104300; Forecast = 109000 }
)
$budget = $budget | Select-Object *, @{ Name = 'Variance'; Expression = { $_.Forecast - $_.Budget } }
$usd = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$total = @{}
foreach ($property in 'Budget', 'Actual', 'Variance') { $total[$property] = ($budget | Measure-Object $property -Sum).Sum }
$overspend = @($budget | Where-Object Variance -gt 0).Department

New-OfficeExcel -Path $path -Force {
    ExcelReportSheet 'Dashboard' {
        ExcelReportTitle 'Department budget review' 'Illustrative annual forecast | USD'
        ExcelReportKpi ([ordered]@{
            'Total budget'      = $total.Budget.ToString('C0', $usd)
            'Actual spend'      = $total.Actual.ToString('C0', $usd)
            'Forecast variance' = $total.Variance.ToString('C0', $usd)
        }) -PerRow 3
        if ($overspend) {
            ExcelReportCallout Warning 'Decision' "Review overspend in $($overspend -join ', ')."
        } else {
            ExcelReportCallout Success 'Decision' 'Confirm the forecast remains within budget.'
        }
        ExcelReportParagraph 'Negative variance means the forecast is below budget.'
        ExcelPageSetup -Orientation Landscape -FitToWidth 1 -FitToHeight 0
    }

    ExcelSheet 'Detail' {
        ExcelTable $budget -Name DepartmentBudget -ColumnFormat @{
            Budget = 'Currency'; Actual = 'Currency'; Forecast = 'Currency'; Variance = 'Currency'
        }
        ExcelConditionalFormat -Table DepartmentBudget -Header Actual -DataBar
        ExcelConditionalFormat -Table DepartmentBudget -Header Forecast -IconSet ThreeTrafficLights1
        ExcelChart "A1:D$($budget.Count + 1)" -Address A8 -Type Column -Title 'Budget, actual and forecast'
    }

    ExcelTableOfContents -Name Index -BackLink -BackLinkText 'Back to Index'
}
