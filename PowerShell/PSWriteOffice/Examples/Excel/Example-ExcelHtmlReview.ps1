Import-Module PSWriteOffice -ErrorAction Stop

$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$workbookPath = Join-Path $documents 'Excel-HtmlReview.xlsx'
$semanticHtmlPath = Join-Path $documents 'Excel-HtmlReview.semantic.html'
$visualHtmlPath = Join-Path $documents 'Excel-HtmlReview.visual.html'

$data = @(
    [PSCustomObject]@{ Service = 'Identity'; Status = 'Healthy'; Incidents = 0; Owner = 'Platform' }
    [PSCustomObject]@{ Service = 'Messaging'; Status = 'Watch'; Incidents = 2; Owner = 'Collaboration' }
    [PSCustomObject]@{ Service = 'Reporting'; Status = 'Healthy'; Incidents = 1; Owner = 'Analytics' }
)

New-OfficeExcel -Path $workbookPath {
    Add-OfficeExcelSheet -Name 'Service Review' -Content {
        Set-OfficeExcelRow -Row 1 -Values 'Service', 'Status', 'Incidents', 'Owner'
        Add-OfficeExcelTable -Data $data -TableName 'ServiceStatus' -TableStyle 'TableStyleMedium4'
        Set-OfficeExcelFormula -Cell 'E2' -Formula '=SUM(ServiceStatus[Incidents])'
        Set-OfficeExcelCell -Cell 'E1' -Value 'Total incidents'
        Set-OfficeExcelColumn -Column 1, 2, 3, 4, 5 -AutoFit
    }
}

ConvertTo-OfficeExcelHtml -Path $workbookPath -OutputPath $semanticHtmlPath -Title 'Service Workbook Review'
ConvertTo-OfficeExcelHtml -Path $workbookPath -Profile VisualReview -OutputPath $visualHtmlPath -Title 'Service Workbook Visual Review'

Write-Host "Workbook saved to $workbookPath"
Write-Host "Semantic HTML saved to $semanticHtmlPath"
Write-Host "Visual HTML saved to $visualHtmlPath"
