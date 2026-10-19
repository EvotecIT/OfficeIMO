Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Excel-ChartFormatting.xlsx'
$rows = @(
    [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
    [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
    [PSCustomObject]@{ Region = 'APAC'; Revenue = 150 }
)

New-OfficeExcel -Path $path {
    Add-OfficeExcelSheet -Name 'Data' -Content {
        Add-OfficeExcelTable -Data $rows -TableName 'Sales' -AutoFit

        $chart = Add-OfficeExcelChart -TableName 'Sales' -Row 6 -Column 1 -Type Pie -Title 'Revenue Mix' -PassThru
        $chart |
            Set-OfficeExcelChartLegend -Position Right -PassThru |
            Set-OfficeExcelChartDataLabels -ShowValue $true -ShowPercent $true -Position OutsideEnd -NumberFormat '0.0%' -SourceLinked:$false -PassThru |
            Set-OfficeExcelChartStyle -StyleId 251 -ColorStyleId 10
    }
}

Write-Host "Workbook saved to $path"
