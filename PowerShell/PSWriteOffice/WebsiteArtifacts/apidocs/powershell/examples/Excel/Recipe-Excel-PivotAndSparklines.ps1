$path = '.\Sales-Analysis.xlsx'
$sales = @(
    [pscustomobject]@{ Region = 'North'; Month = 'Jan'; Revenue = 120; Feb = 128; Mar = 141 }
    [pscustomobject]@{ Region = 'North'; Month = 'Feb'; Revenue = 128; Feb = 132; Mar = 145 }
    [pscustomobject]@{ Region = 'South'; Month = 'Jan'; Revenue = 95; Feb = 104; Mar = 111 }
    [pscustomobject]@{ Region = 'South'; Month = 'Feb'; Revenue = 104; Feb = 109; Mar = 118 }
)

ExcelNew -Path $path {
    ExcelSheet 'Sales' {
        ExcelTable -Data $sales -TableName 'Sales' -AutoFit
        ExcelPivotTable -SourceRange 'A1:C5' -DestinationCell 'G2' -Name 'RevenueByRegion' -RowField Region -ColumnField Month -DataField Revenue
        ExcelSparkline -DataRange 'D2:F5' -LocationRange 'G8:G11' -Type Line -ShowMarkers -ShowHighLow
    }
}
