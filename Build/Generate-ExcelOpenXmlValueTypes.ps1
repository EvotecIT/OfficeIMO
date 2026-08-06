param(
    [string] $AssemblyDirectory = (Join-Path $PSScriptRoot '..\OfficeIMO.Excel.Tests\bin\Debug\net8.0'),
    [string] $OutputPath = (Join-Path $PSScriptRoot '..\OfficeIMO.Excel\ExcelOpenXmlValueTypes.Generated.cs')
)

$typeNames = [ordered]@{
    'DocumentFormat.OpenXml.Office2010.Excel.SparklineTypeValues' = 'ExcelSparklineType'
    'DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues' = 'ExcelBorderStyle'
    'DocumentFormat.OpenXml.Spreadsheet.CellValues' = 'ExcelCellValueType'
    'DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValues' = 'ExcelConditionalFormatType'
    'DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingOperatorValues' = 'ExcelConditionalFormattingOperator'
    'DocumentFormat.OpenXml.Spreadsheet.DataConsolidateFunctionValues' = 'ExcelPivotDataFunction'
    'DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues' = 'ExcelDataValidationErrorStyle'
    'DocumentFormat.OpenXml.Spreadsheet.DataValidationOperatorValues' = 'ExcelDataValidationOperator'
    'DocumentFormat.OpenXml.Spreadsheet.FieldSortValues' = 'ExcelPivotFieldSort'
    'DocumentFormat.OpenXml.Spreadsheet.GroupByValues' = 'ExcelPivotGroupBy'
    'DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues' = 'ExcelHorizontalAlignment'
    'DocumentFormat.OpenXml.Spreadsheet.IconSetValues' = 'ExcelIconSet'
    'DocumentFormat.OpenXml.Spreadsheet.PivotFilterValues' = 'ExcelPivotFilterType'
    'DocumentFormat.OpenXml.Spreadsheet.PivotTableAxisValues' = 'ExcelPivotTableAxis'
    'DocumentFormat.OpenXml.Spreadsheet.ShowDataAsValues' = 'ExcelPivotShowDataAs'
    'DocumentFormat.OpenXml.Spreadsheet.TimePeriodValues' = 'ExcelConditionalTimePeriod'
    'DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues' = 'ExcelTableTotalsFunction'
    'DocumentFormat.OpenXml.Spreadsheet.UnderlineValues' = 'ExcelUnderlineStyle'
    'DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentRunValues' = 'ExcelVerticalTextAlignment'
    'DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues' = 'ExcelVerticalAlignment'
}

$sharedTypeNames = [ordered]@{
    'DocumentFormat.OpenXml.Drawing.Charts.BuiltInUnitValues' = 'OfficeChartDisplayUnit'
    'DocumentFormat.OpenXml.Drawing.Charts.CrossBetweenValues' = 'OfficeChartAxisCrossBetween'
    'DocumentFormat.OpenXml.Drawing.Charts.CrossesValues' = 'OfficeChartAxisCrossingPosition'
    'DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues' = 'OfficeChartDataLabelPosition'
    'DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues' = 'OfficeChartLegendPosition'
    'DocumentFormat.OpenXml.Drawing.Charts.MarkerStyleValues' = 'OfficeChartMarkerShape'
    'DocumentFormat.OpenXml.Drawing.Charts.TickLabelPositionValues' = 'OfficeChartAxisTickLabelPosition'
    'DocumentFormat.OpenXml.Drawing.Charts.TrendlineValues' = 'OfficeChartTrendlineType'
}

$memberNames = @{
    # OpenXML SDK exposes this member as PercentOfRaw even though its serialized token is percentOfRow.
    'DocumentFormat.OpenXml.Spreadsheet.ShowDataAsValues.PercentOfRaw' = 'PercentOfRow'
}

& (Join-Path $PSScriptRoot 'Generate-OpenXmlValueTypes.ps1') `
    -TypeNames $typeNames `
    -SharedTypeNames $sharedTypeNames `
    -Namespace 'OfficeIMO.Excel' `
    -ExtensionClassName 'ExcelOpenXmlValueTypeExtensions' `
    -AssemblyDirectory $AssemblyDirectory `
    -OutputPath $OutputPath `
    -GeneratorName 'Build/Generate-ExcelOpenXmlValueTypes.ps1' `
    -MemberNames $memberNames
