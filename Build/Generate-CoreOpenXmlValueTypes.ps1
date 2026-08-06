param(
    [string] $AssemblyDirectory = (Join-Path $PSScriptRoot '..\OfficeIMO.Excel.Tests\bin\Debug\net8.0'),
    [string] $OutputPath = (Join-Path $PSScriptRoot '..\OfficeIMO.Core\OfficeOpenXmlValueTypes.Generated.cs')
)

$typeNames = [ordered]@{
    'DocumentFormat.OpenXml.Drawing.Charts.BuiltInUnitValues' = 'OfficeChartDisplayUnit'
    'DocumentFormat.OpenXml.Drawing.Charts.CrossBetweenValues' = 'OfficeChartAxisCrossBetween'
    'DocumentFormat.OpenXml.Drawing.Charts.CrossesValues' = 'OfficeChartAxisCrossingPosition'
    'DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues' = 'OfficeChartDataLabelPosition'
    'DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues' = 'OfficeChartLegendPosition'
    'DocumentFormat.OpenXml.Drawing.Charts.MarkerStyleValues' = 'OfficeChartMarkerShape'
    'DocumentFormat.OpenXml.Drawing.Charts.TickLabelPositionValues' = 'OfficeChartAxisTickLabelPosition'
    'DocumentFormat.OpenXml.Drawing.Charts.TrendlineValues' = 'OfficeChartTrendlineType'
    'DocumentFormat.OpenXml.Drawing.LineEndValues' = 'OfficeLineMarkerKind'
    'DocumentFormat.OpenXml.Drawing.ShapeTypeValues' = 'OfficePresetShapeType'
}

& (Join-Path $PSScriptRoot 'Generate-OpenXmlValueTypes.ps1') `
    -TypeNames $typeNames `
    -Namespace 'OfficeIMO.Drawing' `
    -ExtensionClassName 'OfficeOpenXmlValueTypeExtensions' `
    -AssemblyDirectory $AssemblyDirectory `
    -OutputPath $OutputPath `
    -GeneratorName 'Build/Generate-CoreOpenXmlValueTypes.ps1' `
    -SkipMappings
