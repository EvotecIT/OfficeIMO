param(
    [string] $AssemblyDirectory = (Join-Path $PSScriptRoot '..\OfficeIMO.PowerPoint.Tests\bin\Debug\net8.0'),
    [string] $OutputPath = (Join-Path $PSScriptRoot '..\OfficeIMO.PowerPoint\PowerPointOpenXmlValueTypes.Generated.cs')
)

$typeNames = [ordered]@{
    'DocumentFormat.OpenXml.Drawing.Charts.BuiltInUnitValues' = 'PowerPointChartDisplayUnit'
    'DocumentFormat.OpenXml.Drawing.Charts.CrossBetweenValues' = 'PowerPointChartAxisCrossBetween'
    'DocumentFormat.OpenXml.Drawing.Charts.CrossesValues' = 'PowerPointChartAxisCrossing'
    'DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues' = 'PowerPointChartDataLabelPosition'
    'DocumentFormat.OpenXml.Drawing.Charts.GroupingValues' = 'PowerPointChartGrouping'
    'DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues' = 'PowerPointChartLegendPosition'
    'DocumentFormat.OpenXml.Drawing.Charts.MarkerStyleValues' = 'PowerPointChartMarkerStyle'
    'DocumentFormat.OpenXml.Drawing.Charts.TickLabelPositionValues' = 'PowerPointChartTickLabelPosition'
    'DocumentFormat.OpenXml.Drawing.Charts.TrendlineValues' = 'PowerPointChartTrendlineType'
    'DocumentFormat.OpenXml.Drawing.LineEndLengthValues' = 'PowerPointLineEndLength'
    'DocumentFormat.OpenXml.Drawing.LineEndValues' = 'PowerPointLineEndType'
    'DocumentFormat.OpenXml.Drawing.LineEndWidthValues' = 'PowerPointLineEndWidth'
    'DocumentFormat.OpenXml.Drawing.PresetLineDashValues' = 'PowerPointLineDashStyle'
    'DocumentFormat.OpenXml.Drawing.RectangleAlignmentValues' = 'PowerPointRectangleAlignment'
    'DocumentFormat.OpenXml.Drawing.ShapeTypeValues' = 'PowerPointShapeType'
    'DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues' = 'PowerPointTextAlignment'
    'DocumentFormat.OpenXml.Drawing.TextAnchoringTypeValues' = 'PowerPointTextVerticalAlignment'
    'DocumentFormat.OpenXml.Drawing.TextAutoNumberSchemeValues' = 'PowerPointNumberingScheme'
    'DocumentFormat.OpenXml.Drawing.TextUnderlineValues' = 'PowerPointUnderlineStyle'
    'DocumentFormat.OpenXml.Drawing.TextVerticalValues' = 'PowerPointTextDirection'
    'DocumentFormat.OpenXml.Presentation.CommandValues' = 'PowerPointAnimationCommand'
    'DocumentFormat.OpenXml.Presentation.DirectionValues' = 'PowerPointPlaceholderDirection'
    'DocumentFormat.OpenXml.Presentation.OleObjectFollowColorSchemeValues' = 'PowerPointOleFollowColorScheme'
    'DocumentFormat.OpenXml.Presentation.PlaceholderSizeValues' = 'PowerPointPlaceholderSize'
    'DocumentFormat.OpenXml.Presentation.PlaceholderValues' = 'PowerPointPlaceholderType'
    'DocumentFormat.OpenXml.Presentation.SlideSizeValues' = 'PowerPointSlideSizeType'
}

& (Join-Path $PSScriptRoot 'Generate-OpenXmlValueTypes.ps1') `
    -TypeNames $typeNames `
    -Namespace 'OfficeIMO.PowerPoint' `
    -ExtensionClassName 'PowerPointOpenXmlValueTypeExtensions' `
    -AssemblyDirectory $AssemblyDirectory `
    -OutputPath $OutputPath `
    -GeneratorName 'Build/Generate-PowerPointOpenXmlValueTypes.ps1'
