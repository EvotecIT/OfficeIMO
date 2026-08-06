param(
    [string] $AssemblyDirectory = (Join-Path $PSScriptRoot '..\OfficeIMO.PowerPoint.Tests\bin\Debug\net8.0'),
    [string] $OutputPath = (Join-Path $PSScriptRoot '..\OfficeIMO.PowerPoint\PowerPointOpenXmlValueTypes.Generated.cs')
)

$typeNames = [ordered]@{
    'DocumentFormat.OpenXml.Drawing.Charts.GroupingValues' = 'PowerPointChartGrouping'
    'DocumentFormat.OpenXml.Drawing.LineEndLengthValues' = 'PowerPointLineEndLength'
    'DocumentFormat.OpenXml.Drawing.LineEndWidthValues' = 'PowerPointLineEndWidth'
    'DocumentFormat.OpenXml.Drawing.PresetLineDashValues' = 'PowerPointLineDashStyle'
    'DocumentFormat.OpenXml.Drawing.RectangleAlignmentValues' = 'PowerPointRectangleAlignment'
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

$sharedTypeNames = [ordered]@{
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
    -SharedTypeNames $sharedTypeNames `
    -Namespace 'OfficeIMO.PowerPoint' `
    -ExtensionClassName 'PowerPointOpenXmlValueTypeExtensions' `
    -AssemblyDirectory $AssemblyDirectory `
    -OutputPath $OutputPath `
    -GeneratorName 'Build/Generate-PowerPointOpenXmlValueTypes.ps1'
