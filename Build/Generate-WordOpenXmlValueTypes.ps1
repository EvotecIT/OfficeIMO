param(
    [string] $AssemblyDirectory = (Join-Path $PSScriptRoot '..\OfficeIMO.Word.Tests\bin\Debug\net8.0'),
    [string] $OutputPath = (Join-Path $PSScriptRoot '..\OfficeIMO.Word\WordOpenXmlValueTypes.Generated.cs')
)

$typeNames = [ordered]@{
    'DocumentFormat.OpenXml.Bibliography.DataSourceValues' = 'WordBibliographySourceType'
    'DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues' = 'WordImageBlackWhiteMode'
    'DocumentFormat.OpenXml.Drawing.BlipCompressionValues' = 'WordImageCompressionQuality'
    'DocumentFormat.OpenXml.Drawing.Charts.BarDirectionValues' = 'WordChartBarDirection'
    'DocumentFormat.OpenXml.Drawing.Charts.BarGroupingValues' = 'WordChartBarGrouping'
    'DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues' = 'WordChartLegendPosition'
    'DocumentFormat.OpenXml.Drawing.ShapeTypeValues' = 'WordImageShapeType'
    'DocumentFormat.OpenXml.Office2010.Word.Drawing.SizeRelativeHorizontallyValues' = 'WordTextBoxHorizontalSizeReference'
    'DocumentFormat.OpenXml.Wordprocessing.BorderValues' = 'WordBorderStyle'
    'DocumentFormat.OpenXml.Wordprocessing.CharacterSpacingValues' = 'WordCharacterSpacing'
    'DocumentFormat.OpenXml.Wordprocessing.EndnotePositionValues' = 'WordEndnotePosition'
    'DocumentFormat.OpenXml.Wordprocessing.FootnotePositionValues' = 'WordFootnotePosition'
    'DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues' = 'WordHighlightColor'
    'DocumentFormat.OpenXml.Wordprocessing.HorizontalAlignmentValues' = 'WordTableHorizontalAlignment'
    'DocumentFormat.OpenXml.Wordprocessing.HorizontalAnchorValues' = 'WordTableHorizontalAnchor'
    'DocumentFormat.OpenXml.Wordprocessing.LevelJustificationValues' = 'WordListLevelAlignment'
    'DocumentFormat.OpenXml.Wordprocessing.LevelSuffixValues' = 'WordListLevelSuffix'
    'DocumentFormat.OpenXml.Wordprocessing.LineSpacingRuleValues' = 'WordLineSpacingRule'
    'DocumentFormat.OpenXml.Wordprocessing.MergedCellValues' = 'WordCellMerge'
    'DocumentFormat.OpenXml.Wordprocessing.NumberFormatValues' = 'WordNumberFormat'
    'DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues' = 'WordPageOrientation'
    'DocumentFormat.OpenXml.Wordprocessing.PresetZoomValues' = 'WordZoomPreset'
    'DocumentFormat.OpenXml.Wordprocessing.RestartNumberValues' = 'WordNoteNumberRestart'
    'DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues' = 'WordShadingPattern'
    'DocumentFormat.OpenXml.Wordprocessing.TabStopLeaderCharValues' = 'WordTabLeader'
    'DocumentFormat.OpenXml.Wordprocessing.TabStopValues' = 'WordTabAlignment'
    'DocumentFormat.OpenXml.Wordprocessing.TableLayoutValues' = 'WordTableLayoutMode'
    'DocumentFormat.OpenXml.Wordprocessing.TableOverlapValues' = 'WordTableOverlap'
    'DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues' = 'WordTableAlignment'
    'DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues' = 'WordTableWidthUnit'
    'DocumentFormat.OpenXml.Wordprocessing.ThemeColorValues' = 'WordThemeColor'
    'DocumentFormat.OpenXml.Wordprocessing.VerticalAlignmentValues' = 'WordTableVerticalPositionAlignment'
    'DocumentFormat.OpenXml.Wordprocessing.VerticalAnchorValues' = 'WordTableVerticalAnchor'
    'DocumentFormat.OpenXml.Wordprocessing.VerticalPositionValues' = 'WordVerticalTextPosition'
    'DocumentFormat.OpenXml.Wordprocessing.VerticalTextAlignmentValues' = 'WordVerticalCharacterAlignment'
}

& (Join-Path $PSScriptRoot 'Generate-OpenXmlValueTypes.ps1') `
    -TypeNames $typeNames `
    -Namespace 'OfficeIMO.Word' `
    -ExtensionClassName 'WordOpenXmlValueTypeExtensions' `
    -AssemblyDirectory $AssemblyDirectory `
    -OutputPath $OutputPath `
    -GeneratorName 'Build/Generate-WordOpenXmlValueTypes.ps1'
