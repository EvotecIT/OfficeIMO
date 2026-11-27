param(
    # Folder with the ISO/IEC 29500-4 transitional schemas. They are not part of the repository.
    [Parameter(Mandatory)] [string] $SchemaDirectory,
    [string] $OutputPath = (Join-Path $PSScriptRoot '..\OfficeIMO.SharedSource\OpenXml\OfficeChartOpenXmlViews.Generated.cs')
)

# The DrawingML chart types the shared chart reader and writer, and the engines' chart editing, read and write. The
# DrawingML shape and text types inside charts are named Dml* so they sit beside an engine's own DrawingML views.
$types = @(
    'c:CT_Boolean', 'c:CT_Double', 'c:CT_UnsignedInt', 'c:CT_RelId', 'c:CT_Extension', 'c:CT_ExtensionList',
    'c:CT_NumVal', 'c:CT_NumData', 'c:CT_NumRef', 'c:CT_NumDataSource', 'c:CT_StrVal', 'c:CT_StrData',
    'c:CT_StrRef', 'c:CT_Tx', 'c:CT_TextLanguageID', 'c:CT_Lvl', 'c:CT_MultiLvlStrData', 'c:CT_MultiLvlStrRef',
    'c:CT_AxDataSource', 'c:CT_SerTx', 'c:CT_LayoutTarget', 'c:CT_LayoutMode', 'c:CT_ManualLayout',
    'c:CT_Layout', 'c:CT_Title', 'c:CT_RotX', 'c:CT_HPercent', 'c:CT_RotY', 'c:CT_DepthPercent',
    'c:CT_Perspective', 'c:CT_Surface', 'c:CT_Thickness', 'c:CT_DTable', 'c:CT_GapAmount', 'c:CT_Overlap',
    'c:CT_BubbleScale', 'c:CT_SizeRepresents', 'c:CT_FirstSliceAng', 'c:CT_HoleSize', 'c:CT_SplitType',
    'c:CT_CustSplit', 'c:CT_SecondPieSize', 'c:CT_NumFmt', 'c:CT_LblAlgn', 'c:CT_DLblPos', 'c:CT_DLbl',
    'c:CT_DLbls', 'c:CT_MarkerStyle', 'c:CT_MarkerSize', 'c:CT_Marker', 'c:CT_DPt', 'c:CT_TrendlineType',
    'c:CT_Order', 'c:CT_Period', 'c:CT_TrendlineLbl', 'c:CT_Trendline', 'c:CT_ErrDir', 'c:CT_ErrBarType',
    'c:CT_ErrValType', 'c:CT_ErrBars', 'c:CT_UpDownBar', 'c:CT_UpDownBars', 'c:CT_LineSer', 'c:CT_ScatterSer',
    'c:CT_RadarSer', 'c:CT_BarSer', 'c:CT_AreaSer', 'c:CT_PieSer', 'c:CT_BubbleSer', 'c:CT_SurfaceSer',
    'c:CT_Grouping', 'c:CT_ChartLines', 'c:CT_LineChart', 'c:CT_StockChart', 'c:CT_ScatterStyle',
    'c:CT_ScatterChart', 'c:CT_RadarStyle', 'c:CT_RadarChart', 'c:CT_BarGrouping', 'c:CT_BarDir', 'c:CT_Shape',
    'c:CT_BarChart', 'c:CT_AreaChart', 'c:CT_PieChart', 'c:CT_DoughnutChart', 'c:CT_OfPieType',
    'c:CT_OfPieChart', 'c:CT_BubbleChart', 'c:CT_BandFmt', 'c:CT_BandFmts', 'c:CT_SurfaceChart', 'c:CT_AxPos',
    'c:CT_Crosses', 'c:CT_CrossBetween', 'c:CT_TickMark', 'c:CT_TickLblPos', 'c:CT_Skip', 'c:CT_TimeUnit',
    'c:CT_AxisUnit', 'c:CT_BuiltInUnit', 'c:CT_PictureFormat', 'c:CT_PictureStackUnit', 'c:CT_PictureOptions',
    'c:CT_DispUnitsLbl', 'c:CT_DispUnits', 'c:CT_Orientation', 'c:CT_LogBase', 'c:CT_Scaling', 'c:CT_LblOffset',
    'c:CT_CatAx', 'c:CT_DateAx', 'c:CT_SerAx', 'c:CT_ValAx', 'c:CT_PlotArea', 'c:CT_PivotFmt', 'c:CT_PivotFmts',
    'c:CT_LegendPos', 'c:CT_LegendEntry', 'c:CT_Legend', 'c:CT_DispBlanksAs', 'c:CT_Chart', 'c:CT_Style',
    'c:CT_PivotSource', 'c:CT_Protection', 'c:CT_HeaderFooter', 'c:CT_PageMargins', 'c:CT_ExternalData',
    'c:CT_PageSetup', 'c:CT_PrintSettings', 'c:CT_ChartSpace',
    'a:CT_ShapeProperties', 'a:CT_TextBody', 'a:CT_LineProperties', 'a:CT_TextParagraph', 'a:CT_TextParagraphProperties',
    'a:CT_TextCharacterProperties', 'a:CT_TextBodyProperties'
)

& (Join-Path $PSScriptRoot 'Generate-OpenXmlViews.ps1') `
    -SchemaDirectory $SchemaDirectory `
    -Types $types `
    -Namespace 'OfficeIMO.OpenXml.Internal' `
    -OutputPath $OutputPath `
    -GeneratorName 'Build/Generate-ChartOpenXmlViews.ps1' `
    -NodeType 'OfficeXmlNode' `
    -ViewPrefixes @{ a = 'Dml' }
