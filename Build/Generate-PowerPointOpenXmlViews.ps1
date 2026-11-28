param(
    # Folder with the ISO/IEC 29500-4 transitional schemas. They are not part of the repository.
    [Parameter(Mandatory)] [string] $SchemaDirectory,
    [string] $OutputPath = (Join-Path $PSScriptRoot '..\OfficeIMO.PowerPoint\OpenXml\PowerPointOpenXmlViews.Generated.cs')
)

# The PresentationML and DrawingML types the presentation model reads and writes.
$types = @(
    'p:CT_Presentation', 'p:CT_SlideMasterIdList', 'p:CT_SlideMasterIdListEntry', 'p:CT_SlideIdList', 'p:CT_SlideIdListEntry',
    'p:CT_SlideSize', 'a:CT_PositiveSize2D',
    'p:CT_Slide', 'p:CT_SlideLayout', 'p:CT_SlideMaster', 'p:CT_SlideLayoutIdList', 'p:CT_SlideLayoutIdListEntry',
    'p:CT_CommonSlideData', 'p:CT_GroupShape', 'p:CT_GroupShapeNonVisual', 'p:CT_Shape', 'p:CT_ShapeNonVisual',
    'p:CT_ApplicationNonVisualDrawingProps', 'p:CT_Placeholder', 'p:CT_SlideMasterTextStyles',
    'a:CT_NonVisualDrawingProps', 'a:CT_NonVisualDrawingShapeProps', 'a:CT_NonVisualGroupDrawingShapeProps',
    'a:CT_GroupShapeProperties', 'a:CT_GroupTransform2D', 'a:CT_ShapeProperties', 'a:CT_Transform2D', 'a:CT_Point2D',
    'a:CT_PresetGeometry2D', 'a:CT_GeomGuideList', 'a:CT_NoFillProperties', 'a:CT_SolidColorFillProperties', 'a:CT_SRgbColor',
    'a:CT_TextBody', 'a:CT_TextBodyProperties', 'a:CT_TextListStyle', 'a:CT_TextParagraph', 'a:CT_TextParagraphProperties',
    'a:CT_RegularTextRun', 'a:CT_TextCharacterProperties', 'a:CT_TextNormalAutofit', 'a:CT_ColorMapping', 'a:CT_ColorMappingOverride',
    'a:CT_EmptyElement',
    'p:CT_GraphicalObjectFrame', 'p:CT_GraphicalObjectFrameNonVisual', 'a:CT_GraphicalObject', 'a:CT_GraphicalObjectData',
    'a:CT_Table', 'a:CT_TableProperties', 'a:CT_TableGrid', 'a:CT_TableCol', 'a:CT_TableRow', 'a:CT_TableCell',
    'a:CT_TableCellProperties', 'a:CT_LineProperties',
    'p:CT_NotesSlide', 'p:CT_NotesMaster', 'p:CT_NotesMasterIdList', 'p:CT_NotesMasterIdListEntry',
    'a:CT_OfficeStyleSheet', 'a:CT_BaseStyles', 'a:CT_ColorScheme', 'a:CT_FontScheme', 'a:CT_FontCollection'
)

& (Join-Path $PSScriptRoot 'Generate-OpenXmlViews.ps1') `
    -SchemaDirectory $SchemaDirectory `
    -Types $types `
    -Namespace 'OfficeIMO.PowerPoint.OpenXml' `
    -OutputPath $OutputPath `
    -GeneratorName 'Build/Generate-PowerPointOpenXmlViews.ps1' `
    -NodeType 'OfficeXmlNode' `
    -NodeNamespace 'OfficeIMO.OpenXml.Internal'
