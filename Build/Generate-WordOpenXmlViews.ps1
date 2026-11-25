param(
    # Folder with the ISO/IEC 29500-4 transitional schemas. They are not part of the repository.
    [Parameter(Mandatory)] [string] $SchemaDirectory,
    [string] $OutputPath = (Join-Path $PSScriptRoot '..\OfficeIMO.Word\OpenXml\WordOpenXmlViews.Generated.cs')
)

# The WordprocessingML types the Word document model reads and writes without the Open XML SDK.
$types = @(
    'w:CT_Document', 'w:CT_Body', 'w:CT_P', 'w:CT_PPr', 'w:CT_ParaRPr', 'w:CT_R', 'w:CT_RPr', 'w:CT_Text', 'w:CT_Br',
    'w:CT_OnOff', 'w:CT_HpsMeasure', 'w:CT_Color', 'w:CT_Underline', 'w:CT_Highlight', 'w:CT_Fonts', 'w:CT_Jc',
    'w:CT_String', 'w:CT_VerticalAlignRun', 'w:CT_SectPr', 'w:CT_Spacing', 'w:CT_Ind', 'w:CT_SignedTwipsMeasure',
    'w:CT_TextEffect', 'w:CT_Shd', 'w:CT_Em', 'w:CT_Language'
)

& (Join-Path $PSScriptRoot 'Generate-OpenXmlViews.ps1') `
    -SchemaDirectory $SchemaDirectory `
    -Types $types `
    -Namespace 'OfficeIMO.Word.OpenXml' `
    -OutputPath $OutputPath `
    -GeneratorName 'Build/Generate-WordOpenXmlViews.ps1' `
    -OrderedChoiceGroups 'w:EG_RPrBase'
