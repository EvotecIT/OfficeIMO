---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePdf
## SYNOPSIS
Creates a live PDF document or atomically saves a completed composition block.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficePdf [[-Path] <string>] [[-Content] <scriptblock>] [-Force] [-Theme <OfficeThemeInput>] [-Font <string>] [-Size <OfficeLength>] [-FontPath <string>] [-BoldFontPath <string>] [-ItalicFontPath <string>] [-BoldItalicFontPath <string>] [-FileVersion <PdfFileVersion>] [-CreateOutlineFromHeadings] [-OutlineExpansionLevel <Int32>] [-PageMode <PdfCatalogPageMode>] [-PageLayout <PdfCatalogPageLayout>] [-IncludePageLabels] [-PageLabelPrefix <string>] [-OpenActionPage <Int32>] [-OpenActionMode <PdfOpenActionDestinationMode>] [-OpenActionTop <Double>] [-DisplayDocTitle] [-FitWindow] [-CenterWindow] [-HideToolbar] [-HideMenubar] [-HideWindowUI] [-FlattenVisualAnnotations] [-Password <securestring>] [-OwnerPassword <securestring>] [-Permission <PdfStandardPermissions>] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Without Content no file is written and the native document keeps its optional associated path. A completed block saves and closes; Force is required to replace an existing file.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePdf -Path ./Report.pdf -Theme Modern -PassThru {
    PdfHeading 'Service review'
    PdfParagraph 'All services are available.'
}
```


### EXAMPLE 2
```powershell
$pdf = New-OfficePdf -Theme Classic -Size 11pt
Add-OfficePdfParagraph -Document $pdf -Text 'Reviewed and approved'
Save-OfficePdf -Document $pdf -Path ./Review.pdf
```


## PARAMETERS

### -AllowLoss
Gets or sets whether native-save loss is explicitly accepted.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BoldFontPath
Optional bold face used with FontPath.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BoldItalicFontPath
Optional bold italic face used with FontPath.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CenterWindow
Request PDF viewers to center the document window on screen.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Gets or sets a content block; omission returns the live document.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CreateOutlineFromHeadings
Create PDF outline/bookmark entries from heading elements.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DisplayDocTitle
Request PDF viewers to display the document title instead of the file name.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FileVersion
PDF file header version emitted by OfficeIMO.Pdf.

```yaml
Type: PdfFileVersion
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Pdf14, Pdf15, Pdf16, Pdf17, Pdf20

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FitWindow
Request PDF viewers to fit the document window to the first displayed page.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FlattenVisualAnnotations
Flatten generated FreeText and Highlight annotations into static page content.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Body font family; inherited heading defaults follow the document font.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontPath
Regular embedded font file; requires the Font family name.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Allow a completed content block to replace an existing destination.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideMenubar
Request PDF viewers to hide the menu bar.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideToolbar
Request PDF viewers to hide the toolbar.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideWindowUI
Request PDF viewers to hide user-interface elements.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludePageLabels
Emit generated catalog page labels.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ItalicFontPath
Optional italic face used with FontPath.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OpenActionMode
Open-action destination mode.

```yaml
Type: PdfOpenActionDestinationMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Xyz, Fit, FitHorizontal, FitVertical, FitRectangle, FitBoundingBox, FitBoundingBoxHorizontal, FitBoundingBoxVertical

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OpenActionPage
Initial one-based page shown by PDF viewers that honor open actions.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OpenActionTop
Optional open-action top coordinate in the page's native UserUnit space.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutlineExpansionLevel
Initial outline expansion level when heading outlines are created.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OwnerPassword
Optional owner password for the generated encrypted PDF.

```yaml
Type: SecureString
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageLabelPrefix
Optional generated page-label prefix.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageLayout
Catalog page layout hint emitted for generated PDFs.

```yaml
Type: PdfCatalogPageLayout
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: SinglePage, OneColumn, TwoColumnLeft, TwoColumnRight, TwoPageLeft, TwoPageRight

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageMode
Catalog page mode hint emitted for generated PDFs.

```yaml
Type: PdfCatalogPageMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: UseNone, UseOutlines, UseThumbs, FullScreen, UseOC, UseAttachments

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether a completed content block returns the saved FileInfo.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password required to open the generated PDF.

```yaml
Type: SecureString
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional associated destination; required with a content block.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Permission
Operations allowed to readers; defaults to all Standard security permissions.

```yaml
Type: PdfStandardPermissions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: None, Print, ModifyContents, CopyContents, ModifyAnnotations, FillForms, Accessibility, AssembleDocument, HighQualityPrint, All

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Gets or sets whether the saved file opens in the system default application.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Size
Default body font size. A bare number means points.

```yaml
Type: OfficeLength
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Built-in theme, theme object or document/template whose visual theme is read.

```yaml
Type: OfficeThemeInput
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
