---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficePowerPointPdfOptions
## SYNOPSIS
Creates discoverable PowerPoint-to-PDF conversion options for Export-OfficeDocumentPdf.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficePowerPointPdfOptions [-PdfOptions <PdfOptions>] [-FontFamily <string>] [-IncludePictures] [-IncludeAutoShapes] [-IncludeTextBoxes] [-IncludeSlideBackgrounds] [-IncludeTables] [-IncludeCharts] [-IncludeSmartArt] [-IncludeHiddenSlides] [-PageLayout <PowerPointPdfPageLayout>] [-HandoutSlidesPerPage <Int32>] [-IncludeSpeakerNotes] [-MaxGroupShapeDepth <Int32>] [-PictureFit <OfficeImageFit>] [-WarnOnPictureAspectRatioDistortion] [-ChartStyle <OfficeChartStyle>] [-ChartLayout <OfficeChartLayout>] [-AllowSystemFontEmbedding] [-AllowDocumentFontEmbedding] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable PowerPoint-to-PDF conversion options for Export-OfficeDocumentPdf.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficePowerPointPdfOptions -PageLayout Handouts -HandoutSlidesPerPage 3 -IncludeSpeakerNotes -IncludeHiddenSlides
Export-OfficeDocumentPdf -InputPath .\Briefing.pptx -Path .\Briefing.pdf -PowerPointOptions $options
```


## PARAMETERS

### -AllowDocumentFontEmbedding
Allow embedding installed fonts selected by a trusted presentation. System-font embedding must also be enabled.

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

### -AllowSystemFontEmbedding
Allow embedding fonts discovered on the current system.

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

### -ChartLayout
Chart layout override.

```yaml
Type: OfficeChartLayout
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartStyle
Chart visual style override.

```yaml
Type: OfficeChartStyle
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontFamily
Default font family used when the presentation does not specify one.

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

### -HandoutSlidesPerPage
Number of slides on each handout page.

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

### -IncludeAutoShapes
Render automatic shapes.

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

### -IncludeCharts
Render charts.

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

### -IncludeHiddenSlides
Include slides marked hidden.

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

### -IncludePictures
Render pictures.

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

### -IncludeSlideBackgrounds
Render slide backgrounds.

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

### -IncludeSmartArt
Render SmartArt.

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

### -IncludeSpeakerNotes
Include speaker notes.

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

### -IncludeTables
Render tables.

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

### -IncludeTextBoxes
Render text boxes.

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

### -MaxGroupShapeDepth
Maximum nested group-shape depth to render.

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

### -PageLayout
PDF page layout, such as slides, notes, or handouts.

```yaml
Type: PowerPointPdfPageLayout
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Slides, NotesPages, Handouts

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfOptions
Underlying low-level OfficeIMO PDF options.

```yaml
Type: PdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PictureFit
How pictures fit their shape bounds.

```yaml
Type: OfficeImageFit
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Stretch, Contain, Cover

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WarnOnPictureAspectRatioDistortion
Report pictures whose requested fit distorts their aspect ratio.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.PowerPoint.Pdf.PowerPointToPdfOptions`

## RELATED LINKS

- None
