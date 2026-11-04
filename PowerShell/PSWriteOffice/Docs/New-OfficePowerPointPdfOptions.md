---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePowerPointPdfOptions
## SYNOPSIS
Creates native PowerPoint-to-PDF options for Convert-OfficeDocument.

## SYNTAX
### Default (Default)
```powershell
New-OfficePowerPointPdfOptions [-PdfOptions <PdfCreateOptions>] [-Font <string>] [-NoPictures] [-NoAutoShapes] [-NoTextBoxes] [-NoSlideBackgrounds] [-NoTables] [-NoCharts] [-NoSmartArt] [-IncludeHiddenSlides] [-PageLayout <PowerPointPdfPageLayout>] [-HandoutSlidesPerPage <Int32>] [-NoSpeakerNotes] [-MaxGroupShapeDepth <Int32>] [-PictureFit <OfficeImageFit>] [-WarnOnPictureAspectRatioDistortion] [-ChartStyle <OfficeChartStyle>] [-ChartLayout <OfficeChartLayout>] [-AllowSystemFontEmbedding] [-AllowDocumentFontEmbedding] [<CommonParameters>]
```

## DESCRIPTION
Creates native PowerPoint-to-PDF options for Convert-OfficeDocument.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficePowerPointPdfOptions -PageLayout Handouts -HandoutSlidesPerPage 3 -IncludeHiddenSlides
Convert-OfficeDocument -Path ./Briefing.pptx -To Pdf -OutputPath ./Briefing.pdf -Options $options
```


### EXAMPLE 2
```powershell
$options = New-OfficePowerPointPdfOptions -NoCharts -NoSmartArt -Font Aptos
Convert-OfficeDocument -Path ./Briefing.pptx -To Pdf -OutputPath ./Review.pdf -Options $options
```


## PARAMETERS

### -AllowDocumentFontEmbedding
Allow embedding installed fonts selected by a trusted presentation. System-font embedding must also be enabled.

```yaml
Type: SwitchParameter
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Default font family used when the presentation does not specify one.

```yaml
Type: String
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoAutoShapes
Omit automatic shapes.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoCharts
Omit charts.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoPictures
Omit pictures.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoSlideBackgrounds
Omit slide backgrounds.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoSmartArt
Omit SmartArt.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoSpeakerNotes
Omit speaker notes.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoTables
Omit tables.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoTextBoxes
Omit text boxes.

```yaml
Type: SwitchParameter
Parameter Sets: Default
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
Parameter Sets: Default
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
Type: PdfCreateOptions
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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

- `OfficeIMO.PowerPoint.Pdf.PowerPointPdfExportOptions`

## RELATED LINKS

- None
