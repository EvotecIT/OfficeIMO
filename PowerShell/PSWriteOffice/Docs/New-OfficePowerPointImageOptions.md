---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePowerPointImageOptions
## SYNOPSIS
Creates discoverable slide selection and rendering settings for Export-OfficePowerPointImage.

## SYNTAX
### Default (Default)
```powershell
New-OfficePowerPointImageOptions [-IncludeHiddenSlides] [-NoSlideBackground] [-NoSlideContent] [-NoPictures] [-NoAutoShapes] [-NoTextBoxes] [-NoTables] [-NoCharts] [-IncludeHiddenShapes] [-MaximumOutputWidth <Int32>] [-MaximumOutputHeight <Int32>] [-MaximumRasterPixelCount <Int64>] [-RasterOverflowBehavior <OfficeRasterOverflowBehavior>] [-MaximumOutputCount <Int32>] [-MaximumTotalRasterPixelCount <Int64>] [-MaximumTotalEncodedBytes <Int64>] [-RenderTimeoutSeconds <Double>] [-MaximumDegreeOfParallelism <Int32>] [-TextShapingLanguage <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable slide selection and rendering settings for Export-OfficePowerPointImage.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficePowerPointImageOptions -NoPictures
Export-OfficeDocumentImage -Path ./Deck.pptx -Page 1,3 -OutputDirectory ./Slides -Options $options
```


## PARAMETERS

### -IncludeHiddenShapes
Render hidden shapes.

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

### -IncludeHiddenSlides
Include hidden slides.

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

### -MaximumDegreeOfParallelism
Maximum independent renders processed concurrently.

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

### -MaximumOutputCount
Maximum images accepted from one batch export.

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

### -MaximumOutputHeight
Maximum output height in pixels.

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

### -MaximumOutputWidth
Maximum output width in pixels.

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

### -MaximumRasterPixelCount
Maximum pixels allocated for one raster image.

```yaml
Type: Int64
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTotalEncodedBytes
Maximum aggregate encoded bytes accepted from one batch.

```yaml
Type: Int64
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTotalRasterPixelCount
Maximum aggregate raster pixels accepted from one batch.

```yaml
Type: Int64
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
Omit auto shapes.

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

### -NoSlideBackground
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

### -NoSlideContent
Omit slide content.

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

### -RasterOverflowBehavior
Reduce or reject oversized raster output.

```yaml
Type: OfficeRasterOverflowBehavior
Parameter Sets: Default
Aliases: None
Possible values: ReduceScale, Throw

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RenderTimeoutSeconds
Maximum seconds allowed for one render.

```yaml
Type: Double
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextShapingLanguage
BCP 47 language hint for text shaping.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointImageExportOptions`

## RELATED LINKS

- None
