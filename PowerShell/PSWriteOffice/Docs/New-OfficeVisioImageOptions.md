---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeVisioImageOptions
## SYNOPSIS
Creates discoverable page and rendering settings for Export-OfficeVisioImage.

## SYNTAX
### Options (Default)
```powershell
New-OfficeVisioImageOptions [-NoText] [-NoStencilArtwork] [-NoConnectorLabels] [-NoConnectorLabelOverlapResolution] [-Supersampling <Int32>] [-IncludeSvgXmlDeclaration] [-FontPath <string>] [-Font <string>] [-FontCollectionIndex <Int32>] [-MaximumOutputWidth <Int32>] [-MaximumOutputHeight <Int32>] [-MaximumRasterPixelCount <Int64>] [-RasterOverflowBehavior <OfficeRasterOverflowBehavior>] [-MaximumOutputCount <Int32>] [-MaximumTotalRasterPixelCount <Int64>] [-MaximumTotalEncodedBytes <Int64>] [-RenderTimeoutSeconds <Double>] [-MaximumDegreeOfParallelism <Int32>] [-TextShapingLanguage <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable page and rendering settings for Export-OfficeVisioImage.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeVisioImageOptions -Font Arial -NoConnectorLabels
Export-OfficeDocumentImage -Path ./Diagram.vsdx -Page 1 -Format Svg -OutputDirectory ./Preview -Options $options
```


## PARAMETERS

### -Font
Font face name used for text rendering.

```yaml
Type: String
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontCollectionIndex
Face index within a font collection.

```yaml
Type: Int32
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontPath
Literal font file used for text rendering.

```yaml
Type: String
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeSvgXmlDeclaration
Include an XML declaration in SVG output.

```yaml
Type: SwitchParameter
Parameter Sets: Options
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
Parameter Sets: Options
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
Parameter Sets: Options
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
Parameter Sets: Options
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
Parameter Sets: Options
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
Parameter Sets: Options
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
Parameter Sets: Options
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
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoConnectorLabelOverlapResolution
Resolve connector-label overlaps.

```yaml
Type: SwitchParameter
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoConnectorLabels
Render connector labels.

```yaml
Type: SwitchParameter
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoStencilArtwork
Render supported stencil artwork.

```yaml
Type: SwitchParameter
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoText
Render page text.

```yaml
Type: SwitchParameter
Parameter Sets: Options
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
Parameter Sets: Options
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
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Supersampling
Raster supersampling factor.

```yaml
Type: Int32
Parameter Sets: Options
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
Parameter Sets: Options
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

- `OfficeIMO.Visio.VisioImageExportOptions`

## RELATED LINKS

- None
