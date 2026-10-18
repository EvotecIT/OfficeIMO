---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeVisioImageOptions
## SYNOPSIS
Creates discoverable page and rendering settings for Export-OfficeVisioImage.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeVisioImageOptions [-PageIndex <Int32>] [-PageCount <Int32>] [-RenderText] [-RenderStencilArtwork] [-RenderConnectorLabels] [-ResolveConnectorLabelOverlaps] [-Supersampling <Int32>] [-IncludeSvgXmlDeclaration] [-Scale <Double>] [-MaximumOutputWidth <Int32>] [-MaximumOutputHeight <Int32>] [-BackgroundColor <string>] [-TargetDpi <Double>] [-MaximumRasterPixels <Int64>] [-RasterOverflowBehavior <OfficeRasterOverflowBehavior>] [-MaximumOutputCount <Int32>] [-MaximumTotalRasterPixels <Int64>] [-MaximumTotalEncodedBytes <Int64>] [-RenderTimeoutSeconds <Double>] [-MaximumDegreeOfParallelism <Int32>] [-TextShapingLanguage <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable page and rendering settings for Export-OfficeVisioImage.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeVisioImageOptions -PageIndex 0 -PageCount 1 -RenderText -RenderConnectorLabels
Export-OfficeVisioImage -Path .\Diagram.vsdx -OutputPath .\Preview -Format Svg -Options $options
```


## PARAMETERS

### -BackgroundColor
Background color name or hex value.

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

### -IncludeSvgXmlDeclaration
Include an XML declaration in SVG output.

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

### -MaximumDegreeOfParallelism
Maximum independent renders processed concurrently.

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

### -MaximumOutputCount
Maximum images accepted from one batch export.

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

### -MaximumOutputHeight
Maximum output height in pixels.

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

### -MaximumOutputWidth
Maximum output width in pixels.

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

### -MaximumRasterPixels
Maximum pixels allocated for one raster image.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTotalRasterPixels
Maximum aggregate raster pixels accepted from one batch.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageCount
Maximum pages exported.

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

### -PageIndex
Zero-based first page index.

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

### -RasterOverflowBehavior
Reduce or reject oversized raster output.

```yaml
Type: OfficeRasterOverflowBehavior
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: ReduceScale, Throw

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RenderConnectorLabels
Render connector labels.

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

### -RenderStencilArtwork
Render supported stencil artwork.

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

### -RenderText
Render page text.

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

### -RenderTimeoutSeconds
Maximum seconds allowed for one render.

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

### -ResolveConnectorLabelOverlaps
Resolve connector-label overlaps.

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

### -Scale
Output scale multiplier.

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

### -Supersampling
Raster supersampling factor.

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

### -TargetDpi
Target output density in dots per inch.

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

### -TextShapingLanguage
BCP 47 language hint for text shaping.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Visio.VisioImageExportOptions`

## RELATED LINKS

- None
