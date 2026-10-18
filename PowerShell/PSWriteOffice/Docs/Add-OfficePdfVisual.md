---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfVisual
## SYNOPSIS
Adds a ChartForgeX artifact, portable SVG, or converted Office visual to PDF flow content.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfVisual [-InputObject] <Object> [-Align <PdfAlign>] [-SpacingBefore <Double>] [-SpacingAfter <Double>] [-PassThru] [-SvgPolicy <OfficeVisualSvgPolicy>] [-Width <Double>] [-Height <Double>] [-PointsPerPixel <double>] [-MaximumSvgElements <Int32>] [-MaximumSvgViewportDimension <Double>] [-MaximumSvgViewportPixels <Double>] [-Id <string>] [-Title <string>] [-AlternativeText <string>] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfVisual [-InputObject] <Object> -Document <PdfDocument> [-Align <PdfAlign>] [-SpacingBefore <Double>] [-SpacingAfter <Double>] [-PassThru] [-SvgPolicy <OfficeVisualSvgPolicy>] [-Width <Double>] [-Height <Double>] [-PointsPerPixel <double>] [-MaximumSvgElements <Int32>] [-MaximumSvgViewportDimension <Double>] [-MaximumSvgViewportPixels <Double>] [-Id <string>] [-Title <string>] [-AlternativeText <string>] [<CommonParameters>]
```

### PipelineDocument
```powershell
Add-OfficePdfVisual [-InputObject] <Object> -Document <PdfDocument> [-Align <PdfAlign>] [-SpacingBefore <Double>] [-SpacingAfter <Double>] [-PassThru] [-SvgPolicy <OfficeVisualSvgPolicy>] [-Width <Double>] [-Height <Double>] [-PointsPerPixel <double>] [-MaximumSvgElements <Int32>] [-MaximumSvgViewportDimension <Double>] [-MaximumSvgViewportPixels <Double>] [-Id <string>] [-Title <string>] [-AlternativeText <string>] [<CommonParameters>]
```

## DESCRIPTION
Adds a ChartForgeX artifact, portable SVG, or converted Office visual to PDF flow content.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficePdfVisual -InputObject 'Value'
```


### EXAMPLE 2
```powershell
Add-OfficePdfVisual -InputObject 'Value' -Document 'Value'
```


## PARAMETERS

### -Align
Horizontal alignment in PDF flow.

```yaml
Type: PdfAlign
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values: Left, Center, Right, Justify

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AlternativeText
Optional accessible description for an SVG file input.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
PDF document to update outside the DSL context.

```yaml
Type: PdfDocument
Parameter Sets: Document, PipelineDocument
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Height
Optional output height in Office points.

```yaml
Type: Double
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Id
Optional stable identifier for an SVG file input.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
ChartForgeX VisualArtifact, OfficeVisualSource, OfficeVisualConversionResult, or SVG file path.

```yaml
Type: Object
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MaximumSvgElements
Optional OfficeIMO SVG import element limit.

```yaml
Type: Int32
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumSvgViewportDimension
Optional maximum SVG viewport width or height. Increase the safe default only for trusted input.

```yaml
Type: Double
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumSvgViewportPixels
Optional maximum SVG viewport area in pixels. Increase the safe default only for trusted input.

```yaml
Type: Double
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated PDF document.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PointsPerPixel
Conversion factor from ChartForgeX pixels to Office points.

```yaml
Type: Double
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingAfter
Spacing after the visual in points.

```yaml
Type: Double
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingBefore
Spacing before the visual in points.

```yaml
Type: Double
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SvgPolicy
SVG fidelity policy used by OfficeIMO.ChartForgeX.

```yaml
Type: OfficeVisualSvgPolicy
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values: PreserveVector, RasterizeWhenNeeded, RequireVector

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Optional title for an SVG file input.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Optional output width in Office points.

```yaml
Type: Double
Parameter Sets: Context, Document, PipelineDocument
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

- `System.Object`
- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
