---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordVisual
## SYNOPSIS
Adds a ChartForgeX artifact, portable SVG, or converted Office visual to Word.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordVisual [-InputObject] <Object> [-Paragraph <WordParagraph>] [-Wrap <WordImageTextWrapping>] [-PassThru] [-SvgPolicy <OfficeVisualSvgPolicy>] [-Width <Double>] [-Height <Double>] [-PointsPerPixel <double>] [-MaximumSvgElements <Int32>] [-MaximumSvgViewportDimension <Double>] [-MaximumSvgViewportPixels <Double>] [-Id <string>] [-Title <string>] [-AlternativeText <string>] [<CommonParameters>]
```

## DESCRIPTION
Adds a ChartForgeX artifact, portable SVG, or converted Office visual to Word.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> WordParagraph { $artifact | Add-OfficeWordVisual -Width 420 }
```

Embeds the selected SVG or PNG payload with the artifact's accessible description.

## PARAMETERS

### -AlternativeText
Optional accessible description for an SVG file input.

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

### -Height
Optional output height in Office points.

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

### -Id
Optional stable identifier for an SVG file input.

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

### -InputObject
ChartForgeX VisualArtifact, OfficeVisualSource, OfficeVisualConversionResult, or SVG file path.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
Target paragraph. Inside the Word DSL, the current paragraph is used by default.

```yaml
Type: WordParagraph
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the image added to the paragraph.

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

### -PointsPerPixel
Conversion factor from ChartForgeX pixels to Office points.

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

### -SvgPolicy
SVG fidelity policy used by OfficeIMO.ChartForgeX.

```yaml
Type: OfficeVisualSvgPolicy
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Wrap
Word text-wrapping behavior.

```yaml
Type: WordImageTextWrapping
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: InLineWithText, Square, Tight, Through, TopAndBottom, BehindText, InFrontOfText

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

## OUTPUTS

- `OfficeIMO.Word.WordImage`

## RELATED LINKS

- None
