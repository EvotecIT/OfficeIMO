---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeWordImage
## SYNOPSIS
Embeds an image file, SVG or reusable visual in a Word paragraph.

## SYNTAX
### Path (Default)
```powershell
Add-OfficeWordImage [-Path] <string> [-Paragraph <WordParagraph>] [-Wrap <WordImageTextWrapping>] [-Description <string>] [-PassThru] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Title <string>] [<CommonParameters>]
```

### Visual
```powershell
Add-OfficeWordImage -InputObject <Object> [-Paragraph <WordParagraph>] [-Wrap <WordImageTextWrapping>] [-Description <string>] [-PassThru] [-SvgPolicy <OfficeVisualSvgPolicy>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-PointsPerPixel <double>] [-MaximumSvgElements <Int32>] [-MaximumSvgViewportDimension <Double>] [-MaximumSvgViewportPixelCount <Int64>] [-Id <string>] [-Title <string>] [<CommonParameters>]
```

## DESCRIPTION
Numbers mean points; unit suffixes such as 4cm are accepted. A prepared visual is reused without changing its payload or metadata. One dimension preserves proportions; two dimensions set the displayed frame.

## EXAMPLES

### EXAMPLE 1
```powershell
WordParagraph { WordImage -Path ./logo.png -Width 4cm -Description 'Contoso' }
```


### EXAMPLE 2
```powershell
$visual | WordImage -Paragraph $paragraph -Width 12cm -Title 'Availability' -PassThru
```


## PARAMETERS

### -Description
Accessible image description; omitted visual descriptions are retained.

```yaml
Type: String
Parameter Sets: Path, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Optional output height; numbers mean points and unit suffixes are accepted.

```yaml
Type: OfficeLength
Parameter Sets: Path, Visual
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
Parameter Sets: Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
ChartForgeX artifact, OfficeVisualSource, prepared visual or SVG path.

```yaml
Type: Object
Parameter Sets: Visual
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MaximumSvgElements
Optional OfficeIMO SVG import element limit.

```yaml
Type: Int32
Parameter Sets: Visual
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
Parameter Sets: Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumSvgViewportPixelCount
Optional maximum SVG viewport area in pixels. Increase the safe default only for trusted input.

```yaml
Type: Int64
Parameter Sets: Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
Target paragraph; the current DSL paragraph is used when omitted.

```yaml
Type: WordParagraph
Parameter Sets: Path, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Return the native editable WordImage.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Literal local image path.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PointsPerPixel
Conversion factor from ChartForgeX pixels to Office points.

```yaml
Type: Double
Parameter Sets: Visual
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
Parameter Sets: Visual
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
Parameter Sets: Path, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Optional output width; numbers mean points and unit suffixes are accepted.

```yaml
Type: OfficeLength
Parameter Sets: Path, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Wrap
Text wrapping around the embedded image.

```yaml
Type: WordImageTextWrapping
Parameter Sets: Path, Visual
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
- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `OfficeIMO.Word.WordImage`

## RELATED LINKS

- None
