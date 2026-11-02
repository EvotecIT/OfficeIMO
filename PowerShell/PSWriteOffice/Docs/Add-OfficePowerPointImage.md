---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointImage
## SYNOPSIS
Adds a native picture from a local image, SVG or reusable visual, using the slide layout or explicit lengths.

## SYNTAX
### Path (Default)
```powershell
Add-OfficePowerPointImage [-Path] <string> [-AlternativeText <string>] [-Slide <PowerPointSlideSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Column <Int32>] [-Span <Int32>] [-FocalPoint <PowerPointFocalPoint>] [-PassThru] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Title <string>] [<CommonParameters>]
```

### Visual
```powershell
Add-OfficePowerPointImage [-InputObject] <Object> [-AlternativeText <string>] [-Slide <PowerPointSlideSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Column <Int32>] [-Span <Int32>] [-FocalPoint <PowerPointFocalPoint>] [-PassThru] [-SvgPolicy <OfficeVisualSvgPolicy>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-PointsPerPixel <double>] [-MaximumSvgElements <Int32>] [-MaximumSvgViewportDimension <Double>] [-MaximumSvgViewportPixelCount <Int64>] [-Id <string>] [-Title <string>] [<CommonParameters>]
```

## DESCRIPTION
Coordinates use points or unit suffixes. Explicit coordinates, dimensions, grid columns or Slide select fixed placement. Otherwise the image flows through the shared slide layout. FocalPoint controls cover crops.

## EXAMPLES

### EXAMPLE 1
```powershell
PowerPointSlide -Title 'Evidence' {
    PowerPointImage -Path ./evidence.png -Column 2 -Span 5 -Y 3cm -Height 5cm -AlternativeText 'Deployment health'
}
```


### EXAMPLE 2
```powershell
$visual | PowerPointImage -Slide $slide -X 36 -Y 54 -PassThru
```


## PARAMETERS

### -AlternativeText
Accessible description for the visual or placed image.

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

### -Column
One-based starting column in the native twelve-column slide grid.

```yaml
Type: Int32
Parameter Sets: Path, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FocalPoint
Center, Top, Bottom, Left, Right or fractional x,y coordinates kept in view by a cover crop.

```yaml
Type: PowerPointFocalPoint
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
Visual artifact, OfficeVisualSource, prepared conversion or SVG path.

```yaml
Type: Object
Parameter Sets: Visual
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

### -PassThru
Return the native editable picture.

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

### -Slide
Native slide or a one-based position in the current document; omitted inside PowerPointSlide.

```yaml
Type: PowerPointSlideSelector
Parameter Sets: Path, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Span
Number of grid columns occupied by the frame; defaults to twelve with Column.

```yaml
Type: Int32
Parameter Sets: Path, Visual
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

### -X
Left position; numbers mean points.

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

### -Y
Top position; numbers mean points.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Object`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointPicture`

## RELATED LINKS

- None
