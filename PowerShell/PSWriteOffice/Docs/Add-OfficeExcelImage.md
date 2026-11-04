---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelImage
## SYNOPSIS
Embeds a file, HTTP image or reusable visual at a worksheet cell or range.

## SYNTAX
### Path (Default)
```powershell
Add-OfficeExcelImage [-Path] <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Address <string>] [-Scale <Double>] [-OffsetX <OfficeLength>] [-OffsetY <OfficeLength>] [-EndOffsetX <OfficeLength>] [-EndOffsetY <OfficeLength>] [-Name <string>] [-AltText <string>] [-Decorative] [-NoLockAspectRatio] [-Placement <ExcelImagePlacement>] [-Rotation <double>] [-PassThru] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Title <string>] [<CommonParameters>]
```

### Url
```powershell
Add-OfficeExcelImage -Url <uri> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Address <string>] [-Scale <Double>] [-OffsetX <OfficeLength>] [-OffsetY <OfficeLength>] [-EndOffsetX <OfficeLength>] [-EndOffsetY <OfficeLength>] [-Name <string>] [-AltText <string>] [-Decorative] [-NoLockAspectRatio] [-Placement <ExcelImagePlacement>] [-Rotation <double>] [-PassThru] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Title <string>] [<CommonParameters>]
```

### Visual
```powershell
Add-OfficeExcelImage -InputObject <Object> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Address <string>] [-Scale <Double>] [-OffsetX <OfficeLength>] [-OffsetY <OfficeLength>] [-EndOffsetX <OfficeLength>] [-EndOffsetY <OfficeLength>] [-Name <string>] [-AltText <string>] [-Decorative] [-NoLockAspectRatio] [-Placement <ExcelImagePlacement>] [-Rotation <double>] [-PassThru] [-SvgPolicy <OfficeVisualSvgPolicy>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-PointsPerPixel <double>] [-MaximumSvgElements <Int32>] [-MaximumSvgViewportDimension <Double>] [-MaximumSvgViewportPixelCount <Int64>] [-Id <string>] [-Title <string>] [<CommonParameters>]
```

## DESCRIPTION
A cell uses a one-cell anchor; a range uses a two-cell anchor. Numbers mean points and unit suffixes are accepted. With no dimensions or Scale, file images retain their detected size. Range anchors determine their own size. URL downloads use the native network policy and cancellation token.

## EXAMPLES

### EXAMPLE 1
```powershell
ExcelSheet 'Cover' { ExcelImage -Path ./logo.png -Address A1 -AltText 'Contoso' }
```


### EXAMPLE 2
```powershell
$visual | ExcelImage -Document $book -Sheet Summary -Address B2:G12 -Name Health
```


### EXAMPLE 3
```powershell
ExcelImage -Url https://status.example.com/badge.png -Address H1 -Width 4cm
```


## PARAMETERS

### -Address
A1 cell or range; defaults to A1. Use B2:B2 for a two-cell anchor over one cell.

```yaml
Type: String
Parameter Sets: Path, Url, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AltText
Accessible image description; omitted visual descriptions are retained.

```yaml
Type: String
Parameter Sets: Path, Url, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Decorative
Clear accessible title and description and mark the image decorative.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Url, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to modify outside the Excel DSL.

```yaml
Type: ExcelDocument
Parameter Sets: Path, Url, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -EndOffsetX
Horizontal end-marker offset for a range anchor; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Path, Url, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EndOffsetY
Vertical end-marker offset for a range anchor; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Path, Url, Visual
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
Parameter Sets: Path, Url, Visual
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

### -Name
Drawing name in Excel's selection pane.

```yaml
Type: String
Parameter Sets: Path, Url, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLockAspectRatio
Allow interactive aspect-ratio changes in Excel; the default locks the ratio.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Url, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OffsetX
Horizontal start-marker offset; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Path, Url, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OffsetY
Vertical start-marker offset; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Path, Url, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the native editable ExcelImage.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Url, Visual
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

### -Placement
Movement and resizing behavior of a two-cell anchor.

```yaml
Type: ExcelImagePlacement
Parameter Sets: Path, Url, Visual
Aliases: None
Possible values: MoveAndSize, MoveOnly, FreeFloating

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
Parameter Sets: Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rotation
Clockwise rotation in degrees.

```yaml
Type: Double
Parameter Sets: Path, Url, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Scale
Positive percentage of the source size; exclusive with dimensions and range anchors.

```yaml
Type: Double
Parameter Sets: Path, Url, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Sheet name, one-based position or native worksheet belonging to Document.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Path, Url, Visual
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
Parameter Sets: Path, Url, Visual
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Url
HTTP or HTTPS image URI; local files use Path.

```yaml
Type: Uri
Parameter Sets: Url
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Optional output width; numbers mean points and unit suffixes are accepted.

```yaml
Type: OfficeLength
Parameter Sets: Path, Url, Visual
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
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelImage`

## RELATED LINKS

- None
