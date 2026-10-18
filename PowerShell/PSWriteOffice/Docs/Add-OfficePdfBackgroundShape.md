---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfBackgroundShape
## SYNOPSIS
Adds a decorative generated PDF page background shape or band.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfBackgroundShape [-Shape] <OfficePdfBackgroundShapeType> [-X <double>] [-Y <double>] [-Width <Double>] [-Height <Double>] [-CornerRadius <double>] [-InsetX <double>] [-InsetY <double>] [-OffsetY <double>] [-OffsetX <double>] [-FillColor <string>] [-StrokeColor <string>] [-StrokeWidth <double>] [-FillOpacity <Double>] [-StrokeOpacity <Double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfBackgroundShape [-Shape] <OfficePdfBackgroundShapeType> -Document <PdfDocument> [-X <double>] [-Y <double>] [-Width <Double>] [-Height <Double>] [-CornerRadius <double>] [-InsetX <double>] [-InsetY <double>] [-OffsetY <double>] [-OffsetX <double>] [-FillColor <string>] [-StrokeColor <string>] [-StrokeWidth <double>] [-FillOpacity <Double>] [-StrokeOpacity <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Background shapes are intended for subtle page structure such as header bands, side bands, highlight panels, or decorative accents.
They are rendered behind generated content and should usually use restrained opacity values.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Report.pdf {
  PdfBackgroundShape -Shape TopBand -Height 86 -FillColor '#DBEAFE' -FillOpacity 0.75
  PdfBackgroundShape -Shape Ellipse -X 420 -Y 650 -Width 96 -Height 72 -FillColor '#99F6E4' -FillOpacity 0.35
  PdfHeading 'Styled report'
}
```

Creates a polished generated page background without hand-drawing PDF primitives.

## PARAMETERS

### -CornerRadius
Rounded rectangle or band corner radius in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Compatibility parameter. Page composition is supported only inside New-OfficePdf.

```yaml
Type: PdfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FillColor
Fill color. Named colors and hexadecimal values are accepted.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FillOpacity
Fill opacity from 0 to 1.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Shape height in PDF points, or band height for top/bottom bands.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InsetX
Horizontal inset for top/bottom bands in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InsetY
Vertical inset for left/right bands in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OffsetX
Horizontal offset for left/right bands in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OffsetY
Vertical offset for top/bottom bands in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated document.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Shape
Shape type to add.

```yaml
Type: OfficePdfBackgroundShapeType
Parameter Sets: Context, Document
Aliases: None
Possible values: Rectangle, RoundedRectangle, Ellipse, TopBand, BottomBand, LeftBand, RightBand

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StrokeColor
Stroke color. Named colors and hexadecimal values are accepted.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StrokeOpacity
Stroke opacity from 0 to 1.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StrokeWidth
Stroke width in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Shape width in PDF points, or band width for left/right bands.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Shape left coordinate in PDF points for explicit shapes.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Shape bottom coordinate in PDF points for explicit shapes.

```yaml
Type: Double
Parameter Sets: Context, Document
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

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
