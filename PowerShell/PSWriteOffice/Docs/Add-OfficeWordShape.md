---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordShape
## SYNOPSIS
Adds a basic OfficeIMO Word shape to the current paragraph.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordShape [-Type <WordShapeType>] [-Width <double>] [-Height <double>] [-Left <Double>] [-Top <Double>] [-FillColor <string>] [-StrokeColor <string>] [-StrokeWidth <Double>] [-Title <string>] [-Description <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a basic OfficeIMO Word shape to the current paragraph.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\StatusReport.docx {
    Add-OfficeWordParagraph -Text 'Release readiness'
    Add-OfficeWordShape -Type Rectangle -Width 220 -Height 56 -FillColor '#e6fffb' -StrokeColor '#08979c' -StrokeWidth 1.5 -Title 'Status callout' -Description 'Release readiness callout'
}
```

Creates an OfficeIMO Word shape in the current paragraph and sets basic visual and accessibility metadata.

### EXAMPLE 2
```powershell
PS> New-OfficeWord -Path .\Appendix.docx {
    Add-OfficeWordParagraph -Text 'Appendix A'
    Add-OfficeWordShape -Type Rectangle -Width 480 -Height 36 -Left 36 -Top 72 -FillColor '#f0f5ff' -StrokeColor '#adc6ff'
}
```

Positions a shape with explicit offsets when the OfficeIMO anchored-shape API is desired.

## PARAMETERS

### -Description
Optional alternate text metadata.

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

### -FillColor
Fill color. Named colors and hexadecimal values are accepted.

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
Height in points.

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

### -Left
Anchored left position in points. When omitted, the shape is inline.

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

### -PassThru
Emit the created shape.

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

### -StrokeColor
Stroke color. Named colors and hexadecimal values are accepted.

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

### -StrokeWidth
Stroke width in points.

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

### -Title
Optional title metadata.

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

### -Top
Anchored top position in points. When omitted, the shape is inline.

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

### -Type
Shape type to add.

```yaml
Type: WordShapeType
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Rectangle, Ellipse, Line, RoundedRectangle, Triangle, Diamond, Pentagon, Hexagon, Parallelogram, Trapezoid, Chevron, Plus, RightArrow, LeftArrow, UpArrow, DownArrow, LeftRightArrow, Star5, Heart, Cloud, Donut, Can, Cube

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Width in points.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Word.WordShape`

## RELATED LINKS

- None
