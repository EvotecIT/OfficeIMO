---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointShape
## SYNOPSIS
Reads shape summaries from a slide or presentation.

## SYNTAX
### Slide (Default)
```powershell
Get-OfficePowerPointShape [-Slide <PowerPointSlide>] [-ShapeIndex <int[]>] [-Name <string[]>] [-Kind <OfficePowerPointShapeKind[]>] [<CommonParameters>]
```

### Presentation
```powershell
Get-OfficePowerPointShape [-Presentation <PowerPointPresentation>] [-Index <Int32>] [-ShapeIndex <int[]>] [-Name <string[]>] [-Kind <OfficePowerPointShapeKind[]>] [<CommonParameters>]
```

## DESCRIPTION
Returns PowerShell-friendly metadata for text boxes, pictures, tables, charts, and auto shapes.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Get-OfficePowerPointShape
```

Returns shape summaries for the selected slide.

### EXAMPLE 2
```powershell
PS> Get-OfficePowerPointShape -Presentation $ppt -Index 0 -Kind Picture
```

Filters the slide output to picture shapes only.

## PARAMETERS

### -Index
Optional zero-based slide index when reading from a presentation.

```yaml
Type: Int32
Parameter Sets: Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Kind
Optional shape kind filter.

```yaml
Type: OfficePowerPointShapeKind[]
Parameter Sets: Slide, Presentation
Aliases: None
Possible values: TextBox, Picture, Table, Chart, AutoShape, GroupShape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Optional wildcard filter for shape names.

```yaml
Type: String[]
Parameter Sets: Slide, Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Presentation
Presentation whose slides should be inspected.

```yaml
Type: PowerPointPresentation
Parameter Sets: Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ShapeIndex
Optional zero-based shape index filter.

```yaml
Type: Int32[]
Parameter Sets: Slide, Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Slide to inspect (optional inside the DSL).

```yaml
Type: PowerPointSlide
Parameter Sets: Slide
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`
- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `PSWriteOffice.Services.PowerPoint.PowerPointShapeInfo`: PowerShell-friendly description of a PowerPoint shape.

## RELATED LINKS

- None
