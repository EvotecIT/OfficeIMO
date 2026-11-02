---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePowerPointShape
## SYNOPSIS
Gets editable shape summaries, with optional text, name, kind and position filters.

## SYNTAX
### All (Default)
```powershell
Get-OfficePowerPointShape [-Document <PowerPointDocument>] [-Slide <PowerPointSlideSelector>] [-Index <int[]>] [-Name <string[]>] [-Kind <OfficePowerPointShapeKind[]>] [<CommonParameters>]
```

### Text
```powershell
Get-OfficePowerPointShape [-Text] <string> [-Document <PowerPointDocument>] [-Slide <PowerPointSlideSelector>] [-Index <int[]>] [-Name <string[]>] [-Kind <OfficePowerPointShapeKind[]>] [-CaseSensitive] [<CommonParameters>]
```

### Pattern
```powershell
Get-OfficePowerPointShape -Pattern <string> [-Document <PowerPointDocument>] [-Slide <PowerPointSlideSelector>] [-Index <int[]>] [-Name <string[]>] [-Kind <OfficePowerPointShapeKind[]>] [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Searches every slide of the document unless a native slide or one-based slide number is supplied. Inside PowerPointSlide, omission uses the current slide. Text matching includes table cells.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficePowerPointShape -Document $deck -Slide 2 -Kind Picture
```


### EXAMPLE 2
```powershell
Get-OfficePowerPointShape -Document $deck -Text 'Draft' -Kind TextBox |
                Set-OfficePowerPointShapeText -Text 'Ready'
```


## PARAMETERS

### -CaseSensitive
Match text and names with case sensitivity.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Pattern
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Live document; omission uses the enclosing document block.

```yaml
Type: PowerPointDocument
Parameter Sets: All, Text, Pattern
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Index
One-based shape positions within each selected slide.

```yaml
Type: Int32[]
Parameter Sets: All, Text, Pattern
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Kind
Shape kinds to include.

```yaml
Type: OfficePowerPointShapeKind[]
Parameter Sets: All, Text, Pattern
Aliases: None
Possible values: TextBox, Picture, Table, Chart, AutoShape, GroupShape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Wildcard shape-name filters.

```yaml
Type: String[]
Parameter Sets: All, Text, Pattern
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Pattern
Regular expression to find in text boxes or table cells.

```yaml
Type: String
Parameter Sets: Pattern
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Native slide or one-based slide number.

```yaml
Type: PowerPointSlideSelector
Parameter Sets: All, Text, Pattern
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Literal text to find in text boxes or table cells.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointDocument`

## OUTPUTS

- `PSWriteOffice.Services.PowerPoint.PowerPointShapeInfo`: PowerShell-friendly description of a PowerPoint shape.

## RELATED LINKS

- None
