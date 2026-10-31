---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointText
## SYNOPSIS
Adds a live styled run to the enclosing paragraph or list item, or starts a slide paragraph.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointText [-Text] <string> [-Bold] [-Italic] [-Strikethrough] [-Underline <PowerPointUnderlineStyle>] [-Color <OfficeColor>] [-Highlight <OfficeColor>] [-Font <string>] [-Size <OfficeLength>] [-Link <string>] [-Alignment <OfficeAlignment>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a live styled run to the enclosing paragraph or list item, or starts a slide paragraph.

## EXAMPLES

### EXAMPLE 1
```powershell
PowerPointParagraph { PowerPointText 'Service status' -Bold -Link 'https://status.contoso.com' }
```


## PARAMETERS

### -Alignment
Gets or sets alignment when starting a paragraph directly in the slide or column.

```yaml
Type: OfficeAlignment
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Gets or sets bold formatting.

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

### -Color
Gets or sets text color or a theme role.

```yaml
Type: OfficeColor
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Gets or sets the font family.

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

### -Highlight
Gets or sets the highlight color.

```yaml
Type: OfficeColor
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Gets or sets italic formatting.

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

### -Link
Gets or sets the hyperlink target.

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

### -PassThru
Gets or sets whether to emit the native run.

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

### -Size
Gets or sets the font size; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strikethrough
Gets or sets strike-through formatting.

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

### -Text
Gets or sets the text to append.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Gets or sets the native underline style.

```yaml
Type: PowerPointUnderlineStyle
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: None, Words, Single, Double, Heavy, Dotted, HeavyDotted, Dash, DashHeavy, DashLong, DashLongHeavy, DotDash, DotDashHeavy, DotDotDash, DotDotDashHeavy, Wavy, WavyHeavy, WavyDouble

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

- `OfficeIMO.PowerPoint.PowerPointTextRun`

## RELATED LINKS

- None
