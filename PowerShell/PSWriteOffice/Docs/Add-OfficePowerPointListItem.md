---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointListItem
## SYNOPSIS
Adds a native paragraph to the enclosing list, optionally using styled runs.

## SYNTAX
### Text (Default)
```powershell
Add-OfficePowerPointListItem [-Text] <string> [-Level <int>] [-Bold] [-Italic] [-Strikethrough] [-Underline <PowerPointUnderlineStyle>] [-Color <OfficeColor>] [-Highlight <OfficeColor>] [-Font <string>] [-Size <OfficeLength>] [-PassThru] [<CommonParameters>]
```

### Content
```powershell
Add-OfficePowerPointListItem [-Content] <scriptblock> [-Level <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a native paragraph to the enclosing list, optionally using styled runs.

## EXAMPLES

### EXAMPLE 1
```powershell
PowerPointList { PowerPointListItem { PowerPointText 'Capacity' -Bold; PowerPointText ': review weekly' } }
```


## PARAMETERS

### -Bold
Gets or sets whether plain item text is bold.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Gets or sets the text color or theme role.

```yaml
Type: OfficeColor
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Gets or sets the styled-run content block.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Gets or sets the font family.

```yaml
Type: String
Parameter Sets: Text
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
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Gets or sets whether plain item text is italic.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Level
Gets or sets the one-based list nesting level.

```yaml
Type: Int32
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether to emit the live native item paragraph.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Content
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
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strikethrough
Gets or sets whether plain item text has a strike through it.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Gets or sets plain item text.

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

### -Underline
Gets or sets the native underline style.

```yaml
Type: PowerPointUnderlineStyle
Parameter Sets: Text
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

- `OfficeIMO.PowerPoint.PowerPointParagraph`

## RELATED LINKS

- None
