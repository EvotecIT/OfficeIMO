---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointCard
## SYNOPSIS
Adds a narrative card or logo tile to the current designed slide.

## SYNTAX
### Card (Default)
```powershell
Add-OfficePowerPointCard [-Title] <string> [-Text <string>] [-Items <string[]>] [-Image <string>] [-Color <OfficeColor>] [<CommonParameters>]
```

## DESCRIPTION
Adds a narrative card or logo tile to the current designed slide.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficePowerPointCard -Title 'Value'
```


## PARAMETERS

### -Color
Gets or sets the color.

```yaml
Type: OfficeColor
Parameter Sets: Card
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Image
Gets or sets the image.

```yaml
Type: String
Parameter Sets: Card
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Items
Gets or sets the items.

```yaml
Type: String[]
Parameter Sets: Card
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Gets or sets the text.

```yaml
Type: String
Parameter Sets: Card
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Gets or sets the title.

```yaml
Type: String
Parameter Sets: Card
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

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
