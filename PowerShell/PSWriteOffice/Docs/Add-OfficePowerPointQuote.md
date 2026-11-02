---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointQuote
## SYNOPSIS
Adds a quotation and attribution to the current designed slide.

## SYNTAX
### Text (Default)
```powershell
Add-OfficePowerPointQuote [-Text] <string> [[-Attribution] <string>] [<CommonParameters>]
```

## DESCRIPTION
Adds a quotation and attribution to the current designed slide.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficePowerPointQuote -Text 'Value'
```


## PARAMETERS

### -Attribution
Gets or sets the attribution.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Gets or sets the text.

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

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
