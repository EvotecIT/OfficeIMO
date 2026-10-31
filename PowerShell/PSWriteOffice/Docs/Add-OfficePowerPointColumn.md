---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointColumn
## SYNOPSIS
Groups typed content into one column of a designed slide.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointColumn [-Content] <scriptblock> [-Title <string>] [<CommonParameters>]
```

## DESCRIPTION
Groups typed content into one column of a designed slide.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficePowerPointColumn -Content { }
```


## PARAMETERS

### -Content
Gets or sets the column's content block.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Gets or sets the column title.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
