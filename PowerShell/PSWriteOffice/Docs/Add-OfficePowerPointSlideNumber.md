---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointSlideNumber
## SYNOPSIS
Inserts a native slide-number field into the current footer.

## SYNTAX
### Default (Default)
```powershell
Add-OfficePowerPointSlideNumber [-IncludeTotal] [<CommonParameters>]
```

## DESCRIPTION
Inserts a native slide-number field into the current footer.

## EXAMPLES

### EXAMPLE 1
```powershell
PowerPointFooter -Alignment Right { PowerPointSlideNumber -IncludeTotal }
```


## PARAMETERS

### -IncludeTotal
Append the physical slide count, refreshed when the deck is saved.

```yaml
Type: SwitchParameter
Parameter Sets: Default
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
