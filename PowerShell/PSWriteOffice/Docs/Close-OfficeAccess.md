---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Close-OfficeAccess
## SYNOPSIS
Closes a native database without implicitly saving it.

## SYNTAX
### __AllParameterSets
```powershell
Close-OfficeAccess -Document <AccessDocument> [<CommonParameters>]
```

## DESCRIPTION
Closes a native database without implicitly saving it.

## EXAMPLES

### EXAMPLE 1
```powershell
$database | Close-OfficeAccess
```


## PARAMETERS

### -Document
Live database document.

```yaml
Type: AccessDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Access.AccessDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
