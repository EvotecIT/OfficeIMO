---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeAccessTable
## SYNOPSIS
Lists native table definitions and schema without executing their content.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeAccessTable [[-Name] <string[]>] -Document <AccessDocument> [<CommonParameters>]
```

## DESCRIPTION
Lists native table definitions and schema without executing their content.

## EXAMPLES

### EXAMPLE 1
```powershell
$database | Get-OfficeAccessTable | Select-Object Name, RowCount, Columns
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

### -Name
Optional table names.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Access.AccessDocument`

## OUTPUTS

- `OfficeIMO.Access.AccessTable`

## RELATED LINKS

- None
