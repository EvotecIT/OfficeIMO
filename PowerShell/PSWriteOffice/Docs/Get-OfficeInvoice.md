---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeInvoice
## SYNOPSIS
Reads CII or UBL while retaining unmapped-source evidence for later saves.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeInvoice [-Path] <string> [-AsResult] [<CommonParameters>]
```

## DESCRIPTION
Reads CII or UBL while retaining unmapped-source evidence for later saves.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-Item ./invoice.xml | Get-OfficeInvoice
```


## PARAMETERS

### -AsResult
Return the source declaration and unmapped-data diagnostics around the editable model.

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

### -Path
Literal invoice XML path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FullName, FilePath, LiteralPath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.Invoicing.Invoice`

## RELATED LINKS

- None
