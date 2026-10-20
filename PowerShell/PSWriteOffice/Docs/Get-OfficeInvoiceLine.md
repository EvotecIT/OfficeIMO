---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeInvoiceLine
## SYNOPSIS
Returns editable invoice lines, optionally selected by their stable identifier.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeInvoiceLine -Invoice <Invoice> [-Id <string[]>] [<CommonParameters>]
```

## DESCRIPTION
Returns editable invoice lines, optionally selected by their stable identifier.

## EXAMPLES

### EXAMPLE 1
```powershell
$invoice | Get-OfficeInvoiceLine
```


## PARAMETERS

### -Id
Optional exact line identifiers.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Invoice
Editable invoice.

```yaml
Type: Invoice
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

- `OfficeIMO.Invoicing.Invoice`

## OUTPUTS

- `OfficeIMO.Invoicing.InvoiceLine`

## RELATED LINKS

- None
