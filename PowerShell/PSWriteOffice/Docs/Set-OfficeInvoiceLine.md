---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeInvoiceLine
## SYNOPSIS
Edits selected line fields while retaining other invoice content.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeInvoiceLine -Line <InvoiceLine> [-Name <string>] [-Quantity <Decimal>] [-UnitPrice <Decimal>] [-TaxRate <Decimal>] [-TaxCategory <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Edits selected line fields while retaining other invoice content.

## EXAMPLES

### EXAMPLE 1
```powershell
$invoice | Get-OfficeInvoiceLine -Id 1 | Set-OfficeInvoiceLine -Quantity 3
```


## PARAMETERS

### -Line
Editable invoice line.

```yaml
Type: InvoiceLine
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
Description.

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
Return the changed line.

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

### -Quantity
Quantity.

```yaml
Type: Decimal
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TaxCategory
VAT category.

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

### -TaxRate
VAT percentage.

```yaml
Type: Decimal
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UnitPrice
Net unit price.

```yaml
Type: Decimal
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

- `OfficeIMO.Invoicing.InvoiceLine`

## OUTPUTS

- `OfficeIMO.Invoicing.InvoiceLine`

## RELATED LINKS

- None
