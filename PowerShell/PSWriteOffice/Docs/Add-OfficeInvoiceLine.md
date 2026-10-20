---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeInvoiceLine
## SYNOPSIS
Adds an invoice line with explicit quantity, price and VAT.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeInvoiceLine [-Name] <string> -Invoice <Invoice> -Quantity <decimal> -UnitPrice <decimal> [-TaxRate <Decimal>] [-TaxCategory <string>] [-ExemptionReason <string>] [-Id <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds an invoice line with explicit quantity, price and VAT.

## EXAMPLES

### EXAMPLE 1
```powershell
$invoice | Add-OfficeInvoiceLine -Name Consulting -Quantity 1 -UnitPrice 100 -TaxRate 19
```


## PARAMETERS

### -ExemptionReason
Exemption reason when required by the category.

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

### -Id
Optional stable line identifier; omission uses the next line number.

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

### -Name
Line description.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the created or changed object.

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

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TaxCategory
UNCL 5305 VAT category, such as S, Z or E.

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
VAT percentage; omission is retained for categories that prohibit a rate.

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

Required: True
Position: named
Default value: None
Accept pipeline input: False
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
