---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeInvoice
## SYNOPSIS
Updates common invoice header and party fields.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeInvoice -Invoice <Invoice> [-Number <string>] [-IssueDate <DateTime>] [-Currency <string>] [-TaxCurrency <string>] [-TaxAmountInAccountingCurrency <Decimal>] [-Seller <InvoiceParty>] [-Buyer <InvoiceParty>] [-PassThru] [-RecalculateTotals] [<CommonParameters>]
```

## DESCRIPTION
Updates common invoice header and party fields.

## EXAMPLES

### EXAMPLE 1
```powershell
$invoice | Set-OfficeInvoice -Number INV-002 -Currency EUR
```


## PARAMETERS

### -Buyer
Replacement buyer.

```yaml
Type: InvoiceParty
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Currency
Replacement currency code.

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

### -IssueDate
Replacement issue date.

```yaml
Type: DateTime
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Number
Replacement document number.

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

### -RecalculateTotals
Remove source-declared totals and tax breakdowns so the owner recomputes them from the edited model.

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

### -Seller
Replacement seller.

```yaml
Type: InvoiceParty
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TaxAmountInAccountingCurrency
Explicit VAT amount in the accounting currency; no exchange rate is inferred.

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

### -TaxCurrency
Accounting VAT currency. Changing it invalidates its previously declared amount.

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

- `OfficeIMO.Invoicing.Invoice`

## OUTPUTS

- `OfficeIMO.Invoicing.Invoice`

## RELATED LINKS

- None
