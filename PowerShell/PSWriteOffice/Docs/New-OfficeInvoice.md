---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeInvoice
## SYNOPSIS
Creates an editable invoice without serializing or declaring standards compliance.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeInvoice [-Number] <string> -IssueDate <datetime> -Seller <InvoiceParty> -Buyer <InvoiceParty> [-Currency <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates an editable invoice without serializing or declaring standards compliance.

## EXAMPLES

### EXAMPLE 1
```powershell
$invoice = New-OfficeInvoice -Number INV-001 -IssueDate 2026-10-20 -Seller $seller -Buyer $buyer -Currency EUR
```


## PARAMETERS

### -Buyer
Buyer party, created with New-OfficeInvoiceParty or a property hashtable.

```yaml
Type: InvoiceParty
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Currency
ISO currency code.

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

### -IssueDate
Explicit issue date.

```yaml
Type: DateTime
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Number
Seller-assigned document number.

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

### -Seller
Seller party, created with New-OfficeInvoiceParty or a property hashtable.

```yaml
Type: InvoiceParty
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

- `None`

## OUTPUTS

- `OfficeIMO.Invoicing.Invoice`

## RELATED LINKS

- None
