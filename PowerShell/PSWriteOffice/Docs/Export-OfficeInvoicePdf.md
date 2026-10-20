---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficeInvoicePdf
## SYNOPSIS
Creates a visible PDF and embedded invoice XML from the same captured model.

## SYNTAX
### __AllParameterSets
```powershell
Export-OfficeInvoicePdf [-Path] <string> -Invoice <Invoice> [-SpecificationRelease <InvoiceSpecificationRelease>] [-Profile <InvoiceProfile>] [-PdfOptions <PdfOptions>] [-Force] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Creates a visible PDF and embedded invoice XML from the same captured model.

## EXAMPLES

### EXAMPLE 1
```powershell
$invoice | Export-OfficeInvoicePdf -Path ./invoice.pdf
```


## PARAMETERS

### -Force
Replace an existing destination.

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

### -PassThru
Return the saved file.

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
Destination PDF filename.

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

### -PdfOptions
Optional existing PDF rendering options.

```yaml
Type: PdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Profile
Embedded invoice profile.

```yaml
Type: InvoiceProfile
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Minimum, BasicWithoutLines, Basic, En16931, Extended, XRechnung, PeppolBis, ExtendedCtcFr

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpecificationRelease
Exact supported CII specification for the embedded invoice.

```yaml
Type: InvoiceSpecificationRelease
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: En16931_1_3_16, FacturX_1_09_2_Zugferd_2_5_2, XRechnung_3_0_2_2026_08_31, PeppolBis_3_0_21

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

- `System.IO.FileInfo`

## RELATED LINKS

- None
