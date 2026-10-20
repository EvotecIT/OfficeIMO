---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Save-OfficeInvoice
## SYNOPSIS
Saves CII or UBL with strict source and target loss handling.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeInvoice [-Path] <string> -Invoice <Invoice> [-SpecificationRelease <InvoiceSpecificationRelease>] [-Syntax <InvoiceSyntax>] [-Profile <InvoiceProfile>] [-AllowProfileLoss] [-AllowUnmappedDataLoss] [-Force] [-ReportVariable <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves CII or UBL with strict source and target loss handling.

## EXAMPLES

### EXAMPLE 1
```powershell
$invoice | Save-OfficeInvoice -Path ./invoice.xml -Syntax Ubl
```


## PARAMETERS

### -AllowProfileLoss
Permit only profile-defined reductions supported by the target engine.

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

### -AllowUnmappedDataLoss
Explicitly permit discarding unmapped source business data.

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

### -Force
Replace an existing output.

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

### -Path
Destination XML filename.

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

### -Profile
Output profile; supported combinations are listed by Get-OfficeInvoiceTarget.

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

### -ReportVariable
Variable receiving target-mapping diagnostics.

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

### -SpecificationRelease
Exact supported output specification.

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

### -Syntax
Output XML syntax.

```yaml
Type: InvoiceSyntax
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Cii, Ubl

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
