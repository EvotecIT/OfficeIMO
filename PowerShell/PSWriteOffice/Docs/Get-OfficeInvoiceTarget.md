---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeInvoiceTarget
## SYNOPSIS
Lists exact supported invoice authoring contracts from the engine.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeInvoiceTarget [-Syntax <InvoiceSyntax>] [-Profile <InvoiceProfile>] [<CommonParameters>]
```

## DESCRIPTION
Lists exact supported invoice authoring contracts from the engine.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeInvoiceTarget -Syntax Ubl
```


## PARAMETERS

### -Profile
Optional guideline profile filter.

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

### -Syntax
Optional XML syntax filter.

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

- `None`

## OUTPUTS

- `OfficeIMO.Invoicing.InvoiceXmlOptions`

## RELATED LINKS

- None
