---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Test-OfficeInvoice
## SYNOPSIS
Validates an invoice model or exact XML bytes against caller-supplied authority rules.

## SYNTAX
### Model (Default)
```powershell
Test-OfficeInvoice -Invoice <Invoice> [<CommonParameters>]
```

### File
```powershell
Test-OfficeInvoice -Path <string> -RulesArchivePath <string> [-FacturXArchivePath <string>] [-PeppolRulesPath <string>] [-SpecificationRelease <InvoiceSpecificationRelease>] [-SaxonJarPath <string>] [-JavaExecutable <string>] [-TimeoutSeconds <int>] [<CommonParameters>]
```

## DESCRIPTION
Model validation reports semantic completeness and arithmetic. File validation requires PowerShell 7.4 and local pinned rule archives; business rules require a caller-installed Saxon/Java runtime. Missing rule execution remains visible in the report.

## EXAMPLES

### EXAMPLE 1
```powershell
$invoice | Test-OfficeInvoice
```


## PARAMETERS

### -FacturXArchivePath
Optional pinned Factur-X validation archive.

```yaml
Type: String
Parameter Sets: File
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
Parameter Sets: Model
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -JavaExecutable
Caller-installed Java executable used only when SaxonJarPath is supplied.

```yaml
Type: String
Parameter Sets: File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Exact invoice XML bytes to validate against pinned local authority rules. Requires PowerShell 7.4.

```yaml
Type: String
Parameter Sets: File
Aliases: FullName, LiteralPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: False
```

### -PeppolRulesPath
Optional pinned Peppol Schematron file.

```yaml
Type: String
Parameter Sets: File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RulesArchivePath
Pinned KoSIT rule archive. No rules are downloaded automatically.

```yaml
Type: String
Parameter Sets: File
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SaxonJarPath
Optional caller-installed Saxon HE JAR for business rules. Without it, business rules remain NotRun and IsValid is false.

```yaml
Type: String
Parameter Sets: File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpecificationRelease
Exact specification for file validation.

```yaml
Type: InvoiceSpecificationRelease
Parameter Sets: File
Aliases: None
Possible values: En16931_1_3_16, FacturX_1_09_2_Zugferd_2_5_2, XRechnung_3_0_2_2026_08_31, PeppolBis_3_0_21

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TimeoutSeconds
Maximum business-rule process duration in seconds.

```yaml
Type: Int32
Parameter Sets: File
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
- `System.String`

## OUTPUTS

- `OfficeIMO.Invoicing.InvoiceModelValidationResult`
- `OfficeIMO.Invoicing.Validation.InvoiceValidationReport`

## RELATED LINKS

- None
