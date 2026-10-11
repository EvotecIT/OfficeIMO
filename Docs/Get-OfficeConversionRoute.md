---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeConversionRoute
## SYNOPSIS
Lists conversion contracts from the installed OfficeIMO catalog and their availability for PSWriteOffice batch PDF export.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeConversionRoute [-Id <string[]>] [-SourceExtension <string>] [-TargetExtension <string>] [-PdfBatch] [<CommonParameters>]
```

## DESCRIPTION
Each row retains the owning package, API, input kind, fidelity, support evidence, text-formatting contract and known limitations.
PdfBatchSupported describes Export-OfficeDocumentPdf in the current host. Other routes may have format-specific cmdlets or .NET APIs;
a catalog entry alone does not establish that a converter, credentials, fonts or optional providers are ready for a particular input.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeConversionRoute -SourceExtension docx -TargetExtension pdf | Format-List *
```

Extension filters accept a leading dot and compare without regard to case.

### EXAMPLE 2
```powershell
PS> Get-OfficeConversionRoute -PdfBatch | Select-Object Id, SourceExtensions, PackageId, Fidelity, KnownLimitations
```

Batch PDF export requires PowerShell 7.4 or newer. Windows PowerShell 5.1 can inspect the catalog and use supported single-file commands.

## PARAMETERS

### -Id
Exact route identifiers, for example docx-pdf or pdf-docx. Unknown identifiers return no rows.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -PdfBatch
Return only routes supported by Export-OfficeDocumentPdf batch parameter sets in the current host.

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

### -SourceExtension
Accepted source extension, with or without a leading dot.

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

### -TargetExtension
Destination extension, with or without a leading dot.

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

- `System.String[]`

## OUTPUTS

- `OfficeIMO.OfficeConversionCapability`

## RELATED LINKS

- None
