---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePdfSignature
## SYNOPSIS
Gets lightweight PDF signature structure and preservation validation.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePdfSignature [-Path] <string> [-Password <string>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

## DESCRIPTION
Reports signature fields, byte-range structure, DocMDP permissions, DSS/LTV evidence, and append-only preservation markers.
Certificate-chain trust, revocation, digest, and CMS cryptographic verification are intentionally not performed.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $report = Get-OfficePdfSignature -Path .\Signed.pdf
$report.Signatures
$report.Findings
```

Reads signature structure and reports whether OfficeIMO.Pdf found structural issues.

## PARAMETERS

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed usage restrictions.

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

### -Password
Password used to inspect a Standard password-encrypted PDF.

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

### -Path
PDF file path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.Pdf.PdfSignatureValidationReport`

## RELATED LINKS

- None
