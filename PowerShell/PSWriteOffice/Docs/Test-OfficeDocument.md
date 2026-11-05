---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Test-OfficeDocument
## SYNOPSIS
Validates native document structure, schema, security, accessibility or compliance without changing the input.

## SYNTAX
### Path (Default)
```powershell
Test-OfficeDocument [-Path] <string[]> [-Check <OfficeDocumentCheck[]>] [-Untrusted] [-Options <Object[]>] [-Password <securestring>] [-Quiet] [<CommonParameters>]
```

### LiteralPath
```powershell
Test-OfficeDocument -LiteralPath <string[]> [-Check <OfficeDocumentCheck[]>] [-Untrusted] [-Options <Object[]>] [-Password <securestring>] [-Quiet] [<CommonParameters>]
```

### Document
```powershell
Test-OfficeDocument [-Document] <Object> [-Check <OfficeDocumentCheck[]>] [-Untrusted] [-Options <Object[]>] [-Quiet] [<CommonParameters>]
```

## DESCRIPTION
Unsupported checks are reported as not applicable. NativeReports retains native security, capability and PDF proof evidence. Options accepts typed check settings or hashtables; later values replace the matching settings group.

## EXAMPLES

### EXAMPLE 1
```powershell
Test-OfficeDocument -Path .\Reports\Sales.xlsx
```

Returns the common report and the native workbook doctor findings.

### EXAMPLE 2
```powershell
Test-OfficeDocument -LiteralPath .\Uploads\Proposal.docx -Check Security -Untrusted -Quiet
```

Returns a Boolean after applying the untrusted package policy.

## PARAMETERS

### -Check
Checks to perform. Defaults to Structure.

```yaml
Type: OfficeDocumentCheck[]
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values: Structure, Schema, Security, Accessibility, Compliance

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned live native model, including unsaved edits.

```yaml
Type: Object
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
Literal filesystem paths; binds FileInfo.FullName and provider PSPath.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Options
Native check-specific options or flat hashtables, routed to the selected format through the shared binder. For PDF compliance checks, a flat Profile means the compliance profile; use typed PdfLoadOptions for a separate read profile.

```yaml
Type: Object[]
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password for a path input. Accepts SecureString or PSCredential; plain strings are transformed with a warning.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Filesystem paths with wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Quiet
Returns IsValid as a Boolean instead of the report.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Untrusted
Also applies the existing untrusted package-security policy.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
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
- `System.Object`

## OUTPUTS

- `OfficeIMO.OfficeValidationReport`
- `System.Boolean`

## RELATED LINKS

- None
