---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePdfCompliance
## SYNOPSIS
Gets a generated PDF document compliance readiness report.

## SYNTAX
### Document (Default)
```powershell
Get-OfficePdfCompliance [-Document <PdfDocument>] [-Profile <PdfComplianceProfile>] [-Proof] [-ExternalValidator <PdfExternalValidatorKind[]>] [-ExternalValidation <PdfExternalValidationResult[]>] [-ExternalStatus <PdfExternalValidationStatus>] [-ExternalProfile <string>] [-ExternalDiagnostic <string>] [-ExternalValidatorName <string>] [-ExternalValidatorVersion <string>] [-ExternalExitCode <Int32>] [-ExternalSuccessExitCode <int>] [-ExternalExecutablePath <string>] [-ExternalArguments <string>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficePdfCompliance [-Path] <string> -Profile <PdfComplianceProfile> [-Password <string>] [-IgnorePermissionRestrictions] [-Proof] [-ExternalValidator <PdfExternalValidatorKind[]>] [-ExternalValidation <PdfExternalValidationResult[]>] [-ExternalStatus <PdfExternalValidationStatus>] [-ExternalProfile <string>] [-ExternalDiagnostic <string>] [-ExternalValidatorName <string>] [-ExternalValidatorVersion <string>] [-ExternalExitCode <Int32>] [-ExternalSuccessExitCode <int>] [-ExternalExecutablePath <string>] [-ExternalArguments <string>] [<CommonParameters>]
```

## DESCRIPTION
Gets a generated PDF document compliance readiness report.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $pdf = New-OfficePdf {
    Set-OfficePdfCompliance -Profile PdfA3B -Groundwork
    Add-OfficePdfHeading -Text 'Compliance readiness'
} -NoSave
$pdf | Get-OfficePdfCompliance -Profile PdfA3B
```

Returns the OfficeIMO.Pdf readiness report before saving.

## PARAMETERS

### -Document
Generated PDF document to assess outside the DSL context.

```yaml
Type: PdfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ExternalArguments
External validator command-line arguments recorded in the proof evidence.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExternalDiagnostic
Human-readable external validation diagnostic.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExternalExecutablePath
External validator executable path recorded in the proof evidence.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExternalExitCode
External validator process exit code. When provided, status is inferred from -ExternalSuccessExitCode.

```yaml
Type: Int32
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExternalProfile
Profile string reported by the external validator, for example PDF/A-3b.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExternalStatus
Unbound external validator status to attach when -ExternalValidator is provided.

```yaml
Type: PdfExternalValidationStatus
Parameter Sets: Document, Path
Aliases: None
Possible values: NotRun, Passed, Failed, Error

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExternalSuccessExitCode
External validator process exit code that means success.

```yaml
Type: Int32
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExternalValidation
Artifact-bound results produced by the external validation lane.
Use PdfExternalValidationResult.PassedForArtifact or FromExitCodeForArtifact with the exact validated bytes.

```yaml
Type: PdfExternalValidationResult[]
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExternalValidator
External validator families whose result should be attached to the proof report.

```yaml
Type: PdfExternalValidatorKind[]
Parameter Sets: Document, Path
Aliases: None
Possible values: VeraPdf, PdfUaValidator, Mustang, Custom, PdfXValidator

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExternalValidatorName
Human-readable external validator name.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExternalValidatorVersion
External validator version recorded in the artifact-bound proof evidence.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed usage restrictions.

```yaml
Type: SwitchParameter
Parameter Sets: Path
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
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Existing PDF file path to assess after generation.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Profile
Compliance profile to assess. When omitted, the document's configured profile is used.

```yaml
Type: PdfComplianceProfile
Parameter Sets: Document, Path
Aliases: None
Possible values: None, PdfA2B, PdfA2U, PdfA2A, PdfA3B, PdfA3U, PdfA3A, PdfUa1, FacturX, Zugferd, PdfA4, PdfA4E, PdfA4F, PdfUa2, PdfX1A2003, PdfX4

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Proof
Return a proof report that combines readiness with required external validator evidence placeholders.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path
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

- `OfficeIMO.Pdf.PdfDocument`
- `System.String`

## OUTPUTS

- `OfficeIMO.Pdf.PdfComplianceReadinessReport`
- `OfficeIMO.Pdf.PdfComplianceProofReport`

## RELATED LINKS

- None
