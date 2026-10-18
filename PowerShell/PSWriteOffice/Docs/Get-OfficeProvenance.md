---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeProvenance
## SYNOPSIS
Collects structural, text-integrity, and optional provider verification evidence for a file.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeProvenance [-Path] <string> [-Options <OfficeProvenanceAssessmentOptions>] [-C2paToolPath <string>] [-SignalDetector <IOfficeProvenanceSignalDetector[]>] [<CommonParameters>]
```

## DESCRIPTION
The result keeps carrier discovery, cryptographic verification, and provider signals separate; it does not infer authorship from their presence or absence.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $evidence = Get-OfficeProvenance -Path .\Published\cover.png
$evidence.Structural | Select-Object Format, HasC2paManifest
$evidence.TextIntegrity
```


### EXAMPLE 2
```powershell
PS> $evidence = Get-OfficeProvenance -Path .\Published\cover.png -C2paToolPath C:\Tools\c2patool.exe
$evidence.Verification | Select-Object Status, ProviderName, Findings
```

Network access remains disabled unless enabled in the supplied assessment options.

## PARAMETERS

### -C2paToolPath
Explicit c2patool executable path or controlled command name used for cryptographic verification.

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

### -Options
Structural, text-integrity, and provider verification limits and policy.

```yaml
Type: OfficeProvenanceAssessmentOptions
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
Asset or document path to inspect.

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

### -SignalDetector
Optional provider-specific watermark or disclosure detectors.

```yaml
Type: IOfficeProvenanceSignalDetector[]
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Provenance.OfficeProvenanceAssessmentReport`

## RELATED LINKS

- None
