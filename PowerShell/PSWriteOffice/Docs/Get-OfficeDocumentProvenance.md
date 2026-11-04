---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeDocumentProvenance
## SYNOPSIS
Collects structural, text-integrity, and optional provider verification evidence for a file.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeDocumentProvenance [-Path] <string[]> [-Options <Object[]>] [-C2paToolPath <string>] [-SignalDetector <IOfficeProvenanceSignalDetector[]>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeDocumentProvenance -LiteralPath <string[]> [-Options <Object[]>] [-C2paToolPath <string>] [-SignalDetector <IOfficeProvenanceSignalDetector[]>] [<CommonParameters>]
```

## DESCRIPTION
The result keeps carrier discovery, cryptographic verification, and provider signals separate; it does not infer authorship from their presence or absence.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $evidence = Get-OfficeDocumentProvenance -Path .\Published\cover.png
$evidence.Structural | Select-Object Format, HasC2paManifest
$evidence.TextIntegrity
```


### EXAMPLE 2
```powershell
PS> $evidence = Get-OfficeDocumentProvenance -Path .\Published\cover.png -C2paToolPath C:\Tools\c2patool.exe
$evidence.Verification | Select-Object Status, ProviderName, Findings
```

Network access remains disabled unless enabled in the supplied assessment options.

## PARAMETERS

### -C2paToolPath
Explicit c2patool executable path or controlled command name used for cryptographic verification.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal paths, including FileInfo.FullName and provider PSPath pipeline properties.

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
Typed assessment policies or hashtables with Structural, TextIntegrity and Verification groups. The last policy takes precedence.

```yaml
Type: Object[]
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
Asset or document paths, including provider wildcards.

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

### -SignalDetector
Optional provider-specific watermark or disclosure detectors.

```yaml
Type: IOfficeProvenanceSignalDetector[]
Parameter Sets: Path, LiteralPath
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

- `OfficeIMO.Provenance.OfficeProvenanceAssessmentReport`

## RELATED LINKS

- None
