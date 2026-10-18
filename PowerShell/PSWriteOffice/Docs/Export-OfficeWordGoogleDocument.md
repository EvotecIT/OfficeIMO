---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeWordGoogleDocument
## SYNOPSIS
Plans, compiles, or exports a Word document to Google Docs.

## SYNTAX
### Path (Default)
```powershell
Export-OfficeWordGoogleDocument [-Path] <string> [-Options <GoogleDocsSaveOptions>] [-Session <GoogleWorkspaceSession>] [-PlanOnly] [-AsBatch] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Export-OfficeWordGoogleDocument -Document <WordDocument> [-Options <GoogleDocsSaveOptions>] [-Session <GoogleWorkspaceSession>] [-PlanOnly] [-AsBatch] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Plans, compiles, or exports a Word document to Google Docs.

## EXAMPLES

### EXAMPLE 1
```powershell
Export-OfficeWordGoogleDocument -Path 'C:\Path'
```


### EXAMPLE 2
```powershell
Export-OfficeWordGoogleDocument -Document 'Value'
```


## PARAMETERS

### -AsBatch
Return the provider-neutral request batch without contacting Google.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Word document to export.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FailOnLoss
Throw when translation reports a warning or error.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Google Docs translation and destination settings.

```yaml
Type: GoogleDocsSaveOptions
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to a Word document.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PlanOnly
Return the translation plan without compiling requests or contacting Google.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Session
Configured Google Workspace session used for a live export.

```yaml
Type: GoogleWorkspaceSession
Parameter Sets: Path, Document
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

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.GoogleDocs.GoogleDocsTranslationPlan`
- `OfficeIMO.Word.GoogleDocs.GoogleDocsBatch`
- `OfficeIMO.Word.GoogleDocs.GoogleDocumentReference`

## RELATED LINKS

- None
