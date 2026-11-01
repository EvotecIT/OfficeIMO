---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficeWordGoogleDocument
## SYNOPSIS
Plans, compiles or exports Word documents to Google Docs.

## SYNTAX
### Document (Default)
```powershell
Export-OfficeWordGoogleDocument [-Options <Object[]>] [-Session <GoogleWorkspaceSession>] [-Mode <GoogleDocsExportMode>] [-FailOnLoss] [-Document <WordDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Export-OfficeWordGoogleDocument [-Path] <string[]> [-Options <Object[]>] [-Session <GoogleWorkspaceSession>] [-Mode <GoogleDocsExportMode>] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Export-OfficeWordGoogleDocument -LiteralPath <string[]> [-Options <Object[]>] [-Session <GoogleWorkspaceSession>] [-Mode <GoogleDocsExportMode>] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Plans, compiles or exports Word documents to Google Docs.

## EXAMPLES

### EXAMPLE 1
```powershell
Export-OfficeWordGoogleDocument -Path .\Report.docx -Mode Plan
```


## PARAMETERS

### -Document
Caller-owned document; omission uses the enclosing content block.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FailOnLoss
Reject translation warnings or errors before returning a result or mutating Google.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal source paths; FileInfo.FullName and provider PSPath bind here.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Mode
Export by default, or return an offline Plan or Batch.

```yaml
Type: GoogleDocsExportMode
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values: Export, Plan, Batch

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
One native options object and property hashtables applied to an independent snapshot.

```yaml
Type: Object[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Source document paths, including filesystem wildcard patterns.

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

### -Session
Configured Google Workspace session, required for Export mode.

```yaml
Type: GoogleWorkspaceSession
Parameter Sets: Document, Path, LiteralPath
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
- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.GoogleDocs.GoogleDocsTranslationPlan`
- `OfficeIMO.Word.GoogleDocs.GoogleDocsBatch`
- `OfficeIMO.Word.GoogleDocs.GoogleDocumentReference`

## RELATED LINKS

- None
