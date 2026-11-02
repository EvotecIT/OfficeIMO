---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeWordGoogleDocsOptions
## SYNOPSIS
Creates native Word-to-Google-Docs planning and export settings without contacting Google.

## SYNTAX
### Default (Default)
```powershell
New-OfficeWordGoogleDocsOptions [-Title <string>] [-FolderId <string>] [-DriveId <string>] [-ExistingFileId <string>] [-InlineImageMode <GoogleDocsInlineImageMode>] [-Comments <GoogleDocsCommentMode>] [-TabStrategy <GoogleDocsTabStrategy>] [-TabId <string>] [-ConflictMode <GoogleDocsRevisionConflictMode>] [-ExpectedRevisionId <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates native Word-to-Google-Docs planning and export settings without contacting Google.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeWordGoogleDocsOptions -ExistingFileId $fileId -TabStrategy SelectedTab -TabId $tabId -ExpectedRevisionId $revisionId
Export-OfficeWordGoogleDocument -Document $document -Session $session -Options $options
```


## PARAMETERS

### -Comments
How Word comments are transferred; omission preserves the native default.

```yaml
Type: GoogleDocsCommentMode
Parameter Sets: Default
Aliases: None
Possible values: Skip, UnanchoredDriveComments

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConflictMode
Revision conflict handling; omission preserves exact revision checks.

```yaml
Type: GoogleDocsRevisionConflictMode
Parameter Sets: Default
Aliases: None
Possible values: RequireRevision, MergeAgainstTargetRevision, OverwriteLatest

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DriveId
Shared-drive identifier; omission uses the native session location policy.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExistingFileId
Existing document to replace instead of creating a new document.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExpectedRevisionId
Revision captured by an earlier read, checked by the native exporter.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FolderId
Destination Drive folder identifier.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InlineImageMode
Inline image policy. Temporary public staging requires the native session's explicit authorization.

```yaml
Type: GoogleDocsInlineImageMode
Parameter Sets: Default
Aliases: None
Possible values: Placeholder, TemporaryPublicDriveLease

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TabId
Tab identifier for SelectedTab.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TabStrategy
First, selected or every tab when reading or replacing content.

```yaml
Type: GoogleDocsTabStrategy
Parameter Sets: Default
Aliases: None
Possible values: FirstTab, SelectedTab, ReplaceEveryTab

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Title of a newly created document; replacement retains the existing title.

```yaml
Type: String
Parameter Sets: Default
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

- `None`

## OUTPUTS

- `OfficeIMO.Word.GoogleDocs.GoogleDocsSaveOptions`

## RELATED LINKS

- None
