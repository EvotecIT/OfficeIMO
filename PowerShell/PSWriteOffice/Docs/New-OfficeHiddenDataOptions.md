---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeHiddenDataOptions
## SYNOPSIS
Chooses which hidden data Remove-OfficeHiddenData and Optimize-OfficePdf -Sanitize remove.

## SYNTAX
### Default (Default)
```powershell
New-OfficeHiddenDataOptions [[-Kind] <OfficeHiddenDataKind[]>] [-TrackedChanges <OfficeTrackedChangesAction>] [-KeepProvenance <string[]>] [-AnonymousAuthor <string>] [<CommonParameters>]
```

## DESCRIPTION
With no parameters the options select the Default categories: document and custom properties, comments, personal
identities, speaker notes, the thumbnail, custom XML, image metadata, printer settings, external references, document
variables, revision identifiers and active content such as scripts. Tracked changes, hidden content, provenance,
attachments and bookmarks change what readers see or what the file discloses, so they are removed only when you name them.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeHiddenDataOptions -Kind Default, Provenance -TrackedChanges AcceptAll
Remove-OfficeHiddenData -Path .\Proposal.docx -Options $options
```


### EXAMPLE 2
```powershell
PS> $options = New-OfficeHiddenDataOptions -Kind Default, Provenance -KeepProvenance C2pa, External
Remove-OfficeHiddenData -Path .\Poster.pdf -Options $options
```


### EXAMPLE 3
```powershell
PS> Remove-OfficeHiddenData -Path .\Budget.xlsx -Options (New-OfficeHiddenDataOptions -Kind All)
```


## PARAMETERS

### -AnonymousAuthor
Name that replaces authors on comments and revisions that stay in the file. Defaults to Author.

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

### -KeepProvenance
Provenance carriers to keep when Provenance is selected: C2pa (Content Credentials), External (links to Content Credentials) or AiSource (AI-generated labels).

```yaml
Type: String[]
Parameter Sets: Default
Aliases: None
Possible values: C2pa, External, AiSource

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Kind
Categories to remove, combined. Defaults to Default; All selects every category.

```yaml
Type: OfficeHiddenDataKind[]
Parameter Sets: Default
Aliases: None
Possible values: None, DocumentProperties, CustomProperties, Comments, TrackedChanges, PersonalIdentities, HiddenContent, SpeakerNotes, Thumbnail, CustomXml, ImageMetadata, PrinterSettings, ExternalReferences, DocumentVariables, RevisionIdentifiers, Provenance, Attachments, ActiveContent, Default, Bookmarks, All

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TrackedChanges
Keep, accept or reject tracked changes. Defaults to Keep, which reports them and leaves them in place.

```yaml
Type: OfficeTrackedChangesAction
Parameter Sets: Default
Aliases: None
Possible values: Keep, AcceptAll, RejectAll

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

- `OfficeIMO.OfficeHiddenDataOptions`

## RELATED LINKS

- None
