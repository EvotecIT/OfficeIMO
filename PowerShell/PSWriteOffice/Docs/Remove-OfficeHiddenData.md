---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Remove-OfficeHiddenData
## SYNOPSIS
Writes a copy of a file without its metadata and hidden data, and reports what happened to each item.

## SYNTAX
### Path (Default)
```powershell
Remove-OfficeHiddenData [-Path] <string[]> [-OutputPath <string>] [-Options <OfficeHiddenDataOptions>] [-RemoveInvalidatedSignatures] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathDirectory
```powershell
Remove-OfficeHiddenData [-Path] <string[]> -OutputDirectory <string> [-Options <OfficeHiddenDataOptions>] [-RemoveInvalidatedSignatures] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Remove-OfficeHiddenData -LiteralPath <string[]> [-OutputPath <string>] [-Options <OfficeHiddenDataOptions>] [-RemoveInvalidatedSignatures] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathDirectory
```powershell
Remove-OfficeHiddenData -LiteralPath <string[]> -OutputDirectory <string> [-Options <OfficeHiddenDataOptions>] [-RemoveInvalidatedSignatures] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
The input is never changed. The cleaned copy is written as <name>.cleaned<ext> beside the input unless
OutputPath or OutputDirectory says otherwise, and an existing file is replaced only with Force.

Without Options the Default categories are removed: document and custom properties, comments, personal identities,
speaker notes, the thumbnail, custom XML, image metadata, printer settings, external references, document variables,
revision identifiers and active content such as scripts. Tracked changes, hidden content, provenance, attachments and
bookmarks are kept until New-OfficeHiddenDataOptions selects them. Every finding reports an Outcome; a Retained finding
carries a Reason and a Note that say why it stayed.

Saving refuses to invalidate an Office package signature unless RemoveInvalidatedSignatures is present. Signed PDFs
are always refused.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Remove-OfficeHiddenData -Path .\Proposal.docx
```


### EXAMPLE 2
```powershell
PS> $options = New-OfficeHiddenDataOptions -Kind Default, Provenance
Get-ChildItem .\Outgoing -File | Remove-OfficeHiddenData -OutputDirectory .\Share -Options $options
```


### EXAMPLE 3
```powershell
PS> $options = New-OfficeHiddenDataOptions -TrackedChanges AcceptAll
Remove-OfficeHiddenData -Path .\Contract.docx -OutputPath .\Contract-final.docx -Options $options -Force
```


## PARAMETERS

### -Force
Replace an existing output file.

```yaml
Type: SwitchParameter
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Exact file paths, including FileInfo.FullName and provider PSPath pipeline properties.

```yaml
Type: String[]
Parameter Sets: LiteralPath, LiteralPathDirectory
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Options
Categories to remove and how to treat tracked changes and provenance, from New-OfficeHiddenDataOptions. Defaults to the Default categories.

```yaml
Type: OfficeHiddenDataOptions
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
Folder that receives each cleaned copy under its original file name; created when missing.

```yaml
Type: String
Parameter Sets: PathDirectory, LiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
The cleaned copy for a single input. It must keep the input's extension.

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

### -Path
Files to clean; supports wildcards.

```yaml
Type: String[]
Parameter Sets: Path, PathDirectory
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -RemoveInvalidatedSignatures
Allow cleaning to remove Office package signatures that it invalidates; otherwise a signed package is refused.

```yaml
Type: SwitchParameter
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
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

- `PSWriteOffice.Models.OfficeHiddenDataFileReport`: Hidden data found in one file, what removal did with each item, and where the cleaned copy was written.

## RELATED LINKS

- None
