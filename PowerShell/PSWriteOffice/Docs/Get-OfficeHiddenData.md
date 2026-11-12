---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeHiddenData
## SYNOPSIS
Reports the metadata and hidden data that would leave with a file. The file is never changed.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeHiddenData [-Path] <string[]> [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeHiddenData -LiteralPath <string[]> [<CommonParameters>]
```

## DESCRIPTION
Finds author and company properties, comments, tracked changes, hidden text, rows, sheets and slides, speaker notes,
image metadata such as GPS locations, scripts, attachments, Content Credentials and AI labels. Inputs are Word, Excel,
PowerPoint, Visio, OpenDocument, EPUB, PDF, JPEG, PNG, WebP, GIF and TIFF files. Each finding has the Outcome Found.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> (Get-OfficeHiddenData -Path .\Proposal.docx).Findings
```


### EXAMPLE 2
```powershell
PS> Get-ChildItem .\Outgoing -File | Get-OfficeHiddenData |
                Where-Object { $_.Findings.Kind -contains 'Comments' -or $_.Findings.Kind -contains 'TrackedChanges' }
```


## PARAMETERS

### -LiteralPath
Exact file paths, including FileInfo.FullName and provider PSPath pipeline properties.

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

### -Path
Files to inspect; supports wildcards.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`

## OUTPUTS

- `PSWriteOffice.Models.OfficeHiddenDataFileReport`: Hidden data found in one file, what removal did with each item, and where the cleaned copy was written.

## RELATED LINKS

- None
