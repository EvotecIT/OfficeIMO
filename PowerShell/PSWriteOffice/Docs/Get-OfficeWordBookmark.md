---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordBookmark
## SYNOPSIS
Gets bookmarks from a Word document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordBookmark [-Path] <string> [-Name <string[]>] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordBookmark -Document <WordDocument> [-Name <string[]>] [<CommonParameters>]
```

## DESCRIPTION
Returns WordBookmark objects, optionally filtered by name.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $bookmarks = Get-OfficeWordBookmark -Path .\Report.docx -Name 'Appendix*', 'ExecutiveSummary'
$bookmarks |
    Select-Object -Property Name |
    Format-Table -AutoSize
```

Loads the document read-only, filters bookmark names with wildcards, and displays the matching navigation targets.

## PARAMETERS

### -Document
Word document to read.

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

### -Name
Bookmark name filter (wildcards supported).

```yaml
Type: String[]
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Path to the .docx file.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordBookmark`

## RELATED LINKS

- None
