---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordList
## SYNOPSIS
Gets lists from a Word document or section.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordList [-Path] <string> [-IncludeEmpty] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordList -Document <WordDocument> [-IncludeEmpty] [<CommonParameters>]
```

### Section
```powershell
Get-OfficeWordList -Section <WordSection> [-IncludeEmpty] [<CommonParameters>]
```

## DESCRIPTION
Returns existing WordList objects so scripts can inspect list items, report on
checklist content, or pipe a selected list to Add-OfficeWordListItem. Use this command when
the script needs to work with list objects directly instead of using a text search first.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx
$list = $doc | Get-OfficeWordList | Select-Object -First 1
$list.ListItems | Select-Object -Property Text
```

Returns OfficeIMO list objects so existing list items can be reviewed or appended to.

### EXAMPLE 2
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx
Get-OfficeWordList -Document $doc |
    Where-Object { $_.ListItems.Text -contains 'Initial review' } |
    Add-OfficeWordListItem -Text 'Final approval'
$doc | Close-OfficeWord -Save
```

Inspects list objects directly when the caller wants custom selection logic.

## PARAMETERS

### -Document
Open document to inspect. The caller controls the document lifetime.

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

### -IncludeEmpty
Include numbering definitions that do not currently have list items.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document, Section
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to the document to open read-only for list inspection.

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

### -Section
Section to inspect when the caller only wants lists in a specific section.

```yaml
Type: WordSection
Parameter Sets: Section
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`
- `OfficeIMO.Word.WordSection`

## OUTPUTS

- `OfficeIMO.Word.WordList`

## RELATED LINKS

- None
