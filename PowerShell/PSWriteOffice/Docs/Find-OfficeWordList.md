---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Find-OfficeWordList
## SYNOPSIS
Finds Word lists containing matching list-item text.

## SYNTAX
### PathText (Default)
```powershell
Find-OfficeWordList [-Path] <string> [-Text] <string> [-CaseSensitive] [<CommonParameters>]
```

### PathRegex
```powershell
Find-OfficeWordList [-Path] <string> [-Pattern] <string> [-CaseSensitive] [<CommonParameters>]
```

### DocumentText
```powershell
Find-OfficeWordList [-Text] <string> -Document <WordDocument> [-CaseSensitive] [<CommonParameters>]
```

### DocumentRegex
```powershell
Find-OfficeWordList [-Pattern] <string> -Document <WordDocument> [-CaseSensitive] [<CommonParameters>]
```

### SectionText
```powershell
Find-OfficeWordList [-Text] <string> -Section <WordSection> [-CaseSensitive] [<CommonParameters>]
```

### SectionRegex
```powershell
Find-OfficeWordList [-Pattern] <string> -Section <WordSection> [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Searches list items in a Word document or section and returns matching WordList
objects. This is useful when a document has an existing checklist or numbered list and the script
needs to append to that list rather than create a new one elsewhere in the document.

Use -Text for literal contains matching or -Pattern for regular expressions. The
returned list objects can be piped directly to Add-OfficeWordListItem.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx
Find-OfficeWordList -Document $doc -Text 'Initial review' |
    Add-OfficeWordListItem -Text 'Final approval'
```

Searches list-item paragraphs and returns matching OfficeIMO list objects for further editing.

### EXAMPLE 2
```powershell
PS> $doc = Get-OfficeWord -Path .\Handover.docx
$list = Find-OfficeWordList -Document $doc -Text 'Initial review' | Select-Object -First 1
$list | Add-OfficeWordListItem -Text 'Business sign-off'
$list | Add-OfficeWordListItem -Text 'Go-live approval'
$doc | Close-OfficeWord -Save
```

Finds a checklist by an item it already contains and appends new items to the same list.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching.

```yaml
Type: SwitchParameter
Parameter Sets: PathText, PathRegex, DocumentText, DocumentRegex, SectionText, SectionRegex
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Open document to inspect. The caller controls the document lifetime.

```yaml
Type: WordDocument
Parameter Sets: DocumentText, DocumentRegex
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Path
Path to the document to open read-only for searching.

```yaml
Type: String
Parameter Sets: PathText, PathRegex
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Pattern
Regular expression pattern to find in list items.

```yaml
Type: String
Parameter Sets: PathRegex, DocumentRegex, SectionRegex
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Section
Section to inspect when the caller only wants lists in a specific section.

```yaml
Type: WordSection
Parameter Sets: SectionText, SectionRegex
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Text
Literal text to find in list items.

```yaml
Type: String
Parameter Sets: PathText, DocumentText, SectionText
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
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
