---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordFootnote
## SYNOPSIS
Gets footnotes from a Word document or section.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordFootnote [-Path] <string> [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordFootnote -Document <WordDocument> [<CommonParameters>]
```

### Section
```powershell
Get-OfficeWordFootnote -Section <WordSection> [<CommonParameters>]
```

## DESCRIPTION
Gets footnotes from a Word document or section.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $footnotes = Get-OfficeWordFootnote -Path .\PolicyReport.docx
$footnotes |
    Select-Object -Property Kind, ReferenceId, ParentText, Text |
    Export-Csv -Path .\Footnotes.csv -NoTypeInformation
```

Reads footnotes from the document and exports the PowerShell-friendly note snapshot.

## PARAMETERS

### -Document
Document to inspect.

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

### -Path
Path to the document.

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
Section to inspect.

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

- `PSWriteOffice.Services.Word.WordNoteInfo`: Describes a Word footnote or endnote in a document-safe snapshot.

## RELATED LINKS

- None
