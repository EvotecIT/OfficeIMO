---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordImage
## SYNOPSIS
Gets images from a Word document, section, or paragraph.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordImage [-Path] <string> [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordImage -Document <WordDocument> [<CommonParameters>]
```

### Section
```powershell
Get-OfficeWordImage -Section <WordSection> [<CommonParameters>]
```

### Paragraph
```powershell
Get-OfficeWordImage -Paragraph <WordParagraph> [<CommonParameters>]
```

## DESCRIPTION
Gets images from a Word document, section, or paragraph.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $images = Get-OfficeWordImage -Path .\Report.docx
$images |
    Select-Object -Property Title, Description, Width, Height |
    Format-Table -AutoSize
```

Reads images from a document so metadata and sizing can be reviewed before publication.

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

### -Paragraph
Paragraph to inspect.

```yaml
Type: WordParagraph
Parameter Sets: Paragraph
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
- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `OfficeIMO.Word.WordImage`

## RELATED LINKS

- None
