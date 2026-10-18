---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordSection
## SYNOPSIS
Gets sections from a Word document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordSection [-Path] <string> [-Index <int[]>] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordSection -Document <WordDocument> [-Index <int[]>] [<CommonParameters>]
```

## DESCRIPTION
Gets sections from a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $section = Get-OfficeWordSection -Path .\Report.docx -Index 0
$section |
    Get-OfficeWordParagraph |
    Select-Object -First 5 -Property Text
```

Gets a specific section by zero-based index and pipes it to another thin Word reader.

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

### -Index
Optional 0-based section index filter.

```yaml
Type: Int32[]
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordSection`

## RELATED LINKS

- None
