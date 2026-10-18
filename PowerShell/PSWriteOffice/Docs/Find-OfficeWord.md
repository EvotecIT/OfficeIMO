---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Find-OfficeWord
## SYNOPSIS
Finds text matches inside a Word document.

## SYNTAX
### PathText (Default)
```powershell
Find-OfficeWord [-Path] <string> [-Text] <string> [-CaseSensitive] [<CommonParameters>]
```

### PathRegex
```powershell
Find-OfficeWord [-Path] <string> [-Pattern] <string> [-CaseSensitive] [-AsResult] [<CommonParameters>]
```

### DocumentText
```powershell
Find-OfficeWord [-Text] <string> -Document <WordDocument> [-CaseSensitive] [<CommonParameters>]
```

### DocumentRegex
```powershell
Find-OfficeWord [-Pattern] <string> -Document <WordDocument> [-CaseSensitive] [-AsResult] [<CommonParameters>]
```

## DESCRIPTION
Returns matching paragraphs or a WordSearchResult when using regex with -AsResult.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $matches = Find-OfficeWord -Path .\Report.docx -Text 'Quarter'
$matches |
    Select-Object -Property Text |
    Export-Csv -Path .\ReviewMatches.csv -NoTypeInformation
```

Finds paragraphs containing the search text and writes a lightweight review CSV.

### EXAMPLE 2
```powershell
PS> $result = Find-OfficeWord -Path .\Report.docx -Pattern 'FY\d{4}' -AsResult
$result.Paragraphs |
    Select-Object -Property Text
```

Returns the full regex result so paragraphs, table matches, headers, and footers can be inspected separately.

## PARAMETERS

### -AsResult
Emit the full WordSearchResult result for regex searches.

```yaml
Type: SwitchParameter
Parameter Sets: PathRegex, DocumentRegex
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CaseSensitive
Use case-sensitive matching.

```yaml
Type: SwitchParameter
Parameter Sets: PathText, PathRegex, DocumentText, DocumentRegex
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Word document to search.

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
Path to the .docx file.

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
Regular expression pattern to find.

```yaml
Type: String
Parameter Sets: PathRegex, DocumentRegex
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text to find.

```yaml
Type: String
Parameter Sets: PathText, DocumentText
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

## OUTPUTS

- `OfficeIMO.Word.WordParagraph`
- `OfficeIMO.Word.WordSearchResult`

## RELATED LINKS

- None
