---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeWordParagraph
## SYNOPSIS
Gets paragraphs from a document or section, optionally only those matching a pattern.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeWordParagraph [-Pattern <string>] [-SimpleMatch] [-CaseSensitive] [-AsResult] [-Document <WordDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeWordParagraph [-Path] <string[]> [-Pattern <string>] [-SimpleMatch] [-CaseSensitive] [-AsResult] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeWordParagraph -LiteralPath <string[]> [-Pattern <string>] [-SimpleMatch] [-CaseSensitive] [-AsResult] [<CommonParameters>]
```

### Section
```powershell
Get-OfficeWordParagraph -Section <WordSection> [-Pattern <string>] [-SimpleMatch] [-CaseSensitive] [-AsResult] [<CommonParameters>]
```

## DESCRIPTION
Gets paragraphs from a document or section, optionally only those matching a pattern.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeWordParagraph -AsResult
```


### EXAMPLE 2
```powershell
Get-OfficeWordParagraph -LiteralPath @('C:\Path')
```


### EXAMPLE 3
```powershell
Get-OfficeWordParagraph -Path @('C:\Path')
```


## PARAMETERS

### -AsResult
Emit native grouped match details; requires Pattern.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath, Section
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
Parameter Sets: Document, Path, LiteralPath, Section
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned document; omission uses the enclosing Word content block.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
Exact document paths, including literal brackets.

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
Document paths; supports wildcards.

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

### -Pattern
Optional regular expression, or literal text with SimpleMatch.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath, Section
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Section
Caller-owned section to inspect.

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

### -SimpleMatch
Treat Pattern as literal text.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath, Section
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
- `OfficeIMO.Word.WordDocument`
- `OfficeIMO.Word.WordSection`

## OUTPUTS

- `OfficeIMO.Word.WordParagraph`
- `OfficeIMO.Word.WordSearchResult`

## RELATED LINKS

- None
