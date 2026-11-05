---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeWordList
## SYNOPSIS
Gets lists, optionally only those with an item matching a pattern.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeWordList [-IncludeEmpty] [-Pattern <string>] [-SimpleMatch] [-CaseSensitive] [-Document <WordDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeWordList [-Path] <string[]> [-IncludeEmpty] [-Pattern <string>] [-SimpleMatch] [-CaseSensitive] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeWordList -LiteralPath <string[]> [-IncludeEmpty] [-Pattern <string>] [-SimpleMatch] [-CaseSensitive] [<CommonParameters>]
```

### Section
```powershell
Get-OfficeWordList -Section <WordSection> [-IncludeEmpty] [-Pattern <string>] [-SimpleMatch] [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Gets lists, optionally only those with an item matching a pattern.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeWordList -CaseSensitive
```


### EXAMPLE 2
```powershell
Get-OfficeWordList -LiteralPath @('C:\Path')
```


### EXAMPLE 3
```powershell
Get-OfficeWordList -Path @('C:\Path')
```


## PARAMETERS

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

### -IncludeEmpty
Include list definitions with no items.

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

- `OfficeIMO.Word.WordList`

## RELATED LINKS

- None
