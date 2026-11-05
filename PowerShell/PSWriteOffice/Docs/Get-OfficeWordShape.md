---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeWordShape
## SYNOPSIS
Gets shapes from a document, section or paragraph.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeWordShape [-Document <WordDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeWordShape [-Path] <string[]> [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeWordShape -LiteralPath <string[]> [<CommonParameters>]
```

### Section
```powershell
Get-OfficeWordShape -Section <WordSection> [<CommonParameters>]
```

### Paragraph
```powershell
Get-OfficeWordShape -Paragraph <WordParagraph> [<CommonParameters>]
```

## DESCRIPTION
Gets shapes from a document, section or paragraph.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeWordShape -Document 'Value'
```


### EXAMPLE 2
```powershell
Get-OfficeWordShape -LiteralPath @('C:\Path')
```


### EXAMPLE 3
```powershell
Get-OfficeWordShape -Paragraph 'Value'
```


## PARAMETERS

### -Document
Document for this inspection.

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
LiteralPath for this inspection.

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

### -Paragraph
Paragraph for this inspection.

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
Path for this inspection.

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

### -Section
Section for this inspection.

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

- `System.String[]`
- `OfficeIMO.Word.WordDocument`
- `OfficeIMO.Word.WordSection`
- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `OfficeIMO.Word.WordShape`

## RELATED LINKS

- None
