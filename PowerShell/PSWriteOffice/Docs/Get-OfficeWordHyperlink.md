---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeWordHyperlink
## SYNOPSIS
Gets hyperlinks, filtered by text, URL or bookmark anchor.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeWordHyperlink [-Text <string[]>] [-Url <string[]>] [-Anchor <string[]>] [-Document <WordDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeWordHyperlink [-Path] <string[]> [-Text <string[]>] [-Url <string[]>] [-Anchor <string[]>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeWordHyperlink -LiteralPath <string[]> [-Text <string[]>] [-Url <string[]>] [-Anchor <string[]>] [<CommonParameters>]
```

### Section
```powershell
Get-OfficeWordHyperlink -Section <WordSection> [-Text <string[]>] [-Url <string[]>] [-Anchor <string[]>] [<CommonParameters>]
```

### Paragraph
```powershell
Get-OfficeWordHyperlink -Paragraph <WordParagraph> [-Text <string[]>] [-Url <string[]>] [-Anchor <string[]>] [<CommonParameters>]
```

## DESCRIPTION
Gets hyperlinks, filtered by text, URL or bookmark anchor.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeWordHyperlink -Anchor @('Value')
```


### EXAMPLE 2
```powershell
Get-OfficeWordHyperlink -LiteralPath @('C:\Path')
```


### EXAMPLE 3
```powershell
Get-OfficeWordHyperlink -Paragraph 'Value'
```


## PARAMETERS

### -Anchor
Anchor for this inspection.

```yaml
Type: String[]
Parameter Sets: Document, Path, LiteralPath, Section, Paragraph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

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

### -Text
Text for this inspection.

```yaml
Type: String[]
Parameter Sets: Document, Path, LiteralPath, Section, Paragraph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Url
Url for this inspection.

```yaml
Type: String[]
Parameter Sets: Document, Path, LiteralPath, Section, Paragraph
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
- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `OfficeIMO.Word.WordHyperLink`

## RELATED LINKS

- None
