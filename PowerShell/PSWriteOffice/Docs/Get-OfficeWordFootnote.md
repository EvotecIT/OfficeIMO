---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeWordFootnote
## SYNOPSIS
Gets footnotes of a document or section.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeWordFootnote [-Document <WordDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeWordFootnote [-Path] <string[]> [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeWordFootnote -LiteralPath <string[]> [<CommonParameters>]
```

### Section
```powershell
Get-OfficeWordFootnote -Section <WordSection> [<CommonParameters>]
```

## DESCRIPTION
Gets footnotes of a document or section.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeWordFootnote -Document 'Value'
```


### EXAMPLE 2
```powershell
Get-OfficeWordFootnote -LiteralPath @('C:\Path')
```


### EXAMPLE 3
```powershell
Get-OfficeWordFootnote -Path @('C:\Path')
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

## OUTPUTS

- `PSWriteOffice.Services.Word.WordNoteInfo`: Describes a Word footnote or endnote in a document-safe snapshot.

## RELATED LINKS

- None
