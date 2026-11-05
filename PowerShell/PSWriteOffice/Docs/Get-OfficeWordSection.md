---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeWordSection
## SYNOPSIS
Gets sections of a document.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeWordSection [-Index <int[]>] [-Document <WordDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeWordSection [-Path] <string[]> [-Index <int[]>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeWordSection -LiteralPath <string[]> [-Index <int[]>] [<CommonParameters>]
```

## DESCRIPTION
Gets sections of a document.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeWordSection -Document 'Value'
```


### EXAMPLE 2
```powershell
Get-OfficeWordSection -LiteralPath @('C:\Path')
```


### EXAMPLE 3
```powershell
Get-OfficeWordSection -Path @('C:\Path')
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

### -Index
Index for this inspection.

```yaml
Type: Int32[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordSection`

## RELATED LINKS

- None
