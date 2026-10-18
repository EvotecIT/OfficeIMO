---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordDatePicker
## SYNOPSIS
Gets date picker content controls from a Word document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordDatePicker [-Path] <string> [-Alias <string[]>] [-Tag <string[]>] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordDatePicker -Document <WordDocument> [-Alias <string[]>] [-Tag <string[]>] [<CommonParameters>]
```

## DESCRIPTION
Gets date picker content controls from a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $dates = Get-OfficeWordDatePicker -Path .\IntakeForm.docx -Tag 'deadline-*'
$dates |
    Select-Object -Property Alias, Tag |
    Format-Table -AutoSize
```

Filters date picker content controls by tag before a form validation pass.

## PARAMETERS

### -Alias
Filter by date picker alias (wildcards supported).

```yaml
Type: String[]
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Word document to read.

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
Path to the .docx file.

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

### -Tag
Filter by date picker tag (wildcards supported).

```yaml
Type: String[]
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordDatePicker`

## RELATED LINKS

- None
