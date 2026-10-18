---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordStatistics
## SYNOPSIS
Gets document statistics from a Word document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordStatistics [-Path] <string> [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordStatistics -Document <WordDocument> [<CommonParameters>]
```

## DESCRIPTION
Returns a PowerShell-friendly snapshot of OfficeIMO.Word statistics for quick reporting and validation.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $stats = Get-OfficeWordStatistics -Path .\Report.docx
$stats |
    Select-Object -Property Paragraphs, Tables, Images, Charts |
    Format-List
```

Reads OfficeIMO.Word statistics and displays the structural counts that matter for a release artifact.

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

### -Path
Path to the Word document.

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

- `PSWriteOffice.Models.Word.WordDocumentStatisticsInfo`: PowerShell-friendly snapshot of Word document statistics.

## RELATED LINKS

- None
