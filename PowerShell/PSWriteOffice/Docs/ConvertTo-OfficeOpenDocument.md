---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeOpenDocument
## SYNOPSIS
Converts Word, Excel, or PowerPoint content to native OpenDocument with fidelity evidence.

## SYNTAX
### Path (Default)
```powershell
ConvertTo-OfficeOpenDocument [-Path] <string> [-OutputPath] <string> [-WordOptions <WordOpenDocumentConversionOptions>] [-ExcelOptions <ExcelOpenDocumentConversionOptions>] [-PowerPointOptions <PowerPointOpenDocumentConversionOptions>] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Word
```powershell
ConvertTo-OfficeOpenDocument [-OutputPath] <string> -WordDocument <WordDocument> [-WordOptions <WordOpenDocumentConversionOptions>] [-ExcelOptions <ExcelOpenDocumentConversionOptions>] [-PowerPointOptions <PowerPointOpenDocumentConversionOptions>] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Excel
```powershell
ConvertTo-OfficeOpenDocument [-OutputPath] <string> -ExcelDocument <ExcelDocument> [-WordOptions <WordOpenDocumentConversionOptions>] [-ExcelOptions <ExcelOpenDocumentConversionOptions>] [-PowerPointOptions <PowerPointOpenDocumentConversionOptions>] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PowerPoint
```powershell
ConvertTo-OfficeOpenDocument [-OutputPath] <string> -PowerPointPresentation <PowerPointPresentation> [-WordOptions <WordOpenDocumentConversionOptions>] [-ExcelOptions <ExcelOpenDocumentConversionOptions>] [-PowerPointOptions <PowerPointOpenDocumentConversionOptions>] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Converts Word, Excel, or PowerPoint content to native OpenDocument with fidelity evidence.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeWordOpenDocumentOptions -IncludeImages -IncludeHeadersAndFooters
ConvertTo-OfficeOpenDocument -Path .\Report.docx -OutputPath .\Report.odt -WordOptions $options -FailOnLoss
```


## PARAMETERS

### -ExcelDocument
Open Excel workbook.

```yaml
Type: ExcelDocument
Parameter Sets: Excel
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ExcelOptions
Optional Excel-to-ODS conversion settings.

```yaml
Type: ExcelOpenDocumentConversionOptions
Parameter Sets: Path, Word, Excel, PowerPoint
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FailOnLoss
Throw when the conversion approximates, skips, or cannot map a feature.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Word, Excel, PowerPoint
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Destination ODT, ODS, or ODP path.

```yaml
Type: String
Parameter Sets: Path, Word, Excel, PowerPoint
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to a DOCX, XLSX, or PPTX file.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PowerPointOptions
Optional PowerPoint-to-ODP conversion settings.

```yaml
Type: PowerPointOpenDocumentConversionOptions
Parameter Sets: Path, Word, Excel, PowerPoint
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PowerPointPresentation
Open PowerPoint presentation.

```yaml
Type: PowerPointPresentation
Parameter Sets: PowerPoint
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -WordDocument
Open Word document.

```yaml
Type: WordDocument
Parameter Sets: Word
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -WordOptions
Optional Word-to-ODT conversion settings.

```yaml
Type: WordOpenDocumentConversionOptions
Parameter Sets: Path, Word, Excel, PowerPoint
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

- `OfficeIMO.Word.WordDocument`
- `OfficeIMO.Excel.ExcelDocument`
- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.OpenDocument.OdfConversionResult[OfficeIMO.OpenDocument.OdtDocument]`
- `OfficeIMO.OpenDocument.OdfConversionResult[OfficeIMO.OpenDocument.OdsDocument]`
- `OfficeIMO.OpenDocument.OdfConversionResult[OfficeIMO.OpenDocument.OdpPresentation]`

## RELATED LINKS

- None
