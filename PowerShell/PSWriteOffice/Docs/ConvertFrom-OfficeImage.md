---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertFrom-OfficeImage
## SYNOPSIS
Exports a recognized image as editable Excel, Word, text, Markdown, HTML, JSON, CSV or searchable PDF.

## SYNTAX
### __AllParameterSets
```powershell
ConvertFrom-OfficeImage [-InputObject] <PdfSearchableOcrReview> [-OutputPath] <string> [-FirstRowIsHeader] [-ExcelOptions <PdfTablesToExcelOptions>] [-WordOptions <PdfToWordOptions>] [-HtmlOptions <PdfToHtmlOptions>] [-CsvOptions <CsvSaveOptions>] [-Force] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Recognize once with Get-OfficeImageDocument, inspect its evidence, then reuse that snapshot for each output.
Excel and CSV contain detected tables only. CSV requires exactly one table. Empty recognition or missing tables fail before output publication.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $image = Get-OfficeImageDocument -Path .\Ledger.png
$image | ConvertFrom-OfficeImage -OutputPath .\Ledger.xlsx -FirstRowIsHeader -PassThru
$image | ConvertFrom-OfficeImage -OutputPath .\Ledger.docx
```

Excel receives caller-confirmed headers and typed columns. PassThru returns recognition evidence and the format conversion report.

### EXAMPLE 2
```powershell
PS> $csv = @{ Delimiter = ';'; FormulaInjectionPolicy = 'Escape' }
$image | ConvertFrom-OfficeImage -OutputPath .\Ledger.csv -CsvOptions $csv
```

Set Escape explicitly when customizing CSV formatting for spreadsheet consumption. Every recognized table row is written as data without a synthesized header, regardless of IncludeHeader.

## PARAMETERS

### -CsvOptions
CSV field formatting and formula policy. When omitted, formula-like source text is escaped, including negative numeric text. Supplied options retain their FormulaInjectionPolicy; a new CsvSaveOptions object defaults to Preserve. Set Escape explicitly when customizing options for spreadsheet consumption.

Every recognized table row is written as data without a synthesized header, regardless of IncludeHeader. This command controls UTF-8, uncompressed file output and replacement through Force; CsvOptions controls field formatting and formula handling.

```yaml
Type: CsvSaveOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcelOptions
Advanced Excel typing, culture, row limit and table formatting settings.

```yaml
Type: PdfTablesToExcelOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FirstRowIsHeader
Confirm the first row of each Excel table with unknown schema as column headers.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Replace an existing destination after conversion succeeds.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HtmlOptions
Advanced HTML projection settings. Word and HTML defaults exclude the original scan from editable content.

```yaml
Type: PdfToHtmlOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Recognized image review returned by Get-OfficeImageDocument.

```yaml
Type: PdfSearchableOcrReview
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -OutputPath
Output path. The extension selects .xlsx, .docx, .pdf, .txt, .md, .html, .json or .csv.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return output location, source review and format report instead of FileInfo.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WordOptions
Advanced Word projection settings.

```yaml
Type: PdfToWordOptions
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.Pdf.Ocr.PdfSearchableOcrReview`

## OUTPUTS

- `System.IO.FileInfo`
- `PSWriteOffice.Services.Ocr.OfficeImageConversionResult`: Output location, recognition evidence and format projection report from an image export.

## RELATED LINKS

- None
