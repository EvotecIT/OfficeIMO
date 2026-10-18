---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeImageDocument
## SYNOPSIS
Recognizes an image as editable text, reading order and detected tables, with review evidence.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeImageDocument [-Path] <string> [-Provider <IOcrEngine>] [-RecognitionOptions <PdfOcrMergeOptions>] [-MinimumConfidence <Double>] [-Options <TesseractOcrSessionOptions>] [-Language <TesseractOcrLanguage[]>] [-TesseractLanguageExpression <string>] [-TesseractPath <string>] [-TessdataDirectory <string>] [-NoLanguageDownload] [<CommonParameters>]
```

## DESCRIPTION
Uses the original image resolution and the shared OfficeIMO layout engine. Animated and multi-page images are rejected.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $image = Get-OfficeImageDocument -Path .\Ledger.png
$image | ConvertFrom-OfficeImage -OutputPath .\Ledger.xlsx -FirstRowIsHeader
```

Inspect $image.Ocr.Document.Tables and $image.Ocr.Pages before confirming the first row as headers.

## PARAMETERS

### -Language
Friendly OCR languages. Supply more than one value to recognize multilingual content.

```yaml
Type: TesseractOcrLanguage[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: English, Polish, Arabic, ChineseSimplified, ChineseTraditional, Czech, Danish, Dutch, Finnish, French, German, Greek, Hebrew, Hindi, Hungarian, Italian, Japanese, Korean, Norwegian, Portuguese, Romanian, Russian, Slovak, Spanish, Swedish, Turkish, Ukrainian, Vietnamese

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MinimumConfidence
Minimum normalized confidence accepted into the editable document.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLanguageDownload
Do not download checksum-pinned curated language data when a requested language is missing.

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

### -Options
Advanced OfficeIMO OCR options. Convenience parameters override matching values.

```yaml
Type: TesseractOcrSessionOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to a supported single-frame raster image.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Provider
Optional engine-neutral OCR provider. When omitted, uses the configured local Tesseract session.

```yaml
Type: IOcrEngine
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RecognitionOptions
Advanced layout, confidence, orientation, scan preparation and resource limits.

```yaml
Type: PdfOcrMergeOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TessdataDirectory
Explicit directory containing Tesseract trained-data files.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TesseractLanguageExpression
Advanced raw Tesseract expression for caller-installed custom trained-data models.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TesseractPath
Explicit Tesseract executable path. By default OfficeIMO securely discovers an installed runtime.

```yaml
Type: String
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Pdf.Ocr.PdfSearchableOcrReview`

## RELATED LINKS

- None
