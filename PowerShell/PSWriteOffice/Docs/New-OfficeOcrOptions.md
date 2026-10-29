---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeOcrOptions
## SYNOPSIS
Creates optional OCR settings without starting a provider or downloading language data.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeOcrOptions [-Provider <IOcrEngine>] [-ExecutionOptions <OfficeDocumentOcrExecutionOptions>] [-RecognitionOptions <PdfOcrMergeOptions>] [-MinimumConfidence <Double>] [-Options <TesseractOcrSessionOptions>] [-Language <TesseractOcrLanguage[]>] [-TesseractLanguageExpression <string>] [-TesseractPath <string>] [-TessdataDirectory <string>] [-NoLanguageDownload] [<CommonParameters>]
```

## DESCRIPTION
Creates optional OCR settings without starting a provider or downloading language data.

## EXAMPLES

### EXAMPLE 1
```powershell
$ocr = New-OfficeOcrOptions -Language English,Polish -NoLanguageDownload
Read-OfficeDocument ./scan.png -As Text -OcrOptions $ocr
```


## PARAMETERS

### -ExecutionOptions
Bounded recognition and enrichment policy for document views.

```yaml
Type: OfficeDocumentOcrExecutionOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

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
Minimum confidence accepted into an editable OCR review.

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

### -Provider
Optional engine-neutral OCR provider.

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
Layout, confidence and resource policy for original image or PDF reviews.

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

- `None`

## OUTPUTS

- `PSWriteOffice.Models.Reader.OfficeOcrOptions`: Optional OCR provider and existing native policies used by document reads and OCR reviews.

## RELATED LINKS

- None
