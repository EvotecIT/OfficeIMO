---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeOcrOptions
## SYNOPSIS
Creates OCR settings without starting a provider or downloading language data.

## SYNTAX
### Tesseract (Default)
```powershell
New-OfficeOcrOptions [-Language <TesseractOcrLanguage[]>] [-LanguageExpression <string>] [-TesseractPath <string>] [-DataDirectory <string>] [-NoLanguageDownload] [-Dpi <Int32>] [-TimeoutSeconds <Int32>] [-SessionOptions <TesseractOcrSessionOptions>] [-ExecutionOptions <OfficeDocumentOcrExecutionOptions>] [-RecognitionOptions <PdfOcrMergeOptions>] [-MinimumConfidence <Double>] [<CommonParameters>]
```

### Process
```powershell
New-OfficeOcrOptions -ProcessOptions <ProcessOcrEngineOptions> [-ExecutionOptions <OfficeDocumentOcrExecutionOptions>] [-RecognitionOptions <PdfOcrMergeOptions>] [-MinimumConfidence <Double>] [<CommonParameters>]
```

### Engine
```powershell
New-OfficeOcrOptions -Engine <IOcrEngine> [-ExecutionOptions <OfficeDocumentOcrExecutionOptions>] [-RecognitionOptions <PdfOcrMergeOptions>] [-MinimumConfidence <Double>] [<CommonParameters>]
```

## DESCRIPTION
Flat Tesseract parameters override an independent snapshot of SessionOptions. ProcessOptions and Engine select an alternative provider.

## EXAMPLES

### EXAMPLE 1
```powershell
$ocr = New-OfficeOcrOptions -Language English,Polish -Dpi 300 -NoLanguageDownload
Read-OfficeDocument ./scan.png -As Text -OcrOptions $ocr
```


## PARAMETERS

### -DataDirectory
Local directory containing installed trained-data files.

```yaml
Type: String
Parameter Sets: Tesseract
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Dpi
Image resolution supplied to Tesseract.

```yaml
Type: Int32
Parameter Sets: Tesseract
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Engine
Caller-provided OCR engine, used only when recognition is requested.

```yaml
Type: IOcrEngine
Parameter Sets: Engine
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExecutionOptions
Bounded recognition and enrichment policy for document views.

```yaml
Type: OfficeDocumentOcrExecutionOptions
Parameter Sets: Tesseract, Process, Engine
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Language
Friendly languages for the local Tesseract provider.

```yaml
Type: TesseractOcrLanguage[]
Parameter Sets: Tesseract
Aliases: None
Possible values: English, Polish, Arabic, ChineseSimplified, ChineseTraditional, Czech, Danish, Dutch, Finnish, French, German, Greek, Hebrew, Hindi, Hungarian, Italian, Japanese, Korean, Norwegian, Portuguese, Romanian, Russian, Slovak, Spanish, Swedish, Turkish, Ukrainian, Vietnamese

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LanguageExpression
Raw Tesseract language expression for caller-installed trained data.

```yaml
Type: String
Parameter Sets: Tesseract
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MinimumConfidence
Minimum word confidence accepted into an editable OCR review.

```yaml
Type: Double
Parameter Sets: Tesseract, Process, Engine
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLanguageDownload
Disable downloading missing curated language data.

```yaml
Type: SwitchParameter
Parameter Sets: Tesseract
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ProcessOptions
Settings for the native generic JSON file-protocol OCR provider.

```yaml
Type: ProcessOcrEngineOptions
Parameter Sets: Process
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RecognitionOptions
Layout, confidence and resource policy for image or PDF reviews.

```yaml
Type: PdfOcrMergeOptions
Parameter Sets: Tesseract, Process, Engine
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SessionOptions
Advanced Tesseract settings; flat parameters override matching values.

```yaml
Type: TesseractOcrSessionOptions
Parameter Sets: Tesseract
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TesseractPath
Explicit local Tesseract executable path.

```yaml
Type: String
Parameter Sets: Tesseract
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TimeoutSeconds
Maximum duration of a Tesseract process, in seconds.

```yaml
Type: Int32
Parameter Sets: Tesseract
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
