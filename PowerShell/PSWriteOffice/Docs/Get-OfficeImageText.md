---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeImageText
## SYNOPSIS
Recognizes text in an image with automatic local OCR runtime discovery.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeImageText [-Path] <string> [-PassThru] [-Options <TesseractOcrSessionOptions>] [-Language <TesseractOcrLanguage[]>] [-TesseractLanguageExpression <string>] [-TesseractPath <string>] [-TessdataDirectory <string>] [-NoLanguageDownload] [<CommonParameters>]
```

## DESCRIPTION
Recognizes text in an image with automatic local OCR runtime discovery.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeImageText -Path .\Scan.png
```

Returns recognized text and automatically uses an installed Tesseract runtime.

### EXAMPLE 2
```powershell
PS> Get-OfficeImageText -Path .\Scan.png -Language English, Polish -PassThru
```

Returns the full OCR result, including confidence, word geometry, provider, model, and diagnostics.

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

### -PassThru
Return the complete OCR result instead of recognized text only.

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

### -Path
PNG, JPEG, TIFF, BMP, GIF, WebP, or JPEG 2000 image path.

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

- `System.String`
- `OfficeIMO.Ocr.OcrResult`

## RELATED LINKS

- None
