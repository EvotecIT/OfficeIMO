---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeDocumentReader
## SYNOPSIS
Creates an immutable fully configured OfficeIMO document reader.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeDocumentReader [-UseTesseract] [-TesseractExecutablePath <string>] [-OcrLanguage <TesseractOcrLanguage[]>] [-TesseractLanguage <string>] [-TesseractDataPath <string>] [-TesseractDpi <Int32>] [-TesseractTimeoutSeconds <Int32>] [-MaxStoreItems <Int32>] [-AllStoreItems] [-MaxConcurrentReads <Int32>] [-ProcessorFailureBehavior <OfficeDocumentProcessorFailureBehavior>] [<CommonParameters>]
```

## DESCRIPTION
Creates an immutable fully configured OfficeIMO document reader.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $reader = New-OfficeDocumentReader -OcrLanguage English, Polish -MaxStoreItems 5000 -ProcessorFailureBehavior ContinueWithDiagnostic
```

The returned reader can be supplied to every PSWriteOffice Reader command.

## PARAMETERS

### -AllStoreItems
Project every matching item from each email store.

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

### -MaxConcurrentReads
Maximum asynchronous reads allowed in flight.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxStoreItems
Maximum PST, OST, OLM, or EMLX items projected from each store. The default is 1,000.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OcrLanguage
Friendly OCR languages used by the built-in Tesseract adapter.

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

### -ProcessorFailureBehavior
Behavior when a processor fails.

```yaml
Type: OfficeDocumentProcessorFailureBehavior
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Throw, ContinueWithDiagnostic, StopWithDiagnostic

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TesseractDataPath
Optional Tesseract tessdata directory.

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

### -TesseractDpi
Optional input DPI passed to Tesseract.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TesseractExecutablePath
Tesseract executable path or command name.

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

### -TesseractLanguage
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

### -TesseractTimeoutSeconds
Maximum Tesseract process duration in seconds. The default is 120.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseTesseract
Enable the built-in Tesseract command-line OCR adapter with default settings.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Reader.OfficeDocumentReader`

## RELATED LINKS

- None
