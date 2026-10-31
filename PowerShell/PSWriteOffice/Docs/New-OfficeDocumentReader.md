---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeDocumentReader
## SYNOPSIS
Creates an immutable document reader with common policies, format options and optional OCR.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeDocumentReader [-Options <Object[]>] [-Processor <IOfficeDocumentProcessor[]>] [-ProcessorFailureBehavior <OfficeDocumentProcessorFailureBehavior>] [-Ocr] [-OcrOptions <OfficeOcrOptions>] [-MaximumStoreItems <int>] [-MaximumConcurrentReads <int>] [<CommonParameters>]
```

## DESCRIPTION
Creates an immutable document reader with common policies, format options and optional OCR.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeDocumentReaderOptions -Projection TextAndMetadata
$reader = New-OfficeDocumentReader -Options $options -MaximumConcurrentReads 4
Read-OfficeDocument ./Reports -Recurse -Reader $reader -As Text
```


### EXAMPLE 2
```powershell
$ocr = New-OfficeOcrOptions -Language English,Polish -NoLanguageDownload
$reader = New-OfficeDocumentReader -OcrOptions $ocr
Read-OfficeDocument ./Scans -Reader $reader
```


## PARAMETERS

### -MaximumConcurrentReads
Maximum concurrent native document reads.

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

### -MaximumStoreItems
Maximum items projected from each mail store; zero means all.

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

### -Ocr
Enable the default local OCR provider.

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

### -OcrOptions
Optional OCR provider and execution settings. Implies Ocr.

```yaml
Type: OfficeOcrOptions
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
Typed common and format options, or common/format-keyed hashtables.

```yaml
Type: Object[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Processor
Ordered processors applied after native extraction.

```yaml
Type: IOfficeDocumentProcessor[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ProcessorFailureBehavior
How ordinary processor failures affect the operation. Safety limits remain terminal.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Reader.OfficeDocumentReader`

## RELATED LINKS

- None
