---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertFrom-OfficeOpenDocument
## SYNOPSIS
Converts native ODT, ODS, or ODP content to Word, Excel, or PowerPoint with fidelity evidence.

## SYNTAX
### __AllParameterSets
```powershell
ConvertFrom-OfficeOpenDocument [-Path] <string> [-OutputPath] <string> [-WordOptions <WordOpenDocumentConversionOptions>] [-ExcelOptions <ExcelOpenDocumentConversionOptions>] [-PowerPointOptions <PowerPointOpenDocumentConversionOptions>] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Converts native ODT, ODS, or ODP content to Word, Excel, or PowerPoint with fidelity evidence.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeExcelOpenDocumentOptions -IncludeBasicStyles -MaximumExpandedCells 250000
ConvertFrom-OfficeOpenDocument -Path .\Status.ods -OutputPath .\Status.xlsx -ExcelOptions $options
```


## PARAMETERS

### -ExcelOptions
Optional ODS-to-Excel conversion settings.

```yaml
Type: ExcelOpenDocumentConversionOptions
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Destination DOCX, XLSX, or PPTX path.

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

### -Path
Path to an ODT, ODS, or ODP file.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PowerPointOptions
Optional ODP-to-PowerPoint conversion settings.

```yaml
Type: PowerPointOpenDocumentConversionOptions
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
Optional ODT-to-Word conversion settings.

```yaml
Type: WordOpenDocumentConversionOptions
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

- `None`

## RELATED LINKS

- None
