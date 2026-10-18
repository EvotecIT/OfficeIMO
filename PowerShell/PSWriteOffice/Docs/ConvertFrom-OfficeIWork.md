---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertFrom-OfficeIWork
## SYNOPSIS
Converts Pages, Numbers, or Keynote into the matching editable Microsoft Office format.

## SYNTAX
### __AllParameterSets
```powershell
ConvertFrom-OfficeIWork [-Path] <string> [-OutputPath] <string> [-ReadOptions <IWorkReadOptions>] [-ConversionOptions <IWorkConversionOptions>] [-FailOnLoss] [-Force] [-PassThruReport] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
The conversion reports what was reconstructed and what remains iWork-specific instead of implying lossless parity.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $report = ConvertFrom-OfficeIWork -Path .\Quarterly.numbers -OutputPath .\Quarterly.xlsx -PassThruReport
$report | Select-Object SourceKind, ProjectionKind, ReconstructedItemCount, HasLoss
```


## PARAMETERS

### -ConversionOptions
Editable-reconstruction or visual-fallback policy.

```yaml
Type: IWorkConversionOptions
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
Fail when source structures are flattened, omitted, or retained only as preserved records.

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
Overwrite an existing destination.

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
Destination DOCX, XLSX, or PPTX path matching the detected iWork application.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: OutPath
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThruReport
Return the loss-aware conversion report instead of file information.

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
Path to a modern Pages, Numbers, or Keynote package or directory bundle.

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

### -ReadOptions
Bounded OfficeIMO iWork read options.

```yaml
Type: IWorkReadOptions
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

- `System.IO.FileInfo`
- `OfficeIMO.IWork.IWorkConversionReport`

## RELATED LINKS

- None
