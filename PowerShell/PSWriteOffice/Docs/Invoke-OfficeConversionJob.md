---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Invoke-OfficeConversionJob
## SYNOPSIS
Plans or runs a saved conversion job through the same runner as Studio and the CLI.

## SYNTAX
### __AllParameterSets
```powershell
Invoke-OfficeConversionJob [-Path] <string> [-Plan] [-RetryFailed] [-ItemResults] [-SummaryVariable <string>] [-AllowEmpty] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Requires PowerShell 7.4 or newer. Plan performs no output or checkpoint writes. Execution revalidates current inputs and retains completed artifacts after a partial failure or cancellation. Password environment references are resolved only for execution; values are not written to the job or reports.

## EXAMPLES

### EXAMPLE 1
```powershell
Invoke-OfficeConversionJob -Path ./monthly.officejob.json -Plan
            Invoke-OfficeConversionJob -Path ./monthly.officejob.json -ItemResults -SummaryVariable summary -ErrorAction Stop
```


### EXAMPLE 2
```powershell
Invoke-OfficeConversionJob -Path ./monthly.officejob.json -RetryFailed -ErrorAction Stop
```


## PARAMETERS

### -AllowEmpty
Allow execution to select no supported inputs without reporting an error.

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

### -ItemResults
Stream per-file outcomes instead of only the final bounded summary.

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
Literal saved-job JSON path. Relative paths in the job resolve against this file's folder.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: LiteralPath, FullName
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Plan
Inspect the job without converting files or creating output/checkpoint state.

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

### -RetryFailed
Retry failed checkpoint items while verifying completed artifacts.

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

### -SummaryVariable
Variable receiving the completed plan or batch summary, including when ItemResults is selected.

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

- `OfficeIMO.Workflows.OfficeConversionBatchResult`
- `OfficeIMO.Workflows.OfficeConversionBatchItemResult`
- `OfficeIMO.Workflows.OfficeConversionJobPlan`

## RELATED LINKS

- None
