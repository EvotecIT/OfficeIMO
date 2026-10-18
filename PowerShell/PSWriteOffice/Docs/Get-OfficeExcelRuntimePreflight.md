---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelRuntimePreflight
## SYNOPSIS
Inspects the current process for runtime settings that affect Excel workflows.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeExcelRuntimePreflight [<CommonParameters>]
```

## DESCRIPTION
Inspects the current process for runtime settings that affect Excel workflows.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $report = Get-OfficeExcelRuntimePreflight
if (-not $report.IsClean) { $report.Warnings | Write-Warning }
```

Returns framework, operating system, culture, and globalization-invariant mode diagnostics from OfficeIMO.

## PARAMETERS

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
