---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelTimeline
## SYNOPSIS
Adds OfficeIMO-owned workbook timeline binding metadata.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelTimeline -Name <string> [-SourceName <string>] [-PivotTableName <string>] [-Xml <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Add-OfficeExcelTimeline [-Path] <string> -Name <string> [-SourceName <string>] [-PivotTableName <string>] [-Xml <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelTimeline -Document <ExcelDocument> -Name <string> [-SourceName <string>] [-PivotTableName <string>] [-Xml <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Adds OfficeIMO-owned workbook timeline binding metadata.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $timeline = Add-OfficeExcelTimeline -Path .\Report.xlsx -Name OrderDateTimeline -SourceName OrderDate -PivotTableName SalesPivot -PassThru
Get-OfficeExcelDataModel -Path .\Report.xlsx |
    Select-Object -ExpandProperty TimelineCacheCount
```

Writes portable OfficeIMO binding metadata. It does not create native Excel timeline caches or UI shapes.

## PARAMETERS

### -Document
Workbook to update outside the DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Name
Timeline cache name.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit metadata about the added package part.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Workbook path to update.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotTableName
Pivot table name the timeline is intended to filter.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceName
Source date field, table column, or pivot field name.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Xml
Caller-supplied timeline cache XML. When provided, OfficeIMO preserves it within its pivot-interaction metadata envelope.

```yaml
Type: String
Parameter Sets: Context, Path, Document
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

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
