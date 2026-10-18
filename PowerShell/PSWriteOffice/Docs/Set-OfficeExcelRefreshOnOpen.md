---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelRefreshOnOpen
## SYNOPSIS
Configures workbook data refresh metadata for Excel-compatible applications to run when the file opens.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelRefreshOnOpen [-PivotTables] [-Connections] [-Disable] [-SavePivotSourceData] [-NoSavePivotSourceData] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Set-OfficeExcelRefreshOnOpen [-Path] <string> [-PivotTables] [-Connections] [-Disable] [-SavePivotSourceData] [-NoSavePivotSourceData] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelRefreshOnOpen -Document <ExcelDocument> [-PivotTables] [-Connections] [-Disable] [-SavePivotSourceData] [-NoSavePivotSourceData] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Configures workbook data refresh metadata for Excel-compatible applications to run when the file opens.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $refresh = Set-OfficeExcelRefreshOnOpen -Path .\Report.xlsx -PivotTables -Connections -NoSavePivotSourceData -PassThru
Get-OfficeExcelDataModel -Path .\Report.xlsx |
    Select-Object ConnectionPartCount, QueryTablePartCount
```

Sets workbook metadata through OfficeIMO so pivot caches refresh on open.

## PARAMETERS

### -Connections
Update workbook connection refresh-on-open metadata.

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

### -Disable
Disable refresh-on-open instead of enabling it.

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

### -NoSavePivotSourceData
Do not save pivot cache source data.

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

### -PassThru
Emit the update summary.

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

### -PivotTables
Update pivot cache refresh-on-open metadata.

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

### -SavePivotSourceData
Preserve pivot cache source data.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
