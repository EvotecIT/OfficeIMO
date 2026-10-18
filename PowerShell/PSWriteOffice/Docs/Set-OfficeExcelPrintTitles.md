---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelPrintTitles
## SYNOPSIS
Sets or clears repeating print title rows and columns for a worksheet.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelPrintTitles [-Sheet <string>] [-SheetIndex <Int32>] [-FirstRow <Int32>] [-LastRow <Int32>] [-FirstColumn <Int32>] [-LastColumn <Int32>] [-Clear] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Set-OfficeExcelPrintTitles [-Path] <string> [-Sheet <string>] [-SheetIndex <Int32>] [-FirstRow <Int32>] [-LastRow <Int32>] [-FirstColumn <Int32>] [-LastColumn <Int32>] [-Clear] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelPrintTitles -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-FirstRow <Int32>] [-LastRow <Int32>] [-FirstColumn <Int32>] [-LastColumn <Int32>] [-Clear] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Sets or clears repeating print title rows and columns for a worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $proof = @(
    Set-OfficeExcelPrintTitles -Path .\Report.xlsx -Sheet Data -FirstRow 1 -LastRow 1 -FirstColumn 1 -LastColumn 1
    Get-OfficeExcelSummary -Path .\Report.xlsx |
        Select-Object -Property SheetCount, TableCount
)
$proof
```

Stores Excel print titles for the Data worksheet and then reads back workbook structure as a quick proof step.

## PARAMETERS

### -Clear
Clear existing print titles for the worksheet.

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

### -FirstColumn
First 1-based column to repeat.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FirstRow
First 1-based row to repeat.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LastColumn
Last 1-based column to repeat.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LastRow
Last 1-based row to repeat.

```yaml
Type: Int32
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
Emit the worksheet after setting print titles.

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

### -Sheet
Worksheet name. Defaults to the current sheet inside an ExcelSheet block.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: WorksheetName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index when using a workbook object or path.

```yaml
Type: Int32
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

- `None`

## RELATED LINKS

- None
