---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Edit-OfficeExcelRow
## SYNOPSIS
Runs a script block against editable worksheet rows.

## SYNTAX
### Path (Default)
```powershell
Edit-OfficeExcelRow [-Path] <string> [-ScriptBlock] <scriptblock> [-Sheet <string>] [-SheetIndex <Int32>] [-Range <string>] [-NumericAsDecimal] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Edit-OfficeExcelRow [-ScriptBlock] <scriptblock> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Range <string>] [-NumericAsDecimal] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Runs a script block against editable worksheet rows.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Edit-OfficeExcelRow -Path .\Report.xlsx -Sheet Data -ScriptBlock { param($row) if ($row.Get[string]('Status') -eq 'Draft') { $row.Set('Status', 'Ready') } }
```

Loads editable row handles, lets the script update cells, and saves the workbook.

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

### -NumericAsDecimal
Prefer decimals instead of doubles for numeric values.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit each editable row after the script block runs.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
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

### -Range
A1 range to expose as editable rows. Defaults to the worksheet used range.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ScriptBlock
Script block to run once per editable row. The row is passed as the first argument.

```yaml
Type: ScriptBlock
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name. Defaults to the current sheet inside an ExcelSheet block.

```yaml
Type: String
Parameter Sets: Path, Document
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
Parameter Sets: Path, Document
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

- `PSWriteOffice.Services.Excel.ExcelPowerShellRowEdit`: Header-aware editable worksheet row used by Edit-OfficeExcelRow.

## RELATED LINKS

- None
