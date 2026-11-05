---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelSheet
## SYNOPSIS
Adds or reuses a worksheet within the current Excel DSL scope.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelSheet [[-Name] <string>] [[-Content] <scriptblock>] [-Document <ExcelDocument>] [-ValidationMode <ExcelSheetNameValidationMode>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates the sheet when missing, pushes it onto the DSL stack, and executes the nested script block.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Inventory.xlsx {
    ExcelSheet 'Servers' {
        ExcelTable (Import-OfficeCsv -Path .\Data\Servers.csv)
    }
}
```

Adds an Inventory sheet and writes a table of servers into it.

### EXAMPLE 2
```powershell
$staff = Import-OfficeCsv -Path .\HR\Headcount-2026-10.csv
New-OfficeExcel -Path .\Reports\Headcount.xlsx {
    foreach ($group in $staff | Group-Object Department) {
        ExcelSheet $group.Name -ValidationMode Sanitize {
            ExcelTable ($group.Group | Select-Object Name, Title, Location, StartDate)
        }
    }
}
```

Groups the headcount export by department and writes each group to its own sheet, with a cleaned-up name when a department name has characters Excel does not allow.

## PARAMETERS

### -Content
Code to execute inside the worksheet context.

```yaml
Type: ScriptBlock
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook that will receive the worksheet.

```yaml
Type: ExcelDocument
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Name
Name of the worksheet to create or reuse. When omitted the last sheet is reused or a default sheet is created.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the ExcelSheet object after execution.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ValidationMode
Controls how invalid sheet names are handled.

```yaml
Type: ExcelSheetNameValidationMode
Parameter Sets: Default
Aliases: None
Possible values: None, Sanitize, Strict

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

- `OfficeIMO.Excel.ExcelSheet`

## RELATED LINKS

- None
