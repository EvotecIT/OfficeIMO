---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelSheet
## SYNOPSIS
Adds or reuses a worksheet within the current Excel DSL scope.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelSheet [[-Name] <string>] [[-Content] <scriptblock>] [-ValidationMode <ExcelSheetNameValidationMode>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelSheet [[-Name] <string>] [[-Content] <scriptblock>] -Document <ExcelDocument> [-ValidationMode <ExcelSheetNameValidationMode>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates the sheet when missing, pushes it onto the DSL stack, and executes the nested script block.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeExcel -Path .\report.xlsx { Add-OfficeExcelSheet -Name 'Data' { ExcelCell -Address 'A1' -Value 'Region' } }
```

Creates a workbook with a worksheet named Data and writes the header “Region”.

### EXAMPLE 2
```powershell
PS> $sheet = $workbook | Add-OfficeExcelSheet -Name 'Data' -PassThru
$sheet | Set-OfficeExcelCell -Address A1 -Value 'Region'
```

Returns the worksheet so subsequent commands can target it directly.

## PARAMETERS

### -Content
Code to execute inside the worksheet context.

```yaml
Type: ScriptBlock
Parameter Sets: Context, Document
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
Name of the worksheet to create or reuse. When omitted the last sheet is reused or a default sheet is created.

```yaml
Type: String
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
