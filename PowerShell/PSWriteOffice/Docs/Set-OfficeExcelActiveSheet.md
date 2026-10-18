---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelActiveSheet
## SYNOPSIS
Sets the worksheet that opens as the active sheet.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelActiveSheet [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Set-OfficeExcelActiveSheet [-Path] <string> [-Sheet <string>] [-SheetIndex <Int32>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelActiveSheet -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Sets the worksheet that opens as the active sheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $sheet = Set-OfficeExcelActiveSheet -Path .\Report.xlsx -Sheet Summary -PassThru
Get-OfficeExcelWorksheetView -Path .\Report.xlsx -Sheet $sheet.Name |
    Select-Object SheetName, View, TopLeftCell
```

Updates workbook view state so spreadsheet applications open on Summary.

## PARAMETERS

### -Document
Workbook to operate on outside the DSL context.

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

### -PassThru
Emit the activated worksheet.

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
Worksheet name to activate. Defaults to the current sheet inside an ExcelSheet block.

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
Zero-based worksheet index to activate.

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

- `OfficeIMO.Excel.ExcelSheet`
- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
