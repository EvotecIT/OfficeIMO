---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Clear-OfficeExcelRange
## SYNOPSIS
Clears values, formulas, styles and range metadata from a range, a header column or whole sheets.

## SYNTAX
### Range (Default)
```powershell
Clear-OfficeExcelRange [[-Range] <string>] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector[]>] [-Include <ExcelClearOptions[]>] [-PassThru] [<CommonParameters>]
```

### Header
```powershell
Clear-OfficeExcelRange -Header <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector[]>] [-Include <ExcelClearOptions[]>] [-PassThru] [-Table <string>] [-HeaderRow <Int32>] [-IncludeHeader] [<CommonParameters>]
```

## DESCRIPTION
Clears values, formulas, styles and range metadata from a range, a header column or whole sheets.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Templates\Weekly-Report.xlsx {
    ExcelSheet 'Data' {
        Clear-OfficeExcelRange A2:F200 -Include Values
    }
}
```

Removes the values from A2:F200 and keeps formatting, validation and comments.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Weekly-Ops.xlsx {
    ExcelSheet 'Data' {
        $cleared = Clear-OfficeExcelRange -PassThru
        Write-Verbose "Cleared $cleared cells"
        ExcelTable (Import-OfficeCsv -Path .\Data\Ops-Week43.csv) -Name OpsData
    }
}
```

Clears everything from the Data sheet, writes the new extract and reports how many cells were cleared.

## PARAMETERS

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Range, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Header
Header text identifying the target column.

```yaml
Type: String
Parameter Sets: Header
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderRow
One-based row containing the column headers.

```yaml
Type: Int32
Parameter Sets: Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Include
Cell data and metadata categories to clear. Omission clears all supported categories.

```yaml
Type: ExcelClearOptions[]
Parameter Sets: Range, Header
Aliases: None
Possible values: None, Values, Formulas, Styles, Comments, Hyperlinks, DataValidations, ConditionalFormatting, Merges, Sparklines, All

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeHeader
Includes the header cell in the resolved column range.

```yaml
Type: SwitchParameter
Parameter Sets: Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emits the native result after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: Range, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range
A1 cell range on the selected worksheet.

```yaml
Type: String
Parameter Sets: Range
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector[]
Parameter Sets: Range, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
Native Excel table used to resolve the source range or header.

```yaml
Type: String
Parameter Sets: Header
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
