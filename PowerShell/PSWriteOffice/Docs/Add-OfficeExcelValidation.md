---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelValidation
## SYNOPSIS
Adds a data validation rule to a range or a header column.

## SYNTAX
### ListRange (Default)
```powershell
Add-OfficeExcelValidation [-Range] <string> [-Type] <ExcelValidationType> [-Values] <string[]> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NoBlank] [-PromptTitle <string>] [-Prompt <string>] [-ErrorTitle <string>] [-ErrorMessage <string>] [-PassThru] [<CommonParameters>]
```

### CompareRange
```powershell
Add-OfficeExcelValidation [-Range] <string> [-Type] <ExcelValidationType> [-Value] <Object> -Operator <ExcelDataValidationOperator> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NoBlank] [-PromptTitle <string>] [-Prompt <string>] [-ErrorTitle <string>] [-ErrorMessage <string>] [-PassThru] [-SecondValue <Object>] [<CommonParameters>]
```

### FormulaRange
```powershell
Add-OfficeExcelValidation [-Range] <string> [-Type] <ExcelValidationType> [-Formula] <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NoBlank] [-PromptTitle <string>] [-Prompt <string>] [-ErrorTitle <string>] [-ErrorMessage <string>] [-PassThru] [<CommonParameters>]
```

### ListHeader
```powershell
Add-OfficeExcelValidation [-Type] <ExcelValidationType> [-Values] <string[]> -Header <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NoBlank] [-PromptTitle <string>] [-Prompt <string>] [-ErrorTitle <string>] [-ErrorMessage <string>] [-PassThru] [-Table <string>] [-HeaderRow <Int32>] [-IncludeHeader] [<CommonParameters>]
```

### CompareHeader
```powershell
Add-OfficeExcelValidation [-Type] <ExcelValidationType> [-Value] <Object> -Header <string> -Operator <ExcelDataValidationOperator> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NoBlank] [-PromptTitle <string>] [-Prompt <string>] [-ErrorTitle <string>] [-ErrorMessage <string>] [-PassThru] [-Table <string>] [-HeaderRow <Int32>] [-IncludeHeader] [-SecondValue <Object>] [<CommonParameters>]
```

### FormulaHeader
```powershell
Add-OfficeExcelValidation [-Type] <ExcelValidationType> [-Formula] <string> -Header <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NoBlank] [-PromptTitle <string>] [-Prompt <string>] [-ErrorTitle <string>] [-ErrorMessage <string>] [-PassThru] [-Table <string>] [-HeaderRow <Int32>] [-IncludeHeader] [<CommonParameters>]
```

## DESCRIPTION
Adds a data validation rule to a range or a header column.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
    ExcelSheet 'Tickets' {
        ExcelValidation -Range C2:C500 -Type List -Values 'Open', 'In progress', 'On hold', 'Closed'
    }
}
```

Adds a drop-down that allows only the listed statuses in C2:C500.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Forms\Equipment-Requests.xlsx {
    ExcelSheet 'Requests' {
        ExcelValidation -Header Quantity -Table Requests -Type WholeNumber -Operator Between -Value 1 -SecondValue 500 -NoBlank -ErrorTitle 'Quantity' -ErrorMessage 'Enter a whole number from 1 to 500.'
    }
}
```

Allows whole numbers from 1 to 500 in the Quantity column of the Requests table and explains the rule when a wrong value is typed.

### EXAMPLE 3
```powershell
Open-OfficeExcel -Path .\Projects\Plan.xlsx {
    ExcelSheet 'Tasks' {
        ExcelValidation -Range D2:D300 -Type Formula -Formula '=D2>=C2' -ErrorMessage 'The end date must be on or after the start date.'
    }
}
```

Rejects end dates that come before the start date in the same row.

## PARAMETERS

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: ListRange, CompareRange, FormulaRange, ListHeader, CompareHeader, FormulaHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ErrorMessage
Error message shown to the user.

```yaml
Type: String
Parameter Sets: ListRange, CompareRange, FormulaRange, ListHeader, CompareHeader, FormulaHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ErrorTitle
Error title shown to the user.

```yaml
Type: String
Parameter Sets: ListRange, CompareRange, FormulaRange, ListHeader, CompareHeader, FormulaHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Formula
Validation formula.

```yaml
Type: String
Parameter Sets: FormulaRange, FormulaHeader
Aliases: None
Possible values:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header
Header text identifying the target column.

```yaml
Type: String
Parameter Sets: ListHeader, CompareHeader, FormulaHeader
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
Parameter Sets: ListHeader, CompareHeader, FormulaHeader
Aliases: None
Possible values:

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
Parameter Sets: ListHeader, CompareHeader, FormulaHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoBlank
Rejects blank input in the validation rule.

```yaml
Type: SwitchParameter
Parameter Sets: ListRange, CompareRange, FormulaRange, ListHeader, CompareHeader, FormulaHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Operator
Validation operator.

```yaml
Type: ExcelDataValidationOperator
Parameter Sets: CompareRange, CompareHeader
Aliases: None
Possible values: Between, NotBetween, Equal, NotEqual, LessThan, LessThanOrEqual, GreaterThan, GreaterThanOrEqual

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emits the native result after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: ListRange, CompareRange, FormulaRange, ListHeader, CompareHeader, FormulaHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Prompt
Input prompt text. Omit or pass null to clear the prompt.

```yaml
Type: String
Parameter Sets: ListRange, CompareRange, FormulaRange, ListHeader, CompareHeader, FormulaHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PromptTitle
Input prompt title. Omit or pass null to clear the title.

```yaml
Type: String
Parameter Sets: ListRange, CompareRange, FormulaRange, ListHeader, CompareHeader, FormulaHeader
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
Parameter Sets: ListRange, CompareRange, FormulaRange
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SecondValue
Upper bound or second formula used by between comparisons.

```yaml
Type: Object
Parameter Sets: CompareRange, CompareHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector
Parameter Sets: ListRange, CompareRange, FormulaRange, ListHeader, CompareHeader, FormulaHeader
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
Parameter Sets: ListHeader, CompareHeader, FormulaHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Type
Shared chart family. Bubble requires the separate native bubble-data contract.

```yaml
Type: ExcelValidationType
Parameter Sets: ListRange, CompareRange, FormulaRange, ListHeader, CompareHeader, FormulaHeader
Aliases: None
Possible values: List, WholeNumber, Decimal, Date, Time, TextLength, Formula

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Value or values consumed by the selected operation.

```yaml
Type: Object
Parameter Sets: CompareRange, CompareHeader
Aliases: None
Possible values:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Values
Allowed values for the dropdown list.

```yaml
Type: String[]
Parameter Sets: ListRange, ListHeader
Aliases: None
Possible values:

Required: True
Position: 2
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
