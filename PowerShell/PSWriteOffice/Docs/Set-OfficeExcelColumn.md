---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelColumn
## SYNOPSIS
Writes values down a column and sets its width, visibility, number format and style, by position or header.

## SYNTAX
### Column (Default)
```powershell
Set-OfficeExcelColumn [-Column] <Object> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Values <Object[]>] [-StartRow <int>] [-Width <ExcelColumnWidth>] [-Hidden] [-AutoFit] [-NumberFormat <string>] [-Style <OfficeExcelColumnStylePreset>] [-Decimals <int>] [-Culture <cultureinfo>] [-Pattern <string>] [-Bold] [-BackgroundColor <OfficeColor>] [-Color <OfficeColor>] [-Alignment <OfficeExcelColumnAlignment>] [-WrapText] [-BackgroundByText <hashtable>] [-ColorByText <hashtable>] [-BoldByText <string[]>] [-CaseSensitive] [-PassThru] [<CommonParameters>]
```

### Header
```powershell
Set-OfficeExcelColumn -Header <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Values <Object[]>] [-StartRow <int>] [-Width <ExcelColumnWidth>] [-Hidden] [-AutoFit] [-NumberFormat <string>] [-Style <OfficeExcelColumnStylePreset>] [-Decimals <int>] [-Culture <cultureinfo>] [-Pattern <string>] [-Bold] [-BackgroundColor <OfficeColor>] [-Color <OfficeColor>] [-Alignment <OfficeExcelColumnAlignment>] [-WrapText] [-BackgroundByText <hashtable>] [-ColorByText <hashtable>] [-BoldByText <string[]>] [-CaseSensitive] [-PassThru] [-Table <string>] [-IncludeHeader] [-IgnoreMissing] [<CommonParameters>]
```

## DESCRIPTION
Writes values down a column and sets its width, visibility, number format and style, by position or header.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Changes.xlsx {
    ExcelSheet 'Changes' {
        ExcelColumn B -Width 60 -WrapText
    }
}
```

Makes column B wide enough for long descriptions and wraps nothing else.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
    ExcelSheet 'Orders' {
        ExcelColumn -Header Amount -Table Orders -Style Currency -Decimals 2
        ExcelColumn -Header Status -Table Orders -BackgroundByText @{ Overdue = '#FEE2E2'; Paid = '#DCFCE7' }
    }
}
```

Formats the Amount column of the Orders table as currency and colours the Status column cells by their text.

## PARAMETERS

### -Alignment
Align cell content in the resolved column.

```yaml
Type: OfficeExcelColumnAlignment
Parameter Sets: Column, Header
Aliases: None
Possible values: Left, Center, Right

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoFit
Auto-fit the column width after updates.

```yaml
Type: SwitchParameter
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundByText
Background colors keyed by matching cell text.

```yaml
Type: Hashtable
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor
Cell fill colour expressed as a literal OfficeColor or theme role.

```yaml
Type: OfficeColor
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Applies bold text. An explicit false value clears bold; omission preserves the native value.

```yaml
Type: SwitchParameter
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BoldByText
Values that should be bolded when the cell text matches.

```yaml
Type: String[]
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CaseSensitive
Matches names and text using ordinal case sensitivity. Omission ignores case.

```yaml
Type: SwitchParameter
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Text or rule colour expressed as a literal OfficeColor or theme role.

```yaml
Type: OfficeColor
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColorByText
Maps matching cell text to font colours or theme roles.

```yaml
Type: Hashtable
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column
1-based column index.

```yaml
Type: Object
Parameter Sets: Column
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Culture
Culture used for value conversion and number-format presets.

```yaml
Type: CultureInfo
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Decimals
Decimal places for number, percent, and currency styles.

```yaml
Type: Int32
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Column, Header
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

### -Hidden
Hide or show the column.

```yaml
Type: SwitchParameter
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnoreMissing
Do nothing when the header cannot be found instead of throwing.

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

### -NumberFormat
Custom number format. Also used when Style is NumberFormat.

```yaml
Type: String
Parameter Sets: Column, Header
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
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Pattern
Date or DateTime number format pattern.

```yaml
Type: String
Parameter Sets: Column, Header
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
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartRow
Starting row index (1-based) for values.

```yaml
Type: Int32
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Style
Preset number style to apply.

```yaml
Type: OfficeExcelColumnStylePreset
Parameter Sets: Column, Header
Aliases: None
Possible values: Number, Integer, Percent, Currency, Date, DateTime, Time, DurationHours, Text, NumberFormat

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

### -Values
Values to write down the column. Numeric, Boolean, and date values retain their native Excel types; strings remain text.

```yaml
Type: Object[]
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Column width in Excel character units.

```yaml
Type: ExcelColumnWidth
Parameter Sets: Column, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WrapText
Apply or clear wrap text across the target row span.

```yaml
Type: SwitchParameter
Parameter Sets: Column, Header
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
