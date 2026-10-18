---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelConditionalRule
## SYNOPSIS
Adds a conditional formatting rule to the current worksheet.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelConditionalRule [[-Range] <string>] [[-Operator] <ExcelConditionalFormattingOperator>] [[-Formula1] <string>] [-HeaderName <string>] [-TableName <string>] [-PivotTableName <string>] [-PivotWholeTable] [-HeaderRow <int>] [-IncludeHeader] [-RuleType <OfficeExcelConditionalRuleType>] [-Formula2 <string>] [-Text <string>] [-Rank <uint>] [-Percent] [-EqualAverage] [-StandardDeviation <UInt32>] [-TimePeriod <ExcelConditionalTimePeriod>] [-StopIfTrue] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelConditionalRule [[-Range] <string>] [[-Operator] <ExcelConditionalFormattingOperator>] [[-Formula1] <string>] -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-HeaderName <string>] [-TableName <string>] [-PivotTableName <string>] [-PivotWholeTable] [-HeaderRow <int>] [-IncludeHeader] [-RuleType <OfficeExcelConditionalRuleType>] [-Formula2 <string>] [-Text <string>] [-Rank <uint>] [-Percent] [-EqualAverage] [-StandardDeviation <UInt32>] [-TimePeriod <ExcelConditionalTimePeriod>] [-StopIfTrue] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a conditional formatting rule to the current worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Add-OfficeExcelConditionalRule -Range 'C2:C100' -Operator GreaterThan -Formula1 '100' }
```

Applies a conditional rule to column C.

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

### -EqualAverage
Include values equal to the average for average rules.

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

### -Formula1
Primary formula or value.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Formula2
Optional secondary formula or value.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderName
Header or table column name used to resolve the target range.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: ColumnName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderRow
Worksheet header row used when resolving HeaderName without a table. Use 0 for the first row of the used range.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeHeader
Include the header cell in the resolved range.

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

### -Operator
Conditional formatting operator.

```yaml
Type: ExcelConditionalFormattingOperator
Parameter Sets: Context, Document
Aliases: None
Possible values: LessThan, LessThanOrEqual, Equal, NotEqual, GreaterThanOrEqual, GreaterThan, Between, NotBetween, ContainsText, NotContains, BeginsWith, EndsWith

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the range after applying the rule.

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

### -Percent
Treat top/bottom rank as a percent.

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

### -PivotTableName
Pivot table name used to resolve the target range.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotWholeTable
Use the full pivot output range instead of the default data body range.

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

### -Range
A1 range to apply the rule to.

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

### -Rank
Rank used by top/bottom rules.

```yaml
Type: UInt32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RuleType
Rule type to author.

```yaml
Type: OfficeExcelConditionalRuleType
Parameter Sets: Context, Document
Aliases: Type
Possible values: CellIs, Expression, Formula, DuplicateValues, UniqueValues, Top, Top10, Bottom, Bottom10, AboveAverage, BelowAverage, ContainsText, NotContainsText, BeginsWith, EndsWith, ContainsBlanks, NotContainsBlanks, ContainsErrors, NotContainsErrors, TimePeriod

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index (0-based) when using Document.

```yaml
Type: Int32
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StandardDeviation
Optional standard deviation threshold for average rules.

```yaml
Type: UInt32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StopIfTrue
Stop evaluating later rules when this rule is true.

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

### -TableName
Optional table name for header-based range resolution.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text used by text-matching rule types.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TimePeriod
Time period used by time-period rules.

```yaml
Type: ExcelConditionalTimePeriod
Parameter Sets: Context, Document
Aliases: None
Possible values: Today, Yesterday, Tomorrow, Last7Days, ThisMonth, LastMonth, NextMonth, ThisWeek, LastWeek, NextWeek

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
