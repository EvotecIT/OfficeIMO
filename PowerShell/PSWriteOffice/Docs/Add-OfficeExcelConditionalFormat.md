---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelConditionalFormat
## SYNOPSIS
Adds a conditional format: a rule, color scale, data bar or icon set.

## SYNTAX
### RuleRange (Default)
```powershell
Add-OfficeExcelConditionalFormat [-Range] <string> [[-Operator] <ExcelConditionalFormattingOperator>] [[-Value] <string>] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Rule <OfficeExcelConditionalRuleType>] [-SecondValue <string>] [-Text <string>] [-Rank <uint>] [-Percent] [-EqualAverage] [-StandardDeviation <UInt32>] [-TimePeriod <ExcelConditionalTimePeriod>] [-StopIfTrue] [-BackgroundColor <OfficeColor>] [-Color <OfficeColor>] [-PassThru] [<CommonParameters>]
```

### ColorScaleRange
```powershell
Add-OfficeExcelConditionalFormat [-Range] <string> -StartColor <OfficeColor> -EndColor <OfficeColor> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [-MidColor <OfficeColor>] [<CommonParameters>]
```

### DataBarRange
```powershell
Add-OfficeExcelConditionalFormat [-Range] <string> -DataBar [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Color <OfficeColor>] [-PassThru] [<CommonParameters>]
```

### IconSetRange
```powershell
Add-OfficeExcelConditionalFormat [-Range] <string> -IconSet <ExcelIconSet> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [-NoValue] [-Reverse] [-PercentThreshold <double[]>] [-NumberThreshold <double[]>] [<CommonParameters>]
```

### RuleHeader
```powershell
Add-OfficeExcelConditionalFormat [[-Operator] <ExcelConditionalFormattingOperator>] [[-Value] <string>] -Header <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Rule <OfficeExcelConditionalRuleType>] [-SecondValue <string>] [-Text <string>] [-Rank <uint>] [-Percent] [-EqualAverage] [-StandardDeviation <UInt32>] [-TimePeriod <ExcelConditionalTimePeriod>] [-StopIfTrue] [-BackgroundColor <OfficeColor>] [-Color <OfficeColor>] [-PassThru] [-Table <string>] [-HeaderRow <Int32>] [-IncludeHeader] [<CommonParameters>]
```

### RulePivotTable
```powershell
Add-OfficeExcelConditionalFormat [[-Operator] <ExcelConditionalFormattingOperator>] [[-Value] <string>] -PivotTable <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Rule <OfficeExcelConditionalRuleType>] [-SecondValue <string>] [-Text <string>] [-Rank <uint>] [-Percent] [-EqualAverage] [-StandardDeviation <UInt32>] [-TimePeriod <ExcelConditionalTimePeriod>] [-StopIfTrue] [-BackgroundColor <OfficeColor>] [-Color <OfficeColor>] [-PassThru] [-PivotWholeTable] [<CommonParameters>]
```

### DataBarHeader
```powershell
Add-OfficeExcelConditionalFormat -Header <string> -DataBar [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Color <OfficeColor>] [-PassThru] [-Table <string>] [-HeaderRow <Int32>] [-IncludeHeader] [<CommonParameters>]
```

### DataBarPivotTable
```powershell
Add-OfficeExcelConditionalFormat -PivotTable <string> -DataBar [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Color <OfficeColor>] [-PassThru] [-PivotWholeTable] [<CommonParameters>]
```

### ColorScaleHeader
```powershell
Add-OfficeExcelConditionalFormat -Header <string> -StartColor <OfficeColor> -EndColor <OfficeColor> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [-Table <string>] [-HeaderRow <Int32>] [-IncludeHeader] [-MidColor <OfficeColor>] [<CommonParameters>]
```

### IconSetHeader
```powershell
Add-OfficeExcelConditionalFormat -Header <string> -IconSet <ExcelIconSet> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [-Table <string>] [-HeaderRow <Int32>] [-IncludeHeader] [-NoValue] [-Reverse] [-PercentThreshold <double[]>] [-NumberThreshold <double[]>] [<CommonParameters>]
```

### ColorScalePivotTable
```powershell
Add-OfficeExcelConditionalFormat -PivotTable <string> -StartColor <OfficeColor> -EndColor <OfficeColor> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [-PivotWholeTable] [-MidColor <OfficeColor>] [<CommonParameters>]
```

### IconSetPivotTable
```powershell
Add-OfficeExcelConditionalFormat -PivotTable <string> -IconSet <ExcelIconSet> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [-PivotWholeTable] [-NoValue] [-Reverse] [-PercentThreshold <double[]>] [-NumberThreshold <double[]>] [<CommonParameters>]
```

## DESCRIPTION
Adds a conditional format: a rule, color scale, data bar or icon set.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
    ExcelSheet 'Orders' {
        ExcelConditionalFormat -Header Amount -Table Orders -DataBar
    }
}
```

Draws data bars in the Amount column of the Orders table in the first accent colour.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Q3-Performance.xlsx {
    ExcelSheet 'Performance' {
        ExcelConditionalFormat -Header Attainment -Table Actuals -StartColor Danger -MidColor Warning -EndColor Success
        ExcelConditionalFormat -Header Change -Table Actuals -IconSet ThreeArrows
    }
}
```

Applies a three-colour scale from the theme's danger to success colours to target attainment, and arrows to the change against last quarter.

### EXAMPLE 3
```powershell
Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
    ExcelSheet 'Orders' {
        ExcelConditionalFormat E2:E500 -Rule CellIs -Operator LessThan -Value 0 -BackgroundColor '#FEE2E2' -Color Danger
    }
}
```

Highlights negative margins with a light red fill and red text.

### EXAMPLE 4
```powershell
Open-OfficeExcel -Path .\Reports\Users.xlsx {
    ExcelSheet 'Users' {
        ExcelConditionalFormat -Header LastLogon -Table Users -Rule CellIs -Operator LessThan -Value '=TODAY()-90' -BackgroundColor Danger
    }
}
```

Colours the LastLogon cell when the date is more than 90 days old; -Value takes an Excel formula when it starts with '='.

## PARAMETERS

### -BackgroundColor
Cell fill colour expressed as a literal OfficeColor or theme role.

```yaml
Type: OfficeColor
Parameter Sets: RuleRange, RuleHeader, RulePivotTable
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
Parameter Sets: RuleRange, DataBarRange, RuleHeader, RulePivotTable, DataBarHeader, DataBarPivotTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataBar
Selects the native data-bar conditional-format family.

```yaml
Type: SwitchParameter
Parameter Sets: DataBarRange, DataBarHeader, DataBarPivotTable
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: RuleRange, ColorScaleRange, DataBarRange, IconSetRange, RuleHeader, RulePivotTable, DataBarHeader, DataBarPivotTable, ColorScaleHeader, IconSetHeader, ColorScalePivotTable, IconSetPivotTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -EndColor
End color. Named colors and hexadecimal values are accepted.

```yaml
Type: OfficeColor
Parameter Sets: ColorScaleRange, ColorScaleHeader, ColorScalePivotTable
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EqualAverage
Include values equal to the average for average rules.

```yaml
Type: SwitchParameter
Parameter Sets: RuleRange, RuleHeader, RulePivotTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header
Header text identifying the target column.

```yaml
Type: String
Parameter Sets: RuleHeader, DataBarHeader, ColorScaleHeader, IconSetHeader
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
Parameter Sets: RuleHeader, DataBarHeader, ColorScaleHeader, IconSetHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IconSet
Icon set to apply.

```yaml
Type: ExcelIconSet
Parameter Sets: IconSetRange, IconSetHeader, IconSetPivotTable
Aliases: None
Possible values: ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeHeader
Includes the header cell in the resolved column range.

```yaml
Type: SwitchParameter
Parameter Sets: RuleHeader, DataBarHeader, ColorScaleHeader, IconSetHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MidColor
Middle colour of a three-colour scale at the 50th percentile. Omission uses a two-colour scale.

```yaml
Type: OfficeColor
Parameter Sets: ColorScaleRange, ColorScaleHeader, ColorScalePivotTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoValue
Hides cell values behind conditional-format icons.

```yaml
Type: SwitchParameter
Parameter Sets: IconSetRange, IconSetHeader, IconSetPivotTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NumberThreshold
Numeric thresholds for native conditional-format icons.

```yaml
Type: Double[]
Parameter Sets: IconSetRange, IconSetHeader, IconSetPivotTable
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
Parameter Sets: RuleRange, RuleHeader, RulePivotTable
Aliases: None
Possible values: LessThan, LessThanOrEqual, Equal, NotEqual, GreaterThanOrEqual, GreaterThan, Between, NotBetween, ContainsText, NotContains, BeginsWith, EndsWith

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emits the native result after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: RuleRange, ColorScaleRange, DataBarRange, IconSetRange, RuleHeader, RulePivotTable, DataBarHeader, DataBarPivotTable, ColorScaleHeader, IconSetHeader, ColorScalePivotTable, IconSetPivotTable
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
Parameter Sets: RuleRange, RuleHeader, RulePivotTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PercentThreshold
Percentage thresholds for native conditional-format icons.

```yaml
Type: Double[]
Parameter Sets: IconSetRange, IconSetHeader, IconSetPivotTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotTable
Native pivot table whose data or entire range is targeted.

```yaml
Type: String
Parameter Sets: RulePivotTable, DataBarPivotTable, ColorScalePivotTable, IconSetPivotTable
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotWholeTable
Use the full pivot output range instead of the default data body range.

```yaml
Type: SwitchParameter
Parameter Sets: RulePivotTable, DataBarPivotTable, ColorScalePivotTable, IconSetPivotTable
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
Parameter Sets: RuleRange, ColorScaleRange, DataBarRange, IconSetRange
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rank
Rank used by top/bottom rules.

```yaml
Type: UInt32
Parameter Sets: RuleRange, RuleHeader, RulePivotTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Reverse
Reverse the icon order.

```yaml
Type: SwitchParameter
Parameter Sets: IconSetRange, IconSetHeader, IconSetPivotTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rule
Native conditional-format rule family; the default compares cell values.

```yaml
Type: OfficeExcelConditionalRuleType
Parameter Sets: RuleRange, RuleHeader, RulePivotTable
Aliases: None
Possible values: CellIs, Expression, Formula, DuplicateValues, UniqueValues, Top, Top10, Bottom, Bottom10, AboveAverage, BelowAverage, ContainsText, NotContainsText, BeginsWith, EndsWith, ContainsBlanks, NotContainsBlanks, ContainsErrors, NotContainsErrors, TimePeriod

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SecondValue
Upper bound or second formula used by between comparisons.

```yaml
Type: String
Parameter Sets: RuleRange, RuleHeader, RulePivotTable
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
Parameter Sets: RuleRange, ColorScaleRange, DataBarRange, IconSetRange, RuleHeader, RulePivotTable, DataBarHeader, DataBarPivotTable, ColorScaleHeader, IconSetHeader, ColorScalePivotTable, IconSetPivotTable
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
Parameter Sets: RuleRange, RuleHeader, RulePivotTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartColor
Start color. Named colors and hexadecimal values are accepted.

```yaml
Type: OfficeColor
Parameter Sets: ColorScaleRange, ColorScaleHeader, ColorScalePivotTable
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StopIfTrue
Stop evaluating later rules when this rule is true.

```yaml
Type: SwitchParameter
Parameter Sets: RuleRange, RuleHeader, RulePivotTable
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
Parameter Sets: RuleHeader, DataBarHeader, ColorScaleHeader, IconSetHeader
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
Parameter Sets: RuleRange, RuleHeader, RulePivotTable
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
Parameter Sets: RuleRange, RuleHeader, RulePivotTable
Aliases: None
Possible values: Today, Yesterday, Tomorrow, Last7Days, ThisMonth, LastMonth, NextMonth, ThisWeek, LastWeek, NextWeek

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Value or values consumed by the selected operation.

```yaml
Type: String
Parameter Sets: RuleRange, RuleHeader, RulePivotTable
Aliases: None
Possible values:

Required: False
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
