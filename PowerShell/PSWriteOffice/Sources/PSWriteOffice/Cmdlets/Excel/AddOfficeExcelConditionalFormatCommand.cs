using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a conditional format: a rule, color scale, data bar or icon set.</summary>
/// <example>
/// <summary>Add data bars to a column</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     ExcelSheet 'Orders' {
///         ExcelConditionalFormat -Header Amount -Table Orders -DataBar
///     }
/// }</code>
/// <para>Draws data bars in the Amount column of the Orders table in the first accent colour.</para>
/// </example>
/// <example>
/// <summary>Colour attainment from red to green</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Performance.xlsx {
///     ExcelSheet 'Performance' {
///         ExcelConditionalFormat -Header Attainment -Table Actuals -StartColor Danger -MidColor Warning -EndColor Success
///         ExcelConditionalFormat -Header Change -Table Actuals -IconSet ThreeArrows
///     }
/// }</code>
/// <para>Applies a three-colour scale from the theme's danger to success colours to target attainment, and arrows to the change against last quarter.</para>
/// </example>
/// <example>
/// <summary>Flag negative margins with a rule</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     ExcelSheet 'Orders' {
///         ExcelConditionalFormat E2:E500 -Rule CellIs -Operator LessThan -Value 0 -BackgroundColor '#FEE2E2' -Color Danger
///     }
/// }</code>
/// <para>Highlights negative margins with a light red fill and red text.</para>
/// </example>
/// <example>
/// <summary>Highlight accounts with no logon in 90 days</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Users.xlsx {
///     ExcelSheet 'Users' {
///         ExcelConditionalFormat -Header LastLogon -Table Users -Rule CellIs -Operator LessThan -Value '=TODAY()-90' -BackgroundColor Danger
///     }
/// }</code>
/// <para>Colours the LastLogon cell when the date is more than 90 days old; -Value takes an Excel formula when it starts with '='.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelConditionalFormat", DefaultParameterSetName = "RuleRange")]
[Alias("ExcelConditionalFormat")]
public sealed class AddOfficeExcelConditionalFormatCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter(ParameterSetName = "RuleRange", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "ColorScaleRange", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "DataBarRange", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "IconSetRange", Mandatory = true, Position = 0)]
    public string Range { get; set; } = null!;

    /// <summary>Native conditional-format rule family; the default compares cell values.</summary>
    [Parameter(ParameterSetName = "RuleRange")]
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "RulePivotTable")]
    public OfficeExcelConditionalRuleType Rule { get; set; } = OfficeExcelConditionalRuleType.CellIs;

    /// <summary>Conditional formatting operator.</summary>
    [Parameter(ParameterSetName = "RuleRange", Position = 1)]
    [Parameter(ParameterSetName = "RuleHeader", Position = 1)]
    [Parameter(ParameterSetName = "RulePivotTable", Position = 1)]
    public ExcelConditionalFormattingOperator? Operator { get; set; }

    /// <summary>Value or values consumed by the selected operation.</summary>
    [Parameter(ParameterSetName = "RuleRange", Position = 2)]
    [Parameter(ParameterSetName = "RuleHeader", Position = 2)]
    [Parameter(ParameterSetName = "RulePivotTable", Position = 2)]
    public string? Value { get; set; }

    /// <summary>Upper bound or second formula used by between comparisons.</summary>
    [Parameter(ParameterSetName = "RuleRange")]
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "RulePivotTable")]
    public string? SecondValue { get; set; }

    /// <summary>Text used by text-matching rule types.</summary>
    [Parameter(ParameterSetName = "RuleRange")]
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "RulePivotTable")]
    public string? Text { get; set; }

    /// <summary>Rank used by top/bottom rules.</summary>
    [Parameter(ParameterSetName = "RuleRange")]
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "RulePivotTable")]
    public uint Rank { get; set; } = 10;

    /// <summary>Treat top/bottom rank as a percent.</summary>
    [Parameter(ParameterSetName = "RuleRange")]
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "RulePivotTable")]
    public SwitchParameter Percent { get; set; }

    /// <summary>Include values equal to the average for average rules.</summary>
    [Parameter(ParameterSetName = "RuleRange")]
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "RulePivotTable")]
    public SwitchParameter EqualAverage { get; set; }

    /// <summary>Optional standard deviation threshold for average rules.</summary>
    [Parameter(ParameterSetName = "RuleRange")]
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "RulePivotTable")]
    public uint? StandardDeviation { get; set; }

    /// <summary>Time period used by time-period rules.</summary>
    [Parameter(ParameterSetName = "RuleRange")]
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "RulePivotTable")]
    public ExcelConditionalTimePeriod? TimePeriod { get; set; }

    /// <summary>Stop evaluating later rules when this rule is true.</summary>
    [Parameter(ParameterSetName = "RuleRange")]
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "RulePivotTable")]
    public SwitchParameter StopIfTrue { get; set; }

    /// <summary>Cell fill colour expressed as a literal OfficeColor or theme role.</summary>
    [Parameter(ParameterSetName = "RuleRange")]
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "RulePivotTable")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? BackgroundColor { get; set; }

    /// <summary>Text or rule colour expressed as a literal OfficeColor or theme role.</summary>
    [Parameter(ParameterSetName = "RuleRange")]
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "RulePivotTable")]
    [Parameter(ParameterSetName = "DataBarRange")]
    [Parameter(ParameterSetName = "DataBarHeader")]
    [Parameter(ParameterSetName = "DataBarPivotTable")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Header text identifying the target column.</summary>
    [Parameter(ParameterSetName = "RuleHeader", Mandatory = true)]
    [Parameter(ParameterSetName = "ColorScaleHeader", Mandatory = true)]
    [Parameter(ParameterSetName = "DataBarHeader", Mandatory = true)]
    [Parameter(ParameterSetName = "IconSetHeader", Mandatory = true)]
    public string Header { get; set; } = null!;

    /// <summary>Native Excel table used to resolve the source range or header.</summary>
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "ColorScaleHeader")]
    [Parameter(ParameterSetName = "DataBarHeader")]
    [Parameter(ParameterSetName = "IconSetHeader")]
    public string? Table { get; set; }

    /// <summary>One-based row containing the column headers.</summary>
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "ColorScaleHeader")]
    [Parameter(ParameterSetName = "DataBarHeader")]
    [Parameter(ParameterSetName = "IconSetHeader")]
    public int? HeaderRow { get; set; }

    /// <summary>Includes the header cell in the resolved column range.</summary>
    [Parameter(ParameterSetName = "RuleHeader")]
    [Parameter(ParameterSetName = "ColorScaleHeader")]
    [Parameter(ParameterSetName = "DataBarHeader")]
    [Parameter(ParameterSetName = "IconSetHeader")]
    public SwitchParameter IncludeHeader { get; set; }

    /// <summary>Native pivot table whose data or entire range is targeted.</summary>
    [Parameter(ParameterSetName = "RulePivotTable", Mandatory = true)]
    [Parameter(ParameterSetName = "ColorScalePivotTable", Mandatory = true)]
    [Parameter(ParameterSetName = "DataBarPivotTable", Mandatory = true)]
    [Parameter(ParameterSetName = "IconSetPivotTable", Mandatory = true)]
    public string PivotTable { get; set; } = null!;

    /// <summary>Use the full pivot output range instead of the default data body range.</summary>
    [Parameter(ParameterSetName = "RulePivotTable")]
    [Parameter(ParameterSetName = "ColorScalePivotTable")]
    [Parameter(ParameterSetName = "DataBarPivotTable")]
    [Parameter(ParameterSetName = "IconSetPivotTable")]
    public SwitchParameter PivotWholeTable { get; set; }

    /// <summary>Start color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "ColorScaleRange", Mandatory = true)]
    [Parameter(ParameterSetName = "ColorScaleHeader", Mandatory = true)]
    [Parameter(ParameterSetName = "ColorScalePivotTable", Mandatory = true)]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? StartColor { get; set; }

    /// <summary>Middle colour of a three-colour scale at the 50th percentile. Omission uses a two-colour scale.</summary>
    [Parameter(ParameterSetName = "ColorScaleRange")]
    [Parameter(ParameterSetName = "ColorScaleHeader")]
    [Parameter(ParameterSetName = "ColorScalePivotTable")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? MidColor { get; set; }

    /// <summary>End color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "ColorScaleRange", Mandatory = true)]
    [Parameter(ParameterSetName = "ColorScaleHeader", Mandatory = true)]
    [Parameter(ParameterSetName = "ColorScalePivotTable", Mandatory = true)]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? EndColor { get; set; }

    /// <summary>Selects the native data-bar conditional-format family.</summary>
    [Parameter(ParameterSetName = "DataBarRange", Mandatory = true)]
    [Parameter(ParameterSetName = "DataBarHeader", Mandatory = true)]
    [Parameter(ParameterSetName = "DataBarPivotTable", Mandatory = true)]
    public SwitchParameter DataBar { get; set; }

    /// <summary>Icon set to apply.</summary>
    [Parameter(ParameterSetName = "IconSetRange", Mandatory = true)]
    [Parameter(ParameterSetName = "IconSetHeader", Mandatory = true)]
    [Parameter(ParameterSetName = "IconSetPivotTable", Mandatory = true)]
    public ExcelIconSet IconSet { get; set; }

    /// <summary>Hides cell values behind conditional-format icons.</summary>
    [Parameter(ParameterSetName = "IconSetRange")]
    [Parameter(ParameterSetName = "IconSetHeader")]
    [Parameter(ParameterSetName = "IconSetPivotTable")]
    public SwitchParameter NoValue { get; set; }

    /// <summary>Reverse the icon order.</summary>
    [Parameter(ParameterSetName = "IconSetRange")]
    [Parameter(ParameterSetName = "IconSetHeader")]
    [Parameter(ParameterSetName = "IconSetPivotTable")]
    public SwitchParameter Reverse { get; set; }

    /// <summary>Percentage thresholds for native conditional-format icons.</summary>
    [Parameter(ParameterSetName = "IconSetRange")]
    [Parameter(ParameterSetName = "IconSetHeader")]
    [Parameter(ParameterSetName = "IconSetPivotTable")]
    public double[]? PercentThreshold { get; set; }

    /// <summary>Numeric thresholds for native conditional-format icons.</summary>
    [Parameter(ParameterSetName = "IconSetRange")]
    [Parameter(ParameterSetName = "IconSetHeader")]
    [Parameter(ParameterSetName = "IconSetPivotTable")]
    public double[]? NumberThreshold { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (PercentThreshold != null && NumberThreshold != null) throw new PSArgumentException("Choose percent or number thresholds.");
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        string target = ExcelTargetRangeResolver.Resolve(sheet, Range, Header, Table, HeaderRow ?? 0, IncludeHeader, PivotTable, !PivotWholeTable);
        var rules = sheet.Range(target).ConditionalFormat();
        ExcelConditionalFormattingInfo created;
        if (ParameterSetName.StartsWith("ColorScale", StringComparison.Ordinal)) {
            if (MidColor.HasValue) created = rules.ColorScale(StartColor!.Value, MidColor.Value, EndColor!.Value);
            else created = rules.ColorScale(StartColor!.Value, EndColor!.Value);
        } else if (ParameterSetName.StartsWith("DataBar", StringComparison.Ordinal)) created = rules.DataBar(Color ?? OfficeColor.Parse("Accent1"));
        else if (ParameterSetName.StartsWith("IconSet", StringComparison.Ordinal)) created = rules.IconSet(IconSet, !NoValue, Reverse, PercentThreshold, NumberThreshold);
        else created = ApplyRule(sheet, target);
        if (PassThru) WriteObject(created);
    }
    private ExcelConditionalFormattingInfo ApplyRule(ExcelSheet sheet, string targetRange)
    {
        var rules = sheet.Range(targetRange).ConditionalFormat();
        ExcelConditionalFormattingInfo created;
        switch (Rule)
        {
            case OfficeExcelConditionalRuleType.CellIs:
                if (!Operator.HasValue)
                {
                    throw new PSArgumentException("Operator is required for CellIs conditional formatting rules.", nameof(Operator));
                }
                if (string.IsNullOrWhiteSpace(Value))
                {
                    throw new PSArgumentException("Value is required for CellIs conditional formatting rules.", nameof(Value));
                }
                created = rules.CellIs(Operator.Value, Value!, SecondValue, fillColor: BackgroundColor, stopIfTrue: StopIfTrue);
                break;
            case OfficeExcelConditionalRuleType.Expression:
            case OfficeExcelConditionalRuleType.Formula:
                if (string.IsNullOrWhiteSpace(Value))
                {
                    throw new PSArgumentException("Value is required for formula conditional formatting rules.", nameof(Value));
                }
                created = rules.Formula(Value!, stopIfTrue: StopIfTrue.IsPresent);
                break;
            case OfficeExcelConditionalRuleType.DuplicateValues:
                created = rules.DuplicateValues();
                break;
            case OfficeExcelConditionalRuleType.UniqueValues:
                created = rules.UniqueValues();
                break;
            case OfficeExcelConditionalRuleType.Top:
            case OfficeExcelConditionalRuleType.Top10:
                created = rules.TopBottom(Rank, bottom: false, percent: Percent.IsPresent);
                break;
            case OfficeExcelConditionalRuleType.Bottom:
            case OfficeExcelConditionalRuleType.Bottom10:
                created = rules.TopBottom(Rank, bottom: true, percent: Percent.IsPresent);
                break;
            case OfficeExcelConditionalRuleType.AboveAverage:
                created = rules.AboveAverage(aboveAverage: true, equalAverage: EqualAverage.IsPresent, standardDeviation: StandardDeviation, stopIfTrue: StopIfTrue.IsPresent);
                break;
            case OfficeExcelConditionalRuleType.BelowAverage:
                created = rules.AboveAverage(aboveAverage: false, equalAverage: EqualAverage.IsPresent, standardDeviation: StandardDeviation, stopIfTrue: StopIfTrue.IsPresent);
                break;
            case OfficeExcelConditionalRuleType.ContainsText:
                created = rules.Text(ExcelConditionalFormatType.ContainsText, RequireText());
                break;
            case OfficeExcelConditionalRuleType.NotContainsText:
                created = rules.Text(ExcelConditionalFormatType.NotContainsText, RequireText());
                break;
            case OfficeExcelConditionalRuleType.BeginsWith:
                created = rules.Text(ExcelConditionalFormatType.BeginsWith, RequireText());
                break;
            case OfficeExcelConditionalRuleType.EndsWith:
                created = rules.Text(ExcelConditionalFormatType.EndsWith, RequireText());
                break;
            case OfficeExcelConditionalRuleType.ContainsBlanks:
                created = rules.Blanks(containsBlanks: true, stopIfTrue: StopIfTrue.IsPresent);
                break;
            case OfficeExcelConditionalRuleType.NotContainsBlanks:
                created = rules.Blanks(containsBlanks: false, stopIfTrue: StopIfTrue.IsPresent);
                break;
            case OfficeExcelConditionalRuleType.ContainsErrors:
                created = rules.Errors(containsErrors: true, stopIfTrue: StopIfTrue.IsPresent);
                break;
            case OfficeExcelConditionalRuleType.NotContainsErrors:
                created = rules.Errors(containsErrors: false, stopIfTrue: StopIfTrue.IsPresent);
                break;
            case OfficeExcelConditionalRuleType.TimePeriod:
                if (!TimePeriod.HasValue)
                {
                    throw new PSArgumentException("TimePeriod is required for time-period conditional formatting rules.", nameof(TimePeriod));
                }
                created = rules.TimePeriod(TimePeriod.Value, stopIfTrue: StopIfTrue.IsPresent);
                break;
            default:
                throw new PSArgumentException($"Unknown conditional formatting rule type '{Rule}'.", nameof(Rule));
        }
        if (MyInvocation.BoundParameters.ContainsKey(nameof(StopIfTrue))) {
            created.StopIfTrue = StopIfTrue;
            sheet.UpdateConditionalFormattingRule(created);
        }
        return sheet.SetConditionalFormattingColors(created, Color, BackgroundColor);
    }

    private string RequireText()
    {
        if (string.IsNullOrWhiteSpace(Text))
        {
            throw new PSArgumentException("Text is required for text conditional formatting rules.", nameof(Text));
        }

        return Text!;
    }

}
