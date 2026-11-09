using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Writes values and applies row layout to a worksheet in a DSL block or through the pipeline.</summary>
/// <example>
/// <summary>Write a header row</summary>
/// <code>New-OfficeExcel -Path .\Rota.xlsx {
///     ExcelSheet 'On call' {
///         ExcelRow 1 'Week', 'Primary', 'Secondary', 'Escalation' -Bold
///     }
/// }</code>
/// <para>Writes column labels across row 1 and makes them bold.</para>
/// </example>
/// <example>
/// <summary>Write a header band under a title</summary>
/// <code>New-OfficeExcel -Path .\Reports\Budget-2027.xlsx {
///     ExcelSheet 'Departments' {
///         ExcelCell A1 'Budget 2027 by department'
///         ExcelRow 3 'Department', 'Q1', 'Q2', 'Q3', 'Q4', 'Total' -Bold -BackgroundColor '#F1F5F9' -Height 22
///     }
/// }</code>
/// <para>Writes the quarter labels into row 3 below the sheet title, with a light fill and a taller row so the band stands apart from the data.</para>
/// </example>
/// <example>
/// <summary>Write a totals row with formulas</summary>
/// <code>Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
///     ExcelSheet 'Departments' {
///         ExcelRow 20 'Total', '=SUM(B4:B19)', '=SUM(C4:C19)', '=SUM(D4:D19)', '=SUM(E4:E19)' -Formula -Bold
///     }
/// }</code>
/// <para>Adds a totals row whose cells are formulas; without -Formula the same strings would be stored as text.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelRow", DefaultParameterSetName = "Default")]
[Alias("ExcelRow")]
public sealed class SetOfficeExcelRowCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Workbook to modify. Accepts pipeline input; select its worksheet with Sheet.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name when using Document. The first worksheet is used when no selector is supplied.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>1-based row index.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    public int Row { get; set; }

    /// <summary>Values to write across the row. Numeric, Boolean, and date values retain their native Excel types; strings remain text.</summary>
    [Parameter(ParameterSetName = "Default", Position = 1)]
    public object[]? Values { get; set; }

    /// <summary>Starting column index (1-based).</summary>
    [Parameter(ParameterSetName = "Default")]
    public object StartColumn { get; set; } = 1;

    /// <summary>Explicit row height in points.</summary>
    [Parameter(ParameterSetName = "Default")]
    [ExcelColumnWidthTransformation(RowHeight = true)]
    public ExcelColumnWidth? Height { get; set; }

    /// <summary>Clear any custom row height.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ClearHeight { get; set; }

    /// <summary>Auto-fit row height after writing values and applying style.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AutoFit { get; set; }

    /// <summary>Hide or unhide the row. Use -Hidden:$false to unhide.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Hidden { get; set; }

    /// <summary>Apply or clear bold styling across the target row span.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Bold { get; set; }

    /// <summary>Apply or clear italic styling across the target row span.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Italic { get; set; }

    /// <summary>Apply or clear underline styling across the target row span.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Underline { get; set; }

    /// <summary>Apply or clear wrap text across the target row span.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter WrapText { get; set; }

    /// <summary>Apply a font family across the target row span.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Font { get; set; }

    /// <summary>Apply a background color across the target row span.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? BackgroundColor { get; set; }

    /// <summary>First 1-based column affected by style options.</summary>
    [Parameter(ParameterSetName = "Default")]
    public object? FirstColumn { get; set; }

    /// <summary>Last 1-based column affected by style options.</summary>
    [Parameter(ParameterSetName = "Default")]
    public object? LastColumn { get; set; }

    /// <summary>Interpret explicitly prefixed row strings as formulas.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Formula { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);

        int startColumn = ExcelCommandTargets.Column(StartColumn);
        if (Row < 1) {
            throw new ArgumentOutOfRangeException(nameof(Row), "Row index must be 1 or greater.");
        }

        if (startColumn < 1) {
            throw new ArgumentOutOfRangeException(nameof(StartColumn), "StartColumn must be 1 or greater.");
        }

        var values = Values ?? Array.Empty<object>();
        var hasLayout = HasLayoutOptions();
        if (values.Length == 0 && !hasLayout) {
            throw new PSArgumentException("Provide row values or at least one layout/style option.", nameof(Values));
        }

        if (values.Length > 0) {
            var cells = new List<(int Row, int Column, object Value)>(values.Length);
            for (int i = 0; i < values.Length; i++) {
                var value = PowerShellObjectNormalizer.NormalizeCellValueForTable(values[i]) ?? string.Empty;
                cells.Add((Row, startColumn + i, value));
            }

            if (Formula) {
                foreach (var item in cells) {
                    var cell = sheet.Cell(item.Row, item.Column);
                    if (item.Value is string text && text.StartsWith("=", StringComparison.Ordinal)) cell.SetFormula(text);
                    else cell.SetValue(item.Value);
                }
            } else sheet.SetValues(cells);
        }

        if (hasLayout) {
            sheet.SetRowLayout(Row, new ExcelRowLayoutOptions {
                Height = Height,
                ClearHeight = ClearHeight.IsPresent,
                AutoFit = AutoFit.IsPresent,
                Hidden = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Hidden)),
                Bold = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Bold)),
                Italic = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Italic)),
                Underline = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Underline)),
                WrapText = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(WrapText)),
                FontFamily = Font,
                BackgroundColor = BackgroundColor,
                FirstColumn = FirstColumn == null ? (values.Length > 0 ? startColumn : (int?)null) : ExcelCommandTargets.Column(FirstColumn),
                LastColumn = LastColumn == null ? (values.Length > 0 ? startColumn + values.Length - 1 : (int?)null) : ExcelCommandTargets.Column(LastColumn)
            });
        }
        WritePassThru(sheet);
    }

    private bool HasLayoutOptions() {
        return Height.HasValue ||
            ClearHeight.IsPresent ||
            AutoFit.IsPresent ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Hidden)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Bold)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Italic)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Underline)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(WrapText)) ||
            !string.IsNullOrWhiteSpace(Font) ||
            BackgroundColor.HasValue ||
            FirstColumn != null ||
            LastColumn != null;
    }
}
