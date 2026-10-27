using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets a cell value, formula, number format, or fill.</summary>
/// <para>Supports ExcelA1 addresses or row/column coordinates for DSL-style composition.</para>
/// <example>
///   <summary>Write values to ExcelA1 and B1.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelCell -Address 'ExcelA1' -Value 'Region'; Set-OfficeExcelCell -Row 1 -Column 2 -Value 'Revenue' }</code>
///   <para>Writes two headers in the first row.</para>
/// </example>
/// <example>
///   <summary>Write to an explicit worksheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$sheet | Set-OfficeExcelCell -Address B2 -Value 42 -NumberFormat '#,##0'</code>
///   <para>Uses the worksheet pipeline target without an active DSL scope.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelCell")]
[Alias("ExcelCell")]
public sealed class SetOfficeExcelCellCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Worksheet to modify outside a DSL context.</summary>
    [Parameter(ValueFromPipeline = true)]
    [Alias("SheetObject")]
    public ExcelSheet? Worksheet { get; set; }

    /// <summary>Workbook to modify outside a DSL context. Accepts pipeline input.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>1-based row index.</summary>
    [Parameter(ParameterSetName = "Coordinates")]
    public int? Row { get; set; }

    /// <summary>1-based column index.</summary>
    [Parameter(ParameterSetName = "Coordinates")]
    public int? Column { get; set; }

    /// <summary>ExcelA1-style cell address (e.g., ExcelA1, C5).</summary>
    [Parameter(ParameterSetName = "Address")]
    public string? Address { get; set; }

    /// <summary>Cell value to assign. Numeric, Boolean, and date values retain their native Excel types; strings remain text.</summary>
    [Parameter]
    public object? Value { get; set; }

    /// <summary>Formula text (without leading =).</summary>
    [Parameter]
    public string? Formula { get; set; }

    /// <summary>Number format code to apply.</summary>
    [Parameter]
    public string? NumberFormat { get; set; }

    /// <summary>Solid background color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? BackgroundColor { get; set; }

    /// <summary>Gradient start color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? GradientFrom { get; set; }

    /// <summary>Gradient end color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? GradientTo { get; set; }

    /// <summary>Linear gradient angle in degrees.</summary>
    [Parameter]
    public double GradientDegree { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Worksheet, Document, Sheet, SheetIndex);
        var (row, column) = ExcelHostExtensions.ResolveCellAddress(Row, Column, Address);
        var hasValueChange = Value != null || Formula != null || NumberFormat != null;
        var hasStyleChange = !string.IsNullOrWhiteSpace(BackgroundColor)
            || !string.IsNullOrWhiteSpace(GradientFrom)
            || !string.IsNullOrWhiteSpace(GradientTo);

        if (!hasValueChange && !hasStyleChange) {
            throw new PSArgumentException("Provide -Value, -Formula, -NumberFormat, -BackgroundColor, or -GradientFrom/-GradientTo to modify the cell.");
        }

        if (!string.IsNullOrWhiteSpace(GradientFrom) ^ !string.IsNullOrWhiteSpace(GradientTo)) {
            throw new PSArgumentException("Specify both -GradientFrom and -GradientTo for a gradient fill.");
        }

        if (hasValueChange) {
            sheet.Cell(row, column, PowerShellObjectNormalizer.NormalizeCellValueForTable(Value), Formula, NumberFormat);
        }

        if (!string.IsNullOrWhiteSpace(BackgroundColor)) {
            sheet.CellBackground(row, column, BackgroundColor!);
        }

        if (!string.IsNullOrWhiteSpace(GradientFrom) && !string.IsNullOrWhiteSpace(GradientTo)) {
            sheet.CellGradientBackground(row, column, GradientFrom!, GradientTo!, GradientDegree);
        }
        WritePassThru(sheet);
    }
}
