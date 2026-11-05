using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets a cell value, formula, rich text, number format or fill.</summary>
/// <example>
/// <summary>Write a value to a cell</summary>
/// <code>New-OfficeExcel -Path .\Budget.xlsx {
///     ExcelSheet 'Summary' {
///         ExcelCell A1 'Budget 2027'
///     }
/// }</code>
/// <para>Writes a heading into A1 of the Summary sheet.</para>
/// </example>
/// <example>
/// <summary>Lay out a small calculation</summary>
/// <code>New-OfficeExcel -Path .\Estimates\Migration-Cost.xlsx {
///     ExcelSheet 'Estimate' {
///         ExcelCell A1 'Servers'
///         ExcelCell B1 42
///         ExcelCell A2 'Cost per server'
///         ExcelCell B2 1850 -NumberFormat '#,##0.00'
///         ExcelCell A3 'Total'
///         ExcelCell B3 -Formula '=B1*B2' -NumberFormat '#,##0.00' -BackgroundColor '#FEF3C7'
///     }
/// }</code>
/// <para>Writes labels, inputs and formulas for a cost estimate, formatting the result as currency.</para>
/// </example>
/// <example>
/// <summary>Write rich text into a cell</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Status.xlsx {
///     ExcelSheet 'Summary' {
///         ExcelCell B2 {
///             ExcelText 'Status: ' -Bold
///             ExcelText 'On track' -Bold -Color Success
///         }
///     }
/// }</code>
/// <para>Passes a block of text runs as the value, so one cell mixes bold and coloured text.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelCell", DefaultParameterSetName = "Address")]
[Alias("ExcelCell")]
public sealed class SetOfficeExcelCellCommand : PSCmdlet {
    /// <summary>A1 address of the target cell.</summary>
    [Parameter(ParameterSetName = "Address", Mandatory = true, Position = 0)]
    public string Address { get; set; } = null!;

    /// <summary>Value or values consumed by the selected operation.</summary>
    [Parameter(Position = 1)]
    public object? Value { get; set; }

    /// <summary>Formula text (without leading =).</summary>
    [Parameter()]
    public string? Formula { get; set; }

    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Number format code to apply.</summary>
    [Parameter()]
    public string? NumberFormat { get; set; }

    /// <summary>Cell fill colour expressed as a literal OfficeColor or theme role.</summary>
    [Parameter()]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? BackgroundColor { get; set; }

    /// <summary>First fill colour; GradientTo is required with it.</summary>
    [Parameter()]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? GradientFrom { get; set; }

    /// <summary>Second fill colour; GradientFrom is required with it.</summary>
    [Parameter()]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? GradientTo { get; set; }

    /// <summary>Linear fill angle in degrees.</summary>
    [Parameter()]
    public double? GradientAngle { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>1-based row index.</summary>
    [Parameter(ParameterSetName = "Coordinates", Mandatory = true)]
    public int? Row { get; set; }

    /// <summary>1-based column index.</summary>
    [Parameter(ParameterSetName = "Coordinates", Mandatory = true)]
    public object? Column { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (GradientFrom.HasValue != GradientTo.HasValue) throw new PSArgumentException("Supply both GradientFrom and GradientTo.");
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        var (row, column) = ExcelHostExtensions.ResolveCellAddress(Row, Column == null ? null : ExcelCommandTargets.Column(Column), Address);
        var cell = sheet.Cell(row, column);
        object? value = OfficeArgumentValue.Unwrap(Value);
        if (value is ScriptBlock content) cell.SetRichText(ExcelRichTextRunService.ReadContent(content));
        else if (value is ExcelRichTextRun run) cell.SetRichText(run);
        else if (value is IEnumerable sequence && value is not string && sequence.Cast<object>().Select(OfficeArgumentValue.Unwrap).ToArray() is var items && items.Length > 0 && items.All(item => item is ExcelRichTextRun)) cell.SetRichText(items.Cast<ExcelRichTextRun>());
        else if (MyInvocation.BoundParameters.ContainsKey(nameof(Value))) cell.SetValue(PowerShellObjectNormalizer.NormalizeCellValueForTable(value) ?? string.Empty);
        if (Formula != null) cell.SetFormula(Formula);
        if (NumberFormat != null) cell.SetNumberFormat(NumberFormat);
        if (BackgroundColor.HasValue) cell.SetFillColor(BackgroundColor.Value);
        if (GradientFrom.HasValue) cell.SetGradientFill(GradientFrom.Value, GradientTo!.Value, GradientAngle ?? 0);
        if (PassThru) WriteObject(cell);
    }

}
